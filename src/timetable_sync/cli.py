from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path

import httpx
from cryptography.fernet import Fernet

from .auth import Authentication
from .config import Config
from .engine import run
from .graph import Graph
from .models import SyncError
from .store import Store


def parser():
    root = argparse.ArgumentParser(description="Preference-aware timetable to Outlook synchroniser")
    root.add_argument("--env", help="Configuration .env file")
    sub = root.add_subparsers(dest="command", required=True)
    sub.add_parser("init", help="Create a private local .env")
    sub.add_parser("auth", help="Sign in with Microsoft device flow")
    sync = sub.add_parser("sync", help="Preview by default; --apply writes Outlook")
    sync.add_argument("--file", type=Path, help="Read local ICS instead of downloading")
    sync.add_argument("--outlook", action="store_true", help="Include live Outlook in preview")
    sync.add_argument("--apply", action="store_true")
    sync.add_argument("--json", action="store_true")
    sync.add_argument(
        "--summary", action="store_true", help="Print counts instead of event details"
    )
    sync.add_argument("--at", help="Preview at ISO datetime; unavailable with --apply")
    rule = sub.add_parser("rule", help="Set an availability rule; default removes it")
    rule.add_argument(
        "selector", help="activity:Workshop, module:M34521, family:cmis:123, event:<key>, default"
    )
    rule.add_argument("value", choices=["free", "busy", "default"])
    sub.add_parser("rules", help="List availability rules")
    sub.add_parser("status", help="Show the last successful sync")
    health = sub.add_parser("health", help="Exit unsuccessfully if the sync is stale or held")
    health.add_argument("--max-age", type=int, default=900, help="Maximum sync age in seconds")
    link = sub.add_parser(
        "link-move", help="Map a reviewed new occurrence to an existing occurrence"
    )
    link.add_argument("new_key")
    link.add_argument("existing_key")
    sub.add_parser("backup", help="Make an online SQLite backup in the state directory")
    return root


def main(argv=None):
    args = parser().parse_args(argv)
    try:
        if args.command == "init":
            path = Path(args.env or ".env")
            if path.exists():
                raise SyncError("Configuration already exists; it has not been overwritten.")
            template = Path(__file__).resolve().parents[2] / ".env.example"
            content = (
                template.read_text(encoding="utf-8")
                if template.exists()
                else "TSYNC_FEED_URL=\nTSYNC_CLIENT_ID=\nTSYNC_TOKEN_KEY=\n"
            )
            content = content.replace(
                "TSYNC_TOKEN_KEY=", "TSYNC_TOKEN_KEY=" + Fernet.generate_key().decode(), 1
            )
            path.write_text(content, encoding="utf-8")
            if os.name != "nt":
                path.chmod(0o600)
            print(f"Created {path}. Add your private feed URL and Entra client ID.")
            return 0
        config = Config.load(args.env)
        store = Store(config.state_dir)
        if args.command == "auth":
            with store.lock():
                auth = Authentication(config)
                print("Signed in as " + auth.login())
                identifier = auth.account()["home_account_id"]
                if store.meta("account") not in {None, identifier}:
                    raise SyncError(
                        "This state belongs to another account. Use a separate state directory."
                    )
                store.set_meta("account", identifier)
            return 0
        if args.command == "rule":
            selector = args.selector.casefold()
            allowed = ("activity:", "module:", "module-activity:", "family:", "event:")
            if selector != "default" and not selector.startswith(allowed):
                raise SyncError(
                    "Unknown rule selector. Use activity:, module:, module-activity:, family:, event: or default."
                )
            with store.lock():
                store.rule(selector, args.value)
                if selector.startswith("event:"):
                    key = selector.split(":", 1)[1]
                    row = store.events().get(key)
                    if row:
                        row["override"] = None
                        store.save(key, row)
            print(f"{selector} = {args.value}")
            return 0
        if args.command == "rules":
            print(json.dumps(store.rules(), indent=2))
            return 0
        if args.command in {"status", "health"}:
            state = store.meta("last_success")
            print(json.dumps(state, indent=2))
            if args.command == "health":
                if (
                    not state
                    or (
                        datetime.now(timezone.utc) - datetime.fromisoformat(state["at"])
                    ).total_seconds()
                    > args.max_age
                    or state["warnings"]
                ):
                    return 1
            return 0
        if args.command == "backup":
            import sqlite3

            path = config.state_dir / "backup.sqlite3"
            with store.lock(), sqlite3.connect(path) as target:
                store.db.backup(target)
            print(f"Backup saved to {path}")
            return 0
        if args.command == "link-move":
            with store.lock():
                if args.existing_key not in store.events():
                    raise SyncError("The existing key is not in local sync state.")
                aliases = store.meta("aliases", {})
                if (
                    args.new_key in store.events()
                    or args.new_key in aliases
                    or args.new_key == args.existing_key
                ):
                    raise SyncError("New occurrence already has an identity. Review the keys.")
                aliases[args.new_key] = args.existing_key
                store.set_meta("aliases", aliases)
            print("Identity mapping saved. Preview the next sync before applying.")
            return 0
        if args.at and args.apply:
            raise SyncError("--at is for previews only.")
        now = datetime.fromisoformat(args.at) if args.at else None
        if now and now.tzinfo is None:
            raise SyncError("--at needs a timezone offset.")
        if args.apply and args.file:
            raise SyncError(
                "Local feed files are for preview only. Apply always downloads a fresh feed."
            )
        graph = None
        if args.apply or args.outlook:
            auth = Authentication(config)
            account_id = auth.account()["home_account_id"]
            previous = store.meta("account")
            if previous and account_id != previous:
                raise SyncError("The signed-in account differs from the state directory's account.")
            if args.apply and not previous:
                raise SyncError("Run auth to bind this state directory to your account first.")
            graph = Graph(auth.token, config.source_id, config.calendar_id)
        raw = args.file.read_bytes() if args.file else None
        result = run(config, graph=graph, raw=raw, apply=args.apply, now=now, store=store)
        if args.json:
            # No access tokens, private URLs or Graph responses in diagnostic output.
            report = {
                **result,
                "actions": [
                    {k: v for k, v in a.items() if k != "event"} for a in result["actions"]
                ],
            }
            print(json.dumps(report, indent=2))
        else:
            print(
                f"{'APPLY' if args.apply else 'PREVIEW'}: {result['source_events']} occurrences"
                + (" (offline; Outlook has not been checked)" if result["offline"] else "")
            )
            if args.summary:
                from collections import Counter

                print(json.dumps(dict(Counter(a["kind"] for a in result["actions"]))))
            for action in [] if args.summary else result["actions"]:
                session = action["session"]
                print(
                    f"{action['kind']:9} {action['key']} {session['start']} {session['title']} [{action.get('show_as', '')}] â€” {action['reason']}"
                )
            for warning in result["warnings"]:
                print("WARNING: " + warning)
        if args.apply and config.heartbeat_url:
            from urllib.parse import urlparse

            if urlparse(config.heartbeat_url).scheme != "https":
                raise SyncError("Heartbeat URL must use HTTPS.")
            try:
                response = httpx.get(config.heartbeat_url, timeout=10, follow_redirects=False)
                response.raise_for_status()
            except httpx.HTTPError as exc:
                raise SyncError(
                    "Calendar sync succeeded, but the health heartbeat failed."
                ) from exc
        return 0
    except (SyncError, OSError, ValueError) as exc:
        # OSError/ValueError may contain sensitive paths or URL query strings.
        message = (
            str(exc)
            if isinstance(exc, SyncError)
            else "Local configuration or file operation failed."
        )
        print("Error: " + message, file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

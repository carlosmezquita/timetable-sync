from __future__ import annotations

from datetime import datetime, timedelta, timezone

from .calendar import parse_feed
from .feed import download
from .models import Session, SyncError, digest
from .rules import availability
from .store import Store


def payload(session, show_as):
    def utc(value):
        return datetime.fromisoformat(value).astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S")

    return {
        "subject": session.title,
        "start": {"dateTime": utc(session.start), "timeZone": "UTC"},
        "end": {"dateTime": utc(session.end), "timeZone": "UTC"},
        "location": {"displayName": session.location},
        "body": {"contentType": "text", "content": session.description},
        "showAs": show_as,
    }


def signature(session):
    fields = session.json()
    # CMIS changes UID when its recurrence segment changes.
    fields.pop("uid")
    return digest(fields)


def plan(sessions, stored, remote, rules, config, now):
    actions = []
    warnings = []
    present = {s.key: s for s in sessions}
    missing = {}
    new_families = {s.family for s in sessions if s.key not in stored and s.key not in remote}
    absent_families = {
        Session.from_json(row["session"]).family
        for key, row in stored.items()
        if key not in present and datetime.fromisoformat(row["session"]["start"]) > now
    }
    ambiguous = new_families & absent_families
    for session in sessions:
        row = stored.get(session.key)
        event = remote.get(session.key)
        override = row.get("override") if row else None
        if (
            event
            and row
            and event.get("showAs") in {"free", "busy"}
            and event["showAs"] != row.get("show_as")
        ):
            override = event["showAs"]
        if "event:" + session.key in rules:
            override = None
        show_as = availability(session, rules, override)
        action = {
            "key": session.key,
            "session": session.json(),
            "show_as": show_as,
            "override": override,
            "payload": payload(session, show_as),
        }
        if session.activity == "Other":
            warnings.append(f"Unknown activity for {session.key[:12]}; defaults to Busy.")
        if session.cancelled:
            if event:
                actions.append(
                    {**action, "kind": "delete", "reason": "Explicit cancellation", "event": event}
                )
            elif row:
                actions.append(
                    {**action, "kind": "forget", "reason": "Cancelled event already absent"}
                )
            continue
        if session.family in ambiguous and not row and not event:
            actions.append(
                {
                    **action,
                    "kind": "hold",
                    "reason": "Possible move to another day; use link-move after reviewing",
                }
            )
        elif not event:
            actions.append(
                {**action, "kind": "create", "reason": "New or removed managed appointment"}
            )
        else:
            changed = (
                not row
                or signature(session) != signature(Session.from_json(row["session"]))
                or event.get("showAs") != show_as
            )
            actions.append(
                {
                    **action,
                    "kind": "update" if changed else "unchanged",
                    "reason": "Source or availability changed" if changed else "Already current",
                    "event": event,
                }
            )

    horizon = now + timedelta(days=config.lookahead_days)
    for key, row in stored.items():
        if key in present:
            continue
        session = Session.from_json(row["session"])
        starts = datetime.fromisoformat(session.start)
        if starts < now or starts >= horizon:
            continue  # Historical and out-of-window events are retained.
        first = row.get("missing_since")
        count = row.get("missing_count", 0) + 1
        missing[key] = {"missing_count": count, "missing_since": first or now.timestamp()}
        confirmed = (
            first is not None
            and now.timestamp() - first >= config.min_delete_interval
            and count >= 2
        )
        held = session.family in ambiguous or not confirmed
        event = remote.get(key)
        actions.append(
            {
                "kind": "hold" if held else ("delete" if event else "forget"),
                "key": key,
                "session": row["session"],
                "event": event,
                "reason": "Possible move to another day; review and link identities"
                if session.family in ambiguous
                else (
                    "Missing from two successful checks"
                    if confirmed
                    else "Waiting for another successful feed check"
                ),
            }
        )
    deletes = [a for a in actions if a["kind"] in {"delete", "forget"}]
    upcoming = sum(
        datetime.fromisoformat(row["session"]["start"]) >= now for row in stored.values()
    )
    if len(deletes) > config.max_delete_count or (
        upcoming and len(deletes) / upcoming > config.max_delete_fraction
    ):
        warnings.append(
            "Deletion safety limit reached. Review the feed and increase limits only after verifying cancellations."
        )
        for action in deletes:
            action["kind"] = "hold"
            action["reason"] = "Deletion safety limit"
    # Unknown owned events are never removed: recover only when their keys occur in the source.
    orphan_count = len(set(remote) - set(stored) - set(present))
    if orphan_count:
        warnings.append(
            f"{orphan_count} owned Outlook event(s) cannot be matched; retained for review."
        )
    held = sum(a["kind"] == "hold" for a in actions)
    if held:
        warnings.append(f"{held} occurrence(s) held for confirmation or identity review.")
    held = sum(a["kind"] == "hold" for a in actions)
    if held:
        warnings.append(f"{held} occurrence(s) held for confirmation or identity review.")
    return {"actions": actions, "warnings": sorted(set(warnings)), "missing": missing}


def run(config, graph=None, raw=None, apply=False, now=None, store=None):
    now = now or datetime.now(timezone.utc)
    store = store or Store(config.state_dir)
    with store.lock():
        if raw is None:
            if not config.feed_url:
                raise SyncError("Set TSYNC_FEED_URL or provide --file for an offline preview.")
            raw = download(config.feed_url)
        sessions = parse_feed(
            raw, now, now + timedelta(days=config.lookahead_days), config.source_id, config.timezone
        )
        aliases = store.meta("aliases", {})
        if aliases:
            from dataclasses import replace

            sessions = [replace(s, key=aliases.get(s.key, s.key)) for s in sessions]
        if len({s.key for s in sessions}) != len(sessions):
            raise SyncError(
                "Identity mappings collide. Remove the conflicting link before syncing."
            )
        if apply and graph is None:
            raise SyncError("Applying requires an authenticated Outlook connection.")
        binding = store.meta("binding")
        expected = {"source": config.source_id, "calendar": config.calendar_id}
        if binding and binding != expected:
            raise SyncError("State directory is bound to another source or calendar.")
        stored = store.events()
        remote = (
            graph.list_owned()
            if graph
            else {
                key: {"id": row["remote_id"], "showAs": row["show_as"]}
                for key, row in stored.items()
            }
        )
        result = plan(sessions, stored, remote, store.rules(), config, now)
        result["offline"] = graph is None
        result["source_events"] = len(sessions)
        result["applied"] = apply
        if not apply:
            return result
        store.bind(expected)
        for action in result["actions"]:
            key, kind = action["key"], action["kind"]
            if kind == "hold":
                continue
            if kind == "delete":
                graph.change(action["event"], key)
                store.delete(key)
                continue
            if kind == "forget":
                store.delete(key)
                continue
            if kind == "create":
                event = graph.create(action["payload"], key, store.transaction(key))
            elif kind == "update":
                event = graph.change(action["event"], key, action["payload"])
            else:
                event = action["event"]
            store.save(
                key,
                {
                    "session": action["session"],
                    "remote_id": event["id"],
                    "show_as": action["show_as"],
                    "override": action["override"],
                    "missing_count": 0,
                    "missing_since": None,
                },
            )
        for key, changes in result["missing"].items():
            current = store.events().get(key)
            if current:
                store.save(key, {**current, **changes})
        store.set_meta(
            "last_success",
            {"at": now.isoformat(), "source_events": len(sessions), "warnings": result["warnings"]},
        )
        return result

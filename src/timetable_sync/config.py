from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from dotenv import load_dotenv

from .models import SyncError


@dataclass(frozen=True)
class Config:
    state_dir: Path
    source_id: str = "portsmouth-2026-2027"
    timezone: str = "Europe/London"
    calendar_id: str = "primary"
    feed_url: str = ""
    client_id: str = ""
    tenant_id: str = "organizations"
    account_username: str = ""
    token_key: str = ""
    lookahead_days: int = 365
    min_delete_interval: int = 300
    max_delete_fraction: float = 0.25
    max_delete_count: int = 10
    heartbeat_url: str = ""

    @classmethod
    def load(cls, env_file: str | None = None) -> Config:
        load_dotenv(env_file or ".env", override=False)

        def get(name: str, default: str = "") -> str:
            return os.getenv("TSYNC_" + name, default)

        try:
            config = cls(
                state_dir=Path(get("STATE_DIR", ".state")).resolve(),
                source_id=get("SOURCE_ID", "portsmouth-2026-2027"),
                timezone=get("TIMEZONE", "Europe/London"),
                calendar_id=get("CALENDAR_ID", "primary"),
                feed_url=get("FEED_URL"),
                client_id=get("CLIENT_ID"),
                tenant_id=get("TENANT_ID", "organizations"),
                account_username=get("ACCOUNT_USERNAME"),
                token_key=get("TOKEN_KEY"),
                lookahead_days=int(get("LOOKAHEAD_DAYS", "365")),
                min_delete_interval=int(get("MIN_DELETE_INTERVAL_SECONDS", "300")),
                max_delete_fraction=float(get("MAX_DELETE_FRACTION", "0.25")),
                max_delete_count=int(get("MAX_DELETE_COUNT", "10")),
                heartbeat_url=get("HEARTBEAT_URL"),
            )
            ZoneInfo(config.timezone)
            if not config.source_id or not 1 <= config.lookahead_days <= 730:
                raise ValueError("Invalid source ID or lookahead")
            if (
                not 0 < config.max_delete_fraction <= 1
                or config.max_delete_count < 1
                or config.min_delete_interval < 300
            ):
                raise ValueError("Invalid deletion safety limits")
            return config
        except (ValueError, ZoneInfoNotFoundError) as exc:
            raise SyncError("Invalid configuration; check numeric settings and timezone.") from exc

from __future__ import annotations

import hashlib
import json
from dataclasses import asdict, dataclass
from datetime import datetime


class SyncError(Exception):
    """An actionable error whose text is safe to show without credentials."""


def digest(value: object) -> str:
    return hashlib.sha256(json.dumps(value, sort_keys=True).encode()).hexdigest()


@dataclass(frozen=True)
class Session:
    key: str
    family: str
    uid: str
    start: str
    end: str
    identity_date: str
    module: str
    title: str
    activity: str
    location: str
    description: str
    cancelled: bool = False

    def json(self) -> dict:
        return asdict(self)

    @classmethod
    def from_json(cls, data: dict) -> Session:
        return cls(**data)

    @property
    def starts_at(self) -> datetime:
        return datetime.fromisoformat(self.start)

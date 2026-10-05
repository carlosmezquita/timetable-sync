from datetime import datetime, timezone

import pytest

from timetable_sync.config import Config
from timetable_sync.graph import KEY_PROPERTY, SOURCE_PROPERTY
from timetable_sync.store import Store


@pytest.fixture
def now():
    return datetime(2026, 10, 5, tzinfo=timezone.utc)


@pytest.fixture
def config(tmp_path):
    return Config(tmp_path / "state", max_delete_fraction=1, max_delete_count=100)


@pytest.fixture
def store(config):
    return Store(config.state_dir)


def feed(events):
    return (
        "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//Tests//EN\r\n"
        + "\r\n".join(events)
        + "\r\nEND:VCALENDAR\r\n"
    ).encode()


def event(
    uid="0ID123:2026-09-28@timetabling.port.ac.uk",
    start="20261012T100000Z",
    end="20261012T120000Z",
    extra="",
    description="M34521/311\\nAI CONCEPTS AND APPLICATIONS\\nTeaching Event\\nWorkshop",
):
    return f"BEGIN:VEVENT\r\nUID:{uid}\r\nDTSTART:{start}\r\nDTEND:{end}\r\nSUMMARY:Demo room\r\nLOCATION:Room 1\r\nDESCRIPTION:{description}\r\n{extra}\r\nEND:VEVENT"


class FakeGraph:
    def __init__(self, source="portsmouth-2026-2027"):
        self.events = {}
        self.transactions = {}
        self.source = source
        self.creates = 0
        self.updates = 0
        self.deletes = 0
        self.fail_after_create = False

    def list_owned(self):
        return {key: dict(value) for key, value in self.events.items()}

    def create(self, payload, key, transaction):
        if transaction in self.transactions:
            return self.transactions[transaction]
        self.creates += 1
        remote = {
            **payload,
            "id": str(self.creates),
            "@odata.etag": str(self.creates),
            "singleValueExtendedProperties": [
                {"id": SOURCE_PROPERTY, "value": self.source},
                {"id": KEY_PROPERTY, "value": key},
            ],
        }
        self.events[key] = remote
        self.transactions[transaction] = remote
        if self.fail_after_create:
            self.fail_after_create = False
            from timetable_sync.models import SyncError

            raise SyncError("Response lost")
        return remote

    def change(self, remote, key, payload=None):
        if payload is None:
            self.deletes += 1
            del self.events[key]
            return {}
        self.updates += 1
        self.events[key].update(payload)
        return self.events[key]

from datetime import timedelta

import pytest
from conftest import FakeGraph, event, feed

from timetable_sync.calendar import parse_feed
from timetable_sync.engine import run
from timetable_sync.models import SyncError
from timetable_sync.store import Store


def test_recurrence_exception_preserves_original_identity_when_moved(config, now):
    master = event(uid="generic", extra="RRULE:FREQ=WEEKLY;COUNT=3")
    exception = event(
        uid="generic",
        start="20261020T140000Z",
        end="20261020T160000Z",
        extra="RECURRENCE-ID:20261019T100000Z",
    )
    original = parse_feed(
        feed([master]), now, now + timedelta(days=40), config.source_id, config.timezone
    )
    changed = parse_feed(
        feed([master, exception]), now, now + timedelta(days=40), config.source_id, config.timezone
    )
    old = next(s for s in original if s.start.startswith("2026-10-19"))
    moved = next(s for s in changed if s.start.startswith("2026-10-20"))
    assert old.key == moved.key
    assert len(changed) == 3


def test_process_lock_prevents_overlap_and_releases_after_failure(store, config):
    other = Store(config.state_dir)
    with store.lock():
        with pytest.raises(SyncError):
            with other.lock():
                pytest.fail("Overlapping process acquired lock")
    with other.lock():
        pass


def test_persistent_rules_survive_restart(config, store, now):
    store.rule("activity:workshop", "free")
    reopened = Store(config.state_dir)
    result = run(config, raw=feed([event()]), now=now, store=reopened)
    assert result["actions"][0]["show_as"] == "free"


def test_source_reappearance_clears_missing_evidence(config, store, now):
    graph = FakeGraph()
    complete = feed([event(), event(uid="other")])
    incomplete = feed([event(uid="other")])
    run(config, graph, complete, True, now, store)
    run(config, graph, incomplete, True, now, store)
    run(config, graph, complete, True, now + timedelta(minutes=6), store)
    assert all(row["missing_count"] == 0 for row in store.events().values())
    run(config, graph, incomplete, True, now + timedelta(minutes=12), store)
    assert graph.deletes == 0


def test_subdaily_recurrence_rejected_before_expansion(config, now):
    with pytest.raises(SyncError, match="Sub-daily"):
        parse_feed(
            feed([event(extra="RRULE:FREQ=SECONDLY")]),
            now,
            now + timedelta(days=40),
            config.source_id,
            config.timezone,
        )


def test_historical_events_are_retained(config, store, now):
    graph = FakeGraph()
    run(config, graph, feed([event()]), True, now, store)
    run(
        config,
        graph,
        feed([event(uid="future", start="20261115T100000Z", end="20261115T120000Z")]),
        True,
        now + timedelta(days=20),
        store,
    )
    assert graph.deletes == 0 and len(graph.events) == 2


def test_cli_local_apply_rejected_before_authentication(tmp_path, capsys, monkeypatch):
    from timetable_sync.cli import main

    monkeypatch.chdir(tmp_path)
    path = tmp_path / "demo.ics"
    path.write_bytes(feed([event()]))
    assert main(["sync", "--apply", "--file", str(path)]) == 1
    assert "local feed" in capsys.readouterr().err.lower()

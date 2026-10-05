from dataclasses import replace
from datetime import timedelta

import pytest
from conftest import FakeGraph, event, feed

from timetable_sync.calendar import parse_feed
from timetable_sync.engine import run
from timetable_sync.models import SyncError


def test_recurring_sessions_have_distinct_keys(config, now):
    raw = feed([event(extra="RRULE:FREQ=WEEKLY;COUNT=3")])
    sessions = parse_feed(raw, now, now + timedelta(days=50), config.source_id, config.timezone)
    assert len(sessions) == len({s.key for s in sessions}) == 3
    assert all(s.module == "M34521" and s.activity == "Workshop" for s in sessions)


def test_repeat_sync_and_time_edit(config, store, now):
    graph = FakeGraph()
    raw = feed([event()])
    run(config, graph, raw, True, now, store)
    run(config, graph, raw, True, now, store)
    assert graph.creates == 1 and graph.updates == 0
    run(
        config,
        graph,
        feed([event(start="20261012T110000Z", end="20261012T130000Z")]),
        True,
        now,
        store,
    )
    assert graph.creates == 1 and graph.updates == 1


def test_manual_outlook_free_choice_survives_source_change(config, store, now):
    graph = FakeGraph()
    run(config, graph, feed([event()]), True, now, store)
    key = next(iter(graph.events))
    graph.events[key]["showAs"] = "free"
    run(config, graph, feed([event(extra="SEQUENCE:2")]), True, now, store)
    run(config, graph, feed([event(end="20261012T130000Z")]), True, now, store)
    assert graph.events[key]["showAs"] == "free"
    assert store.events()[key]["override"] == "free"


def test_explicit_rule_overrides_manual_choice(config, store, now):
    graph = FakeGraph()
    run(config, graph, feed([event()]), True, now, store)
    key = next(iter(graph.events))
    graph.events[key]["showAs"] = "free"
    store.rule("event:" + key, "busy")
    run(config, graph, feed([event()]), True, now, store)
    assert graph.events[key]["showAs"] == "busy"


def test_missing_requires_two_spaced_successful_checks(config, store, now):
    graph = FakeGraph()
    both = feed([event(), event(uid="other", start="20261013T100000Z", end="20261013T120000Z")])
    remaining = feed([event(uid="other", start="20261013T100000Z", end="20261013T120000Z")])
    run(config, graph, both, True, now, store)
    run(config, graph, remaining, True, now, store)
    run(config, graph, remaining, True, now + timedelta(seconds=299), store)
    assert graph.deletes == 0
    run(config, graph, remaining, True, now + timedelta(seconds=300), store)
    assert graph.deletes == 1


def test_preview_does_not_confirm_deletion(config, store, now):
    graph = FakeGraph()
    run(config, graph, feed([event(), event(uid="other")]), True, now, store)
    run(config, graph, feed([event(uid="other")]), False, now, store)
    run(config, graph, feed([event(uid="other")]), True, now + timedelta(minutes=6), store)
    assert graph.deletes == 0


@pytest.mark.parametrize(
    "raw", [b"", b"<html>expired login</html>", b"BEGIN:VCALENDAR\r\nVERSION:2.0\r\nEND:VCALENDAR"]
)
def test_failed_or_empty_feed_never_mutates(config, store, now, raw):
    graph = FakeGraph()
    run(config, graph, feed([event()]), True, now, store)
    with pytest.raises(SyncError):
        run(config, graph, raw, True, now, store)
    assert len(graph.events) == 1 and graph.deletes == 0


def test_mass_deletion_guard(config, store, now):
    graph = FakeGraph()
    run(config, graph, feed([event(), event(uid="other")]), True, now, store)
    guarded = replace(config, max_delete_fraction=0.25)
    remaining = feed([event(uid="other")])
    run(guarded, graph, remaining, True, now, store)
    result = run(guarded, graph, remaining, True, now + timedelta(minutes=6), store)
    assert graph.deletes == 0 and result["warnings"]


def test_response_lost_recovers_without_duplicate(config, store, now):
    graph = FakeGraph()
    graph.fail_after_create = True
    with pytest.raises(SyncError):
        run(config, graph, feed([event()]), True, now, store)
    run(config, graph, feed([event()]), True, now, store)
    assert graph.creates == 1 and len(store.events()) == 1


def test_changed_cmis_segment_uid_preserves_identity(config, store, now):
    graph = FakeGraph()
    run(config, graph, feed([event()]), True, now, store)
    run(
        config,
        graph,
        feed([event(uid="1ID123:2026-10-01@timetabling.port.ac.uk")]),
        True,
        now,
        store,
    )
    assert graph.creates == 1 and graph.deletes == 0


def test_cross_day_move_is_held_until_identity_link(config, store, now):
    graph = FakeGraph()
    run(config, graph, feed([event()]), True, now, store)
    old = next(iter(graph.events))
    moved = feed([event(start="20261013T100000Z", end="20261013T120000Z")])
    result = run(config, graph, moved, True, now, store)
    assert graph.creates == 1 and all(a["kind"] == "hold" for a in result["actions"])
    new = next(a["key"] for a in result["actions"] if a["key"] != old)
    store.set_meta("aliases", {new: old})
    run(config, graph, moved, True, now, store)
    assert graph.updates == 1 and graph.creates == 1


def test_dst_keeps_local_start_time(config, now):
    raw = feed(
        [
            event()
            .replace("DTSTART:20261012T100000Z", "DTSTART;TZID=Europe/London:20261019T110000")
            .replace("DTEND:20261012T120000Z", "DTEND;TZID=Europe/London:20261019T130000")
            .replace("LOCATION:Room 1", "RRULE:FREQ=WEEKLY;COUNT=2")
        ]
    )
    sessions = parse_feed(raw, now, now + timedelta(days=40), config.source_id, config.timezone)
    assert [s.start for s in sessions] == ["2026-10-19T11:00:00+01:00", "2026-10-26T11:00:00+00:00"]


def test_exdate_removes_single_occurrence(config, now):
    raw = feed([event(extra="RRULE:FREQ=WEEKLY;COUNT=3\r\nEXDATE:20261019T100000Z")])
    assert (
        len(parse_feed(raw, now, now + timedelta(days=40), config.source_id, config.timezone)) == 2
    )


def test_self_directed_defaults_free(config, store, now):
    result = run(
        config,
        raw=feed(
            [
                event(
                    description="M33132/1\\nCOMPLEX PROBLEM SOLVING\\nTeaching Event\\nSelf Directive Studies"
                )
            ]
        ),
        now=now,
        store=store,
    )
    assert result["actions"][0]["show_as"] == "free"


def test_source_binding_prevents_accidental_calendar_switch(config, store, now):
    graph = FakeGraph()
    run(config, graph, feed([event()]), True, now, store)
    with pytest.raises(SyncError):
        run(replace(config, calendar_id="other"), graph, feed([event()]), True, now, store)


def test_explicit_cancelled_occurrence(config, store, now):
    graph = FakeGraph()
    run(config, graph, feed([event()]), True, now, store)
    result = run(config, graph, feed([event(extra="STATUS:CANCELLED")]), True, now, store)
    assert graph.deletes == 1, result

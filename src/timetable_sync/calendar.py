from __future__ import annotations

import re
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import recurring_ical_events
from icalendar import Calendar

from .models import Session, SyncError, digest

ACTIVITIES = {
    "tutorial",
    "workshop",
    "lecture",
    "problem based learning",
    "self directive studies",
    "self directed studies",
    "computer aided teaching",
    "drop-in",
    "induction wk event",
}


def parse_feed(
    raw: bytes, start: datetime, end: datetime, source_id: str, timezone: str
) -> list[Session]:
    """Expand RFC recurrence fields, never the descriptive teaching-week list.

    CMISGo's embedded event ID plus original local date survives segment changes
    and same-day time edits. Generic feeds use UID plus original occurrence time.
    Ambiguous identities abort; unexplained cross-day changes require review.
    """
    if (
        len(raw) > 5_000_000
        or not raw.strip().startswith(b"BEGIN:VCALENDAR")
        or not raw.strip().endswith(b"END:VCALENDAR")
    ):
        raise SyncError("Feed is oversized, incomplete, or not an iCalendar document.")
    zone = ZoneInfo(timezone)
    try:
        calendar = Calendar.from_ical(raw)
        if calendar.name != "VCALENDAR" or not calendar.get("VERSION"):
            raise ValueError("Missing calendar version")
        definitions = calendar.walk("VEVENT")
        if not definitions:
            raise SyncError("Empty feed refused; existing Outlook events are preserved.")
        if len(definitions) > 1000:
            raise SyncError("Too many event definitions; feed refused.")
        for event in definitions:
            if event.errors or not event.get("UID") or not event.get("DTSTART"):
                raise ValueError("Incomplete event")
            if not isinstance(event.decoded("DTSTART"), datetime):
                raise SyncError("All-day source events are unsupported; nothing was applied.")
            rule = event.get("RRULE")
            if rule and str(rule.get("FREQ", [""])[0]).upper() not in {
                "DAILY",
                "WEEKLY",
                "MONTHLY",
                "YEARLY",
            }:
                raise SyncError("Sub-daily recurrence is unsupported; feed refused.")
            recurrence_id = event.get("RECURRENCE-ID")
            if recurrence_id and recurrence_id.params.get("RANGE") == "THISANDFUTURE":
                raise SyncError("THISANDFUTURE exceptions require source review.")
            if not event.get("DTEND") and not event.get("DURATION"):
                raise SyncError("Source event has no end time or duration.")
        expanded = recurring_ical_events.of(calendar).between(start, end)
        sessions: dict[str, Session] = {}
        for event in expanded:
            beginning = event.decoded("DTSTART")
            ending = event.decoded("DTEND", beginning + event.decoded("DURATION", timedelta(0)))
            if not isinstance(beginning, datetime) or not isinstance(ending, datetime):
                raise ValueError("Invalid dates")
            if beginning.tzinfo is None:
                beginning = beginning.replace(tzinfo=zone)
            if ending.tzinfo is None:
                ending = ending.replace(tzinfo=zone)
            beginning, ending = beginning.astimezone(zone), ending.astimezone(zone)
            if ending <= beginning or (ending - beginning).total_seconds() > 86400:
                raise SyncError("Source session duration is invalid; nothing was applied.")
            if not start <= beginning < end:
                continue
            uid = str(event["UID"])
            cmis = re.match(r"^\d+ID(\d+):.*@timetabling\.port\.ac\.uk$", uid)
            family = "cmis:" + cmis[1] if cmis else "ical:" + uid
            original = event.decoded("RECURRENCE-ID", beginning)
            if isinstance(original, datetime):
                if original.tzinfo is None:
                    original = original.replace(tzinfo=zone)
                identity_date = original.astimezone(zone).date().isoformat()
            else:
                identity_date = original.isoformat()
            identity = identity_date if cmis else original.isoformat()
            key = digest([source_id, family, identity])
            description = str(event.get("DESCRIPTION", "")).strip()
            lines = [line.strip() for line in description.splitlines() if line.strip()]
            values = event.get("CATEGORIES", [])
            if not isinstance(values, list):
                values = [values]
            categories = [str(c).strip() for group in values for c in getattr(group, "cats", [])]
            module_match = re.search(
                r"\b([MI]\d{5})(?:/\w+)?\b", description + " " + " ".join(categories)
            )
            module = module_match[1] if module_match else ""
            activity = next(
                (line for line in lines + categories if line.casefold() in ACTIVITIES), "Other"
            )
            module_title = (
                lines[1]
                if module_match and len(lines) > 1
                else str(event.get("SUMMARY", "Class")).strip()
            )
            title = " Â· ".join(part for part in [module, module_title] if part)
            if activity != "Other":
                title += " â€” " + activity
            session = Session(
                key,
                family,
                uid,
                beginning.isoformat(),
                ending.isoformat(),
                identity_date,
                module,
                title,
                activity,
                str(event.get("LOCATION", "")).strip(),
                description,
                str(event.get("STATUS", "")).upper() == "CANCELLED",
            )
            if key in sessions and sessions[key] != session:
                raise SyncError("Ambiguous duplicate occurrence identity; nothing was applied.")
            sessions[key] = session
            if len(sessions) > 5000:
                raise SyncError("Too many expanded sessions; feed refused.")
        return sorted(sessions.values(), key=lambda item: (item.start, item.key))
    except SyncError:
        raise
    except Exception as exc:
        raise SyncError(
            "Could not validate or expand the timetable feed; nothing was applied."
        ) from exc

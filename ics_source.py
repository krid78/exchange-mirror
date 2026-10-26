"""Offline conversion of a complete ICS snapshot into importer records."""

from datetime import date, datetime, timedelta, timezone
from uuid import UUID

from icalendar import Calendar

from ics_model import EventFields, EventIdentity, EventStatus, SourceEvent


class IcsParseError(ValueError):
    """The entire snapshot is unusable; callers must not infer any deletions."""


def _single(component, name, *, required=False):
    value = component.get(name)
    if isinstance(value, list):
        raise IcsParseError(f"Duplicate {name} property")
    if required and value is None:
        raise IcsParseError(f"Missing {name} property")
    return value


def _utc_datetime(component, name, *, required=False):
    prop = _single(component, name, required=required)
    if prop is None:
        return None
    value = prop.dt
    if not isinstance(value, datetime) or value.utcoffset() != timedelta(0):
        raise IcsParseError(f"{name} must be a UTC datetime")
    return value.astimezone(timezone.utc)


def _event_time(component, name):
    prop = _single(component, name, required=True)
    if prop.params.get("TZID"):
        raise IcsParseError("TZID support is not implemented yet")
    value = prop.dt
    if isinstance(value, datetime):
        if value.utcoffset() != timedelta(0):
            raise IcsParseError("Only UTC event timestamps are supported yet")
    elif not isinstance(value, date):
        raise IcsParseError(f"Invalid {name} value")
    return value


def _parse_event(event, profile_id):
    for name in ("RRULE", "RDATE", "EXDATE", "RECURRENCE-ID", "EXRULE"):
        if name in event:
            raise IcsParseError("Recurring events are not supported yet")
    identity = EventIdentity(profile_id, str(_single(event, "UID", required=True)))
    status_prop = _single(event, "STATUS")
    status_text = str(status_prop).upper() if status_prop is not None else "CONFIRMED"
    try:
        status = EventStatus(status_text.lower())
    except ValueError:
        raise IcsParseError("Unsupported event STATUS") from None
    sequence_prop = _single(event, "SEQUENCE")
    sequence = int(sequence_prop) if sequence_prop is not None else 0
    stamp = _utc_datetime(event, "DTSTAMP", required=True)
    modified = _utc_datetime(event, "LAST-MODIFIED")
    _utc_datetime(event, "CREATED")
    if "DTEND" in event and "DURATION" in event:
        raise IcsParseError("DTEND and DURATION are mutually exclusive")
    fields = None
    if "DTSTART" in event:
        start = _event_time(event, "DTSTART")
        if "DTEND" in event:
            end = _event_time(event, "DTEND")
            if isinstance(start, datetime) == isinstance(end, datetime) and end <= start:
                raise IcsParseError("Explicit DTEND must follow DTSTART")
        elif "DURATION" in event:
            duration = _single(event, "DURATION").dt
            if not isinstance(duration, timedelta) or duration <= timedelta(0):
                raise IcsParseError("DURATION must be positive")
            if not isinstance(start, datetime) and duration.seconds:
                raise IcsParseError("DATE events require a whole-day duration")
            end = start + duration
        else:
            end = start if isinstance(start, datetime) else start + timedelta(days=1)
        categories = event.get("CATEGORIES", [])
        if not isinstance(categories, list):
            categories = [categories]
        fields = EventFields(
            start, end,
            subject=str(_single(event, "SUMMARY") or ""),
            body=str(_single(event, "DESCRIPTION") or ""),
            location=str(_single(event, "LOCATION") or ""),
            categories=tuple(str(category) for group in categories for category in group.cats),
        )
    elif status is not EventStatus.CANCELLED or "DTEND" in event or "DURATION" in event:
        raise IcsParseError("Missing DTSTART property")
    return SourceEvent(identity, fields, status, sequence, modified or stamp)


def _check_envelope(text):
    """Guard against tolerant parsing accepting truncated or trailing content."""
    stack = []
    finished = False
    for line in text.splitlines():
        if not line:
            continue
        if line.startswith((" ", "\t")):
            if not stack:
                raise IcsParseError("Content outside VCALENDAR")
            continue
        upper = line.upper()
        if finished or (not stack and upper != "BEGIN:VCALENDAR"):
            raise IcsParseError("Expected one complete VCALENDAR")
        if upper.startswith("BEGIN:"):
            name = upper[6:]
            if name == "VCALENDAR" and stack:
                raise IcsParseError("Nested VCALENDAR")
            stack.append(name)
        elif upper.startswith("END:"):
            if not stack or stack.pop() != upper[4:]:
                raise IcsParseError("Mismatched component boundary")
            if not stack:
                finished = True
    if stack or not finished:
        raise IcsParseError("Incomplete VCALENDAR")


def parse_ics(data: bytes | str, profile_id: UUID) -> tuple[SourceEvent, ...]:
    """Parse all events or fail the complete snapshot; performs no I/O.

    Step 3 supports UTC and DATE singleton events, not floating/TZID timestamps,
    VTIMEZONE definitions, recurrence rules or multiple revisions of one UID.
    """
    if not isinstance(profile_id, UUID):
        raise TypeError("profile_id must be a UUID")
    if not isinstance(data, (bytes, str)):
        raise TypeError("data must be bytes or text")
    try:
        text = data.decode("utf-8-sig") if isinstance(data, bytes) else data.removeprefix("\ufeff")
        _check_envelope(text)
        calendar = Calendar.from_ical(text)
        if calendar.name != "VCALENDAR":
            raise IcsParseError("Expected VCALENDAR")
        if str(_single(calendar, "VERSION", required=True)) != "2.0":
            raise IcsParseError("Only VERSION:2.0 is supported")
        _single(calendar, "PRODID", required=True)
        method = str(_single(calendar, "METHOD") or "PUBLISH").upper()
        if method != "PUBLISH":
            raise IcsParseError("Only publication snapshots are supported yet")
        for component in calendar.walk():
            if component.errors:
                raise IcsParseError("Invalid calendar property")
            if component.name not in ("VCALENDAR", "VEVENT", "VALARM"):
                raise IcsParseError("Unsupported calendar component")
            if component.name == "VEVENT" and component.subcomponents:
                if any(child.name != "VALARM" for child in component.subcomponents):
                    raise IcsParseError("Unsupported nested event component")
        if any(child.name != "VEVENT" for child in calendar.subcomponents):
            raise IcsParseError("Expected top-level VEVENT components")
        events = tuple(_parse_event(event, profile_id) for event in calendar.subcomponents)
        if len({event.identity for event in events}) != len(events):
            raise IcsParseError("Duplicate source identity; revision resolution is not implemented yet")
        return events
    except IcsParseError:
        raise
    except (ValueError, TypeError, AttributeError, KeyError, OverflowError):
        # Library diagnostics may contain raw source content. Do not propagate it.
        raise IcsParseError("Invalid ICS snapshot") from None

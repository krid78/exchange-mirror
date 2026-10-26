"""Offline conversion of a complete ICS snapshot into importer records."""

from datetime import date, datetime, timedelta, timezone
import re
from uuid import UUID
from zoneinfo import ZoneInfo

from dateutil.tz import datetime_ambiguous, datetime_exists
from icalendar import Calendar
from icalendar.parser import Contentline
from icalendar.timezone import TZP

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


def _resolve_zone(name, zones):
    if name in zones:
        return zones[name]
    try:
        return ZoneInfo(name)
    except (KeyError, ValueError):
        raise IcsParseError("Unknown timezone") from None


def _localize(value, zone):
    value = value.replace(tzinfo=zone, fold=0)
    if not datetime_exists(value):
        raise IcsParseError("Nonexistent local time at a timezone transition")
    if datetime_ambiguous(value):
        raise IcsParseError("Ambiguous local time; specify UTC instead")
    return value


def _event_time(component, name, zones, floating_zone, raw_properties):
    prop = _single(component, name, required=True)
    tzid = prop.params.get("TZID")
    raw_params, raw_value = raw_properties[name]
    if tzid is not None and (raw_params.get("VALUE") == "DATE" or raw_value.endswith("Z")):
        raise IcsParseError("TZID is only allowed on local DATE-TIME values")
    value = prop.dt
    if isinstance(value, datetime):
        if tzid is not None:
            # Re-resolve against this snapshot, not icalendar's global TZID cache.
            value = _localize(value.replace(tzinfo=None), _resolve_zone(str(tzid), zones))
        elif value.utcoffset() is None:
            if floating_zone is None:
                raise IcsParseError("Floating time requires an explicit profile timezone")
            value = _localize(value, floating_zone)
    elif not isinstance(value, date):
        raise IcsParseError(f"Invalid {name} value")
    elif tzid is not None:
        raise IcsParseError("DATE values must not have TZID")
    return value


def _parse_event(event, profile_id, zones, floating_zone, raw_properties):
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
        start = _event_time(event, "DTSTART", zones, floating_zone, raw_properties)
        if "DTEND" in event:
            end = _event_time(event, "DTEND", zones, floating_zone, raw_properties)
            comparable_start = start.astimezone(timezone.utc) if isinstance(start, datetime) else start
            comparable_end = end.astimezone(timezone.utc) if isinstance(end, datetime) else end
            if isinstance(start, datetime) == isinstance(end, datetime) and comparable_end <= comparable_start:
                raise IcsParseError("Explicit DTEND must follow DTSTART")
        elif "DURATION" in event:
            duration = _single(event, "DURATION").dt
            if not isinstance(duration, timedelta) or duration <= timedelta(0):
                raise IcsParseError("DURATION must be positive")
            if not isinstance(start, datetime) and duration.seconds:
                raise IcsParseError("DATE events require a whole-day duration")
            end = _duration_end(start, raw_properties["DURATION"][1])
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


def _duration_end(start, token):
    # Preserve nominal days versus exact hours: P1D and PT24H differ across DST.
    match = re.fullmatch(r"\+?P(?:(\d+)W|(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?)", token or "")
    if match is None:
        raise IcsParseError("Unsupported DURATION syntax")
    weeks, days, hours, minutes, seconds = (int(v or 0) for v in match.groups())
    clock = timedelta(hours=hours, minutes=minutes, seconds=seconds)
    if not isinstance(start, datetime) and (hours or minutes or seconds):
        raise IcsParseError("DATE events require nominal days or weeks")
    end = start + timedelta(days=7 * weeks + days)
    if isinstance(end, datetime):
        end = _localize(end.replace(tzinfo=None), start.tzinfo)
        end = end.astimezone(timezone.utc) + clock
    return end


def _event_time_tokens(text):
    """Retain lexical time details lost by the library's normalization."""
    lines = re.sub(r"\r?\n[ \t]", "", text).splitlines()
    stack, properties = [], []
    for line in lines:
        upper = line.upper()
        if upper.startswith("BEGIN:"):
            stack.append(upper[6:])
            if stack == ["VCALENDAR", "VEVENT"]:
                properties.append({})
        elif upper.startswith("END:"):
            stack.pop()
        elif stack == ["VCALENDAR", "VEVENT"]:
            name, params, value = Contentline(line).parts()
            name = name.upper()
            if name in ("DTSTART", "DTEND", "DURATION"):
                properties[-1][name] = (params, value)
    return properties


def _snapshot_zones(calendar):
    zones = {}
    for component in calendar.subcomponents:
        if component.name != "VTIMEZONE":
            continue
        name = str(_single(component, "TZID", required=True))
        if not name or name in zones:
            raise IcsParseError("Empty or duplicate VTIMEZONE TZID")
        if not component.subcomponents:
            raise IcsParseError("VTIMEZONE requires an observance")
        for observance in component.subcomponents:
            if observance.name not in ("STANDARD", "DAYLIGHT") or observance.subcomponents:
                raise IcsParseError("Invalid VTIMEZONE observance")
            for field in ("DTSTART", "TZOFFSETFROM", "TZOFFSETTO"):
                _single(observance, field, required=True)
        zones[name] = component.to_tz(TZP(), lookup_tzid=False)
    return zones


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


def parse_ics(data: bytes | str, profile_id: UUID, *, floating_timezone: str | None = None) -> tuple[SourceEvent, ...]:
    """Parse all events or fail the complete snapshot; performs no I/O.

    Floating timestamps require a profile timezone; no machine-local fallback.
    Event recurrence and multiple revisions of one UID remain unsupported.
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
            if component.name not in ("VCALENDAR", "VEVENT", "VALARM", "VTIMEZONE", "STANDARD", "DAYLIGHT"):
                raise IcsParseError("Unsupported calendar component")
            if component.name == "VEVENT" and component.subcomponents:
                if any(child.name != "VALARM" for child in component.subcomponents):
                    raise IcsParseError("Unsupported nested event component")
        if any(child.name not in ("VEVENT", "VTIMEZONE") for child in calendar.subcomponents):
            raise IcsParseError("Expected top-level VEVENT or VTIMEZONE components")
        zones = _snapshot_zones(calendar)
        floating_zone = _resolve_zone(floating_timezone, zones) if floating_timezone is not None else None
        raw_properties = iter(_event_time_tokens(text))
        events = tuple(_parse_event(event, profile_id, zones, floating_zone, next(raw_properties))
                       for event in calendar.subcomponents if event.name == "VEVENT")
        if len({event.identity for event in events}) != len(events):
            raise IcsParseError("Duplicate source identity; revision resolution is not implemented yet")
        return events
    except IcsParseError:
        raise
    except (ValueError, TypeError, AttributeError, KeyError, OverflowError):
        # Library diagnostics may contain raw source content. Do not propagate it.
        raise IcsParseError("Invalid ICS snapshot") from None

"""Offline conversion of a complete ICS snapshot into importer records."""

from dataclasses import dataclass, replace
from datetime import date, datetime, timedelta, timezone
import re
from uuid import UUID
from zoneinfo import ZoneInfo

from dateutil.tz import datetime_ambiguous, datetime_exists
from icalendar import Calendar, vDDDTypes
from icalendar.parser import Contentline
from icalendar.timezone import TZP

from ics_model import EventFields, EventIdentity, EventStatus, SourceEvent
from ics_recurrence import expand_starts


class IcsParseError(ValueError):
    """The entire snapshot is unusable; callers must not infer any deletions."""


@dataclass(frozen=True)
class _ParsedEvent:
    component: object
    source: SourceEvent
    raw: dict

    @property
    def rank(self):
        return self.source.sequence, self.source.revision_time


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
    raw_params, raw_value = raw_properties[name]
    return _time_value(prop.dt, raw_params, raw_value, zones, floating_zone)


def _time_value(value, raw_params, raw_value, zones, floating_zone):
    tzid = raw_params.get("TZID")
    if tzid is not None and (raw_params.get("VALUE") == "DATE" or raw_value.endswith("Z")):
        raise IcsParseError("TZID is only allowed on local DATE-TIME values")
    if isinstance(value, datetime):
        if tzid is not None:
            # Re-resolve against this snapshot, not icalendar's global TZID cache.
            value = _localize(value.replace(tzinfo=None), _resolve_zone(str(tzid), zones))
        elif value.utcoffset() is None:
            if floating_zone is None:
                raise IcsParseError("Floating time requires an explicit profile timezone")
            value = _localize(value, floating_zone)
    elif not isinstance(value, date):
        raise IcsParseError("Invalid event time value")
    elif tzid is not None:
        raise IcsParseError("DATE values must not have TZID")
    return value


def _parse_event(event, profile_id, zones, floating_zone, raw_properties):
    if "EXRULE" in event:
        raise IcsParseError("EXRULE is not supported")
    identity = EventIdentity(profile_id, str(_single(event, "UID", required=True)))
    recurrence = _single(event, "RECURRENCE-ID")
    if recurrence is not None:
        if "RANGE" in recurrence.params:
            raise IcsParseError("RECURRENCE-ID RANGE is not supported")
        if any(name in event for name in ("RRULE", "RDATE", "EXDATE")):
            raise IcsParseError("Detached exceptions cannot define a recurrence set")
        slot = _event_time(event, "RECURRENCE-ID", zones, floating_zone, raw_properties)
        identity = replace(identity, recurrence_id=slot)
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
        if recurrence is not None and isinstance(start, datetime) != isinstance(slot, datetime):
            raise IcsParseError("Exception DTSTART must match RECURRENCE-ID value type")
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
            if name in ("RDATE", "EXDATE"):
                properties[-1].setdefault(name, []).append((params, value))
            elif name in ("DTSTART", "DTEND", "DURATION", "RRULE", "RECURRENCE-ID"):
                properties[-1][name] = (params, value)
    return properties


def _expand_event(component, source, zones, floating_zone, raw, window_start,
                  window_end, max_occurrences, max_candidates):
    if not any(name in component for name in ("RRULE", "RDATE", "EXDATE")):
        return (source,)
    if source.fields is None or source.status is EventStatus.CANCELLED:
        raise IcsParseError("Recurring cancellations are not supported yet")
    start = _event_time(component, "DTSTART", zones, floating_zone, raw)
    dates = {}
    for name in ("RDATE", "EXDATE"):
        dates[name] = []
        for params, values in raw.get(name, []):
            if params.get("VALUE") == "PERIOD":
                raise IcsParseError("RDATE periods are not supported yet")
            for value in values.split(","):
                decoded = vDDDTypes.from_ical(value)
                value = _time_value(decoded, params, value, zones, floating_zone)
                if isinstance(value, datetime) != isinstance(start, datetime):
                    raise IcsParseError("Recurrence dates must match DTSTART value type")
                if isinstance(value, datetime):
                    # Use the same zone for set arithmetic and local window membership.
                    value = value.astimezone(start.tzinfo)
                    value = _localize(value.replace(tzinfo=None), start.tzinfo)
                dates[name].append(value)
    rule = _single(component, "RRULE")
    rule_text = raw["RRULE"][1] if rule is not None else None
    starts = expand_starts(start, rule_text, dates["RDATE"], dates["EXDATE"],
                           window_start, window_end, max_occurrences=max_occurrences,
                           max_candidates=max_candidates)
    result = []
    for slot in starts:
        if isinstance(slot, datetime):
            slot = _localize(slot.replace(tzinfo=None), slot.tzinfo)
        if "DURATION" in component:
            end = _duration_end(slot, raw["DURATION"][1])
        else:
            # An explicit DTEND supplies an exact duration, not a wall-clock end.
            end = slot + (source.fields.end - source.fields.start)
            if isinstance(slot, datetime):
                end = slot.astimezone(timezone.utc) + (source.fields.end - source.fields.start)
        identity = replace(source.identity, recurrence_id=slot)
        result.append(replace(source, identity=identity,
                              fields=replace(source.fields, start=slot, end=end)))
    return tuple(result)


def _select_revisions(records):
    groups = {}
    for record in records:
        groups.setdefault(record.source.identity, []).append(record)
    selected = []
    for versions in groups.values():
        rank = max(record.rank for record in versions)
        winners = [record for record in versions if record.rank == rank]
        first = winners[0]
        for other in winners[1:]:
            if (other.source != first.source or other.raw != first.raw
                    or other.component.to_ical() != first.component.to_ical()):
                raise IcsParseError("Conflicting equal-rank source revisions")
        selected.append(first)
    return selected


def _in_exception_window(source, zone, window_start, window_end):
    if type(window_start) is not date or type(window_end) is not date:
        raise IcsParseError("Recurrence exceptions require an explicit date window")
    values = [source.identity.recurrence_id]
    if source.fields is not None:
        values.append(source.fields.start)
    for value in values:
        local_date = value.astimezone(zone).date() if isinstance(value, datetime) else value
        if window_start <= local_date < window_end:
            return True
    return False


def _assemble_events(records, zones, floating_zone, window_start, window_end,
                     max_occurrences, max_candidates):
    groups = {}
    for record in _select_revisions(records):
        groups.setdefault(record.source.identity.uid, []).append(record)
    output = []
    for group in groups.values():
        master = next((r for r in group if r.source.identity.recurrence_id is None), None)
        exceptions = [r for r in group if r.source.identity.recurrence_id is not None]
        if master is not None and master.source.status is EventStatus.CANCELLED:
            if any(r.rank > master.rank for r in exceptions):
                raise IcsParseError("Exception is newer than the series cancellation")
            # UID-wide tombstone; the future planner still enforces target ownership/scope.
            output.append(replace(master.source, fields=None))
        else:
            zone = None
            if master is not None:
                if exceptions and not any(k in master.component for k in ("RRULE", "RDATE", "EXDATE")):
                    raise IcsParseError("Detached exception has a nonrecurring master")
                start = _event_time(master.component, "DTSTART", zones, floating_zone, master.raw)
                zone = start.tzinfo if isinstance(start, datetime) else None
                expanded = _expand_event(master.component, master.source, zones, floating_zone,
                                         master.raw, window_start, window_end,
                                         max_occurrences, max_candidates)
            else:
                expanded = ()
            # Remove the original slot even when the replacement moves outside the window.
            overridden = {r.source.identity for r in exceptions}
            output.extend(event for event in expanded if event.identity not in overridden)
            for record in exceptions:
                source = record.source
                slot = source.identity.recurrence_id
                if master is not None and isinstance(slot, datetime) != isinstance(start, datetime):
                    raise IcsParseError("Exception slot must match master DTSTART value type")
                exception_zone = zone
                if master is None and isinstance(slot, datetime):
                    local_slot = _event_time(record.component, "RECURRENCE-ID", zones,
                                             floating_zone, record.raw)
                    exception_zone = local_slot.tzinfo
                if _in_exception_window(source, exception_zone, window_start, window_end):
                    output.append(source)
        if len(output) > max_occurrences:
            raise IcsParseError("Snapshot occurrence limit exceeded")
    return tuple(output)


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


def parse_ics(data: bytes | str, profile_id: UUID, *, floating_timezone: str | None = None,
              window_start: date | None = None, window_end: date | None = None,
              max_occurrences: int = 10000, max_candidates: int = 100000) -> tuple[SourceEvent, ...]:
    """Parse all events or fail the complete snapshot; performs no I/O.

    Floating timestamps require a profile timezone; no machine-local fallback.
    Recurrences need a half-open local date window. Detached exceptions retain
    their original slot; revision conflicts fail the complete snapshot.
    """
    if not isinstance(profile_id, UUID):
        raise TypeError("profile_id must be a UUID")
    if not isinstance(data, (bytes, str)):
        raise TypeError("data must be bytes or text")
    try:
        for limit in (max_occurrences, max_candidates):
            if isinstance(limit, bool) or not isinstance(limit, int) or limit <= 0:
                raise IcsParseError("Expansion limits must be positive integers")
        if window_start is not None or window_end is not None:
            if (type(window_start) is not date or type(window_end) is not date
                    or window_end <= window_start):
                raise IcsParseError("Expansion window must contain increasing dates")
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
        records = []
        for component in calendar.subcomponents:
            if component.name != "VEVENT":
                continue
            raw = next(raw_properties)
            source = _parse_event(component, profile_id, zones, floating_zone, raw)
            records.append(_ParsedEvent(component, source, raw))
        events = _assemble_events(records, zones, floating_zone, window_start, window_end,
                                  max_occurrences, max_candidates)
        if len({event.identity for event in events}) != len(events):
            raise IcsParseError("Conflicting expanded event identities")
        return tuple(events)
    except IcsParseError:
        raise
    except (ValueError, TypeError, AttributeError, KeyError, OverflowError):
        # Library diagnostics may contain raw source content. Do not propagate it.
        raise IcsParseError("Invalid ICS snapshot") from None

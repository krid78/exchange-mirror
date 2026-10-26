"""Bounded recurrence starts, independent of ICS parsing and Outlook."""

from datetime import date, datetime, time

from dateutil.rrule import rrulestr


_RULE_PARTS = {
    "FREQ", "UNTIL", "COUNT", "INTERVAL", "BYSECOND", "BYMINUTE", "BYHOUR",
    "BYDAY", "BYMONTHDAY", "BYYEARDAY", "BYWEEKNO", "BYMONTH", "BYSETPOS", "WKST",
}


def expand_starts(start, rule, rdates, exdates, window_start, window_end,
                  *, max_occurrences=10000, max_candidates=100000):
    """Return original slots whose local start date is in [start, end).

    Limits fail the entire expansion. DTSTART is part of the recurrence set,
    EXDATE takes precedence, and repeated RDATE values collapse to one slot.
    """
    if (type(window_start) is not date or type(window_end) is not date
            or window_end <= window_start):
        raise ValueError("Recurrence requires an increasing date window")
    for limit in (max_occurrences, max_candidates):
        if isinstance(limit, bool) or not isinstance(limit, int) or limit <= 0:
            raise ValueError("Recurrence limits must be positive integers")
    all_day = not isinstance(start, datetime)

    def slot_date(value):
        if all_day:
            return value
        return value.astimezone(start.tzinfo).date()

    selected = set()
    excluded = set(exdates)

    def include(value):
        if value not in excluded and window_start <= slot_date(value) < window_end:
            selected.add(value)
            if len(selected) > max_occurrences:
                raise ValueError("Recurrence occurrence limit exceeded")

    include(start)
    if len(rdates) + len(exdates) > max_candidates:
        raise ValueError("Recurrence date-list limit exceeded")
    for value in rdates:
        include(value)
    if rule is not None:
        parts = {}
        for field in rule.upper().split(";"):
            key, value = field.split("=", 1)
            if key not in _RULE_PARTS or key in parts or not value:
                raise ValueError("Unsupported or duplicate RRULE part")
            parts[key] = value
        if "FREQ" not in parts or {"UNTIL", "COUNT"} <= parts.keys():
            raise ValueError("Invalid RRULE termination")
        for key in ("COUNT", "INTERVAL"):
            if key in parts and int(parts[key]) <= 0:
                raise ValueError("RRULE count/interval must be positive")
        if all_day and {"BYSECOND", "BYMINUTE", "BYHOUR"} & parts.keys():
            raise ValueError("DATE recurrences cannot specify clock fields")
        until = parts.get("UNTIL")
        if until is not None:
            if all_day and (len(until) != 8 or not until.isdigit()):
                raise ValueError("DATE recurrence needs a DATE UNTIL")
            if not all_day and not until.endswith("Z"):
                raise ValueError("Timed recurrence needs a UTC UNTIL")
        dtstart = datetime.combine(start, time()) if all_day else start
        iterator = rrulestr(rule, dtstart=dtstart, cache=False)
        for index, value in enumerate(iterator, 1):
            if index > max_candidates:
                raise ValueError("Recurrence candidate limit exceeded")
            value = value.date() if all_day else value
            if slot_date(value) >= window_end:
                break
            include(value)
    return tuple(sorted(selected))

"""Snapshot timezone isolation, daylight-saving boundaries and all-day semantics."""

from datetime import date, datetime, timedelta, timezone
import unittest

from ics_source import IcsParseError, parse_ics
from test_source import calendar, EVENT, FIXTURES, PROFILE


def event_at(start, ending=""):
    return EVENT.replace("DTSTART:20261027T130000Z", start) + ("\r\n" + ending if ending else "")


def zone_component(offset="+0300", name="Synthetic/Local"):
    return ("BEGIN:VTIMEZONE\r\nTZID:" + name + "\r\nBEGIN:STANDARD\r\n"
            "DTSTART:19700101T000000\r\nTZOFFSETFROM:" + offset + "\r\n"
            "TZOFFSETTO:" + offset + "\r\nEND:STANDARD\r\nEND:VTIMEZONE\r\n")


def with_zone(event, component):
    return calendar(event).replace("BEGIN:VEVENT", component + "BEGIN:VEVENT", 1)


class TimezoneTests(unittest.TestCase):
    def test_timetable_snapshot_and_dst_definition(self):
        events = parse_ics((FIXTURES / "timetable.ics").read_bytes(), PROFILE)
        self.assertEqual(len(events), 2)
        self.assertEqual(events[0].fields.start, datetime(2026, 10, 27, 13, tzinfo=timezone.utc))
        self.assertEqual(events[1].fields.start, datetime(2026, 10, 28, 10, 30, tzinfo=timezone.utc))

    def test_iana_zone_winter_and_summer(self):
        for month, utc_hour in ((1, 11), (7, 10)):
            with self.subTest(month=month):
                event = event_at(f"DTSTART;TZID=Europe/Berlin:2026{month:02}15T120000")
                parsed = parse_ics(calendar(event), PROFILE)[0]
                self.assertEqual(parsed.fields.start,
                                 datetime(2026, month, 15, utc_hour, tzinfo=timezone.utc))

    def test_explicit_end_across_dst_uses_elapsed_time(self):
        event = event_at("DTSTART;TZID=Europe/Berlin:20260329T013000",
                         "DTEND;TZID=Europe/Berlin:20260329T033000")
        fields = parse_ics(calendar(event), PROFILE)[0].fields
        self.assertEqual(fields.end - fields.start, timedelta(hours=1))

    def test_nominal_days_differ_from_exact_hours(self):
        for day, hours in (("20260328", 23), ("20261024", 25)):
            with self.subTest(day=day):
                start = f"DTSTART;TZID=Europe/Berlin:{day}T120000"
                nominal = parse_ics(calendar(event_at(start, "DURATION:P1D")), PROFILE)[0].fields
                exact = parse_ics(calendar(event_at(start, "DURATION:PT24H")), PROFILE)[0].fields
                self.assertEqual(nominal.end - nominal.start, timedelta(hours=hours))
                self.assertEqual(exact.end - exact.start, timedelta(hours=24))

    def test_floating_time_requires_explicit_profile_zone(self):
        source = calendar(event_at("DTSTART:20260715T120000"))
        with self.assertRaises(IcsParseError):
            parse_ics(source, PROFILE)

        fields = parse_ics(source, PROFILE, floating_timezone="Europe/Berlin")[0].fields
        self.assertEqual(fields.start, datetime(2026, 7, 15, 10, tzinfo=timezone.utc))
        # Calendar-level display metadata must not implicitly choose a profile zone.
        source = source.replace("VERSION:2.0", "VERSION:2.0\r\nX-WR-TIMEZONE:Europe/Berlin")
        with self.assertRaises(IcsParseError):
            parse_ics(source, PROFILE)


    def test_custom_floating_zone_and_mixed_duration_units(self):
        source = with_zone(event_at("DTSTART:20260715T120000"), zone_component())
        fields = parse_ics(source, PROFILE, floating_timezone="Synthetic/Local")[0].fields
        self.assertEqual(fields.start, datetime(2026, 7, 15, 9, tzinfo=timezone.utc))
        source = calendar(event_at("DTSTART;TZID=Europe/Berlin:20260328T120000",
                                   "DURATION:P1DT2H"))
        fields = parse_ics(source, PROFILE)[0].fields
        self.assertEqual(fields.end - fields.start, timedelta(hours=25))

    def test_all_day_week_duration_and_missing_end(self):
        start = "DTSTART;VALUE=DATE:20261024"
        default = parse_ics(calendar(event_at(start)), PROFILE)[0].fields
        week = parse_ics(calendar(event_at(start, "DURATION:P1W")), PROFILE)[0].fields
        self.assertEqual(default.end, date(2026, 10, 25))
        self.assertEqual(week.end, date(2026, 10, 31))
    def test_unknown_zone_never_falls_back_to_profile_zone(self):
        source = calendar(event_at("DTSTART;TZID=Unknown/Nowhere:20260715T120000"))
        with self.assertRaises(IcsParseError):
            parse_ics(source, PROFILE, floating_timezone="Europe/Berlin")
        with self.assertRaises(IcsParseError):
            parse_ics(calendar(EVENT), PROFILE, floating_timezone="Unknown/Nowhere")

    def test_custom_vtimezone_overrides_and_is_snapshot_local(self):
        for name in ("Synthetic/Local", "Europe/Berlin"):
            for offset, expected_hour in (("+0300", 9), ("+0400", 8), ("+0300", 9)):
                with self.subTest(name=name, offset=offset):
                    source = with_zone(event_at(f"DTSTART;TZID={name}:20260715T120000"),
                                       zone_component(offset, name))
                    fields = parse_ics(source, PROFILE)[0].fields
                    self.assertEqual(fields.start, datetime(2026, 7, 15, expected_hour,
                                                           tzinfo=timezone.utc))
        # The library may cache the synthetic name; our resolver must not use that cache.
        with self.assertRaises(IcsParseError):
            parse_ics(calendar(event_at("DTSTART;TZID=Synthetic/Local:20260715T120000")), PROFILE)
        normal = parse_ics(calendar(event_at("DTSTART;TZID=Europe/Berlin:20260715T120000")), PROFILE)
        self.assertEqual(normal[0].fields.start.hour, 10)

    def test_nonexistent_and_ambiguous_local_times_fail(self):
        for stamp in ("20260329T023000", "20261025T023000"):
            for start, kwargs in ((f"DTSTART;TZID=Europe/Berlin:{stamp}", {}),
                                  (f"DTSTART:{stamp}", {"floating_timezone": "Europe/Berlin"})):
                with self.subTest(start=start), self.assertRaises(IcsParseError):
                    parse_ics(calendar(event_at(start)), PROFILE, **kwargs)

    def test_all_day_dates_remain_exclusive_across_dst(self):
        source = calendar(event_at("DTSTART;VALUE=DATE:20260328", "DTEND;VALUE=DATE:20260331"))
        fields = parse_ics(source, PROFILE, floating_timezone="Europe/Berlin")[0].fields
        self.assertEqual(fields.start, date(2026, 3, 28))
        self.assertEqual(fields.end, date(2026, 3, 31))
        self.assertTrue(fields.all_day)
        for ending in ("DURATION:PT24H", "DURATION:PT1H", "DTEND;VALUE=DATE:20260328",
                       "DTEND:20260331T000000Z"):
            with self.subTest(ending=ending), self.assertRaises(IcsParseError):
                parse_ics(calendar(event_at("DTSTART;VALUE=DATE:20260328", ending)), PROFILE)

    def test_illegal_tzid_and_broken_zone_definitions(self):
        for start in ("DTSTART;VALUE=DATE;TZID=Europe/Berlin:20260715",
                      "DTSTART;TZID=Europe/Berlin:20260715T120000Z"):
            with self.subTest(start=start), self.assertRaises(IcsParseError):
                parse_ics(calendar(event_at(start)), PROFILE)
        for component in (zone_component() + zone_component(),
                          zone_component().replace("TZOFFSETTO:+0300\r\n", ""),
                          "BEGIN:VTIMEZONE\r\nTZID:Empty/Zone\r\nEND:VTIMEZONE\r\n"):
            with self.subTest(component=component), self.assertRaises(IcsParseError):
                parse_ics(with_zone(EVENT, component), PROFILE)


if __name__ == "__main__":
    unittest.main()

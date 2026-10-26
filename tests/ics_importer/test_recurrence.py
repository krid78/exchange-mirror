"""Finite recurrence sets and stable original slots, without Outlook or network."""

from datetime import date, datetime, timedelta, timezone
import unittest

from ics_source import IcsParseError, parse_ics
from test_source import calendar, EVENT, PROFILE
from test_timezones import event_at


WINDOW = dict(window_start=date(2026, 10, 27), window_end=date(2026, 10, 30))


class RecurrenceTests(unittest.TestCase):
    def parse(self, event, **kwargs):
        return parse_ics(calendar(event), PROFILE, **(WINDOW | kwargs))

    def test_infinite_daily_rule_is_limited_by_exclusive_window(self):
        events = self.parse(EVENT + "\r\nRRULE:FREQ=DAILY")
        self.assertEqual([e.fields.start.day for e in events], [27, 28, 29])
        self.assertTrue(all(e.identity.recurrence_id == e.fields.start for e in events))
        self.assertEqual(len({e.identity.key for e in events}), 3)
        self.assertEqual(events, self.parse(EVENT + "\r\nRRULE:FREQ=DAILY"))

    def test_count_until_and_weekly_byday(self):
        for rule in ("FREQ=DAILY;COUNT=2", "FREQ=DAILY;UNTIL=20261028T130000Z"):
            with self.subTest(rule=rule):
                self.assertEqual(len(self.parse(EVENT + "\r\nRRULE:" + rule)), 2)
        events = self.parse(EVENT + "\r\nRRULE:FREQ=WEEKLY;BYDAY=TU,TH")
        self.assertEqual([e.fields.start.day for e in events], [27, 29])

    def test_rdate_union_exdate_precedence_and_duplicates(self):
        source = (EVENT + "\r\nRRULE:FREQ=DAILY;COUNT=3"
                  "\r\nRDATE:20261027T130000Z,20261028T130000Z,20261029T150000Z"
                  "\r\nRDATE:20261029T150000Z,20261030T130000Z"
                  "\r\nEXDATE:20261027T130000Z,20261028T130000Z")
        events = self.parse(source)
        self.assertEqual([(e.fields.start.day, e.fields.start.hour) for e in events], [(29, 13), (29, 15)])
        self.assertEqual(len({e.identity.key for e in events}), 2)

    def test_rdate_without_rule_and_exdate_without_rule(self):
        events = self.parse(EVENT + "\r\nRDATE:20261028T130000Z")
        self.assertEqual(len(events), 2)
        events = self.parse(EVENT + "\r\nEXDATE:20261027T130000Z")
        self.assertEqual(events, ())

    def test_all_day_recurrence_and_type_validation(self):
        source = event_at("DTSTART;VALUE=DATE:20261027", "DURATION:P2D")
        events = self.parse(source + "\r\nRRULE:FREQ=DAILY;UNTIL=20261029"
                           "\r\nEXDATE;VALUE=DATE:20261028")
        self.assertEqual([e.fields.start for e in events], [date(2026, 10, 27), date(2026, 10, 29)])
        self.assertTrue(all(e.fields.end - e.fields.start == timedelta(days=2) for e in events))
        self.assertTrue(all(type(e.identity.recurrence_id) is date for e in events))
        for extra in ("RDATE:20261028T130000Z", "RRULE:FREQ=DAILY;BYHOUR=12",
                      "RRULE:FREQ=DAILY;UNTIL=20261029T130000Z"):
            with self.subTest(extra=extra), self.assertRaises(IcsParseError):
                self.parse(source + "\r\n" + extra)

    def test_daily_wall_time_survives_summer_time_transition(self):
        source = event_at("DTSTART;TZID=Europe/Berlin:20261024T120000", "DURATION:PT1H")
        events = self.parse(source + "\r\nRRULE:FREQ=DAILY;COUNT=3",
                            window_start=date(2026, 10, 24), window_end=date(2026, 10, 27))
        self.assertEqual([e.fields.start.hour for e in events], [10, 11, 11])
        self.assertTrue(all(e.fields.end - e.fields.start == timedelta(hours=1) for e in events))

    def test_duration_days_reapplied_for_each_occurrence(self):
        source = event_at("DTSTART;TZID=Europe/Berlin:20260328T120000", "DURATION:P1D")
        events = self.parse(source + "\r\nRRULE:FREQ=DAILY;COUNT=2",
                            window_start=date(2026, 3, 28), window_end=date(2026, 3, 30))
        self.assertEqual([e.fields.end - e.fields.start for e in events],
                         [timedelta(hours=23), timedelta(hours=24)])

    def test_rdate_exdate_equivalent_instants_in_other_zone(self):
        source = event_at("DTSTART;TZID=Europe/Berlin:20261027T140000")
        events = self.parse(source + "\r\nRDATE:20261027T130000Z,20261028T130000Z"
                            "\r\nEXDATE;TZID=Europe/Berlin:20261028T140000")
        self.assertEqual(len(events), 1)
        self.assertEqual(events[0].identity.recurrence_id,
                         datetime(2026, 10, 27, 13, tzinfo=timezone.utc))

    def test_window_uses_local_dates_not_utc_dates(self):
        source = event_at("DTSTART;TZID=Asia/Tokyo:20261027T003000")
        events = self.parse(source + "\r\nRRULE:FREQ=DAILY;COUNT=1",
                            window_end=date(2026, 10, 28))
        self.assertEqual(len(events), 1)
        self.assertEqual(events[0].fields.start.date(), date(2026, 10, 26))

    def test_limits_fail_instead_of_returning_truncated_results(self):
        for source, options in ((EVENT + "\r\nRRULE:FREQ=DAILY", {"max_occurrences": 2}),
                                (EVENT + "\r\nRRULE:FREQ=SECONDLY", {"max_candidates": 3}),
                                (EVENT + "\r\nRDATE:20261028T130000Z,20261029T130000Z",
                                 {"max_candidates": 1})):
            with self.subTest(options=options), self.assertRaises(IcsParseError):
                self.parse(source, **options)
        historical = EVENT.replace("20261027T130000Z", "20000101T130000Z")
        with self.assertRaises(IcsParseError):
            self.parse(historical + "\r\nRRULE:FREQ=DAILY", max_candidates=10)

    def test_invalid_or_missing_window_and_unsupported_rules(self):
        source = calendar(EVENT + "\r\nRRULE:FREQ=DAILY")
        with self.assertRaises(IcsParseError):
            parse_ics(source, PROFILE)
        for options in ({"window_end": WINDOW["window_start"]}, {"max_occurrences": 0},
                        {"max_candidates": True}, {"window_start": datetime(2026, 10, 27)}):
            with self.subTest(options=options), self.assertRaises(IcsParseError):
                self.parse(EVENT + "\r\nRRULE:FREQ=DAILY", **options)
        for rule in ("FREQ=DAILY;COUNT=0", "FREQ=DAILY;INTERVAL=0",
                     "FREQ=DAILY;COUNT=2;UNTIL=20261028T130000Z",
                     "FREQ=DAILY;UNTIL=20261028T130000", "FREQ=YEARLY;BYEASTER=1",
                     "FREQ=DAILY;FREQ=WEEKLY"):
            with self.subTest(rule=rule), self.assertRaises(IcsParseError):
                self.parse(EVENT + "\r\nRRULE:" + rule)

    def test_unsupported_ranges_periods_and_conflicting_masters_are_rejected(self):
        for prop in ("RECURRENCE-ID;RANGE=THISANDFUTURE:20261027T130000Z", "EXRULE:FREQ=DAILY",
                     "RDATE;VALUE=PERIOD:20261028T130000Z/PT1H"):
            with self.subTest(prop=prop), self.assertRaises(IcsParseError):
                self.parse(EVENT + "\r\n" + prop)
        with self.assertRaises(IcsParseError):
            parse_ics(calendar(EVENT + "\r\nRRULE:FREQ=DAILY", EVENT), PROFILE, **WINDOW)

    def test_singletons_remain_unfiltered_and_keep_single_identity(self):
        event = self.parse(EVENT, window_start=date(2027, 1, 1), window_end=date(2027, 2, 1))[0]
        self.assertIsNone(event.identity.recurrence_id)

    def test_total_snapshot_limit_and_generated_transition_errors(self):
        one = EVENT + "\r\nRRULE:FREQ=DAILY;COUNT=2"
        two = one.replace("UID:synthetic", "UID:another")
        with self.assertRaises(IcsParseError):
            parse_ics(calendar(one, two), PROFILE, **WINDOW, max_occurrences=3)
        for stamp in ("20260328T023000", "20261024T023000"):
            event = event_at(f"DTSTART;TZID=Europe/Berlin:{stamp}")
            with self.subTest(stamp=stamp), self.assertRaises(IcsParseError):
                self.parse(event + "\r\nRRULE:FREQ=DAILY;COUNT=2",
                           window_start=date(2026, 1, 1), window_end=date(2027, 1, 1))


if __name__ == "__main__":
    unittest.main()

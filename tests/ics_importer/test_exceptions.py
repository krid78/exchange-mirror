"""Published exceptions, cancellations and order-independent revision selection."""

from datetime import date, datetime, timezone
import unittest

from ics_model import EventStatus
from ics_source import IcsParseError, parse_ics
from test_source import calendar, EVENT, PROFILE


WINDOW = dict(window_start=date(2026, 10, 27), window_end=date(2026, 10, 30))
MASTER = EVENT + "\r\nRRULE:FREQ=DAILY;COUNT=3"
SLOT = "RECURRENCE-ID:20261028T130000Z"
MOVED = EVENT.replace("20261027T130000Z", "20261028T150000Z") + "\r\n" + SLOT
CANCELLED = ("UID:synthetic@example.invalid\r\nDTSTAMP:20260101T120000Z"
             "\r\nSTATUS:CANCELLED")


class ExceptionTests(unittest.TestCase):
    def parse(self, *events, **kwargs):
        return parse_ics(calendar(*events), PROFILE, **(WINDOW | kwargs))

    def test_moved_exception_replaces_original_and_preserves_key(self):
        events = self.parse(MASTER, MOVED)
        self.assertEqual(len(events), 3)
        changed = next(e for e in events if e.identity.recurrence_id.day == 28)
        self.assertEqual(changed.fields.start.hour, 15)
        self.assertEqual(changed.identity.recurrence_id.hour, 13)
        original = self.parse(MASTER)[1]
        self.assertEqual(changed.identity.key, original.identity.key)
        self.assertEqual(len({e.identity.key for e in events}), 3)

    def test_exception_can_move_out_of_or_into_window(self):
        outside = MOVED.replace("DTSTART:20261028T150000Z", "DTSTART:20261105T150000Z")
        events = self.parse(MASTER, outside)
        changed = next(e for e in events if e.identity.recurrence_id.day == 28)
        self.assertEqual(changed.fields.start.month, 11)
        incoming = MOVED.replace(SLOT, "RECURRENCE-ID:20261026T130000Z")
        events = self.parse(MASTER, incoming)
        self.assertTrue(any(e.identity.recurrence_id.day == 26 for e in events))

    def test_cancelled_occurrence_is_a_tombstone_not_an_active_copy(self):
        events = self.parse(MASTER, CANCELLED + "\r\n" + SLOT)
        self.assertEqual(len(events), 3)
        removed = next(e for e in events if e.identity.recurrence_id.day == 28)
        self.assertEqual(removed.status, EventStatus.CANCELLED)
        self.assertIsNone(removed.fields)
        self.assertEqual(sum(e.status is not EventStatus.CANCELLED for e in events), 2)

    def test_series_cancellation_is_uid_wide_without_expansion(self):
        events = self.parse(MASTER, CANCELLED + "\r\nSEQUENCE:1", MOVED)
        self.assertEqual(len(events), 1)
        self.assertEqual(events[0].status, EventStatus.CANCELLED)
        self.assertIsNone(events[0].fields)
        self.assertIsNone(events[0].identity.recurrence_id)
        # A CANCELLED master with RRULE is also one UID-wide tombstone.
        self.assertEqual(len(self.parse(MASTER + "\r\nSTATUS:CANCELLED")), 1)
        with self.assertRaises(IcsParseError):
            self.parse(CANCELLED, MOVED + "\r\nSEQUENCE:1")

    def test_newer_active_revision_can_supersede_cancellation(self):
        active = MASTER + "\r\nSEQUENCE:2"
        events = self.parse(CANCELLED + "\r\nSEQUENCE:1", active)
        self.assertEqual(len(events), 3)
        self.assertTrue(all(e.status is EventStatus.CONFIRMED for e in events))
        events = self.parse(MASTER, MOVED, CANCELLED + "\r\n" + SLOT + "\r\nSEQUENCE:1")
        removed = next(e for e in events if e.identity.recurrence_id.day == 28)
        self.assertEqual(removed.status, EventStatus.CANCELLED)

    def test_timezone_equivalent_slot_and_exdate_override(self):
        moved = MOVED.replace(SLOT, "RECURRENCE-ID;TZID=Europe/Berlin:20261028T140000")
        master = MASTER + "\r\nEXDATE:20261028T130000Z"
        events = self.parse(master, moved)
        self.assertEqual(len(events), 3)
        changed = next(e for e in events if e.identity.recurrence_id.day == 28)
        self.assertEqual(changed.identity.key, self.parse(MASTER)[1].identity.key)
        self.assertEqual(changed.fields.start.hour, 15)

    def test_orphan_exception_and_outside_scope(self):
        self.assertEqual(len(self.parse(MOVED)), 1)
        self.assertEqual(len(self.parse(CANCELLED + "\r\n" + SLOT)), 1)
        outside = MOVED.replace("20261028", "20261105")
        self.assertEqual(self.parse(outside), ())
        with self.assertRaises(IcsParseError):
            parse_ics(calendar(MOVED), PROFILE)

    def test_equal_rank_conflicts_fail_but_identical_revisions_collapse(self):
        self.assertEqual(len(self.parse(EVENT, EVENT)), 1)
        for conflict in (EVENT + "\r\nSUMMARY:Different", EVENT + "\r\nSTATUS:CANCELLED"):
            with self.subTest(conflict=conflict), self.assertRaises(IcsParseError):
                self.parse(EVENT, conflict)
        with self.assertRaises(IcsParseError):
            self.parse(MASTER, MOVED, MOVED + "\r\nSUMMARY:Different")

    def test_revision_rank_is_sequence_then_timestamp(self):
        old = EVENT + "\r\nSEQUENCE:1\r\nLAST-MODIFIED:20261001T120000Z"
        new = EVENT + "\r\nSEQUENCE:2\r\nSUMMARY:New"
        for versions in ((old, new), (new, old)):
            self.assertEqual(self.parse(*versions)[0].fields.subject, "New")
        newer_time = old.replace("20261001", "20261002") + "\r\nSUMMARY:Later"
        self.assertEqual(self.parse(newer_time, old)[0].fields.subject, "Later")
        newer_stamp = EVENT.replace("20260101T120000Z", "20260102T120000Z")
        self.assertEqual(self.parse(EVENT, newer_stamp)[0].revision_time,
                         datetime(2026, 1, 2, 12, tzinfo=timezone.utc))

    def test_obsolete_conflicts_do_not_make_resolution_order_dependent(self):
        conflict = EVENT + "\r\nSUMMARY:Obsolete conflict"
        winner = EVENT + "\r\nSEQUENCE:1\r\nSUMMARY:Winner"
        for versions in ((EVENT, conflict, winner), (winner, conflict, EVENT)):
            self.assertEqual(self.parse(*versions)[0].fields.subject, "Winner")

    def test_exception_revisions_are_separate_from_master_revision(self):
        master = MASTER + "\r\nSEQUENCE:5"
        changed = MOVED + "\r\nSEQUENCE:1"
        latest = MOVED.replace("150000Z", "160000Z") + "\r\nSEQUENCE:2"
        events = self.parse(changed, master, latest)
        event = next(e for e in events if e.identity.recurrence_id.day == 28)
        self.assertEqual(event.fields.start.hour, 16)
        self.assertEqual(event.sequence, 2)

    def test_all_day_exception_and_mixed_type_rejection(self):
        master = MASTER.replace("DTSTART:20261027T130000Z", "DTSTART;VALUE=DATE:20261027")
        moved = (EVENT.replace("DTSTART:20261027T130000Z", "DTSTART;VALUE=DATE:20261029")
                 + "\r\nRECURRENCE-ID;VALUE=DATE:20261028")
        events = self.parse(master, moved)
        event = next(e for e in events if e.identity.recurrence_id == date(2026, 10, 28))
        self.assertEqual(event.fields.start, date(2026, 10, 29))
        for exception in (moved, CANCELLED + "\r\nRECURRENCE-ID;VALUE=DATE:20261028"):
            with self.subTest(exception=exception), self.assertRaises(IcsParseError):
                self.parse(MASTER, exception)
        with self.assertRaises(IcsParseError):
            self.parse(MOVED.replace(SLOT, "RECURRENCE-ID;VALUE=DATE:20261028"))

    def test_unsupported_ranges_and_invalid_detached_rules_fail(self):
        for exception in (MOVED.replace(SLOT, SLOT.replace(":", ";RANGE=THISANDFUTURE:", 1)),
                          MOVED + "\r\nRRULE:FREQ=DAILY",
                          MOVED + "\r\nRECURRENCE-ID:20261029T130000Z"):
            with self.subTest(exception=exception), self.assertRaises(IcsParseError):
                self.parse(MASTER, exception)
        with self.assertRaises(IcsParseError):
            self.parse(EVENT, MOVED)


if __name__ == "__main__":
    unittest.main()

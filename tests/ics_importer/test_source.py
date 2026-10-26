"""Offline singleton parsing; no network or Outlook access."""

from datetime import date, datetime, timedelta, timezone
from pathlib import Path
import unittest
from uuid import UUID

from ics_model import EventStatus
from ics_source import IcsParseError, parse_ics


PROFILE = UUID(int=1)
FIXTURES = Path(__file__).resolve().parent / "fixtures"


def calendar(*events):
    return ("BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:synthetic-tests\r\n"
            + "".join("BEGIN:VEVENT\r\n" + event + "\r\nEND:VEVENT\r\n" for event in events)
            + "END:VCALENDAR\r\n")


EVENT = ("UID:synthetic@example.invalid\r\nDTSTAMP:20260101T120000Z\r\n"
         "DTSTART:20261027T130000Z")


class SourceTests(unittest.TestCase):
    def test_minimal_fixture(self):
        events = parse_ics((FIXTURES / "minimal.ics").read_bytes(), PROFILE)
        self.assertEqual(len(events), 1)
        event = events[0]
        self.assertEqual(event.identity.profile_id, PROFILE)
        self.assertEqual(event.identity.uid, "minimal-event@example.invalid")
        self.assertIsNone(event.identity.recurrence_id)
        self.assertEqual(event.fields.start, datetime(2026, 10, 27, 13, tzinfo=timezone.utc))
        self.assertEqual(event.fields.end - event.fields.start, timedelta(hours=1))
        self.assertEqual(event.fields.subject, "Synthetic appointment")

    def test_utf8_folding_escapes_and_categories(self):
        source = calendar(EVENT + '\r\nSUMMARY:Übung: langer\r\n Titel'
                          '\r\nDESCRIPTION:Zeile 1\\nZeile 2\\, mit Komma\\; Semikolon\\\\'
                          '\r\nLOCATION:Raum Ü\r\nCATEGORIES:Lehre,Übung'
                          '\r\nCATEGORIES:Extra')
        event = parse_ics(source.encode("utf-8"), PROFILE)[0]
        self.assertEqual(event.fields.subject, "Übung: langerTitel")
        self.assertEqual(event.fields.body, 'Zeile 1\nZeile 2, mit Komma; Semikolon\\')
        self.assertEqual(event.fields.location, "Raum Ü")
        self.assertEqual(event.fields.categories, ("Lehre", "Übung", "Extra"))

    def test_defaults_and_zero_duration(self):
        event = parse_ics(calendar(EVENT), PROFILE)[0]
        self.assertEqual(event.fields.start, event.fields.end)
        self.assertEqual((event.fields.subject, event.fields.body, event.fields.location), ("", "", ""))
        self.assertEqual(event.fields.categories, ())
        self.assertEqual(event.status, EventStatus.CONFIRMED)
        self.assertEqual(event.sequence, 0)

    def test_duration_and_exclusive_date_defaults(self):
        timed = parse_ics(calendar(EVENT + "\r\nDURATION:PT90M"), PROFILE)[0]
        self.assertEqual(timed.fields.end - timed.fields.start, timedelta(minutes=90))
        source = EVENT.replace("DTSTART:20261027T130000Z", "DTSTART;VALUE=DATE:20261027")
        all_day = parse_ics(calendar(source), PROFILE)[0]
        self.assertEqual(all_day.fields.end, date(2026, 10, 28))
        self.assertTrue(all_day.fields.all_day)
        multi = parse_ics(calendar(source + "\r\nDURATION:P2D"), PROFILE)[0]
        self.assertEqual(multi.fields.end, date(2026, 10, 29))

    def test_status_sequence_and_revision(self):
        event = parse_ics(calendar(EVENT + "\r\nSTATUS:TENTATIVE\r\nSEQUENCE:3"
                                  "\r\nLAST-MODIFIED:20261001T120000Z"), PROFILE)[0]
        self.assertEqual(event.status, EventStatus.TENTATIVE)
        self.assertEqual(event.sequence, 3)
        self.assertEqual(event.revision_time, datetime(2026, 10, 1, 12, tzinfo=timezone.utc))
        fallback = parse_ics(calendar(EVENT), PROFILE)[0]
        self.assertEqual(fallback.revision_time, datetime(2026, 1, 1, 12, tzinfo=timezone.utc))

    def test_cancelled_tombstone(self):
        source = "UID:cancelled@example.invalid\r\nDTSTAMP:20260101T120000Z\r\nSTATUS:CANCELLED"
        event = parse_ics(calendar(source), PROFILE)[0]
        self.assertEqual(event.status, EventStatus.CANCELLED)
        self.assertIsNone(event.fields)

    def test_multiple_events_and_valid_empty_snapshot(self):
        events = parse_ics(calendar(EVENT, EVENT.replace("UID:synthetic", "UID:second")), PROFILE)
        self.assertEqual(len(events), 2)
        self.assertEqual(parse_ics(calendar(), PROFILE), ())

    def test_missing_and_duplicate_single_properties(self):
        invalid = [EVENT.replace("UID:synthetic@example.invalid\r\n", ""),
                   EVENT.replace("UID:synthetic@example.invalid", "UID:"),
                   EVENT.replace("DTSTAMP:20260101T120000Z\r\n", ""),
                   EVENT.replace("DTSTART:20261027T130000Z", ""),
                   EVENT + "\r\nUID:second@example.invalid",
                   EVENT + "\r\nSUMMARY:first\r\nSUMMARY:second",
                   EVENT + "\r\nDTEND:20261027T140000Z\r\nDURATION:PT1H"]
        for source in invalid:
            with self.subTest(source=source), self.assertRaises(IcsParseError):
                parse_ics(calendar(source), PROFILE)

    def test_invalid_property_values_fail_complete_snapshot(self):
        invalid = [EVENT.replace("20261027T130000Z", "INVALID"),
                   EVENT + "\r\nDTEND:20261027T120000Z",
                   EVENT + "\r\nDTEND:20261027T130000Z",
                   EVENT + "\r\nDTEND;VALUE=DATE:20261028",
                   EVENT + "\r\nSTATUS:UNKNOWN",
                   EVENT + "\r\nSTATUS:",
                   EVENT + "\r\nSEQUENCE:-1",
                   EVENT + "\r\nSEQUENCE:INVALID",
                   EVENT + "\r\nDURATION:-PT1H",
                   EVENT + "\r\nDURATION:PT0S",
                   EVENT + "\r\nDURATION:INVALID",
                   EVENT + "\r\nLAST-MODIFIED:INVALID"]
        for source in invalid:
            with self.subTest(source=source), self.assertRaises(IcsParseError):
                parse_ics(calendar(EVENT.replace("UID:synthetic", "UID:valid"), source), PROFILE)

    def test_recurrence_and_unconfigured_floating_time_are_rejected(self):
        for prop in ("RRULE:FREQ=DAILY", "RDATE:20261028T130000Z",
                     "EXDATE:20261028T130000Z", "RECURRENCE-ID:20261027T130000Z",
                     "EXRULE:FREQ=DAILY"):
            with self.subTest(prop=prop), self.assertRaises(IcsParseError):
                parse_ics(calendar(EVENT + "\r\n" + prop), PROFILE)
        for start in ("DTSTART:20261027T130000",):
            with self.subTest(start=start), self.assertRaises(IcsParseError):
                parse_ics(calendar(EVENT.replace("DTSTART:20261027T130000Z", start)), PROFILE)

    def test_invalid_envelopes_and_other_components(self):
        good = calendar(EVENT)
        invalid = ["", "<html>login</html>", good.replace("END:VCALENDAR", ""),
                   good.replace("END:VEVENT", "END:VTODO"), good + good,
                   "unexpected\r\n" + good, good + "unexpected",
                   good.replace("VERSION:2.0", "VERSION:1.0"),
                   good.replace("PRODID:synthetic-tests\r\n", ""),
                   good.replace("VERSION:2.0", "VERSION:2.0\r\nMETHOD:REQUEST"),
                   good.replace("VEVENT", "VTODO"),
                   good.replace("VERSION:2.0", "VERSION:2.0\r\nVERSION:2.0")]
        for source in invalid:
            with self.subTest(source=source), self.assertRaises(IcsParseError):
                parse_ics(source, PROFILE)
        with self.assertRaises(IcsParseError):
            parse_ics(b"\xff", PROFILE)

    def test_duplicate_identity_selects_newer_revision(self):
        events = parse_ics(calendar(EVENT, EVENT + "\r\nSEQUENCE:1"), PROFILE)
        self.assertEqual(len(events), 1)
        self.assertEqual(events[0].sequence, 1)

    def test_source_alarm_and_attendees_are_not_copied(self):
        source = EVENT + ("\r\nATTENDEE:mailto:synthetic@example.invalid"
                          "\r\nBEGIN:VALARM\r\nACTION:DISPLAY\r\n"
                          "TRIGGER:-PT15M\r\nDESCRIPTION:Example\r\nEND:VALARM")
        event = parse_ics(calendar(source), PROFILE)[0]
        self.assertEqual(event.fields, parse_ics(calendar(EVENT), PROFILE)[0].fields)

    def test_lf_crlf_and_bom(self):
        source = calendar(EVENT)
        expected = parse_ics(source, PROFILE)
        self.assertEqual(parse_ics(source.replace("\r\n", "\n"), PROFILE), expected)
        self.assertEqual(parse_ics(b"\xef\xbb\xbf" + source.encode(), PROFILE), expected)


if __name__ == "__main__":
    unittest.main()

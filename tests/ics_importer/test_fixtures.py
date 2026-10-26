"""Offline integrity checks for the initial synthetic test assets, not an ICS parser."""

import unittest
from pathlib import Path


FIXTURES = Path(__file__).resolve().parent / "fixtures"


class FixtureTests(unittest.TestCase):
    def test_fixture_envelopes_and_synthetic_identities(self):
        for filename, expected_events in (("minimal.ics", 1), ("timetable.ics", 2)):
            with self.subTest(filename=filename):
                # Accept LF in source checkouts and CRLF after Windows Git conversion.
                lines = (FIXTURES / filename).read_text(encoding="utf-8").splitlines()
                self.assertEqual(lines[0], "BEGIN:VCALENDAR")
                self.assertEqual(lines[-1], "END:VCALENDAR")
                self.assertIn("VERSION:2.0", lines)
                stack = []
                for line in lines:
                    if line.startswith("BEGIN:"):
                        stack.append(line.removeprefix("BEGIN:"))
                    elif line.startswith("END:"):
                        self.assertTrue(stack, "Unexpected component end")
                        self.assertEqual(stack.pop(), line.removeprefix("END:"))
                self.assertEqual(stack, [], "Unclosed component")
                self.assertEqual(lines.count("BEGIN:VEVENT"), expected_events)
                uids = [line.removeprefix("UID:") for line in lines if line.startswith("UID:")]
                self.assertEqual(len(uids), expected_events)
                self.assertEqual(len(set(uids)), expected_events)
                self.assertTrue(all(uid.endswith("@example.invalid") for uid in uids))

    def test_timetable_captures_observed_source_characteristics(self):
        lines = (FIXTURES / "timetable.ics").read_text(encoding="utf-8").splitlines()
        self.assertIn("METHOD:PUBLISH", lines)
        self.assertIn("TZID:Europe/Berlin", lines)
        self.assertIn("BEGIN:VTIMEZONE", lines)
        self.assertEqual(sum(line.startswith("RRULE:") for line in lines), 2)
        inside_event = False
        for line in lines:
            if line == "BEGIN:VEVENT":
                inside_event = True
            elif line == "END:VEVENT":
                inside_event = False
            elif inside_event:
                self.assertFalse(line.startswith(("RRULE:", "RECURRENCE-ID")))
        self.assertEqual(lines.count("STATUS:TENTATIVE"), 2)
        self.assertEqual(lines.count("X-MICROSOFT-CDO-BUSYSTATUS:TENTATIVE"), 2)
        self.assertEqual(lines.count("TRANSP:OPAQUE"), 2)
        self.assertIn("SUMMARY:Übung: Synthetische Systeme", lines)


if __name__ == "__main__":
    unittest.main()

"""Identity and normalized-record contracts; no ICS parser or Outlook required."""

from dataclasses import FrozenInstanceError, replace
from datetime import date, datetime, timedelta, timezone
import unittest
from uuid import UUID

from ics_model import (
    EventFields, EventIdentity, EventStatus, OperationKind, SourceEvent,
    SyncOperation, TargetState,
)


PROFILE = UUID("00000000-0000-0000-0000-000000000001")
START = datetime(2026, 10, 27, 13, tzinfo=timezone.utc)


class IdentityTests(unittest.TestCase):
    def test_v1_serialization_contract(self):
        identity = EventIdentity(PROFILE, "event@example.invalid")
        self.assertEqual(identity.serialize(),
                         '["ics-event-v1","00000000-0000-0000-0000-000000000001",'
                         '"event@example.invalid",["single"]]')
        self.assertEqual(len(identity.key), 64)
        self.assertEqual(identity.key, EventIdentity(PROFILE, identity.uid).key)

    def test_equivalent_instants_share_identity(self):
        local = START.astimezone(timezone(timedelta(hours=1)))
        a = EventIdentity(PROFILE, "event", START)
        b = EventIdentity(PROFILE, "event", local)
        self.assertEqual(a, b)
        self.assertEqual(a.serialize(), b.serialize())
        self.assertEqual(a.key, b.key)

    def test_slots_profiles_and_uids_remain_distinct(self):
        identities = [
            EventIdentity(PROFILE, "event"),
            EventIdentity(PROFILE, "Event"),
            EventIdentity(UUID(int=2), "event"),
            EventIdentity(PROFILE, "event", START.date()),
            EventIdentity(PROFILE, "event", START.replace(hour=0)),
            EventIdentity(PROFILE, "event", START),
            EventIdentity(PROFILE, "event", START.replace(microsecond=1)),
        ]
        self.assertEqual(len({identity.key for identity in identities}), len(identities))

    def test_opaque_uid_survives_serialization(self):
        import json
        uid = ' Übung::["date"]\\\n '
        identity = EventIdentity(PROFILE, uid)
        self.assertEqual(json.loads(identity.serialize())[2], uid)
        self.assertNotEqual(identity.key, EventIdentity(PROFILE, uid.strip()).key)

    def test_invalid_identity_inputs(self):
        for uid in ("", " ", None):
            with self.subTest(uid=uid), self.assertRaises(ValueError):
                EventIdentity(PROFILE, uid)
        with self.assertRaises(TypeError):
            EventIdentity(str(PROFILE), "event")
        with self.assertRaises(ValueError):
            EventIdentity(PROFILE, "event", START.replace(tzinfo=None))
        with self.assertRaises(TypeError):
            EventIdentity(PROFILE, "event", "2026-10-27")


class ModelTests(unittest.TestCase):
    def setUp(self):
        self.identity = EventIdentity(PROFILE, "event")
        self.fields = EventFields(START, START + timedelta(hours=1), "Original")

    def test_edits_and_rescheduling_do_not_change_identity(self):
        for identity in (self.identity, replace(self.identity, recurrence_id=START)):
            original = SourceEvent(identity, self.fields)
            changed = replace(original, fields=replace(
                self.fields, subject="Changed", start=START + timedelta(days=1),
                end=START + timedelta(days=1, hours=1)))
            self.assertEqual(original.identity.key, changed.identity.key)
            self.assertNotEqual(original.fields, changed.fields)

    def test_date_and_timestamp_fields(self):
        all_day = EventFields(date(2026, 10, 27), date(2026, 10, 29))
        self.assertTrue(all_day.all_day)
        self.assertEqual(all_day.end, date(2026, 10, 29))
        local = START.astimezone(timezone(timedelta(hours=1)))
        timed = EventFields(local, local + timedelta(hours=1))
        self.assertFalse(timed.all_day)
        self.assertEqual(timed, replace(self.fields, subject=""))
        self.assertEqual(timed.start.tzinfo, timezone.utc)
        EventFields(START, START)  # RFC permits a zero-duration timed event.

    def test_invalid_times_are_rejected(self):
        for start, end in ((START, START.date()), (START, START - timedelta(seconds=1)),
                           (START.date(), START.date()),
                           (START.replace(tzinfo=None), START), ("today", "tomorrow")):
            with self.subTest(start=start, end=end), self.assertRaises((TypeError, ValueError)):
                EventFields(start, end)

    def test_cancellation_and_revisions(self):
        cancelled = SourceEvent(self.identity, None, EventStatus.CANCELLED)
        self.assertIsNone(cancelled.fields)
        with self.assertRaises(ValueError):
            SourceEvent(self.identity, None)
        for sequence in (-1, True, 1.5):
            with self.subTest(sequence=sequence), self.assertRaises(ValueError):
                SourceEvent(self.identity, self.fields, sequence=sequence)
        local = START.astimezone(timezone(timedelta(hours=1)))
        event = SourceEvent(self.identity, self.fields, revision_time=local)
        self.assertEqual(event.revision_time.tzinfo, timezone.utc)
        with self.assertRaises(ValueError):
            replace(event, revision_time=START.replace(tzinfo=None))

    def test_records_are_immutable_and_copy_collection_inputs(self):
        categories = ["Example"]
        fields = replace(self.fields, categories=categories)
        categories.append("Later")
        self.assertEqual(fields.categories, ("Example",))
        with self.assertRaises(FrozenInstanceError):
            fields.subject = "Changed"
        with self.assertRaises(FrozenInstanceError):
            self.identity.uid = "Changed"

    def test_target_state_keeps_observed_flags_and_locator_out_of_key(self):
        target = TargetState(self.identity, "entry-one", self.fields, 1, True)
        other = replace(target, entry_id="entry-two", busy_status=2)
        self.assertEqual(target.busy_status, 1)
        self.assertEqual(target.identity.key, other.identity.key)

    def test_operations_retain_duplicate_evidence(self):
        first = TargetState(self.identity, "one", self.fields, 2, False)
        targets = [first, replace(first, entry_id="two")]
        operation = SyncOperation(OperationKind.CONFLICT, self.identity, targets=targets,
                                  reason="Duplicate target identity")
        targets.clear()
        self.assertEqual(len(operation.targets), 2)
        self.assertEqual({kind.value for kind in OperationKind},
                         {"create", "update", "delete", "unchanged", "conflict"})

    def test_operations_reject_mismatched_identities(self):
        other = replace(self.identity, uid="other")
        with self.assertRaises(ValueError):
            SyncOperation(OperationKind.CREATE, self.identity,
                          source=SourceEvent(other, self.fields))
        with self.assertRaises(ValueError):
            SyncOperation(OperationKind.UPDATE, self.identity,
                          targets=(TargetState(other, "one", self.fields, 2, False),))


if __name__ == "__main__":
    unittest.main()

"""Pure upsert plans from normalized records, without Outlook or I/O."""

from dataclasses import replace
from datetime import datetime, timedelta, timezone
from unittest import TestCase
from uuid import UUID

from ics_model import EventFields, EventIdentity, EventStatus, OperationKind, SourceEvent, TargetState
from ics_profiles import Destination
from ics_sync import TargetInventory, UnmanagedItem, plan_upserts, project_event
from test_profiles import profile


START = datetime(2026, 10, 27, 13, tzinfo=timezone.utc)


def source(uid="event"):
    return SourceEvent(EventIdentity(profile().profile_id, uid),
                       EventFields(START, START + timedelta(hours=1), "Public", "Details", "Room", ("Lecture",)),
                       EventStatus.TENTATIVE)


def target(event=None, entry_id="one", **changes):
    event = event or source()
    return replace(TargetState(event.identity, entry_id, event.fields, 2, False), **changes)


def inventory(*items):
    return TargetInventory(profile().destination, items)


class SyncTests(TestCase):
    def kinds(self, result):
        return [op.kind for op in result.operations]

    def test_create_then_repeated_snapshot_is_unchanged(self):
        created = plan_upserts(profile(), [source()], inventory())
        self.assertEqual(self.kinds(created), [OperationKind.CREATE])
        desired = project_event(created.operations[0].source)
        observed = target(fields=desired.fields, busy_status=desired.busy_status,
                          reminder_set=desired.reminder_set, meeting_status=desired.meeting_status)
        repeated = plan_upserts(profile(), [source()], inventory(observed))
        self.assertEqual(self.kinds(repeated), [OperationKind.UNCHANGED])
        self.assertEqual(desired.fields.subject, "Public")
        self.assertEqual(desired.fields.body, "Details")

    def test_managed_fields_and_fixed_flags_trigger_one_update(self):
        original = target()
        for changed in (replace(original, busy_status=1), replace(original, reminder_set=True),
                        replace(original, meeting_status=3),
                        replace(original, fields=replace(original.fields, subject="Old")),
                        replace(original, fields=replace(original.fields, location="Elsewhere")),
                        replace(original, fields=replace(original.fields, start=START - timedelta(hours=1)))):
            with self.subTest(changed=changed):
                result = plan_upserts(profile(), [source()], inventory(changed))
                self.assertEqual(self.kinds(result), [OperationKind.UPDATE])
                self.assertEqual(result.operations[0].targets[0].entry_id, "one")

    def test_body_and_category_normalization_avoids_false_updates(self):
        observed = target(fields=replace(source().fields, body="Details\r\n", categories=("lecture",)))
        self.assertEqual(self.kinds(plan_upserts(profile(), [source()], inventory(observed))),
                         [OperationKind.UNCHANGED])

    def test_duplicates_and_unknown_schema_are_conflicts(self):
        for items in ((target(), target(entry_id="two")), (target(marker_schema_version=2),)):
            result = plan_upserts(profile(), [source()], inventory(*items))
            self.assertTrue(result.has_conflicts)
            self.assertEqual(self.kinds(result), [OperationKind.CONFLICT])
        self.assertTrue(plan_upserts(profile(), [source(), source()], inventory()).has_conflicts)

    def test_other_profiles_unmarked_and_missing_sources_never_delete(self):
        foreign = target(entry_id="foreign", identity=EventIdentity(UUID(int=2), "event"))
        result = plan_upserts(profile(), [], inventory(foreign, UnmanagedItem("manual"), target()))
        self.assertEqual(result.operations, ())
        cancelled = replace(source(), status=EventStatus.CANCELLED, fields=None)
        self.assertEqual(plan_upserts(profile(), [cancelled], inventory(target())).operations, ())

    def test_destination_and_source_profile_must_match(self):
        wrong = TargetInventory(Destination("other", "folder", "Same display name"), ())
        with self.assertRaises(ValueError):
            plan_upserts(profile(), [source()], wrong)
        with self.assertRaises(ValueError):
            plan_upserts(profile(), [replace(source(), identity=EventIdentity(UUID(int=2), "event"))], inventory())
        renamed = TargetInventory(replace(profile().destination, display_name="Renamed"), ())
        self.assertEqual(self.kinds(plan_upserts(profile(), [source()], renamed)), [OperationKind.CREATE])

    def test_outside_window_and_rescheduled_known_target(self):
        moved = replace(source(), fields=replace(source().fields, start=START + timedelta(days=400),
                                                 end=START + timedelta(days=400, hours=1)))
        self.assertEqual(plan_upserts(profile(), [moved], inventory()).operations, ())
        self.assertEqual(self.kinds(plan_upserts(profile(), [moved], inventory(target()))), [OperationKind.UPDATE])
        occurrence = replace(moved, identity=replace(source().identity, recurrence_id=START))
        self.assertEqual(self.kinds(plan_upserts(profile(), [occurrence], inventory())), [OperationKind.CREATE])

    def test_invalid_inventory_and_cancelled_projection(self):
        with self.assertRaises(ValueError):
            inventory(target(), target())
        with self.assertRaises(ValueError):
            project_event(replace(source(), status=EventStatus.CANCELLED))

"""Deletion authority and safety gates over pure source/target snapshots."""

from dataclasses import replace
from datetime import date, timedelta
from unittest import TestCase
from uuid import UUID

from ics_model import EventIdentity, EventStatus, OperationKind
from ics_profiles import Destination, ExportCoverage
from ics_sync import TargetInventory, UnmanagedItem, plan_sync
from test_profiles import profile
from test_sync import START, source, target


def covered_profile():
    return replace(profile(), source_coverage=ExportCoverage(date(2026, 1, 1), date(2027, 1, 1), True))


def inventory(*items):
    return TargetInventory(profile().destination, items)


def deletes(plan):
    return [op for op in plan.operations if op.kind is OperationKind.DELETE]


class DeletionTests(TestCase):
    def plan(self, sources=(), items=(), selected=None, **options):
        return plan_sync(selected or covered_profile(), sources, inventory(*items),
                         **(dict(snapshot_complete=True, confirm_deletions=True,
                                 confirm_empty_snapshot=True, confirm_large_removal=True) | options))

    def test_missing_owned_event_with_confirmed_coverage(self):
        plan = self.plan(items=[target()])
        self.assertEqual(len(deletes(plan)), 1)
        self.assertEqual(deletes(plan)[0].targets[0].entry_id, "one")
        self.assertEqual(plan.pending_deletions, ())

    def test_unknown_unconfirmed_or_partial_coverage_never_authorizes_absence(self):
        for coverage in (None, ExportCoverage(date(2026, 1, 1), date(2027, 1, 1)),
                         ExportCoverage(date(2026, 11, 1), date(2027, 1, 1), True)):
            with self.subTest(coverage=coverage):
                plan = self.plan(items=[target()], selected=replace(profile(), source_coverage=coverage))
                self.assertEqual(deletes(plan), [])
                self.assertEqual(len(plan.pending_deletions), 1)
                self.assertTrue(plan.notices)

    def test_empty_large_and_normal_confirmation_are_independent(self):
        for flag in ("confirm_deletions", "confirm_empty_snapshot", "confirm_large_removal"):
            with self.subTest(flag=flag):
                plan = self.plan(items=[target()], **{flag: False})
                self.assertEqual(deletes(plan), [])
                self.assertTrue(plan.requires_confirmation)
        plan = plan_sync(covered_profile(), (), inventory(target()), snapshot_complete=True)
        self.assertEqual(deletes(plan), [])
        self.assertEqual(len(plan.pending_deletions[0].blockers), 3)

    def test_large_drop_with_nonempty_source(self):
        keep, gone, other = source("keep"), source("gone"), source("other")
        items = [target(keep, "keep"), target(gone, "gone"), target(other, "other")]
        plan = self.plan([keep], items, confirm_large_removal=False)
        self.assertEqual(deletes(plan), [])
        self.assertEqual(len(plan.pending_deletions), 2)
        self.assertEqual(len(deletes(self.plan([keep], items))), 2)

    def test_foreign_unmarked_and_historical_items_survive(self):
        foreign = target(entry_id="foreign", identity=EventIdentity(UUID(int=2), "same"))
        old = target(entry_id="old", fields=replace(source().fields,
                      start=START - timedelta(days=400), end=START - timedelta(days=400, hours=-1)))
        plan = self.plan(items=[foreign, old, UnmanagedItem("manual")])
        self.assertEqual(deletes(plan), [])
        self.assertEqual(plan.pending_deletions, ())

    def test_window_shrink_and_moved_occurrence_do_not_prove_absence(self):
        narrowed = replace(covered_profile(), window_start=date(2026, 11, 1))
        self.assertEqual(deletes(self.plan(items=[target()], selected=narrowed)), [])
        moved = target(identity=replace(source().identity, recurrence_id=START - timedelta(days=400)))
        plan = self.plan(items=[moved])
        self.assertEqual(deletes(plan), [])
        self.assertIn("Original slot", plan.pending_deletions[0].blockers[0])
        moved_out = replace(source(), fields=replace(source().fields,
                            start=START + timedelta(days=400), end=START + timedelta(days=400, hours=1)))
        plan = self.plan([moved_out], [target()])
        self.assertEqual(deletes(plan), [])
        self.assertEqual(plan.operations[0].kind, OperationKind.UPDATE)

    def test_explicit_occurrence_cancellation_does_not_need_absence_coverage(self):
        event = replace(source(), identity=replace(source().identity, recurrence_id=START))
        cancelled = replace(event, status=EventStatus.CANCELLED, fields=None)
        plan = self.plan([cancelled], [target(event)], selected=profile())
        self.assertEqual(len(deletes(plan)), 1)
        self.assertEqual(deletes(plan)[0].source, cancelled)

    def test_uid_wide_cancellation_only_deletes_owned_in_scope_instances(self):
        cancel = replace(source(), status=EventStatus.CANCELLED, fields=None)
        event = replace(source(), identity=replace(source().identity, recurrence_id=START))
        different = source("unrelated")
        plan = self.plan([cancel, different], [target(event, "occurrence"), target(different, "other")],
                         selected=profile())
        self.assertEqual([op.targets[0].entry_id for op in deletes(plan)], ["occurrence"])
        self.assertIsNone(deletes(plan)[0].source)

    def test_duplicates_and_unknown_schemas_block_all_deletion(self):
        for extras in ((target(source("duplicate"), "two"), target(source("duplicate"), "three")),
                       (target(source("unknown"), "two", marker_schema_version=2),),
                       (target(source("unknown"), "two", marker_schema_version=True),)):
            with self.subTest(extras=extras):
                plan = self.plan(items=[target(), *extras])
                self.assertTrue(plan.has_conflicts)
                self.assertEqual(deletes(plan), [])
                self.assertTrue(plan.pending_deletions)

    def test_duplicate_or_contradictory_source_blocks_deletions(self):
        event = source("keep")
        plan = self.plan([event, event], [target(source("gone"))])
        self.assertTrue(plan.has_conflicts)
        self.assertEqual(deletes(plan), [])
        occurrence = replace(source(), identity=replace(source().identity, recurrence_id=START))
        cancel = replace(source(), status=EventStatus.CANCELLED, fields=None)
        plan = self.plan([cancel, occurrence], [target(source("gone"))])
        self.assertTrue(plan.has_conflicts)
        self.assertEqual(deletes(plan), [])

    def test_incomplete_snapshot_wrong_destination_and_invalid_threshold_fail(self):
        with self.assertRaises(ValueError):
            self.plan(items=[target()], snapshot_complete=False)
        with self.assertRaises(ValueError):
            plan_sync(covered_profile(), (), TargetInventory(Destination("wrong", "wrong", "Calendar"), ()),
                      snapshot_complete=True)
        for fraction in (0, 2, True, float("nan")):
            with self.subTest(fraction=fraction), self.assertRaises(ValueError):
                self.plan(maximum_removal_fraction=fraction)

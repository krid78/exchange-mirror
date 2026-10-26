"""Coordinator ordering, freshness, recovery and cancellation with in-memory state."""

from contextlib import nullcontext
from dataclasses import replace
from datetime import datetime, timezone
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from unittest import TestCase
from unittest.mock import Mock

from ics_http import FeedSnapshot
from ics_lock import RunBusyError, target_lock
from ics_model import OperationKind, TargetState
from ics_outlook import OutlookCreateError, OutlookUpdateError, OutlookReadError
from ics_run import ImportCoordinator, PreviewError, RunApprovals, RunStatus
from ics_source import parse_ics
from ics_sync import TargetInventory, UnmanagedItem, project_event
from test_deletions import covered_profile
from test_source import EVENT, calendar


def event(uid, title="Current"):
    return EVENT.replace("synthetic@example.invalid", uid) + "\r\nDURATION:PT1H\r\nSUMMARY:" + title


def snapshot(body):
    return FeedSnapshot(body, "example.invalid", datetime.now(timezone.utc), "ignored-untrusted-digest")


class MemoryAdapter:
    def __init__(self, selected):
        self.profile, self.items, self.calls = selected, [], []
        self.counter = 0
        self.after_create = self.after_update = self.after_delete = None

    def read_inventory(self, destination):
        return TargetInventory(destination, tuple(self.items))

    def create_appointment(self, destination, source):
        self.calls.append("create")
        desired = project_event(source)
        self.counter += 1
        state = TargetState(source.identity, f"new-{self.counter}", desired.fields, 2, False)
        self.items.append(state)
        if self.after_create:
            self.after_create()
        return state

    def update_appointment(self, destination, source, expected):
        self.calls.append("update")
        state = replace(expected, fields=source.fields, busy_status=2, reminder_set=False, meeting_status=0)
        self.items[self.items.index(expected)] = state
        if self.after_update:
            self.after_update()
        return state

    def delete_appointment(self, destination, operation, *, confirmed):
        assert confirmed is True
        self.calls.append("delete")
        expected = operation.targets[0]
        self.items.remove(expected)
        if self.after_delete:
            self.after_delete()
        return SimpleNamespace(entry_id=expected.entry_id)


class CoordinatorTests(TestCase):
    def test_read_failure_diagnostic_survives_without_retry_or_raw_error_text(self):
        preview = self.runner.preview(self.profile)
        self.adapter.create_appointment = Mock(side_effect=OutlookReadError("PRIVATE", code="inventory_changed"))
        with self.assertLogs("ics_importer.run", level="WARNING") as captured:
            result = self.runner.execute(self.profile, preview)
        self.assertEqual(result.status, RunStatus.FAILED)
        self.assertEqual(result.failure.diagnostic, "inventory_changed")
        self.assertFalse(result.failure.outcome_unknown)
        self.adapter.create_appointment.assert_called_once()
        self.assertIn("inventory_changed", str(captured.output))
        self.assertNotIn("PRIVATE", str(captured.output) + repr(result.failure))

    def setUp(self):
        self.temp = TemporaryDirectory(dir=Path(__file__).resolve().parent)
        self.addCleanup(self.temp.cleanup)
        self.profile = covered_profile()
        self.adapter = MemoryAdapter(self.profile)
        self.body = calendar(event("a")).encode()
        self.fetcher = Mock(side_effect=lambda url: snapshot(self.body))
        self.write_modes = []
        def session(*, write):
            self.write_modes.append(write)
            return nullcontext(self.adapter)
        self.runner = ImportCoordinator(fetcher=self.fetcher, session_factory=session,
                                        lock_directory=self.temp.name)
        self.approvals = RunApprovals(True, True, True)

    def seed(self, *events):
        sources = parse_ics(calendar(*events), self.profile.profile_id)
        for source in sources:
            self.adapter.create_appointment(self.profile.destination, source)
        self.adapter.calls.clear()

    def test_preview_is_read_only_and_second_run_is_unchanged(self):
        preview = self.runner.preview(self.profile)
        self.assertEqual(self.write_modes, [False])
        self.assertEqual(self.adapter.calls, [])
        result = self.runner.execute(self.profile, preview)
        self.assertEqual(result.status, RunStatus.COMPLETED)
        self.assertEqual(result.counts["create"], 1)
        repeated = self.runner.execute(self.profile, self.runner.preview(self.profile))
        self.assertEqual(repeated.unchanged, 1)
        self.assertEqual(len(repeated.completed), 0)
        self.assertEqual(self.adapter.calls, ["create"])

    def test_creates_and_updates_precede_deletes(self):
        self.seed(event("a", "Old"), event("z"))
        self.body = calendar(event("a"), event("b")).encode()
        result = self.runner.execute(self.profile, self.runner.preview(self.profile), approvals=self.approvals)
        self.assertEqual(result.status, RunStatus.COMPLETED)
        self.assertEqual(self.adapter.calls, ["create", "update", "delete"])
        self.assertEqual(result.counts, {"create": 1, "update": 1, "delete": 1, "unchanged": 0})

    def test_stale_source_target_or_profile_returns_new_preview_without_writes(self):
        preview = self.runner.preview(self.profile)
        self.body = calendar(event("a", "Changed after preview")).encode()
        result = self.runner.execute(self.profile, preview, approvals=self.approvals)
        self.assertEqual(result.status, RunStatus.REVIEW_REQUIRED)
        self.assertNotEqual(result.preview.digest, preview.digest)
        self.assertEqual(self.adapter.calls, [])
        preview = result.preview
        self.adapter.items.append(UnmanagedItem("manual"))
        result = self.runner.execute(self.profile, preview)
        self.assertEqual(result.status, RunStatus.REVIEW_REQUIRED)
        changed_profile = replace(self.profile, name="New name")
        result = self.runner.execute(changed_profile, result.preview)
        self.assertEqual(result.status, RunStatus.REVIEW_REQUIRED)
        self.assertEqual(self.adapter.calls, [])

    def test_collection_order_does_not_invalidate_preview(self):
        self.seed(event("missing-a"), event("missing-b"))
        preview = self.runner.preview(self.profile)
        self.adapter.items.reverse()
        result = self.runner.execute(self.profile, preview)
        self.assertEqual(result.status, RunStatus.PENDING)
        self.assertEqual(self.adapter.calls, ["create"])

    def test_publication_metadata_only_changes_do_not_trap_user_in_review_loop(self):
        preview = self.runner.preview(self.profile)
        self.body = self.body.replace(b"DTSTAMP:20260101", b"DTSTAMP:20260102")
        result = self.runner.execute(self.profile, preview)
        self.assertEqual(result.status, RunStatus.COMPLETED)
        self.assertEqual(result.counts["create"], 1)

    def test_pending_deletions_do_not_execute_without_approvals(self):
        self.seed(event("missing"))
        result = self.runner.execute(self.profile, self.runner.preview(self.profile))
        self.assertEqual(result.status, RunStatus.PENDING)
        self.assertEqual(self.adapter.calls, ["create"])
        self.assertEqual(result.counts["delete"], 0)

    def test_empty_snapshot_needs_its_own_per_call_approval(self):
        self.seed(event("a"))
        self.body = calendar().encode()
        preview = self.runner.preview(self.profile)
        result = self.runner.execute(self.profile, preview, approvals=RunApprovals(True, False, True))
        self.assertEqual(result.status, RunStatus.PENDING)
        self.assertEqual(self.adapter.calls, [])
        result = self.runner.execute(self.profile, preview, approvals=self.approvals)
        self.assertEqual(result.status, RunStatus.COMPLETED)
        self.assertEqual(result.counts["delete"], 1)
        self.assertEqual(self.adapter.calls, ["delete"])

    def test_any_conflict_blocks_all_mutations(self):
        self.seed(event("a"))
        self.adapter.items.append(replace(self.adapter.items[0], entry_id="duplicate"))
        result = self.runner.execute(self.profile, self.runner.preview(self.profile), approvals=self.approvals)
        self.assertEqual(result.status, RunStatus.BLOCKED)
        self.assertEqual(self.adapter.calls, [])

    def test_update_error_preserves_receipts_and_prevents_deletion(self):
        self.seed(event("a", "Old"), event("z"))
        self.body = calendar(event("a"), event("b")).encode()
        def fail():
            raise OutlookUpdateError(save_attempted=True, entry_id="new-1")
        self.adapter.after_update = fail
        with self.assertLogs("ics_importer.run", level="WARNING"):
            result = self.runner.execute(self.profile, self.runner.preview(self.profile), approvals=self.approvals)
        self.assertEqual(result.status, RunStatus.FAILED)
        self.assertTrue(result.failure.outcome_unknown)
        self.assertEqual(result.counts["create"], 1)
        self.assertEqual(result.counts["delete"], 0)
        self.assertNotIn("delete", self.adapter.calls)

    def test_failed_pre_delete_fetch_preserves_upserts_without_deleting(self):
        self.seed(event("z"))
        def break_fetch():
            self.fetcher.side_effect = OSError("private query token")
        self.adapter.after_create = break_fetch
        preview = self.runner.preview(self.profile)
        with self.assertLogs("ics_importer.run", level="WARNING"):
            result = self.runner.execute(self.profile, preview, approvals=self.approvals)
        self.assertEqual(result.status, RunStatus.FAILED)
        self.assertEqual(result.failure.phase, "pre_delete_capture")
        self.assertEqual(result.counts["create"], 1)
        self.assertNotIn("delete", self.adapter.calls)

    def test_first_delete_failure_stops_remaining_deletes(self):
        self.seed(event("y"), event("z"))
        def fail():
            from ics_outlook import OutlookDeleteError
            raise OutlookDeleteError(entry_id="uncertain")
        self.adapter.after_delete = fail
        preview = self.runner.preview(self.profile)
        with self.assertLogs("ics_importer.run", level="WARNING"):
            result = self.runner.execute(self.profile, preview, approvals=self.approvals)
        self.assertEqual(result.status, RunStatus.FAILED)
        self.assertEqual(self.adapter.calls.count("delete"), 1)
        self.assertEqual(result.failure.operation_kind, OperationKind.DELETE)
        self.assertTrue(result.failure.outcome_unknown)

    def test_uncertain_create_is_not_retried_and_restart_uses_markers(self):
        def fail():
            raise OutlookCreateError(save_attempted=True)
        self.adapter.after_create = fail
        with self.assertLogs("ics_importer.run", level="WARNING"):
            result = self.runner.execute(self.profile, self.runner.preview(self.profile))
        self.assertEqual(result.status, RunStatus.FAILED)
        self.assertTrue(result.failure.outcome_unknown)
        self.assertEqual(self.adapter.calls, ["create"])
        self.adapter.after_create = None
        result = self.runner.execute(self.profile, self.runner.preview(self.profile))
        self.assertEqual(result.status, RunStatus.COMPLETED)
        self.assertEqual(result.unchanged, 1)
        self.assertEqual(self.adapter.calls, ["create"])

    def test_interruption_after_persistence_before_receipt_needs_no_local_status(self):
        def interrupt():
            raise SystemExit("simulated process interruption")
        self.adapter.after_create = interrupt
        with self.assertRaises(SystemExit):
            self.runner.execute(self.profile, self.runner.preview(self.profile))
        self.adapter.after_create = None
        result = self.runner.execute(self.profile, self.runner.preview(self.profile))
        self.assertEqual(result.unchanged, 1)
        self.assertEqual(len(self.adapter.items), 1)
        self.assertEqual(self.adapter.calls, ["create"])

    def test_changed_source_before_delete_requires_new_review_after_upserts(self):
        self.seed(event("z"))
        def change_source():
            self.body = calendar(event("a", "New remote revision")).encode()
        self.adapter.after_create = change_source
        result = self.runner.execute(self.profile, self.runner.preview(self.profile), approvals=self.approvals)
        self.assertEqual(result.status, RunStatus.REVIEW_REQUIRED)
        self.assertEqual(result.counts["create"], 1)
        self.assertNotIn("delete", self.adapter.calls)

    def test_external_inventory_change_before_delete_prevents_deletion(self):
        self.seed(event("z"))
        self.adapter.after_create = lambda: self.adapter.items.append(UnmanagedItem("concurrent"))
        result = self.runner.execute(self.profile, self.runner.preview(self.profile), approvals=self.approvals)
        self.assertEqual(result.status, RunStatus.REVIEW_REQUIRED)
        self.assertNotIn("delete", self.adapter.calls)

    def test_cancellation_between_operations_keeps_partial_receipts(self):
        self.body = calendar(event("a"), event("b")).encode()
        preview = self.runner.preview(self.profile)
        result = self.runner.execute(self.profile, preview, cancel_requested=lambda: len(self.adapter.calls) >= 1)
        self.assertEqual(result.status, RunStatus.CANCELLED)
        self.assertEqual(result.counts["create"], 1)
        self.assertEqual(self.adapter.calls, ["create"])

    def test_errors_and_logs_do_not_echo_source_url(self):
        preview = self.runner.preview(self.profile)
        self.fetcher.side_effect = OSError(self.profile.source_url)
        with self.assertRaises(PreviewError) as caught:
            self.runner.preview(self.profile)
        self.assertNotIn(self.profile.source_url, str(caught.exception))
        with self.assertLogs("ics_importer.run", level="WARNING") as logs:
            result = self.runner.execute(self.profile, preview)
        self.assertEqual(result.status, RunStatus.FAILED)
        self.assertEqual(result.completed, ())
        self.assertNotIn(self.profile.source_url, str(logs.output))
        self.assertNotIn(self.profile.source_url, repr(result.failure))

    def test_busy_run_cannot_preview_or_execute(self):
        preview = self.runner.preview(self.profile)
        with target_lock(self.profile, directory=self.temp.name):
            with self.assertRaises(RunBusyError):
                self.runner.preview(self.profile)
            result = self.runner.execute(self.profile, preview)
        self.assertEqual(result.status, RunStatus.BUSY)
        self.assertEqual(self.adapter.calls, [])

    def test_embedded_profile_timezone_and_bom_work_through_coordinator(self):
        from test_timezones import with_zone, zone_component
        selected = replace(self.profile, floating_timezone="Synthetic/Local")
        self.body = b"\xef\xbb\xbf" + with_zone(event("a"), zone_component()).encode()
        result = self.runner.execute(selected, self.runner.preview(selected))
        self.assertEqual(result.status, RunStatus.COMPLETED)
        self.assertEqual(result.counts["create"], 1)

    def test_generator_parser_is_materialized_before_planning(self):
        self.runner.parser = lambda *args, **kwargs: iter(parse_ics(*args, **kwargs))
        result = self.runner.execute(self.profile, self.runner.preview(self.profile))
        self.assertEqual(result.status, RunStatus.COMPLETED)
        self.assertEqual(result.counts["create"], 1)

    def test_progress_reports_verified_counts_and_observer_cannot_abort_writes(self):
        self.seed(event("a", "Old"), event("z"))
        self.body = calendar(event("a"), event("b")).encode()
        updates = []
        result = self.runner.execute(self.profile, self.runner.preview(self.profile), approvals=self.approvals, progress=updates.append)
        self.assertEqual(result.status, RunStatus.COMPLETED)
        self.assertEqual([update.completed for update in updates if update.phase == "verified"], [1, 2, 3])
        self.assertEqual([(update.operation_kind, update.identity_key) for update in updates if update.phase == "verified"],
                         [(receipt.operation.kind, receipt.operation.identity.key) for receipt in result.completed])
        self.assertTrue(any(update.phase == "rechecking" and update.completed == 2 for update in updates))
        self.body = calendar(event("a", "Another update"), event("b")).encode()
        def broken(update):
            raise RuntimeError("private source URL")
        with self.assertLogs("ics_importer.run", level="WARNING") as logs:
            result = self.runner.execute(self.profile, self.runner.preview(self.profile), progress=broken)
        self.assertEqual(result.status, RunStatus.COMPLETED)
        self.assertNotIn("private source URL", str(logs.output))

"""GUI confirmations, worker execution, progress and partial results; fake backends."""

from dataclasses import replace
from threading import Event
import time
from unittest import TestCase

from ics_gui import confirm_preview, run_summary
from ics_model import OperationKind, SyncOperation
from ics_run import OperationReceipt, RunFailure, RunProgress, RunResult, RunStatus
from ics_sync import DeletionCandidate
from test_gui import GuiHarness, preview
from test_profiles import profile
from test_sync import target


class ConfirmationTests(TestCase):
    def test_declined_confirmation_never_grants_execution(self):
        self.assertIsNone(confirm_preview(preview(profile()), lambda *args: False))

    def test_coverage_blocker_cannot_be_overridden_by_confirmation(self):
        value = preview(profile())
        candidate = DeletionCandidate(target(), "Source identity absent", ("Source export coverage is unconfirmed", "Deletion preview requires explicit confirmation"))
        value = replace(value, plan=replace(value.plan, pending_deletions=(candidate,)))
        messages = []
        approvals = confirm_preview(value, lambda title, text: messages.append(text) or True)
        self.assertFalse(approvals.deletions)
        self.assertIn("Deletions: 0", messages[0])
        self.assertIn("remain pending: 1", messages[0])

    def test_empty_and_large_removal_confirmations_are_separate(self):
        value = preview(profile())
        candidate = DeletionCandidate(target(), "Source identity absent", ("Deletion preview requires explicit confirmation", "Empty snapshot requires separate confirmation", "Large removal requires separate confirmation"))
        value = replace(value, plan=replace(value.plan, pending_deletions=(candidate,)))
        prompts = []
        approvals = confirm_preview(value, lambda title, text: prompts.append(title) or True)
        self.assertEqual(len(prompts), 3)
        self.assertTrue(approvals.deletions and approvals.empty_snapshot and approvals.large_removal)
        answers = iter([True, False])
        self.assertIsNone(confirm_preview(value, lambda *args: next(answers)))

    def test_unknown_result_is_not_reported_as_rolled_back_or_success(self):
        result = RunResult(RunStatus.FAILED, failure=RunFailure("write", "SECRET", True))
        text = run_summary(result)
        self.assertIn("may already have been saved", text)
        self.assertNotIn("SECRET", text)


class GuiRunTests(GuiHarness, TestCase):
    def test_coverage_help_explains_missing_imports_without_granting_permission(self):
        self.app.complete_export.set(False)
        self.app.explain_coverage()
        title, message = self.dialogs.showinfo.call_args.args
        self.assertIn("missing imports", title)
        self.assertIn("every appointment", message)
        self.assertIn("New and changed appointments still synchronize", message)
        self.assertIn("still require confirmation", message)
        self.assertFalse(self.app.complete_export.get())
        self.coordinator.execute.assert_not_called()
        self.coordinator.preview.assert_not_called()

    def prepare(self):
        self.fill()
        self.app.save_profile()
        self.app.build_preview()
        self.complete_jobs()
        self.dialogs.askyesno.return_value = True

    def result(self, selected, reviewed, **kwargs):
        operation = reviewed.plan.operations[0]
        kwargs["progress"](RunProgress("verified", 1, 1, operation.kind))
        return RunResult(RunStatus.COMPLETED, reviewed, (OperationReceipt(operation, "saved"),))

    def test_run_shows_verified_result_and_requires_new_preview(self):
        self.prepare()
        self.coordinator.execute.side_effect = self.result
        self.app.synchronize()
        self.complete_jobs()
        self.assertEqual(self.app.last_result.status, RunStatus.COMPLETED)
        self.assertIn("1 new", self.app.report.get())
        row = self.app.tree.get_children()[0]
        self.assertEqual(self.app.tree.set(row, "action"), "New verified")
        self.assertIsNone(self.app.last_preview)
        self.assertEqual(str(self.app.execute_button.cget("state")), "disabled")

    def test_declining_dialog_performs_no_execution(self):
        self.prepare()
        self.dialogs.askyesno.return_value = False
        self.app.synchronize()
        self.assertFalse(self.app.jobs.busy)
        self.coordinator.execute.assert_not_called()

    def test_double_click_cannot_start_second_run(self):
        self.prepare()
        entered, release = Event(), Event()
        def slow(selected, reviewed, **kwargs):
            entered.set()
            if not release.wait(5):
                raise TimeoutError
            return self.result(selected, reviewed, **kwargs)
        self.coordinator.execute.side_effect = slow
        self.app.synchronize()
        self.assertTrue(entered.wait(3))
        try:
            self.app.synchronize()
            self.assertEqual(self.coordinator.execute.call_count, 1)
        finally:
            release.set()
        self.complete_jobs()

    def test_changed_snapshot_renders_new_preview_without_auto_resume(self):
        self.prepare()
        refreshed = self.app.last_preview
        changed = replace(refreshed.sources[0], fields=replace(refreshed.sources[0].fields, subject="Remote change"))
        refreshed = replace(refreshed, sources=(changed,), plan=replace(refreshed.plan, operations=(replace(refreshed.plan.operations[0], source=changed),)))
        self.coordinator.execute.side_effect = lambda *args, **kwargs: RunResult(RunStatus.REVIEW_REQUIRED, refreshed)
        self.app.synchronize()
        self.complete_jobs()
        self.assertEqual(self.coordinator.execute.call_count, 1)
        self.assertIs(self.app.last_preview, refreshed)
        self.assertIn("not carried forward", self.app.status.get())
        self.assertEqual(str(self.app.execute_button.cget("state")), "normal")

    def test_partial_failure_marks_verified_and_uncertain_rows(self):
        self.fill()
        self.app.save_profile()
        def two(selected):
            value = preview(selected)
            first = value.plan.operations[0]
            second_source = replace(first.source, identity=replace(first.identity, uid="second"))
            second = SyncOperation(OperationKind.CREATE, second_source.identity, source=second_source)
            return replace(value, plan=replace(value.plan, operations=(first, second)), sources=(first.source, second_source))
        self.coordinator.preview.side_effect = two
        self.app.build_preview()
        self.complete_jobs()
        self.dialogs.askyesno.return_value = True
        value = self.app.last_preview
        failure = RunFailure("write", "OutlookCreateError", True, value.sources[1].identity.key, OperationKind.CREATE)
        self.coordinator.execute.side_effect = lambda *args, **kwargs: RunResult(RunStatus.FAILED, value, (OperationReceipt(value.plan.operations[0], "saved"),), failure)
        self.app.synchronize()
        self.complete_jobs()
        actions = [self.app.tree.set(row, "action") for row in self.app.tree.get_children()]
        self.assertEqual(actions, ["New verified", "Uncertain"])
        self.assertIn("1 new", self.app.report.get())
        self.assertIn("may already have been saved", self.app.report.get())

    def test_cancel_and_close_wait_for_current_operation_then_report_counts(self):
        self.prepare()
        entered, release = Event(), Event()
        def blocked(selected, reviewed, **kwargs):
            entered.set()
            if not release.wait(5):
                raise TimeoutError
            self.assertTrue(kwargs["cancel_requested"]())
            return RunResult(RunStatus.CANCELLED, reviewed, (OperationReceipt(reviewed.plan.operations[0], "saved"),))
        self.coordinator.execute.side_effect = blocked
        self.app.synchronize()
        self.assertTrue(entered.wait(3))
        self.app.close()
        self.assertTrue(self.app._cancel.is_set())
        release.set()
        deadline = time.monotonic() + 5
        while self.app._alive and time.monotonic() < deadline:
            self.root.update()
            time.sleep(.01)
        self.assertFalse(self.app._alive)
        self.assertIn("1 new", self.dialogs.showinfo.call_args.args[1])

    def test_source_or_period_edit_resets_coverage_assertion(self):
        self.prepare()
        self.app.complete_export.set(True)
        self.assertIsNone(self.app.last_preview)
        self.app.values["end"].set("2027-02-01")
        self.assertFalse(self.app.complete_export.get())
        self.app.complete_export.set(True)
        self.app.values["url"].set(profile().source_url + "changed")
        self.assertFalse(self.app.complete_export.get())
        self.app.complete_export.set(True)
        self.app.values["zone"].set("UTC")
        self.assertFalse(self.app.complete_export.get())

    def test_unsaved_identity_cannot_execute_and_saving_keeps_valid_preview(self):
        self.fill()
        self.app.build_preview()
        self.complete_jobs()
        reviewed = self.app.last_preview
        self.assertEqual(str(self.app.execute_button.cget("state")), "disabled")
        self.app.synchronize()
        self.coordinator.execute.assert_not_called()
        self.assertIn("Save this profile", self.app.status.get())
        self.app.save_profile()
        self.assertIs(self.app.last_preview, reviewed)
        self.assertEqual(str(self.app.execute_button.cget("state")), "normal")

    def test_removed_persisted_profile_cannot_start_run(self):
        self.prepare()
        self.path.unlink()
        self.app.synchronize()
        self.coordinator.execute.assert_not_called()
        self.assertIn("Save this profile", self.app.status.get())

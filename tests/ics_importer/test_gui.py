"""Preview-only form, worker and withdrawn Tk tests; no network or real Outlook."""

from dataclasses import replace
from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory
from threading import Event, get_ident
import time
import tkinter as tk
from types import SimpleNamespace
from unittest import TestCase
from unittest.mock import Mock

from ics_gui import (FormError, ImporterWindow, MappingChangeRequired, ProfileForm,
                     ReadJobs, build_profile, new_form, preview_rows, read_error_message)
from ics_model import OperationKind, SyncOperation
from ics_outlook import CalendarChoice
from ics_profiles import ExportCoverage, load_profiles, save_profiles, update_profile
from ics_run import ImportPreview
from ics_sync import DeletionCandidate, SyncPlan, TargetInventory
from test_profiles import profile
from test_sync import source, target


def form(selected=None):
    selected = selected or profile()
    return ProfileForm(selected.profile_id, selected.name, selected.source_url, selected.destination,
                       selected.floating_timezone, selected.window_start.isoformat(), selected.window_end.isoformat())


def preview(selected):
    event = replace(source(), identity=replace(source().identity, profile_id=selected.profile_id))
    return ImportPreview(selected, "digest", TargetInventory(selected.destination, ()),
                         SyncPlan((SyncOperation(OperationKind.CREATE, event.identity, source=event, reason="Source present"),)), (event,))


class FormTests(TestCase):
    def test_current_year_defaults_and_stable_draft_uuid(self):
        draft = new_form(date(2026, 10, 26))
        self.assertEqual((draft.start_text, draft.end_text), ("2026-01-01", "2027-01-01"))
        complete = replace(draft, source_url=profile().source_url, destination=profile().destination)
        self.assertEqual(build_profile(complete).profile_id, build_profile(complete).profile_id)

    def test_validation_is_specific_and_does_not_echo_url_tokens(self):
        cases = [(replace(form(), name=" "), "profile name"),
                 (replace(form(), destination=None), "Choose"),
                 (replace(form(), timezone_name=""), "timezone"),
                 (replace(form(), start_text="20260101"), "YYYY-MM-DD"),
                 (replace(form(), end_text="2026-01-01"), "Until"),
                 (replace(form(), source_url="http://example.invalid/?key=SECRET"), "HTTPS")]
        for draft, text in cases:
            with self.subTest(draft=draft), self.assertRaises(FormError) as caught:
                build_profile(draft)
            self.assertIn(text, str(caught.exception))
            self.assertNotIn("SECRET", str(caught.exception))

    def test_mapping_confirmation_and_coverage_reset(self):
        selected = replace(profile(), source_coverage=ExportCoverage(date(2026, 1, 1), date(2027, 1, 1), True))
        changed = replace(form(selected), source_url=selected.source_url + "new")
        with self.assertRaises(MappingChangeRequired):
            build_profile(changed, base=selected)
        accepted = build_profile(changed, base=selected, accept_mapping_change=True)
        self.assertEqual(accepted.profile_id, selected.profile_id)
        self.assertIsNone(accepted.source_coverage)

    def test_display_rename_does_not_change_destination_mapping(self):
        renamed = replace(profile().destination, display_name="Renamed calendar")
        result = build_profile(replace(form(), destination=renamed), base=profile())
        self.assertEqual(result.destination, renamed)
        self.assertEqual(update_profile(profile(), destination=renamed).destination, renamed)

    def test_rows_show_changes_and_pending_reasons_without_marker_details(self):
        value = preview(profile())
        candidate = DeletionCandidate(target(), "Source identity absent", ("Source export coverage is unconfirmed",))
        value = replace(value, plan=replace(value.plan, pending_deletions=(candidate,)))
        rows, zone = preview_rows(value)
        self.assertEqual(rows[0][0], "New")
        self.assertEqual(rows[1][0], "Deletion candidate")
        self.assertIn("export period", rows[1][-1])
        self.assertEqual(zone, "Europe/Berlin")
        custom = replace(value, profile=replace(profile(), floating_timezone="Custom/Unknown"))
        self.assertIn("UTC", preview_rows(custom)[1])

    def test_errors_are_static_and_do_not_expose_raw_exception_text(self):
        for error_type in ("FeedDownloadError", "IcsParseError", "OutlookReadError", "RunBusyError", "SECRET"):
            self.assertNotIn("SECRET", read_error_message("preview", error_type))


class WorkerTests(TestCase):
    def test_busy_guard_and_plain_result_delivery(self):
        jobs = ReadJobs()
        entered, release = Event(), Event()
        main_thread = get_ident()
        def read():
            entered.set()
            if not release.wait(5):
                raise TimeoutError
            return get_ident()
        jobs.start("preview", 1, read)
        self.assertTrue(entered.wait(3))
        try:
            self.assertIsNone(jobs.poll())
            with self.assertRaises(RuntimeError):
                jobs.start("preview", 1, lambda: None)
        finally:
            release.set()
        deadline = time.monotonic() + 5
        result = None
        while result is None and time.monotonic() < deadline:
            result = jobs.poll()
            time.sleep(0.005)
        self.assertIsNotNone(result)
        self.assertNotEqual(result.value, main_thread)
        self.assertFalse(jobs.busy)


class GuiHarness:
    def setUp(self):
        self.directory = TemporaryDirectory(dir=Path(__file__).resolve().parent)
        self.addCleanup(self.directory.cleanup)
        self.path = Path(self.directory.name) / "profiles.json"
        try:
            self.root = tk.Tk()
        except tk.TclError as error:
            self.skipTest(f"Tk display unavailable: {type(error).__name__}")
        self.root.withdraw()
        self.addCleanup(self._cleanup)
        self.dialogs = SimpleNamespace(showerror=Mock(), showinfo=Mock(), askyesno=Mock(return_value=False))
        self.coordinator = SimpleNamespace(preview=Mock(side_effect=preview),
                                           execute=Mock(side_effect=AssertionError("No execution in step 16")))
        self.choice = CalendarChoice(profile().destination, "Synthetic account", False)
        self.app = ImporterWindow(self.root, profiles_path=self.path, coordinator=self.coordinator,
                                  picker=Mock(return_value=self.choice), inspector=Mock(return_value=self.choice),
                                  dialogs=self.dialogs)
        self.root.update()

    def _cleanup(self):
        if self.app._alive:
            deadline = time.monotonic() + 5
            while self.app.jobs.busy and time.monotonic() < deadline:
                self.root.update()
                time.sleep(0.01)
            self.app.close()

    def complete_jobs(self):
        deadline = time.monotonic() + 5
        while self.app.jobs.busy and time.monotonic() < deadline:
            self.root.update()
            time.sleep(0.01)
        self.root.update()
        self.assertFalse(self.app.jobs.busy)

    def fill(self):
        selected = profile()
        self.app.destination = selected.destination
        for key, value in (("name", selected.name), ("url", selected.source_url),
                           ("start", "2026-01-01"), ("end", "2027-01-01")):
            self.app.values[key].set(value)

class TkPreviewTests(GuiHarness, TestCase):
    def test_startup_never_reads_outlook_or_source_and_requires_preview_for_execution(self):
        self.coordinator.preview.assert_not_called()
        self.app.picker.assert_not_called()
        self.assertFalse(self.path.exists())
        self.assertEqual(str(self.app.execute_button.cget("state")), "disabled")
        self.coordinator.execute.assert_not_called()

    def test_invalid_input_does_not_schedule_preview(self):
        self.app.build_preview()
        self.assertFalse(self.app.jobs.busy)
        self.dialogs.showerror.assert_called_once()
        self.coordinator.preview.assert_not_called()

    def test_picker_cancel_preserves_destination(self):
        self.fill()
        self.app.picker.return_value = None
        self.app.choose_calendar()
        self.assertIn("Alt+Tab", self.app.picker_notice.get())
        self.complete_jobs()
        self.assertEqual(self.app.destination, profile().destination)
        self.assertIn("cancelled", self.app.status.get())

    def test_preview_populates_table_and_warns_for_additional_calendar(self):
        self.fill()
        self.app.build_preview()
        self.complete_jobs()
        self.assertIsNotNone(self.app.last_preview)
        self.assertEqual(len(self.app.tree.get_children()), 1)
        self.assertIn("not confirmed", self.app.calendar_hint.get())
        self.assertIn("1 new", self.app.summary.get())
        self.app.tree.selection_set(self.app.tree.get_children()[0])
        self.root.update()
        self.assertIn("Present in source", self.app.selection_note.get())
        self.coordinator.execute.assert_not_called()
        self.app.values["url"].set(profile().source_url + "changed")
        self.assertIsNone(self.app.last_preview)
        self.assertEqual(self.app.tree.get_children(), ())

    def test_save_and_load_profile_in_temporary_path(self):
        self.fill()
        self.app.save_profile()
        self.assertEqual(len(load_profiles(self.path)), 1)
        self.assertEqual(load_profiles(self.path)[0].profile_id, self.app._draft_id)
        self.coordinator.preview.assert_not_called()
        self.coordinator.execute.assert_not_called()

    def test_read_failure_is_visible_and_does_not_keep_old_preview(self):
        from ics_run import PreviewError
        self.fill()
        self.coordinator.preview.side_effect = PreviewError("FeedDownloadError")
        self.app.build_preview()
        self.complete_jobs()
        self.assertIsNone(self.app.last_preview)
        self.assertIn("Cannot download", self.dialogs.showerror.call_args.args[1])

    def test_worker_does_not_block_tk_and_late_result_after_input_change_is_discarded(self):
        self.fill()
        entered, release = Event(), Event()
        def slow(selected):
            entered.set()
            if not release.wait(5):
                raise TimeoutError
            return preview(selected)
        self.coordinator.preview.side_effect = slow
        self.app.build_preview()
        self.assertTrue(entered.wait(3))
        try:
            responsive = []
            self.root.after(0, lambda: responsive.append(True))
            self.root.update()
            self.assertEqual(responsive, [True])
            self.app.values["name"].set("Changed programmatically during read")
        finally:
            release.set()
        self.complete_jobs()
        self.assertIsNone(self.app.last_preview)
        self.assertIn("Inputs changed", self.app.status.get())

    def test_mapping_change_cancellation_prevents_read(self):
        self.fill()
        self.app._base = replace(profile(), profile_id=self.app._draft_id)
        self.app.values["url"].set(profile().source_url + "changed")
        self.app.build_preview()
        self.dialogs.askyesno.assert_called_once()
        self.assertFalse(self.app.jobs.busy)
        self.coordinator.preview.assert_not_called()

    def test_close_waits_for_read_without_saving_or_executing(self):
        self.fill()
        entered, release = Event(), Event()
        def slow(selected):
            entered.set()
            if not release.wait(5):
                raise TimeoutError
            return preview(selected)
        self.coordinator.preview.side_effect = slow
        self.app.build_preview()
        self.assertTrue(entered.wait(3))
        self.app.close()
        self.assertTrue(self.app._alive)
        self.assertTrue(self.app._closing)
        release.set()
        deadline = time.monotonic() + 5
        while self.app._alive and time.monotonic() < deadline:
            self.root.update()
            time.sleep(0.01)
        self.assertFalse(self.app._alive)
        self.assertFalse(self.path.exists())
        self.coordinator.execute.assert_not_called()

    def test_corrupt_profile_file_is_preserved_and_saving_disabled(self):
        self.app.close()
        self.path.write_text("{broken", encoding="utf-8")
        self.root = tk.Tk()
        self.root.withdraw()
        self.app = ImporterWindow(self.root, profiles_path=self.path, coordinator=self.coordinator,
                                  dialogs=self.dialogs)
        self.assertEqual(str(self.app.save_button.cget("state")), "disabled")
        self.fill()
        self.app.save_profile()
        self.assertEqual(self.path.read_text(), "{broken")

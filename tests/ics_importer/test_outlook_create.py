"""Single-create roundtrip and uncertain Save outcomes using fake COM only."""

from dataclasses import replace
from datetime import date, timedelta, timezone
from types import SimpleNamespace
from unittest import TestCase
from unittest.mock import Mock

from ics_model import EventIdentity, EventStatus
from ics_outlook import (MARKERS, MARKER_IDENTITY, MARKER_SCHEMA, OutlookCreateConflict,
                          OutlookCreateError, OutlookReadError, OutlookWriter, outlook_session)
from test_outlook_read import Items, folder
from test_profiles import profile
from test_sync import source


class WritableProperties:
    def __init__(self, values=None):
        self.values = {name: SimpleNamespace(Type=1, Value=value) for name, value in (values or {}).items()}
        self.add_calls = []

    def Find(self, name, custom):
        assert custom is True
        return self.values.get(name)

    def Add(self, name, kind, add_to_folder):
        self.add_calls.append((name, kind, add_to_folder))
        prop = SimpleNamespace(Type=kind, Value="")
        self.values[name] = prop
        return prop


class WritableItems(Items):
    def __init__(self):
        super().__init__([])
        self.add_count = 0
        self.fail_after_persist = False
        self.fail_before_persist = False
        self.post_save = None
        self.recipient_count = 0

    def Add(self, kind):
        assert kind == 1
        self.add_count += 1
        self.pending = SimpleNamespace(EntryID="", Class=26, IsRecurring=False,
                                       UserProperties=WritableProperties(),
                                       Recipients=SimpleNamespace(Count=self.recipient_count),
                                       Close=Mock(), Send=Mock(side_effect=AssertionError("No Send")))
        self.pending.Save = Mock(side_effect=self._save)
        return self.pending

    def _save(self):
        if self.fail_before_persist:
            raise OSError("simulated Save failure")
        pending = self.pending
        self.at_first_save = {name: prop.Value for name, prop in pending.UserProperties.values.items()}
        pending.EntryID = f"created-{self.add_count}"
        values = {key: value for key, value in vars(pending).items()
                  if key not in ("Save", "Close", "Send", "UserProperties")}
        self.fresh = SimpleNamespace(**values, UserProperties=WritableProperties(self.at_first_save),
                                     Parent=self.parent)
        self.records.append(self.fresh)
        if self.post_save:
            self.post_save(self.fresh)
        if self.fail_after_persist:
            raise OSError("Save returned private error after persistence")


class OutlookCreateTests(TestCase):
    def setUp(self):
        self.items = WritableItems()
        self.folder = folder(Items=self.items)
        self.items.parent = self.folder
        self.namespace = SimpleNamespace(GetFolderFromID=Mock(return_value=self.folder),
                                         GetItemFromID=Mock(side_effect=lambda entry, store: self.items.fresh))
        self.writer = OutlookWriter(self.namespace, category_separator=";")

    def create(self, event=None):
        return self.writer.create_appointment(profile().destination, event or source())

    def test_first_save_contains_every_marker_and_fixed_flag(self):
        result = self.create()
        self.assertEqual(set(self.items.at_first_save), set(MARKERS))
        self.assertEqual(self.items.at_first_save[MARKER_IDENTITY], source().identity.serialize())
        self.assertEqual(self.items.at_first_save[MARKER_SCHEMA], "1")
        self.assertTrue(all(kind == 1 and not register for _, kind, register in self.items.pending.UserProperties.add_calls))
        self.items.pending.Save.assert_called_once()
        self.items.pending.Send.assert_not_called()
        self.namespace.GetItemFromID.assert_called_once_with("created-1", "store")
        self.assertEqual((result.busy_status, result.reminder_set, result.meeting_status), (2, False, 0))
        self.assertEqual(result.fields, source().fields)
        self.assertEqual(self.items.pending.StartUTC.tzinfo, timezone.utc)
        self.assertEqual(self.items.pending.EndUTC.tzinfo, timezone.utc)

    def test_repeat_reads_existing_identity_without_add_or_save(self):
        first = self.create()
        second = self.create()
        self.assertEqual(second, first)
        self.assertEqual(self.items.add_count, 1)
        self.assertEqual(len(self.writer.read_inventory(profile().destination).items), 1)
        self.items.pending.Save.assert_called_once()

    def test_outlook_crlf_body_normalization_is_not_a_verification_error(self):
        event = replace(source(), fields=replace(source().fields, body="Line one\nLine two"))
        self.items.post_save = lambda fresh: setattr(fresh, "Body", fresh.Body.replace("\n", "\r\n") + "\r\n")
        result = self.create(event)
        self.assertIn("\r\n", result.fields.body)
        self.assertEqual(self.create(event).entry_id, result.entry_id)
        self.assertEqual(self.items.add_count, 1)

    def test_all_day_exclusive_end_and_long_unicode_identity_roundtrip(self):
        event = replace(source(), identity=EventIdentity(profile().profile_id, "Übung" * 100, date(2026, 10, 27)),
                        fields=replace(source().fields, start=date(2026, 10, 27), end=date(2026, 10, 30)))
        result = self.create(event)
        self.assertEqual(result.identity, event.identity)
        self.assertTrue(result.fields.all_day)
        self.assertEqual(result.fields.end, date(2026, 10, 30))
        self.assertEqual(self.items.pending.Start.tzinfo, timezone.utc)
        self.assertEqual(self.items.pending.End.tzinfo, timezone.utc)

    def test_changed_duplicate_or_unknown_existing_identity_never_creates(self):
        self.create()
        with self.assertRaises(OutlookCreateConflict):
            self.create(replace(source(), fields=replace(source().fields, subject="Changed")))
        self.items.fresh.UserProperties.values[MARKER_SCHEMA].Value = "2"
        with self.assertRaises(OutlookCreateConflict):
            self.create()
        self.items.fresh.UserProperties.values[MARKER_SCHEMA].Value = "1"
        duplicate = SimpleNamespace(**vars(self.items.fresh))
        duplicate.EntryID = "duplicate"
        self.items.records.append(duplicate)
        with self.assertRaises(OutlookCreateConflict):
            self.create()
        self.assertEqual(self.items.add_count, 1)

    def test_cancelled_unrepresentable_categories_and_subsecond_times_fail_before_add(self):
        for event in (replace(source(), status=EventStatus.CANCELLED),
                      replace(source(), fields=replace(source().fields, categories=("A;B",))),
                      replace(source(), fields=replace(source().fields, start=source().fields.start + timedelta(microseconds=1)))):
            with self.subTest(event=event), self.assertRaises(ValueError):
                self.create(event)
        self.assertEqual(self.items.add_count, 0)

    def test_save_error_after_persistence_does_not_retry_or_delete(self):
        self.items.fail_after_persist = True
        with self.assertRaises(OutlookCreateError) as caught:
            self.create()
        self.assertTrue(caught.exception.save_attempted)
        self.assertNotIn("private", str(caught.exception))
        self.assertEqual(len(self.items.records), 1)
        self.items.pending.Close.assert_not_called()
        # Recovery is a fresh scan, not another Add after an uncertain Save.
        self.items.fail_after_persist = False
        self.assertEqual(self.create().entry_id, "created-1")
        self.assertEqual(self.items.add_count, 1)

    def test_verification_failure_reports_known_entry_id_without_retry(self):
        for failure in ("read", "flags", "parent", "markers"):
            with self.subTest(failure=failure):
                self.setUp()
                if failure == "read":
                    self.namespace.GetItemFromID.side_effect = OSError("unavailable")
                elif failure == "flags":
                    self.items.post_save = lambda fresh: setattr(fresh, "BusyStatus", 1)
                elif failure == "parent":
                    self.items.post_save = lambda fresh: setattr(fresh, "Parent", folder(StoreID="other"))
                else:
                    self.items.post_save = lambda fresh: setattr(fresh, "UserProperties", WritableProperties())
                with self.assertRaises(OutlookCreateError) as caught:
                    self.create()
                self.assertTrue(caught.exception.save_attempted)
                self.assertEqual(caught.exception.entry_id, "created-1")
                self.assertEqual(self.items.add_count, 1)
                self.items.pending.Close.assert_not_called()

    def test_unexpected_recipients_discard_unsaved_object(self):
        self.items.recipient_count = 1
        with self.assertRaises(OutlookCreateError) as caught:
            self.create()
        self.assertFalse(caught.exception.save_attempted)
        self.items.pending.Save.assert_not_called()
        self.items.pending.Close.assert_called_once_with(1)
        self.assertEqual(self.items.records, [])

    def test_save_failure_before_persistence_is_still_not_retried(self):
        self.items.fail_before_persist = True
        with self.assertRaises(OutlookCreateError) as caught:
            self.create()
        self.assertTrue(caught.exception.save_attempted)
        self.assertEqual(self.items.add_count, 1)
        self.items.pending.Save.assert_called_once()
        self.items.pending.Close.assert_not_called()
        self.assertEqual(self.items.records, [])

    def test_write_session_requires_boolean_opt_in(self):
        with self.assertRaises(ValueError):
            with outlook_session(write="yes"):
                self.fail("Invalid opt-in must not attach to Outlook")

    def test_failed_inventory_prevents_add(self):
        self.namespace.GetFolderFromID.side_effect = OSError("read failed")
        with self.assertRaises(OutlookReadError):
            self.create()
        self.assertEqual(self.items.add_count, 0)

"""Optimistic single-item update contracts using independently reopened fakes."""

from dataclasses import replace
from datetime import date, timedelta
from types import SimpleNamespace
from unittest import TestCase
from unittest.mock import Mock

from ics_outlook import (MARKER_KEY, OutlookReadError, OutlookUpdateError,
                          OutlookWriteConflict, OutlookWriter)
from test_outlook_create import WritableItems, WritableProperties
from test_outlook_read import folder
from test_profiles import profile
from test_sync import source


class EditableStore:
    def __init__(self):
        self.items = WritableItems()
        self.folder = folder(Items=self.items)
        self.items.parent = self.folder
        self.namespace = SimpleNamespace(GetFolderFromID=Mock(return_value=self.folder),
                                         GetItemFromID=Mock(side_effect=self.get))
        self.writer = OutlookWriter(self.namespace, category_separator=";")
        self.fail_save_before = self.fail_save_after = False
        self.after_save = None

    def get(self, entry_id, store_id):
        matches = [item for item in self.items.records if item.EntryID == entry_id]
        if len(matches) != 1:
            raise OSError("not found")
        return matches[0]

    def seed(self, event=None):
        result = self.writer.create_appointment(profile().destination, event or source())
        self.bind(self.items.records[-1])
        self.namespace.GetItemFromID.reset_mock()
        return result

    def bind(self, appointment):
        appointment.Saved = True
        appointment.Close = Mock()
        appointment.Save = Mock(side_effect=lambda: self.save(appointment))
        appointment.Delete = Mock(side_effect=AssertionError("No Delete during update"))
        appointment.Send = Mock(side_effect=AssertionError("No Send during update"))

    def save(self, appointment):
        if self.fail_save_before:
            raise OSError("private save failure")
        fields = {key: value for key, value in vars(appointment).items()
                  if key not in ("Save", "Close", "Delete", "UserProperties")}
        props = {name: prop.Value for name, prop in appointment.UserProperties.values.items()}
        fresh = SimpleNamespace(**fields, UserProperties=WritableProperties(props))
        self.bind(fresh)
        index = next(i for i, item in enumerate(self.items.records) if item.EntryID == appointment.EntryID)
        self.items.records[index] = fresh
        if self.after_save:
            self.after_save(fresh)
        if self.fail_save_after:
            raise OSError("private error after persistence")


class OutlookUpdateTests(TestCase):
    def setUp(self):
        self.store = EditableStore()
        self.expected = self.store.seed()
        self.original = self.store.items.records[0]
        self.changed = replace(source(), fields=replace(source().fields, subject="Updated", location="New room",
                               start=source().fields.start + timedelta(days=1),
                               end=source().fields.end + timedelta(days=1)))

    def update(self, event=None, expected=None):
        return self.store.writer.update_appointment(profile().destination, event or self.changed,
                                                    expected or self.expected)

    def test_updates_fields_in_place_and_retains_identity(self):
        result = self.update()
        self.assertEqual(result.entry_id, self.expected.entry_id)
        self.assertEqual(result.identity.key, self.expected.identity.key)
        self.assertEqual(result.fields, self.changed.fields)
        self.assertEqual(self.store.items.add_count, 1)
        self.original.Save.assert_called_once()
        self.original.Delete.assert_not_called()
        self.original.Send.assert_not_called()
        self.assertEqual(self.update(self.changed, result), result)
        self.store.items.records[0].Save.assert_not_called()

    def test_unchanged_record_does_not_save(self):
        self.assertEqual(self.update(source()), self.expected)
        self.original.Save.assert_not_called()

    def test_all_day_transition_and_busy_reminder_repair(self):
        self.original.BusyStatus = 1
        self.original.ReminderSet = True
        expected = self.store.writer.read_inventory(profile().destination).items[0]
        all_day = replace(source(), fields=replace(source().fields, start=date(2026, 10, 27), end=date(2026, 10, 30)))
        result = self.update(all_day, expected)
        self.assertTrue(result.fields.all_day)
        self.assertEqual(result.fields.end, date(2026, 10, 30))
        self.assertEqual((result.busy_status, result.reminder_set), (2, False))
        result = self.update(source(), result)
        self.assertFalse(result.fields.all_day)

    def test_stale_expected_unsaved_edits_and_duplicates_prevent_mutation(self):
        self.original.Subject = "Manual edit"
        with self.assertRaises(OutlookWriteConflict):
            self.update()
        self.original.Save.assert_not_called()
        self.original.Subject = source().fields.subject
        self.original.Saved = False
        with self.assertRaises(OutlookWriteConflict):
            self.update()
        self.original.Saved = True
        duplicate = SimpleNamespace(**vars(self.original))
        duplicate.EntryID = "duplicate"
        self.store.items.records.append(duplicate)
        with self.assertRaises(OutlookWriteConflict):
            self.update()
        self.original.Save.assert_not_called()

    def test_ownership_and_parent_are_rechecked_after_inventory(self):
        def moved(entry_id, store_id):
            self.original.Parent = folder(StoreID="other")
            return self.original
        self.store.namespace.GetItemFromID.side_effect = moved
        with self.assertRaises(OutlookWriteConflict):
            self.update()
        self.original.Save.assert_not_called()
        self.original.Parent = self.store.folder
        def corrupt(entry_id, store_id):
            self.original.UserProperties.values[MARKER_KEY].Value = "changed"
            return self.original
        self.store.namespace.GetItemFromID.side_effect = corrupt
        with self.assertRaises(OutlookReadError):
            self.update()
        self.original.Save.assert_not_called()

    def test_unexpected_meeting_recipients_prevent_any_field_changes(self):
        self.original.Recipients = SimpleNamespace(Count=1)
        with self.assertRaises(OutlookWriteConflict):
            self.update()
        self.assertEqual(self.original.Subject, source().fields.subject)
        self.original.Save.assert_not_called()

    def test_save_failures_are_uncertain_and_never_create_replacements(self):
        for field in ("fail_save_before", "fail_save_after"):
            with self.subTest(field=field):
                store = EditableStore()
                expected = store.seed()
                setattr(store, field, True)
                with self.assertRaises(OutlookUpdateError) as caught:
                    store.writer.update_appointment(profile().destination, self.changed, expected)
                self.assertTrue(caught.exception.save_attempted)
                self.assertEqual(caught.exception.entry_id, expected.entry_id)
                self.assertNotIn("private", str(caught.exception))
                self.assertEqual(store.items.add_count, 1)

    def test_readback_mismatch_is_not_success(self):
        self.store.after_save = lambda fresh: setattr(fresh, "BusyStatus", 1)
        with self.assertRaises(OutlookUpdateError) as caught:
            self.update()
        self.assertTrue(caught.exception.save_attempted)
        self.assertEqual(self.store.items.add_count, 1)

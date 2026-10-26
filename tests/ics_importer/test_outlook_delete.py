"""Single DELETE authority, stale-state guards and uncertain outcomes, fake COM only."""

from dataclasses import replace
from types import SimpleNamespace
from unittest import TestCase
from unittest.mock import Mock, patch
from uuid import UUID

from ics_model import EventIdentity, EventStatus, OperationKind, SyncOperation
from ics_outlook import (MARKER_KEY, OutlookDeleteError, OutlookReadError, OutlookWriteConflict)
from test_outlook_read import folder
from test_outlook_update import EditableStore
from test_profiles import profile
from test_sync import source


class DeletableStore(EditableStore):
    def __init__(self):
        super().__init__()
        self.fail_delete_before = self.fail_delete_after = False
        self.after_delete = None

    def bind(self, appointment):
        super().bind(appointment)
        appointment.Delete = Mock(side_effect=lambda: self.delete(appointment))

    def delete(self, appointment):
        if self.fail_delete_before:
            raise OSError("private delete failure")
        self.items.records[:] = [item for item in self.items.records if item.EntryID != appointment.EntryID]
        if self.after_delete:
            self.after_delete(appointment)
        if self.fail_delete_after:
            raise OSError("private delete error after removal")


class OutlookDeleteTests(TestCase):
    def setUp(self):
        self.store = DeletableStore()
        self.expected = self.store.seed()
        self.original = self.store.items.records[0]
        self.operation = SyncOperation(OperationKind.DELETE, self.expected.identity,
                                       targets=(self.expected,), reason="Confirmed absence")

    def delete(self, operation=None, **kwargs):
        return self.store.writer.delete_appointment(profile().destination, operation or self.operation,
                                                     **(dict(confirmed=True) | kwargs))

    def test_delete_only_expected_item_and_verify_inventory(self):
        unmanaged = SimpleNamespace(EntryID="unmarked", UserProperties=SimpleNamespace(Find=lambda *args: None))
        self.store.items.records.append(unmanaged)
        result = self.delete()
        self.assertEqual(result.entry_id, self.expected.entry_id)
        self.assertEqual(result.identity, self.expected.identity)
        self.assertEqual([item.EntryID for item in self.store.items.records], ["unmarked"])
        self.original.Delete.assert_called_once()
        self.original.Save.assert_not_called()
        self.original.Send.assert_not_called()

    def test_confirmation_and_delete_shape_are_required(self):
        for confirmed in (False, None, "yes", 1):
            with self.subTest(confirmed=confirmed), self.assertRaises(ValueError):
                self.delete(confirmed=confirmed)
        invalid = [replace(self.operation, kind=OperationKind.UPDATE),
                   replace(self.operation, targets=()),
                   replace(self.operation, source=source())]
        for operation in invalid:
            with self.subTest(operation=operation), self.assertRaises(ValueError):
                self.delete(operation)
        self.original.Delete.assert_not_called()

    def test_stale_or_missing_target_never_deletes(self):
        self.original.Subject = "Manual change"
        with self.assertRaises(OutlookWriteConflict):
            self.delete()
        self.original.Delete.assert_not_called()
        self.store.items.records.clear()
        with self.assertRaises(OutlookWriteConflict):
            self.delete()
        self.original.Delete.assert_not_called()

    def test_explicit_cancellation_operation_is_supported(self):
        cancelled = replace(source(), status=EventStatus.CANCELLED, fields=None)
        result = self.delete(replace(self.operation, source=cancelled))
        self.assertEqual(result.identity, cancelled.identity)

    def test_wrong_destination_prevents_delete(self):
        with self.assertRaises(OutlookReadError):
            self.store.writer.delete_appointment(replace(profile().destination, store_id="different"),
                                                 self.operation, confirmed=True)
        self.original.Delete.assert_not_called()

    def test_parent_and_marker_rechecked_after_inventory(self):
        def wrong_parent(entry, store):
            self.original.Parent = folder(StoreID="other")
            return self.original
        self.store.namespace.GetItemFromID.side_effect = wrong_parent
        with self.assertRaises(OutlookWriteConflict):
            self.delete()
        self.original.Delete.assert_not_called()
        self.original.Parent = self.store.folder
        def wrong_marker(entry, store):
            self.original.UserProperties.values[MARKER_KEY].Value = "changed"
            return self.original
        self.store.namespace.GetItemFromID.side_effect = wrong_marker
        with self.assertRaises(OutlookReadError):
            self.delete()
        self.original.Delete.assert_not_called()

    def test_other_profile_schema_and_duplicate_targets_are_rejected(self):
        foreign = replace(self.expected, identity=EventIdentity(UUID(int=2), "event"))
        with self.assertRaises(OutlookWriteConflict):
            self.delete(SyncOperation(OperationKind.DELETE, foreign.identity, targets=(foreign,)))
        with self.assertRaises(ValueError):
            self.delete(replace(self.operation, targets=(replace(self.expected, marker_schema_version=2),)))
        duplicate = SimpleNamespace(**vars(self.original))
        duplicate.EntryID = "duplicate"
        self.store.items.records.append(duplicate)
        with self.assertRaises(OutlookWriteConflict):
            self.delete()
        self.original.Delete.assert_not_called()

    def test_meetings_recipients_and_unsaved_edits_are_not_deleted(self):
        for mode in ("meeting", "recipients", "unsaved"):
            with self.subTest(mode=mode):
                store = DeletableStore()
                expected = store.seed()
                item = store.items.records[0]
                if mode == "meeting":
                    item.MeetingStatus = 1
                    expected = store.writer.read_inventory(profile().destination).items[0]
                elif mode == "recipients":
                    item.Recipients = SimpleNamespace(Count=1)
                else:
                    item.Saved = False
                operation = SyncOperation(OperationKind.DELETE, expected.identity, targets=(expected,))
                with self.assertRaises(OutlookWriteConflict):
                    store.writer.delete_appointment(profile().destination, operation, confirmed=True)
                item.Delete.assert_not_called()

    def test_delete_failures_before_or_after_removal_never_retry(self):
        for mode in ("fail_delete_before", "fail_delete_after"):
            with self.subTest(mode=mode):
                store = DeletableStore()
                expected = store.seed()
                item = store.items.records[0]
                setattr(store, mode, True)
                operation = SyncOperation(OperationKind.DELETE, expected.identity, targets=(expected,))
                with self.assertRaises(OutlookDeleteError) as caught:
                    store.writer.delete_appointment(profile().destination, operation, confirmed=True)
                self.assertTrue(caught.exception.delete_attempted)
                self.assertEqual(caught.exception.entry_id, expected.entry_id)
                self.assertNotIn("private", str(caught.exception))
                item.Delete.assert_called_once()
                self.assertEqual(store.items.add_count, 1)

    def test_failed_post_delete_read_is_not_success(self):
        inventory = self.store.writer.read_inventory(profile().destination)
        with patch.object(self.store.writer, "read_inventory", side_effect=[inventory, OutlookReadError("unavailable")]):
            with self.assertRaises(OutlookDeleteError):
                self.delete()
        self.original.Delete.assert_called_once()
        self.assertEqual(self.store.items.records, [])

    def test_reappearing_identity_is_not_deleted_again(self):
        def reappear(original):
            replacement = SimpleNamespace(**vars(original))
            replacement.EntryID = "new-concurrent-copy"
            self.store.items.records.append(replacement)
        self.store.after_delete = reappear
        with self.assertRaises(OutlookDeleteError):
            self.delete()
        self.original.Delete.assert_called_once()
        self.assertEqual(len(self.store.items.records), 1)

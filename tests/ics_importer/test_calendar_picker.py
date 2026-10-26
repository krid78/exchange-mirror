"""Native picker metadata contract with fake Outlook only; no dialogs opened."""

from contextlib import nullcontext
from types import SimpleNamespace
from unittest import TestCase
from unittest.mock import Mock, patch

from ics_outlook import OutlookReadError, inspect_calendar, pick_calendar
from test_outlook_read import folder
from test_profiles import profile


class CalendarPickerTests(TestCase):
    def test_cancel_returns_none_without_choosing_a_fallback(self):
        namespace = SimpleNamespace(PickFolder=Mock(return_value=None))
        with patch("ics_outlook.outlook_session", return_value=nullcontext(SimpleNamespace(_namespace=namespace))):
            self.assertIsNone(pick_calendar())
        namespace.PickFolder.assert_called_once()

    def test_selected_account_default_and_additional_calendar(self):
        for default_id, expected in (("folder", True), ("different", False)):
            selected = folder()
            selected.Store = SimpleNamespace(DisplayName="Synthetic account", GetDefaultFolder=Mock(return_value=folder(EntryID=default_id)))
            namespace = SimpleNamespace(PickFolder=Mock(return_value=selected))
            with patch("ics_outlook.outlook_session", return_value=nullcontext(SimpleNamespace(_namespace=namespace))):
                choice = pick_calendar()
            self.assertEqual(choice.store_name, "Synthetic account")
            self.assertEqual(choice.is_default, expected)
            self.assertEqual(choice.destination.folder_id, "folder")

    def test_missing_default_folder_keeps_classification_unknown(self):
        selected = folder()
        selected.Store = SimpleNamespace(DisplayName="Shared store", GetDefaultFolder=Mock(side_effect=OSError()))
        reader = SimpleNamespace(resolve_calendar=Mock(return_value=selected))
        with patch("ics_outlook.outlook_session", return_value=nullcontext(reader)):
            self.assertIsNone(inspect_calendar(profile().destination).is_default)
        reader.resolve_calendar.assert_called_once_with(profile().destination)

    def test_non_calendar_or_missing_folder_fail_without_guessing(self):
        namespace = SimpleNamespace(PickFolder=Mock(return_value=folder(DefaultItemType=0)))
        with patch("ics_outlook.outlook_session", return_value=nullcontext(SimpleNamespace(_namespace=namespace))):
            with self.assertRaises(OutlookReadError):
                pick_calendar()
        reader = SimpleNamespace(resolve_calendar=Mock(side_effect=OutlookReadError("missing")))
        with patch("ics_outlook.outlook_session", return_value=nullcontext(reader)):
            with self.assertRaises(OutlookReadError):
                inspect_calendar(profile().destination)

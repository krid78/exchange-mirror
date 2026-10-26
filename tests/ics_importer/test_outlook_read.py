"""Read adapter contracts against fake COM objects; never attaches to Outlook."""

from contextlib import nullcontext
from dataclasses import replace
from datetime import date, datetime, timedelta, timezone
from types import ModuleType, SimpleNamespace
from threading import Thread
from unittest import TestCase
from unittest.mock import Mock, patch
from uuid import UUID

from ics_model import EventIdentity
from ics_outlook import (MARKER_IDENTITY, MARKER_KEY, MARKER_PROFILE, MARKER_SCHEMA,
                          OutlookReader, OutlookReadError, describe_calendar, outlook_session)
from ics_sync import UnmanagedItem, plan_upserts
from test_profiles import profile
from test_sync import source


def markers(identity=None, schema="1"):
    identity = identity or EventIdentity(UUID(int=1), "event")
    return {MARKER_PROFILE: str(identity.profile_id), MARKER_IDENTITY: identity.serialize(),
            MARKER_KEY: identity.key, MARKER_SCHEMA: schema}


class Properties:
    def __init__(self, values):
        self.values = values

    def Find(self, name, custom):
        assert custom is True
        value = self.values.get(name)
        return SimpleNamespace(Type=1, Value=value) if value is not None else None


def item(entry_id="one", values=None, **changes):
    fields = dict(EntryID=entry_id, UserProperties=Properties(markers() if values is None else values),
                  Class=26, IsRecurring=False, AllDayEvent=False,
                  StartUTC=datetime(2026, 10, 27, 13), EndUTC=datetime(2026, 10, 27, 14),
                  Start=datetime(2026, 10, 27), End=datetime(2026, 10, 28),
                  Subject="Public", Body="Details", Location="Room", Categories="Lecture;Other, category",
                  BusyStatus=2, ReminderSet=False, MeetingStatus=0,
                  Save=Mock(side_effect=AssertionError("No Save allowed")),
                  Delete=Mock(side_effect=AssertionError("No Delete allowed")))
    fields.update(changes)
    return SimpleNamespace(**fields)


class Items:
    def __init__(self, records, *, count=None, fail_next=False):
        self.records, self.count, self.fail_next = records, count, fail_next
        self.IncludeRecurrences = True

    @property
    def Count(self):
        return len(self.records) if self.count is None else self.count

    def GetFirst(self):
        self.index = 0
        return self.records[0] if self.records else None

    def GetNext(self):
        if self.fail_next:
            raise OSError("private item details")
        self.index += 1
        return self.records[self.index] if self.index < len(self.records) else None


def folder(records=(), **changes):
    values = dict(DefaultItemType=1, StoreID="store", EntryID="folder", FolderPath="Calendar",
                  Items=Items(list(records)))
    values.update(changes)
    return SimpleNamespace(**values)


def reader(selected):
    namespace = SimpleNamespace(GetFolderFromID=Mock(return_value=selected))
    return OutlookReader(namespace, category_separator=";")


class OutlookReadTests(TestCase):
    def test_complete_read_without_date_filter_or_recurrence_expansion(self):
        first, second = item(), item("two", values={})
        selected = folder([first, second])
        adapter = reader(selected)
        result = adapter.read_inventory(profile().destination)
        self.assertEqual(len(result.items), 2)
        self.assertEqual(result.items[0].identity, source().identity)
        self.assertEqual(result.items[0].fields.categories, ("Lecture", "Other, category"))
        self.assertEqual(result.items[1], UnmanagedItem("two"))
        self.assertFalse(selected.Items.IncludeRecurrences)
        adapter._namespace.GetFolderFromID.assert_called_once_with("folder", "store")
        first.Save.assert_not_called()
        first.Delete.assert_not_called()

    def test_unmarked_objects_do_not_read_private_content(self):
        unmarked = SimpleNamespace(EntryID="manual", UserProperties=Properties({}))
        result = reader(folder([unmarked])).read_inventory(profile().destination)
        self.assertEqual(result.items, (UnmanagedItem("manual"),))

    def test_date_and_utc_property_semantics(self):
        timed = item(StartUTC=datetime(2026, 10, 27, 13, tzinfo=timezone(timedelta(hours=2))))
        all_day = item("all-day", values=markers(EventIdentity(UUID(int=1), "all-day", date(2026, 10, 27))),
                       AllDayEvent=True, End=datetime(2026, 10, 30))
        result = reader(folder([timed, all_day])).read_inventory(profile().destination)
        self.assertEqual(result.items[0].fields.start, datetime(2026, 10, 27, 13, tzinfo=timezone.utc))
        self.assertEqual(result.items[1].fields.end, date(2026, 10, 30))
        self.assertTrue(result.items[1].fields.all_day)

    def test_marker_identity_types_roundtrip(self):
        for slot in (None, date(2026, 10, 27), datetime(2026, 10, 27, 13, tzinfo=timezone.utc)):
            identity = EventIdentity(UUID(int=1), 'Übung::["opaque"]', slot)
            with self.subTest(slot=slot):
                result = reader(folder([item(values=markers(identity))])).read_inventory(profile().destination)
                self.assertEqual(result.items[0].identity, identity)

    def test_incomplete_corrupt_and_inconsistent_markers_fail_whole_read(self):
        good = markers()
        cases = [{MARKER_KEY: good[MARKER_KEY]}, good | {MARKER_KEY: "wrong"},
                 good | {MARKER_PROFILE: str(UUID(int=2))}, good | {MARKER_IDENTITY: "broken"},
                 good | {MARKER_IDENTITY: good[MARKER_IDENTITY] + " "},
                 good | {MARKER_SCHEMA: "01"}, good | {MARKER_SCHEMA: "0"}]
        for values in cases:
            with self.subTest(values=values), self.assertRaises(OutlookReadError):
                reader(folder([item("valid"), item("bad", values=values)])).read_inventory(profile().destination)

    def test_marker_property_and_field_errors_are_not_missing_markers(self):
        broken_props = SimpleNamespace(Find=Mock(side_effect=OSError("private content")))
        for broken in (item(UserProperties=broken_props), item(StartUTC="invalid"),
                       item(Class=43), item(IsRecurring=True), item(EntryID="")):
            with self.subTest(broken=broken), self.assertRaises(OutlookReadError) as caught:
                reader(folder([broken])).read_inventory(profile().destination)
            self.assertNotIn("private content", str(caught.exception))
        props = SimpleNamespace(Find=Mock(return_value=SimpleNamespace(Type=3, Value=1)))
        with self.assertRaises(OutlookReadError):
            reader(folder([item(UserProperties=props)])).read_inventory(profile().destination)

    def test_duplicates_are_preserved_for_planner_conflict(self):
        result = reader(folder([item(), item("two")])).read_inventory(profile().destination)
        self.assertEqual(len(result.items), 2)
        self.assertTrue(plan_upserts(profile(), [source()], result).has_conflicts)
        result = reader(folder([item(values=markers(schema="2"))])).read_inventory(profile().destination)
        self.assertTrue(plan_upserts(profile(), [source()], result).has_conflicts)

    def test_collection_limits_count_changes_and_iteration_failures(self):
        for records in (Items([item()], count=2), Items([item()], fail_next=True),
                        Items([item(), item()])):
            with self.subTest(records=records), self.assertRaises(OutlookReadError):
                reader(folder(Items=records)).read_inventory(profile().destination)
        with self.assertRaises(OutlookReadError):
            reader(folder([item(), item("two")])).read_inventory(profile().destination, max_items=1)
        self.assertEqual(reader(folder()).read_inventory(profile().destination).items, ())

    def test_folder_identity_type_missing_and_closed_session(self):
        for selected in (folder(DefaultItemType=0), folder(StoreID="other"), folder(EntryID="other")):
            with self.subTest(selected=selected), self.assertRaises(OutlookReadError):
                reader(selected).read_inventory(profile().destination)
        missing = OutlookReader(SimpleNamespace(GetFolderFromID=Mock(side_effect=OSError("missing"))),
                                category_separator=";")
        with self.assertRaises(OutlookReadError):
            missing.resolve_calendar(profile().destination)
        adapter = reader(folder())
        adapter.close()
        with self.assertRaises(OutlookReadError):
            adapter.read_inventory(profile().destination)
        self.assertEqual(describe_calendar(folder()).folder_id, "folder")

    def test_session_initializes_and_releases_com_even_on_body_failure(self):
        com = ModuleType("pythoncom")
        com.CoInitialize, com.CoUninitialize = Mock(), Mock()
        win32 = ModuleType("win32com")
        win32.client = ModuleType("win32com.client")
        namespace = SimpleNamespace(GetFolderFromID=Mock(return_value=folder()))
        win32.client.GetActiveObject = Mock(return_value=SimpleNamespace(GetNamespace=Mock(return_value=namespace)))
        registry = ModuleType("winreg")
        registry.HKEY_CURRENT_USER = "HKCU"
        registry.OpenKey = Mock(return_value=nullcontext("key"))
        registry.QueryValueEx = Mock(return_value=(";", 1))
        modules = {"pythoncom": com, "win32com": win32, "win32com.client": win32.client, "winreg": registry}
        with patch.dict("sys.modules", modules):
            with self.assertRaisesRegex(RuntimeError, "body error"):
                with outlook_session() as adapter:
                    self.assertEqual(adapter.read_inventory(profile().destination).items, ())
                    with self.assertRaises(OutlookReadError):
                        with outlook_session():
                            self.fail("Nested sessions must be rejected before COM initialization")
                    raise RuntimeError("body error")
            self.assertIsNone(adapter._namespace)
            com.CoInitialize.assert_called_once()
            com.CoUninitialize.assert_called_once()

            win32.client.GetActiveObject.assert_called_once_with("Outlook.Application")
            com.CoInitialize.reset_mock()
            com.CoUninitialize.reset_mock()
            win32.client.GetActiveObject.side_effect = OSError("private error")
            with self.assertRaises(OutlookReadError):
                with outlook_session():
                    self.fail("Must not yield on connection failure")
            com.CoUninitialize.assert_called_once()
            com.CoUninitialize.reset_mock()
            win32.client.GetActiveObject.side_effect = SystemExit("interrupted during connection")
            with self.assertRaises(SystemExit):
                with outlook_session():
                    self.fail("Must not yield on interrupted connection")
            com.CoUninitialize.assert_called_once()
            win32.client.GetActiveObject.side_effect = None
            with outlook_session() as reopened:
                self.assertEqual(reopened.read_inventory(profile().destination).items, ())

    def test_reader_rejects_cross_thread_com_access(self):
        adapter = reader(folder())
        errors = []
        def worker():
            try:
                adapter.read_inventory(profile().destination)
            except OutlookReadError as error:
                errors.append(error)
        thread = Thread(target=worker)
        thread.start()
        thread.join()
        self.assertEqual(len(errors), 1)
        adapter._namespace.GetFolderFromID.assert_not_called()

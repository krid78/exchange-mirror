"""Read-only classic Outlook adapter. No mirror initialization or item writes."""

from contextlib import contextmanager
from datetime import date, datetime, timezone
import json
import threading
from uuid import UUID

from ics_model import EventFields, EventIdentity, TargetState
from ics_profiles import Destination
from ics_sync import TargetInventory, UnmanagedItem
from sync_policy import parse_categories


MARKER_PROFILE = "IcsImportProfileId"
MARKER_IDENTITY = "IcsImportIdentity"
MARKER_KEY = "IcsImportEventKey"
MARKER_SCHEMA = "IcsImportSchemaVersion"
MARKERS = (MARKER_PROFILE, MARKER_IDENTITY, MARKER_KEY, MARKER_SCHEMA)


class OutlookReadError(RuntimeError):
    """No usable complete inventory was obtained; no raw calendar data in errors."""


def _decode_identity(serialized):
    value = json.loads(serialized)
    if (not isinstance(value, list) or len(value) != 4 or value[0] != "ics-event-v1"
            or not isinstance(value[1], str) or not isinstance(value[2], str)
            or not isinstance(value[3], list)):
        raise ValueError("Invalid identity encoding")
    slot = value[3]
    if slot == ["single"]:
        recurrence = None
    elif len(slot) == 2 and isinstance(slot[1], str):
        if slot[0] == "date":
            recurrence = date.fromisoformat(slot[1])
        elif slot[0] == "date-time":
            recurrence = datetime.fromisoformat(slot[1])
        else:
            raise ValueError("Invalid identity slot tag")
    else:
        raise ValueError("Invalid identity slot")
    identity = EventIdentity(UUID(value[1]), value[2], recurrence)
    if identity.serialize() != serialized:
        raise ValueError("Identity must use the canonical v1 encoding")
    return identity


def _read_markers(item):
    values = {}
    props = item.UserProperties
    for name in MARKERS:
        prop = props.Find(name, True)
        if prop is not None:
            value = prop.Value
            if prop.Type != 1 or not isinstance(value, str):  # olText
                raise ValueError("Importer markers must be text properties")
            values[name] = value
    if not values:
        return None
    if len(values) != len(MARKERS):
        raise ValueError("Incomplete importer markers")
    identity = _decode_identity(values[MARKER_IDENTITY])
    if (values[MARKER_PROFILE] != str(identity.profile_id)
            or values[MARKER_KEY] != identity.key):
        raise ValueError("Inconsistent importer identity markers")
    schema = int(values[MARKER_SCHEMA])
    if schema <= 0 or str(schema) != values[MARKER_SCHEMA]:
        raise ValueError("Invalid importer marker schema")
    return identity, schema


def _utc_property(value):
    if not isinstance(value, datetime):
        raise ValueError("Expected an Outlook UTC date-time property")
    # COM DATE has no zone. StartUTC/EndUTC are UTC wall-clock values by contract;
    # do not reinterpret the timezone label attached by the Python COM wrapper.
    return datetime(value.year, value.month, value.day, value.hour, value.minute,
                    value.second, value.microsecond, tzinfo=timezone.utc)


def _read_item(item, separator):
    entry_id = item.EntryID
    if not isinstance(entry_id, str) or not entry_id:
        raise ValueError("Missing item EntryID")
    markers = _read_markers(item)
    if markers is None:
        return UnmanagedItem(entry_id)
    if item.Class != 26 or item.IsRecurring:
        raise ValueError("Importer-owned items must be nonrecurring appointments")
    identity, schema = markers
    if item.AllDayEvent:
        start, end = item.Start.date(), item.End.date()
    else:
        start, end = _utc_property(item.StartUTC), _utc_property(item.EndUTC)
    fields = EventFields(start, end, item.Subject or "", item.Body or "", item.Location or "",
                         tuple(parse_categories(item.Categories or "", separator)))
    return TargetState(identity, entry_id, fields, int(item.BusyStatus), bool(item.ReminderSet),
                       schema, int(item.MeetingStatus))


def describe_calendar(folder):
    """Read IDs/display path; no account-count requirement or name-based matching."""
    try:
        if folder.DefaultItemType != 1:  # olAppointmentItem
            raise OutlookReadError("Selected folder is not a calendar")
        return Destination(folder.StoreID, folder.EntryID, folder.FolderPath)
    except OutlookReadError:
        raise
    except Exception:
        raise OutlookReadError("Cannot describe the selected calendar") from None


class OutlookReader:
    """Caller-owned COM namespace; use only on its initialized owner thread."""

    def __init__(self, namespace, *, category_separator):
        if not isinstance(category_separator, str) or not category_separator.strip():
            raise ValueError("A Windows category separator is required")
        self._namespace = namespace
        self._separator = category_separator
        self._owner_thread = threading.get_ident()

    def _check_thread(self):
        if threading.get_ident() != self._owner_thread:
            raise OutlookReadError("Outlook reader must stay on its COM owner thread")

    def close(self):
        self._check_thread()
        self._namespace = None

    def resolve_calendar(self, destination):
        try:
            self._check_thread()
            if self._namespace is None:
                raise OutlookReadError("Outlook session is closed")
            folder = self._namespace.GetFolderFromID(destination.folder_id, destination.store_id)
            actual = describe_calendar(folder)
            if (actual.store_id.casefold(), actual.folder_id.casefold()) != (
                    destination.store_id.casefold(), destination.folder_id.casefold()):
                raise OutlookReadError("Resolved calendar identity differs from the destination")
            return folder
        except OutlookReadError:
            raise
        except Exception:
            raise OutlookReadError("Cannot resolve the selected Outlook calendar") from None

    def read_inventory(self, destination, *, max_items=100000):
        if isinstance(max_items, bool) or not isinstance(max_items, int) or max_items <= 0:
            raise ValueError("max_items must be a positive integer")
        try:
            folder = self.resolve_calendar(destination)
            items = folder.Items
            # Collection view only: never expand potentially infinite native series.
            items.IncludeRecurrences = False
            expected = items.Count
            if expected < 0 or expected > max_items:
                raise OutlookReadError("Calendar inventory exceeds its item limit")
            result = []
            item = items.GetFirst()
            while item is not None:
                if len(result) >= max_items:
                    raise OutlookReadError("Calendar inventory exceeds its item limit")
                result.append(_read_item(item, self._separator))
                item = items.GetNext()
            if len(result) != expected or items.Count != expected:
                raise OutlookReadError("Calendar changed during inventory; retry the read")
            # Duplicate source keys remain distinct records, enabling planner conflicts.
            return TargetInventory(destination, tuple(result))
        except OutlookReadError:
            raise
        except Exception:
            raise OutlookReadError("Cannot read a complete Outlook calendar inventory") from None


@contextmanager
def outlook_session():
    """Attach to running classic Outlook on this thread, release COM on exit."""
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        raise OutlookReadError("pywin32 is required for Outlook access") from None
    initialized = False
    reader = None
    application = namespace = None
    try:
        pythoncom.CoInitialize()
        initialized = True
        application = win32com.client.GetActiveObject("Outlook.Application")
        namespace = application.GetNamespace("MAPI")
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Control Panel\International") as key:
            separator, _ = winreg.QueryValueEx(key, "sList")
        reader = OutlookReader(namespace, category_separator=separator)
    except Exception:
        namespace = application = None
        if initialized:
            pythoncom.CoUninitialize()
        raise OutlookReadError("Cannot attach to running classic Outlook in this Windows session") from None
    try:
        yield reader
    finally:
        reader.close()
        reader = namespace = application = None
        pythoncom.CoUninitialize()

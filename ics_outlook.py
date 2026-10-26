"""Classic Outlook inventory and guarded single-item writes; no mirror setup."""

from contextlib import contextmanager
from dataclasses import dataclass
from datetime import date, datetime, time, timezone
import json
import threading
from uuid import UUID

from ics_model import EventFields, EventIdentity, EventStatus, OperationKind, SyncOperation, TargetState
from ics_profiles import Destination
from ics_sync import TargetInventory, UnmanagedItem, project_event, target_matches
from sync_policy import parse_categories


MARKER_PROFILE = "IcsImportProfileId"
MARKER_IDENTITY = "IcsImportIdentity"
MARKER_KEY = "IcsImportEventKey"
MARKER_SCHEMA = "IcsImportSchemaVersion"
MARKERS = (MARKER_PROFILE, MARKER_IDENTITY, MARKER_KEY, MARKER_SCHEMA)
_SESSION_STATE = threading.local()


class OutlookReadError(RuntimeError):
    """No usable complete inventory was obtained; no raw calendar data in errors."""


class OutlookCreateError(RuntimeError):
    """A create may have persisted: re-inventory before any retry, never assume rollback."""

    def __init__(self, *, save_attempted, entry_id=None):
        super().__init__("Appointment creation or verification failed; re-read inventory before retrying")
        self.save_attempted = save_attempted
        self.entry_id = entry_id


class OutlookCreateConflict(RuntimeError):
    """An existing identity prevents creation; no writes were attempted."""


class OutlookWriteConflict(RuntimeError):
    """Expected ownership or state changed; no mutation was attempted."""


class OutlookUpdateError(RuntimeError):
    def __init__(self, *, save_attempted, entry_id):
        super().__init__("Appointment update or verification failed; re-read inventory before retrying")
        self.save_attempted = save_attempted
        self.entry_id = entry_id


class OutlookDeleteError(RuntimeError):
    def __init__(self, *, entry_id):
        super().__init__("Appointment deletion or verification failed; re-read inventory before retrying")
        self.delete_attempted = True
        self.entry_id = entry_id


@dataclass(frozen=True)
class DeletedAppointment:
    entry_id: str
    identity: EventIdentity


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


class OutlookWriter(OutlookReader):
    """Explicit single-item writes; run-level locking remains a later step."""

    def _desired(self, source):
        desired = project_event(source)
        fields = desired.fields
        for value in (fields.subject, fields.body, fields.location, *fields.categories):
            if not isinstance(value, str):
                raise ValueError("Appointment content must be text")
        if any(not category.strip() or self._separator in category for category in fields.categories):
            raise ValueError("Source category cannot be represented with the Windows separator")
        if not fields.all_day and (fields.start.microsecond or fields.end.microsecond):
            raise ValueError("Outlook appointment times require whole-second precision")
        return desired

    def _apply_fields(self, appointment, desired):
        fields = desired.fields
        appointment.AllDayEvent = fields.all_day
        if fields.all_day:
            # UTC tags preserve date-only midnight wall values through COM marshaling.
            appointment.Start = datetime.combine(fields.start, time(), timezone.utc)
            appointment.End = datetime.combine(fields.end, time(), timezone.utc)
        else:
            appointment.StartUTC = fields.start
            appointment.EndUTC = fields.end
        appointment.Subject = fields.subject
        appointment.Body = fields.body
        appointment.Location = fields.location
        appointment.Categories = (self._separator + " ").join(fields.categories)
        appointment.MeetingStatus = desired.meeting_status
        appointment.BusyStatus = desired.busy_status
        appointment.ReminderSet = desired.reminder_set
        if appointment.Recipients.Count != 0:
            raise ValueError("Import unexpectedly has recipients")

    def _read_opened(self, appointment, destination, entry_id, identity):
        parent = describe_calendar(appointment.Parent)
        if (parent.store_id.casefold(), parent.folder_id.casefold()) != (
                destination.store_id.casefold(), destination.folder_id.casefold()):
            raise OutlookWriteConflict("Appointment is outside the selected destination")
        observed = _read_item(appointment, self._separator)
        if (not isinstance(observed, TargetState) or observed.identity != identity
                or observed.entry_id.casefold() != entry_id.casefold()
                or observed.marker_schema_version != 1):
            raise OutlookWriteConflict("Appointment identity or marker schema changed")
        return observed

    def _open_expected(self, destination, expected):
        if (not isinstance(expected, TargetState) or type(expected.marker_schema_version) is not int
                or expected.marker_schema_version != 1):
            raise ValueError("A known-schema target snapshot is required")
        inventory = self.read_inventory(destination)
        matches = [item for item in inventory.items
                   if isinstance(item, TargetState) and item.identity == expected.identity]
        if len(matches) != 1 or matches[0] != expected:
            raise OutlookWriteConflict("Target snapshot is stale or ambiguous; re-plan")
        try:
            appointment = self._namespace.GetItemFromID(expected.entry_id, destination.store_id)
            observed = self._read_opened(appointment, destination, expected.entry_id, expected.identity)
            if observed != expected or not appointment.Saved:
                raise OutlookWriteConflict("Target has changed or has unsaved edits; re-plan")
            return appointment, observed
        except OutlookWriteConflict:
            raise
        except Exception:
            raise OutlookReadError("Cannot reopen and verify the expected target") from None

    def create_appointment(self, destination, source):
        desired = self._desired(source)
        inventory = self.read_inventory(destination)
        matches = [item for item in inventory.items
                   if isinstance(item, TargetState) and item.identity == source.identity]
        if matches:
            if (len(matches) == 1 and matches[0].marker_schema_version == 1
                    and target_matches(desired, matches[0])):
                return matches[0]
            raise OutlookCreateConflict("Existing source identity requires conflict resolution or an update")
        if len(inventory.items) >= 100000:
            raise OutlookCreateConflict("Destination inventory limit would be exceeded")
        folder = self.resolve_calendar(destination)
        appointment = None
        save_attempted = False
        entry_id = None
        try:
            appointment = folder.Items.Add(1)  # olAppointmentItem, unsaved
            self._apply_fields(appointment, desired)
            values = {MARKER_PROFILE: str(source.identity.profile_id),
                      MARKER_IDENTITY: source.identity.serialize(), MARKER_KEY: source.identity.key,
                      MARKER_SCHEMA: "1"}
            props = appointment.UserProperties
            for name, value in values.items():
                prop = props.Find(name, True)
                if prop is None:
                    prop = props.Add(name, 1, False)  # olText, no folder field registration
                if prop.Type != 1:
                    raise ValueError("Unexpected new-item marker property type")
                prop.Value = value
            save_attempted = True
            appointment.Save()  # First Save includes content AND every ownership marker.
            entry_id = appointment.EntryID
            if not isinstance(entry_id, str) or not entry_id:
                raise ValueError("Saved appointment has no EntryID")
            fresh = self._namespace.GetItemFromID(entry_id, destination.store_id)
            observed = self._read_opened(fresh, destination, entry_id, source.identity)
            if not target_matches(desired, observed):
                raise ValueError("Saved appointment does not match the requested state")
            return observed
        except Exception:
            if appointment is not None and not save_attempted:
                try:
                    appointment.Close(1)  # olDiscard: discard the unsaved object only.
                except Exception:
                    pass
            raise OutlookCreateError(save_attempted=save_attempted, entry_id=entry_id) from None

    def update_appointment(self, destination, source, expected):
        desired = self._desired(source)
        if not isinstance(expected, TargetState) or source.identity != expected.identity:
            raise ValueError("Source and expected target identities must match")
        appointment, observed = self._open_expected(destination, expected)
        if target_matches(desired, observed):
            return observed
        try:
            recipient_count = appointment.Recipients.Count
        except Exception:
            raise OutlookReadError("Cannot verify the target's recipient state") from None
        if recipient_count != 0:
            raise OutlookWriteConflict("Cannot update an import with meeting recipients")
        save_attempted = False
        try:
            self._apply_fields(appointment, desired)
            save_attempted = True
            appointment.Save()
            fresh = self._namespace.GetItemFromID(expected.entry_id, destination.store_id)
            updated = self._read_opened(fresh, destination, expected.entry_id, source.identity)
            if not target_matches(desired, updated):
                raise ValueError("Saved update does not match the desired state")
            return updated
        except Exception:
            if not save_attempted:
                try:
                    appointment.Close(1)  # Discard only unsaved edits; never roll back a Save.
                except Exception:
                    pass
            raise OutlookUpdateError(save_attempted=save_attempted, entry_id=expected.entry_id) from None

    def delete_appointment(self, destination, operation, *, confirmed=False):
        """Execute one confirmed planned DELETE; never infer deletion authority here."""
        if confirmed is not True:
            raise ValueError("Deleting requires explicit confirmation of the planned operation")
        if (not isinstance(operation, SyncOperation) or operation.kind is not OperationKind.DELETE
                or len(operation.targets) != 1
                or (operation.source is not None and operation.source.status is not EventStatus.CANCELLED)):
            raise ValueError("Expected one planned DELETE with no active source")
        expected = operation.targets[0]
        appointment, observed = self._open_expected(destination, expected)
        try:
            recipient_count = appointment.Recipients.Count
        except Exception:
            raise OutlookReadError("Cannot verify the target's recipient state") from None
        if observed.meeting_status != 0 or recipient_count != 0:
            raise OutlookWriteConflict("Only ordinary recipient-free imports can be deleted")
        try:
            appointment.Delete()
            inventory = self.read_inventory(destination)
            if any(item.entry_id.casefold() == expected.entry_id.casefold()
                   or (isinstance(item, TargetState) and item.identity == expected.identity)
                   for item in inventory.items):
                raise ValueError("Deleted identity is still present or has reappeared")
            return DeletedAppointment(expected.entry_id, expected.identity)
        except Exception:
            # Delete may already have moved the item to Deleted Items. Never retry,
            # re-create it, or interpret a failed verification read as success.
            raise OutlookDeleteError(entry_id=expected.entry_id) from None


@contextmanager
def outlook_session(*, write=False):
    """Attach to running classic Outlook on this thread, release COM on exit."""
    if type(write) is not bool:
        raise ValueError("write must be a boolean")
    if getattr(_SESSION_STATE, "active", False):
        raise OutlookReadError("Close the current Outlook session before opening another on this thread")
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        raise OutlookReadError("pywin32 is required for Outlook access") from None
    initialized = False
    reader = None
    application = namespace = None
    _SESSION_STATE.active = True
    try:
        pythoncom.CoInitialize()
        initialized = True
        application = win32com.client.GetActiveObject("Outlook.Application")
        namespace = application.GetNamespace("MAPI")
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Control Panel\International") as key:
            separator, _ = winreg.QueryValueEx(key, "sList")
        reader_class = OutlookWriter if write else OutlookReader
        reader = reader_class(namespace, category_separator=separator)
    except BaseException as error:
        namespace = application = None
        try:
            if initialized:
                pythoncom.CoUninitialize()
        finally:
            _SESSION_STATE.active = False
        if not isinstance(error, Exception):
            raise
        raise OutlookReadError("Cannot attach to running classic Outlook in this Windows session") from None
    try:
        yield reader
    finally:
        try:
            reader.close()
        finally:
            reader = namespace = application = None
            try:
                pythoncom.CoUninitialize()
            finally:
                _SESSION_STATE.active = False

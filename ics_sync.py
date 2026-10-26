"""Pure synchronization plans. No COM, filesystem, network or mutation."""

from dataclasses import dataclass
from datetime import datetime
from zoneinfo import ZoneInfo

from ics_model import EventFields, EventStatus, OperationKind, SyncOperation, TargetState
from ics_profiles import Destination


@dataclass(frozen=True)
class UnmanagedItem:
    entry_id: str


@dataclass(frozen=True)
class TargetInventory:
    """Complete, successfully read contents of one concrete destination.

    Unmarked items are explicit but never synchronized. Adapter read/marker errors
    must fail inventory construction rather than become UnmanagedItem records.
    """

    destination: Destination
    items: tuple[TargetState | UnmanagedItem, ...]

    def __post_init__(self):
        object.__setattr__(self, "items", tuple(self.items))
        if any(not isinstance(item, (TargetState, UnmanagedItem)) for item in self.items):
            raise ValueError("Invalid target inventory record")
        ids = [item.entry_id for item in self.items]
        if any(not isinstance(value, str) or not value for value in ids) or len(ids) != len(set(ids)):
            raise ValueError("Target inventory requires unique, nonempty EntryIDs")


@dataclass(frozen=True)
class DesiredState:
    fields: EventFields
    busy_status: int = 2
    reminder_set: bool = False
    meeting_status: int = 0


@dataclass(frozen=True)
class SyncPlan:
    operations: tuple[SyncOperation, ...]
    notices: tuple[str, ...] = ()

    @property
    def has_conflicts(self):
        return any(op.kind is OperationKind.CONFLICT for op in self.operations)


def project_event(source):
    if source.status is EventStatus.CANCELLED or source.fields is None:
        raise ValueError("Cancelled events have no desired target")
    return DesiredState(source.fields)


def _matches(desired, target):
    def categories(fields):
        return {value.strip().casefold() for value in fields.categories if value.strip()}
    left, right = desired.fields, target.fields
    return (left.start == right.start and left.end == right.end and left.subject == right.subject
            and left.body.rstrip() == right.body.rstrip() and left.location == right.location
            and categories(left) == categories(right)
            and desired.busy_status == target.busy_status
            and desired.reminder_set == target.reminder_set
            and desired.meeting_status == target.meeting_status)


def _in_window(value, profile, zone):
    if value is None:
        return False
    day = value.astimezone(zone).date() if isinstance(value, datetime) else value
    return profile.window_start <= day < profile.window_end


def _context(profile, sources, inventory, window_timezone):
    if ((profile.destination.store_id, profile.destination.folder_id)
            != (inventory.destination.store_id, inventory.destination.folder_id)):
        raise ValueError("Target inventory does not match profile destination")
    sources = tuple(sources)
    if any(source.identity.profile_id != profile.profile_id for source in sources):
        raise ValueError("Source identity belongs to another profile")
    zone = window_timezone if window_timezone is not None else ZoneInfo(profile.floating_timezone)
    source_index, target_index = {}, {}
    for source in sources:
        source_index.setdefault(source.identity, []).append(source)
    for target in inventory.items:
        if isinstance(target, TargetState) and target.identity.profile_id == profile.profile_id:
            target_index.setdefault(target.identity, []).append(target)
    return source_index, target_index, zone


def plan_upserts(profile, sources, inventory, *, window_timezone=None):
    """Create/update/unchanged/conflict only; missing and cancelled events never delete.

    Sources must already be fully parsed and revision-resolved. Inventories are
    complete folder reads, including matches outside the current window.
    """
    source_index, target_index, zone = _context(profile, sources, inventory, window_timezone)
    operations = []
    for identity in sorted(source_index, key=lambda value: value.serialize()):
        versions = source_index[identity]
        targets = tuple(target_index.get(identity, ()))
        if len(versions) != 1 or len(targets) > 1 or any(t.marker_schema_version != 1 for t in targets):
            operations.append(SyncOperation(OperationKind.CONFLICT, identity, targets=targets,
                                            reason="Duplicate identity or unsupported target marker schema"))
            continue
        source = versions[0]
        if source.status is EventStatus.CANCELLED:
            continue
        if not (_in_window(source.fields.start, profile, zone)
                or _in_window(identity.recurrence_id, profile, zone)
                or any(_in_window(t.fields.start, profile, zone) for t in targets)):
            continue
        desired = project_event(source)
        kind = (OperationKind.CREATE if not targets else
                OperationKind.UNCHANGED if _matches(desired, targets[0]) else OperationKind.UPDATE)
        operations.append(SyncOperation(kind, identity, source, targets,
                                        reason="Source projection differs" if kind is OperationKind.UPDATE else "Source present"))
    return SyncPlan(tuple(operations))

"""Pure synchronization plans. No COM, filesystem, network or mutation."""

from dataclasses import dataclass
from datetime import datetime
import math
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
class DeletionCandidate:
    target: TargetState
    reason: str
    blockers: tuple[str, ...]


@dataclass(frozen=True)
class SyncPlan:
    operations: tuple[SyncOperation, ...]
    notices: tuple[str, ...] = ()
    pending_deletions: tuple[DeletionCandidate, ...] = ()

    @property
    def has_conflicts(self):
        return any(op.kind is OperationKind.CONFLICT for op in self.operations)

    @property
    def requires_confirmation(self):
        return bool(self.pending_deletions)


def project_event(source):
    if source.status is EventStatus.CANCELLED or source.fields is None:
        raise ValueError("Cancelled events have no desired target")
    return DesiredState(source.fields)


def target_matches(desired, target):
    def body(value):
        return value.replace("\r\n", "\n").replace("\r", "\n").rstrip()
    def categories(fields):
        return {value.strip().casefold() for value in fields.categories if value.strip()}
    left, right = desired.fields, target.fields
    return (left.start == right.start and left.end == right.end and left.subject == right.subject
            and body(left.body) == body(right.body) and left.location == right.location
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
        if len(versions) != 1 or len(targets) > 1 or any(not _known_schema(t) for t in targets):
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
                OperationKind.UNCHANGED if target_matches(desired, targets[0]) else OperationKind.UPDATE)
        operations.append(SyncOperation(kind, identity, source, targets,
                                        reason="Source projection differs" if kind is OperationKind.UPDATE else "Source present"))
    return SyncPlan(tuple(operations))


def _known_schema(target):
    return type(target.marker_schema_version) is int and target.marker_schema_version == 1


def plan_sync(profile, sources, inventory, *, snapshot_complete,
              confirm_deletions=False, confirm_empty_snapshot=False,
              confirm_large_removal=False, maximum_removal_fraction=0.5,
              window_timezone=None):
    """Plan scoped deletion only from a complete validated snapshot and explicit grants.

    This is still data, never execution. Pending candidates are informational;
    only confirmed/unblocked ones become DELETE operations. No grant is persisted.
    """
    if snapshot_complete is not True:
        raise ValueError("Planning requires a complete successfully parsed snapshot")
    if (isinstance(maximum_removal_fraction, bool)
            or not isinstance(maximum_removal_fraction, (float, int))
            or not math.isfinite(maximum_removal_fraction)
            or not 0 < maximum_removal_fraction <= 1):
        raise ValueError("Removal fraction must be in (0, 1]")
    sources = tuple(sources)
    base = plan_upserts(profile, sources, inventory, window_timezone=window_timezone)
    source_index, targets, zone = _context(profile, sources, inventory, window_timezone)
    operations = list(base.operations)
    conflicted = {op.identity for op in operations if op.kind is OperationKind.CONFLICT}
    wide_cancellations = {s.identity.uid: s for s in sources
                          if s.status is EventStatus.CANCELLED and s.identity.recurrence_id is None}
    for identity, items in targets.items():
        if identity not in conflicted and (len(items) != 1 or any(not _known_schema(t) for t in items)):
            operations.append(SyncOperation(OperationKind.CONFLICT, identity, targets=tuple(items),
                                            reason="Ambiguous target identity or unsupported marker schema"))
            conflicted.add(identity)
    for source in sources:
        if source.status is not EventStatus.CANCELLED and source.identity.uid in wide_cancellations:
            identity = wide_cancellations[source.identity.uid].identity
            if identity not in conflicted:
                operations.append(SyncOperation(OperationKind.CONFLICT, identity,
                                                reason="Active source contradicts UID cancellation"))
                conflicted.add(identity)

    def inside(value):
        return _in_window(value, profile, zone)

    candidates, scoped_count = [], 0
    coverage = profile.source_coverage
    for identity, items in targets.items():
        if identity in conflicted:
            continue
        target = items[0]
        original = identity.recurrence_id if identity.recurrence_id is not None else target.fields.start
        if not (inside(original) or inside(target.fields.start)):
            continue
        scoped_count += 1
        exact = source_index.get(identity, ())
        cancelled = (identity.uid in wide_cancellations
                     or any(s.status is EventStatus.CANCELLED for s in exact))
        if exact and not cancelled:
            continue
        blockers = []
        reason = "Explicit source cancellation" if cancelled else "Source identity absent"
        if not cancelled:
            # Both dates must be covered: partial recurrence expansion cannot prove
            # absence of an original slot outside the synchronization window.
            if not inside(original) or not inside(target.fields.start):
                blockers.append("Original slot and target start must both be in scope")
            if coverage is None or coverage.confirmed is not True:
                blockers.append("Source export coverage is unconfirmed")
            else:
                for value in (original, target.fields.start):
                    day = value.astimezone(zone).date() if isinstance(value, datetime) else value
                    if not coverage.start <= day < coverage.end:
                        blockers.append("Source absence is outside confirmed export coverage")
                        break
        candidates.append(DeletionCandidate(target, reason, tuple(blockers)))

    global_blockers = []
    if conflicted:
        global_blockers.append("Resolve all source/target conflicts before deleting")
    if not sources and confirm_empty_snapshot is not True:
        global_blockers.append("Empty snapshot requires separate confirmation")
    if scoped_count and len(candidates) / scoped_count > maximum_removal_fraction and confirm_large_removal is not True:
        global_blockers.append("Large removal requires separate confirmation")
    if confirm_deletions is not True:
        global_blockers.append("Deletion preview requires explicit confirmation")
    pending = []
    for candidate in candidates:
        blockers = candidate.blockers + tuple(global_blockers)
        if blockers:
            pending.append(DeletionCandidate(candidate.target, candidate.reason, blockers))
        else:
            identity = candidate.target.identity
            exact = source_index.get(identity, ())
            source = exact[0] if len(exact) == 1 else None
            operations.append(SyncOperation(OperationKind.DELETE, identity, source,
                                            (candidate.target,), candidate.reason))
    notices = tuple(dict.fromkeys(blocker for candidate in pending for blocker in candidate.blockers))
    return SyncPlan(tuple(operations), notices, tuple(pending))

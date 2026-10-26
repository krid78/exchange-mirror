"""Pure importer records and versioned ICS identities; no Outlook or I/O."""

from dataclasses import dataclass
from datetime import date, datetime, timezone
from enum import Enum
from hashlib import sha256
import json
from uuid import UUID


def _canonical_time(value: date | datetime) -> date | datetime:
    if isinstance(value, datetime):
        if value.utcoffset() is None:
            raise ValueError("Floating times must be resolved before entering the model")
        return value.astimezone(timezone.utc)
    if isinstance(value, date):
        return value
    raise TypeError("Expected a date or timezone-aware datetime")


@dataclass(frozen=True)
class EventIdentity:
    """Original recurrence slot, never the occurrence's rescheduled start."""

    profile_id: UUID
    uid: str
    recurrence_id: date | datetime | None = None

    def __post_init__(self):
        if not isinstance(self.profile_id, UUID):
            raise TypeError("profile_id must be a UUID")
        if not isinstance(self.uid, str) or not self.uid.strip():
            raise ValueError("uid must be nonempty text")
        # UID is opaque and case-sensitive: do not strip or normalize its contents.
        if self.recurrence_id is not None:
            object.__setattr__(self, "recurrence_id", _canonical_time(self.recurrence_id))

    def serialize(self) -> str:
        """Stable v1 encoding with explicit tags and no delimiter ambiguity."""
        slot = self.recurrence_id
        if isinstance(slot, datetime):
            recurrence = ["date-time", slot.isoformat(timespec="microseconds")]
        elif isinstance(slot, date):
            recurrence = ["date", slot.isoformat()]
        else:
            recurrence = ["single"]
        return json.dumps(
            ["ics-event-v1", str(self.profile_id), self.uid, recurrence],
            ensure_ascii=True, separators=(",", ":"),
        )

    @property
    def key(self) -> str:
        return sha256(self.serialize().encode("utf-8")).hexdigest()


@dataclass(frozen=True)
class EventFields:
    """Normalized content; all-day end dates are exclusive."""

    start: date | datetime
    end: date | datetime
    subject: str = ""
    body: str = ""
    location: str = ""
    categories: tuple[str, ...] = ()

    def __post_init__(self):
        start, end = _canonical_time(self.start), _canonical_time(self.end)
        if isinstance(start, datetime) != isinstance(end, datetime):
            raise ValueError("start and end must both be dates or both be datetimes")
        if end < start:
            raise ValueError("end must not precede start")
        if not isinstance(start, datetime) and end == start:
            raise ValueError("An all-day event requires an exclusive end after its start")
        object.__setattr__(self, "start", start)
        object.__setattr__(self, "end", end)
        object.__setattr__(self, "categories", tuple(self.categories))

    @property
    def all_day(self) -> bool:
        return not isinstance(self.start, datetime)


class EventStatus(Enum):
    CONFIRMED = "confirmed"
    TENTATIVE = "tentative"
    CANCELLED = "cancelled"


@dataclass(frozen=True)
class SourceEvent:
    identity: EventIdentity
    # Cancellation tombstones may have no dates/content.
    fields: EventFields | None
    status: EventStatus = EventStatus.CONFIRMED
    sequence: int = 0
    revision_time: datetime | None = None

    def __post_init__(self):
        if self.fields is None and self.status is not EventStatus.CANCELLED:
            raise ValueError("Active source events require content")
        if isinstance(self.sequence, bool) or not isinstance(self.sequence, int) or self.sequence < 0:
            raise ValueError("sequence must be a nonnegative integer")
        if self.revision_time is not None:
            if not isinstance(self.revision_time, datetime):
                raise TypeError("revision_time must be a datetime")
            object.__setattr__(self, "revision_time", _canonical_time(self.revision_time))


@dataclass(frozen=True)
class TargetState:
    """Observed target, including flags which later planning may need to repair.

    The enclosing adapter identifies the destination folder/store. EntryID is a
    locator within that destination, not part of the stable event identity.
    """

    identity: EventIdentity
    entry_id: str
    fields: EventFields
    busy_status: int
    reminder_set: bool
    marker_schema_version: int = 1


class OperationKind(Enum):
    CREATE = "create"
    UPDATE = "update"
    DELETE = "delete"
    UNCHANGED = "unchanged"
    CONFLICT = "conflict"


@dataclass(frozen=True)
class SyncOperation:
    """Plan record only; constructing it never changes a calendar.

    Source/target may be absent (creation, removal, or conflict). Multiple targets
    preserve duplicate-identity evidence rather than choosing an arbitrary item.
    Action-specific validation belongs to the planner and executor.
    """

    kind: OperationKind
    identity: EventIdentity
    source: SourceEvent | None = None
    targets: tuple[TargetState, ...] = ()
    reason: str = ""

    def __post_init__(self):
        object.__setattr__(self, "targets", tuple(self.targets))
        if self.source is not None and self.source.identity != self.identity:
            raise ValueError("Source identity must match the operation")
        if any(target.identity != self.identity for target in self.targets):
            raise ValueError("Target identities must match the operation")

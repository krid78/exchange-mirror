"""Manual preview/execute orchestration with lock, freshness checks and no retry."""

from dataclasses import dataclass, replace
from enum import Enum
from hashlib import sha256
import logging

from ics_http import fetch_ics
from ics_lock import RunBusyError, target_lock
from ics_model import OperationKind, SourceEvent, SyncOperation
from ics_outlook import OutlookCreateConflict, OutlookReadError, OutlookWriteConflict, outlook_session
from ics_profiles import ImportProfile
from ics_source import calendar_timezone, parse_ics
from ics_sync import SyncPlan, TargetInventory, plan_sync


log = logging.getLogger("ics_importer.run")


class PreviewError(RuntimeError):
    def __init__(self, error_type):
        super().__init__("Cannot build a complete importer preview")
        self.error_type = error_type


@dataclass(frozen=True)
class RunApprovals:
    deletions: bool = False
    empty_snapshot: bool = False
    large_removal: bool = False

    def __post_init__(self):
        if any(type(value) is not bool for value in (self.deletions, self.empty_snapshot, self.large_removal)):
            raise ValueError("Run approvals must be explicit booleans")


@dataclass(frozen=True)
class ImportPreview:
    profile: ImportProfile
    digest: str
    inventory: TargetInventory
    plan: SyncPlan
    sources: tuple[SourceEvent, ...]


class RunStatus(Enum):
    COMPLETED = "completed"
    PENDING = "completed_with_pending_deletions"
    REVIEW_REQUIRED = "review_required"
    FAILED = "failed"
    CANCELLED = "cancelled"
    BLOCKED = "blocked_by_conflicts"
    BUSY = "busy"


@dataclass(frozen=True)
class OperationReceipt:
    operation: SyncOperation
    entry_id: str


@dataclass(frozen=True)
class RunProgress:
    phase: str
    completed: int
    total: int
    operation_kind: OperationKind | None = None


@dataclass(frozen=True)
class RunFailure:
    phase: str
    error_type: str
    outcome_unknown: bool = False
    identity_key: str | None = None
    operation_kind: OperationKind | None = None
    diagnostic: str | None = None


@dataclass(frozen=True)
class RunResult:
    status: RunStatus
    preview: ImportPreview | None = None
    completed: tuple[OperationReceipt, ...] = ()
    failure: RunFailure | None = None
    unchanged: int = 0

    @property
    def counts(self):
        return {kind.value: sum(receipt.operation.kind is kind for receipt in self.completed)
                for kind in (OperationKind.CREATE, OperationKind.UPDATE, OperationKind.DELETE)} | {"unchanged": self.unchanged}


def _inventory_key(inventory):
    return tuple(sorted(inventory.items, key=lambda item: item.entry_id))


def _source_key(sources):
    # Compare resolved effects, not generation timestamps/SEQUENCE alone. Some
    # publishers refresh DTSTAMP on every download without changing any event.
    return tuple(sorted((replace(source, sequence=0, revision_time=None) for source in sources),
                        key=lambda source: source.identity.serialize()))


def _operation_key(operations):
    return tuple(sorted((replace(op, source=replace(op.source, sequence=0, revision_time=None)
                                 if op.source is not None else None) for op in operations),
                        key=lambda op: (op.kind.value, op.identity.serialize())))


def _plan_key(plan):
    return (_operation_key(plan.operations),
            tuple(sorted(plan.pending_deletions, key=lambda candidate: candidate.target.entry_id)),
            tuple(sorted(plan.notices)))


class ImportCoordinator:
    """No retained COM objects. Session/transport seams permit offline harnesses."""

    def __init__(self, *, session_factory=outlook_session, fetcher=fetch_ics,
                 parser=parse_ics, timezone_resolver=calendar_timezone,
                 lock_directory=None, maximum_removal_fraction=0.5):
        self.session_factory = session_factory
        self.fetcher = fetcher
        self.parser = parser
        self.timezone_resolver = timezone_resolver
        self.lock_directory = lock_directory
        self.maximum_removal_fraction = maximum_removal_fraction

    def _capture(self, profile, adapter, approvals):
        snapshot = self.fetcher(profile.source_url)
        # Recompute rather than trusting a transport's supplied digest.
        digest = sha256(snapshot.data).hexdigest()
        sources = tuple(self.parser(snapshot.data, profile.profile_id,
                              floating_timezone=profile.floating_timezone,
                              window_start=profile.window_start, window_end=profile.window_end))
        zone = self.timezone_resolver(snapshot.data, profile.floating_timezone)
        inventory = adapter.read_inventory(profile.destination)
        plan = self._plan(profile, sources, inventory, approvals, zone)
        return ImportPreview(profile, digest, inventory, plan, sources), sources, zone

    def _plan(self, profile, sources, inventory, approvals, zone):
        return plan_sync(profile, sources, inventory, snapshot_complete=True,
                         confirm_deletions=approvals.deletions,
                         confirm_empty_snapshot=approvals.empty_snapshot,
                         confirm_large_removal=approvals.large_removal,
                         maximum_removal_fraction=self.maximum_removal_fraction,
                         window_timezone=zone)

    def preview(self, profile):
        """Dry-run: fetch/parse/read/plan only; acquires no write-enabled session."""
        try:
            with target_lock(profile, directory=self.lock_directory):
                with self.session_factory(write=False) as adapter:
                    preview, _, _ = self._capture(profile, adapter, RunApprovals())
                    return preview
        except RunBusyError:
            raise
        except Exception as error:
            raise PreviewError(type(error).__name__) from None

    def execute(self, profile, preview, *, approvals=None, cancel_requested=None, progress=None):
        """Accept only the exact preview; stale state returns a fresh preview and no new writes.

        A review-required result after some verified upserts preserves their receipts.
        All approvals are per-call. No write exception triggers an automatic retry.
        """
        if not isinstance(preview, ImportPreview) or not isinstance(profile, ImportProfile):
            raise TypeError("A profile and its preview are required")
        approvals = approvals if approvals is not None else RunApprovals()
        if not isinstance(approvals, RunApprovals):
            raise TypeError("Expected explicit run approvals")
        cancel = cancel_requested if cancel_requested is not None else lambda: False
        receipts = []
        current = None
        phase = "lock"
        active = None
        unchanged = 0
        total = 0
        def notify(stage):
            if progress is not None:
                try:
                    progress(RunProgress(stage, len(receipts), total, active.kind if active else None))
                except Exception:
                    log.warning("Progress observer failed; synchronization state is unaffected")
        try:
            with target_lock(profile, directory=self.lock_directory):
                if cancel():
                    return RunResult(RunStatus.CANCELLED)
                phase = "session"
                with self.session_factory(write=True) as adapter:
                    phase = "capture"
                    notify("reading")
                    current, sources, zone = self._capture(profile, adapter, RunApprovals())
                    if cancel():
                        return RunResult(RunStatus.CANCELLED, current)
                    if (profile != preview.profile or _source_key(current.sources) != _source_key(preview.sources)
                            or _inventory_key(current.inventory) != _inventory_key(preview.inventory)
                            or _plan_key(current.plan) != _plan_key(preview.plan)):
                        return RunResult(RunStatus.REVIEW_REQUIRED, current)
                    plan = self._plan(profile, sources, current.inventory, approvals, zone)
                    if plan.has_conflicts:
                        return RunResult(RunStatus.BLOCKED, current)
                    current = replace(current, plan=plan)
                    unchanged = sum(op.kind is OperationKind.UNCHANGED for op in plan.operations)
                    total = sum(op.kind in (OperationKind.CREATE, OperationKind.UPDATE, OperationKind.DELETE) for op in plan.operations)
                    notify("planned")
                    expected_items = {item.entry_id: item for item in current.inventory.items}
                    for kind in (OperationKind.CREATE, OperationKind.UPDATE):
                        for operation in plan.operations:
                            if operation.kind is not kind:
                                continue
                            if cancel():
                                return RunResult(RunStatus.CANCELLED, current, tuple(receipts), unchanged=unchanged)
                            phase, active = "write", operation
                            notify(kind.value)
                            if kind is OperationKind.CREATE:
                                state = adapter.create_appointment(profile.destination, operation.source)
                            else:
                                state = adapter.update_appointment(profile.destination, operation.source, operation.targets[0])
                            expected_items[state.entry_id] = state
                            receipts.append(OperationReceipt(operation, state.entry_id))
                            notify("verified")
                            log.info("Verified %s %s", kind.value, operation.identity.key)
                            active = None
                    deletions = tuple(op for op in plan.operations if op.kind is OperationKind.DELETE)
                    if deletions:
                        if cancel():
                            return RunResult(RunStatus.CANCELLED, current, tuple(receipts), unchanged=unchanged)
                        phase = "pre_delete_capture"
                        notify("rechecking")
                        fresh, fresh_sources, fresh_zone = self._capture(profile, adapter, RunApprovals())
                        if cancel():
                            return RunResult(RunStatus.CANCELLED, fresh, tuple(receipts), unchanged=unchanged)
                        expected_inventory = TargetInventory(profile.destination, tuple(expected_items.values()))
                        if (_source_key(fresh.sources) != _source_key(current.sources)
                                or _inventory_key(fresh.inventory) != _inventory_key(expected_inventory)):
                            return RunResult(RunStatus.REVIEW_REQUIRED, fresh, tuple(receipts), unchanged=unchanged)
                        fresh_plan = self._plan(profile, fresh_sources, fresh.inventory, approvals, fresh_zone)
                        fresh_deletes = tuple(op for op in fresh_plan.operations if op.kind is OperationKind.DELETE)
                        if (fresh_plan.has_conflicts or _operation_key(fresh_deletes) != _operation_key(deletions)
                                or any(op.kind in (OperationKind.CREATE, OperationKind.UPDATE) for op in fresh_plan.operations)):
                            return RunResult(RunStatus.REVIEW_REQUIRED, fresh, tuple(receipts), unchanged=unchanged)
                        for operation in fresh_deletes:
                            if cancel():
                                return RunResult(RunStatus.CANCELLED, fresh, tuple(receipts), unchanged=unchanged)
                            phase, active = "write", operation
                            notify("delete")
                            state = adapter.delete_appointment(profile.destination, operation, confirmed=True)
                            receipts.append(OperationReceipt(operation, state.entry_id))
                            notify("verified")
                            log.info("Verified delete %s", operation.identity.key)
                            active = None
                    status = RunStatus.PENDING if plan.pending_deletions else RunStatus.COMPLETED
                    return RunResult(status, current, tuple(receipts), unchanged=unchanged)
        except RunBusyError:
            return RunResult(RunStatus.BUSY)
        except Exception as error:
            unknown = active is not None and not isinstance(error, (OutlookReadError, OutlookWriteConflict,
                                                                     OutlookCreateConflict, ValueError))
            failure = RunFailure(phase, type(error).__name__, unknown,
                                 active.identity.key if active is not None else None,
                                 active.kind if active is not None else None,
                                 error.diagnostic if isinstance(error, OutlookReadError) else None)
            log.warning("Importer stopped in %s (%s%s)", phase, type(error).__name__,
                        ": " + failure.diagnostic if failure.diagnostic else "")
            return RunResult(RunStatus.FAILED, current, tuple(receipts), failure, unchanged)

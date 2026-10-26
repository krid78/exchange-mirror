"""Opt-in destructive acceptance test, restricted to the approved empty fixture.

Run with --run; --cleanup recovers only this run's marked synthetic appointments.
No network, messages, account synchronization or production-calendar writes.
"""
import argparse
from contextlib import contextmanager
from datetime import date, datetime, timezone
import json
import logging
import os
from pathlib import Path
import subprocess
import sys
from types import SimpleNamespace
from uuid import UUID, uuid4

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
from ics_http import FeedSnapshot
from ics_model import TargetState
from ics_outlook import describe_calendar, outlook_session, OutlookReadError
from ics_profiles import Destination, ExportCoverage, ImportProfile
from ics_run import ImportCoordinator, RunApprovals, RunStatus

CALENDAR = "Exchange Mirror Importer Test"


def fixture(adapter):
    matches = [folder for folder in adapter._namespace.GetDefaultFolder(9).Folders
               if folder.Name == CALENDAR]
    if len(matches) != 1:
        raise RuntimeError("Exactly one approved test calendar must already exist")
    return describe_calendar(matches[0])


def load_profile(path):
    state = json.loads(path.read_text(encoding="utf-8"))
    destination = Destination(**state["destination"])
    with outlook_session() as adapter:
        if fixture(adapter) != destination:
            raise RuntimeError("Test destination changed")
    return ImportProfile(UUID(state["profile_id"]), "Synthetic acceptance",
                         "https://example.invalid/acceptance", destination, "Europe/Berlin",
                         date(2030, 1, 1), date(2030, 2, 1),
                         ExportCoverage(date(2030, 1, 1), date(2030, 2, 1), True))


def feed(events):
    lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Acceptance//EN"]
    for uid, title, status in events:
        lines += ["BEGIN:VEVENT", f"UID:acceptance-{uid}", "DTSTAMP:20300101T000000Z",
                  "DTSTART:20300115T100000Z", "DTEND:20300115T110000Z",
                  f"SUMMARY:{title}", f"STATUS:{status}", "TRANSP:TRANSPARENT", "END:VEVENT"]
    return ("\r\n".join(lines + ["END:VCALENDAR", ""]).encode())


def coordinator(profile, events, path, crash=False):
    @contextmanager
    def session(*, write=False):
        with outlook_session(write=write) as adapter:
            if fixture(adapter) != profile.destination:
                raise RuntimeError("Test destination changed")
            if crash and write:
                original = adapter.create_appointment
                def create(*args, **kwargs):
                    original(*args, **kwargs)
                    os._exit(23)  # Verified Save, before coordinator receipt/finally.
                adapter.create_appointment = create
            yield adapter
    return ImportCoordinator(session_factory=session,
        fetcher=lambda url: FeedSnapshot(feed(events), "example.invalid", datetime.now(timezone.utc), ""),
        lock_directory=path.parent / "acceptance-locks")


def run_sync(profile, events, path, expected, crash=False):
    runner = coordinator(profile, events, path, crash)
    preview = runner.preview(profile)
    result = runner.execute(profile, preview, approvals=RunApprovals(True, True, True))
    assert result.status is RunStatus.COMPLETED, result.status
    assert result.counts == expected, result.counts
    return result


def originals(profile):
    with outlook_session() as adapter:
        items = adapter.read_inventory(profile.destination).items
        own = [item for item in items if isinstance(item, TargetState)
               and item.identity.profile_id == profile.profile_id]
        assert all(item.busy_status == 2 and not item.reminder_set and item.meeting_status == 0
                   for item in own)
        return {item.identity.uid: item.entry_id for item in own}


def mirror_action(profile, action, source_id=None):
    import mirror
    mirror._initialized = True
    mirror.config = SimpleNamespace(DRY_RUN=False, MIRROR_SUBJECT_PREFIX="[Mirror] ",
        MIRROR_CATEGORY="Mirrored", DO_NOT_SYNC_CATEGORY="DoNotSync", SECRET_CATEGORY="Secret",
        ANONYMIZE_ALL=False, ANONYMIZED_SUBJECT="Belegt")
    mirror.log = logging.getLogger("acceptance.mirror")
    email = f"acceptance-{profile.profile_id}@example.invalid"
    with outlook_session(write=True) as adapter:
        folder = adapter.resolve_calendar(profile.destination)
        copies = [item for item in folder.Items
                  if (mirror.get_mirror_source_id(item) or "").startswith(email + "::")]
        if action == "create":
            source = adapter._namespace.GetItemFromID(source_id, profile.destination.store_id)
            assert mirror.get_mirror_source_id(source) is None
            mirror.create_mirror(source, folder, email)
        elif action == "update":
            assert len(copies) == 1
            source = adapter._namespace.GetItemFromID(source_id, profile.destination.store_id)
            mirror.update_mirror(source, copies[0])
            assert copies[0].Subject == "[Mirror] C updated"
        else:
            assert len(copies) == 1
            assert not copies[0].UserProperties.Find("IcsImportProfileId")
            mirror.delete_mirror(copies[0])


def cleanup(path):
    profile = load_profile(path)
    prefix = f"acceptance-{profile.profile_id}@example.invalid::"
    with outlook_session(write=True) as adapter:
        folder = adapter.resolve_calendar(profile.destination)
        ids = []
        for item in folder.Items:
            marker = item.UserProperties.Find("IcsImportProfileId")
            mirrored = item.UserProperties.Find("MirrorSourceID")
            if ((marker and marker.Value == str(profile.profile_id))
                    or (mirrored and str(mirrored.Value).startswith(prefix))):
                ids.append(item.EntryID)
        for entry_id in ids:
            adapter._namespace.GetItemFromID(entry_id, profile.destination.store_id).Delete()
        assert folder.Items.Count == 0, "Unexpected items retained; never delete foreign data"
    path.unlink()


def phase(name, path):
    profile = load_profile(path)
    a = ("a", "A", "TENTATIVE")
    c = ("c", "C", "CONFIRMED")
    if name == "seed":
        run_sync(profile, [a], path, {}, crash=True)
        raise AssertionError("Crash hook was not reached")
    before = originals(profile)
    assert len(before) == 1
    run_sync(profile, [a], path, dict(create=0, update=0, delete=0, unchanged=1))
    run_sync(profile, [("a", "A updated", "CONFIRMED"), ("b", "B", "CONFIRMED"), c],
             path, dict(create=2, update=1, delete=0, unchanged=0))
    assert originals(profile)["acceptance-a"] == before["acceptance-a"]
    run_sync(profile, [("a", "A", "CANCELLED"), c], path,
             dict(create=0, update=0, delete=2, unchanged=1))
    c_id = originals(profile)["acceptance-c"]
    mirror_action(profile, "create", c_id)
    run_sync(profile, [("c", "C updated", "CONFIRMED")], path,
             dict(create=0, update=1, delete=0, unchanged=0))
    assert originals(profile)["acceptance-c"] == c_id
    mirror_action(profile, "update", c_id)
    run_sync(profile, [], path, dict(create=0, update=0, delete=1, unchanged=0))
    assert originals(profile) == {}
    mirror_action(profile, "delete")  # Survived importer deletion as an unmanaged item.


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--run", action="store_true")
    parser.add_argument("--cleanup", action="store_true")
    parser.add_argument("--phase", choices=("seed", "resume"), help=argparse.SUPPRESS)
    parser.add_argument("--state", type=Path, default=ROOT / ".venv/tmp/ics-acceptance-state.json")
    args = parser.parse_args()
    if args.phase:
        phase(args.phase, args.state)
    elif args.cleanup:
        cleanup(args.state)
    elif args.run:
        if args.state.exists():
            raise RuntimeError("Previous state exists; inspect and use --cleanup first")
        from dataclasses import asdict
        with outlook_session() as adapter:
            destination = fixture(adapter)
            assert not adapter.read_inventory(destination).items, "Fixture must be empty"
        args.state.parent.mkdir(parents=True, exist_ok=True)
        args.state.write_text(json.dumps(dict(profile_id=str(uuid4()), destination=asdict(destination))), encoding="utf-8")
        try:
            for name, expected in (("seed", 23), ("resume", 0)):
                result = subprocess.run([sys.executable, str(Path(__file__).resolve()),
                    "--phase", name, "--state", str(args.state)], cwd=ROOT, timeout=60)
                assert result.returncode == expected, (name, result.returncode)
            print("PASS: crash/restart, idempotency, update, cancellation, absence deletion, Busy=2, mirror isolation")
            print("OPEN: Exchange server free/busy and Outlook-client restart")
        finally:
            cleanup(args.state)
            print("Cleanup verified: fixture empty")
    else:
        parser.print_help()


if __name__ == "__main__":
    try:
        main()
    except OutlookReadError as error:
        print("Read diagnostic:", error.diagnostic, flush=True)
        raise

# Manual ICS importer: merge-request draft

## Proposed title

Add a manual one-way ICS calendar importer for classic Outlook

## Proposed description

Users can manually synchronize a public HTTPS ICS feed into an explicitly selected
classic Outlook calendar. A saved profile gives each source a stable identity:
repeated runs preserve imported appointments, update changed events in place and
remove only owned, confirmed cancellations or events missing from confirmed source
coverage. The source is authoritative; the tool never writes back to it.

The separate GUI starts idle and presents a reviewable preview before changes.
Imported appointments are Busy, recipient-free and reminder-free. Scope/ownership
checks, per-run deletion approvals, optimistic state checks and fresh verification
guard writes; uncertain operations stop without automatic retry. A new preview
with the same profile supports recovery after an interrupted run. Unsupported or
ambiguous input fails rather than becoming a partial deletion snapshot.

The implementation includes bounded recurrence/exception/timezone parsing, an
atomic per-user profile store, complete Outlook inventory reads, process locks,
background execution with live verified rows and cancellation, and a separate
portable PyInstaller build. Change-key snapshots avoid reopening unchanged items
in large destinations while checking additions, removals and edits on every read.

The existing Mirror application, installer, build and autostart behavior are
unchanged relative to the dependency branch. The shared category parsing helper
is reused from that branch. Legacy/unmarked appointments are not adopted;
the README covers first use, migration, coverage assertions and error recovery.

Validation:

- 214 offline importer tests pass; dependency consistency and whitespace checks pass.
- Opt-in real Outlook acceptance covers create, unchanged rerun, stable in-place
  update, cancellation/removal, interrupted-process recovery, persisted Busy=2
  and direct Mirror create/update/delete interoperability in an isolated calendar.
- The frozen EXE passes offline parser/timezone/profile/idle-GUI checks as a normal
  user from a copied directory with spaces and with Python/Tcl removed from PATH.
- The user accepted the GUI, bounded real import (approximately one second per
  appointment) and the compiled EXE on 2026-10-26.

Evidence and reproducible commands: [acceptance](../tests/ics_importer/ACCEPTANCE.md),
[build](../tests/ics_importer/BUILD.md), [release checks](../tests/ics_importer/RELEASE.md).

Remaining verification: server-side Exchange free/busy visibility to another
account, restarting Outlook itself, full simultaneous cross-account Mirror sync,
and a separate Windows machine with no Python installed. An additional calendar's
Busy=2 is not counted as proof of account-level free/busy publication. These checks
remain open; the recorded user acceptance does not establish them.

## Local branch and dependency preparation

Source: `codex/ics-outlook-importer`. Local dependency/base:
`sync-filter-DoNotSync-and-Secret` at `8eca081`.
The local `main` at `3c7763e` does not contain those privacy changes.
A comparison to local `main` therefore also includes the dependency's four commits.
The importer uses `sync_policy` helpers introduced there; dropping the dependency
without adapting/reviewing those imports would leave an incomplete change.

Prepare a dependent MR against `sync-filter-DoNotSync-and-Secret` where that branch
is available, or wait for its integration into the original project. After that,
refresh the intended upstream `main`, rebase the importer as appropriate, review
the resulting diff and rerun validation before targeting `main`. This file does not
assert that the remote original project's current integration state was checked.
Only the fork's `origin` remote is configured locally; no network fetch, push or MR
creation was performed for this preparation.

The independent `codex/installer-build-environment` branch's commits `72cacd3`
and `727ca3a` are not ancestors of this importer branch. `installer.iss`,
`build.bat`, `requirements.txt`, Mirror runtime/configuration and privacy helpers
have no importer changes relative to the dependency base. The importer is a
portable distribution; installer inclusion is a separate follow-up.

Feature/feature-documentation commits and test/test-documentation commits are
separate throughout the importer series. Keep them separate when moving the
series; tests depend on the corresponding feature layers. Do not squash the
dependency into the importer MR. The temporary local plan, real downloaded feeds,
profile URLs and generated distributions are not committed. The source URL example
in the README is synthetic and cannot be fetched.

Final review and any remote publication remain separate actions after step 20.

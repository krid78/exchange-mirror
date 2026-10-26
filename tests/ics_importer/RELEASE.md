# Step 20 documentation and delivery checks

Recorded 2026-10-26. This step changes documentation only. No new behavioral
regression test is needed; feature documentation and this verification record
are committed separately, matching the existing feature/test separation.

## Usage walkthrough from an empty profile file

The README flow was followed with a real withdrawn Tk window and an explicit,
initially nonexistent temporary `--profiles`-equivalent path. The actual GUI,
profile persistence and coordinator were used with the existing synthetic minimal
ICS fixture, in-memory Outlook adapter and injected calendar picker/confirmations.
This is an offline documentation check, not a new real-calendar acceptance run.

| Action | Observed result |
| --- | --- |
| Open with no profiles | Empty profile list, no source read, no job, synchronization disabled. |
| Enter source/date/timezone and choose target | Synthetic destination selected explicitly; coverage remains unchecked. |
| Save profile | File created with a stable profile UUID. |
| Build preview | One proposed creation; target remains empty. |
| Confirm synchronization | One verified creation. |
| Preview and synchronize again | Zero writes, one unchanged event; profile UUID preserved. |
| Close and cleanup | Window closed; all temporary profiles/locks removed. No network or real Outlook access. |

Parser, migration boundaries, coverage/deletion, conflict and restart behavior are
covered by the existing focused tests described in [README.md](README.md).
To repeat those checks and the fresh isolated profile/runtime check:

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests/ics_importer -p 'test_*.py' -q
.\.venv\Scripts\python.exe -m pip check
.\.venv\Scripts\python.exe tests/ics_importer/check_frozen_runtime.py
```

The documentation step reran **214 tests**, all passed with no reported skips;
`pip check` and `git diff --check` passed. The existing compiled runtime was not
rebuilt for documentation-only edits; its verification is in [BUILD.md](BUILD.md).
All 18 local Markdown link targets in the edited documentation resolve to files.
The user explicitly accepted the importer and compiled EXE on 2026-10-26. That
acceptance does not close the separate server/free-busy and clean-machine checks.

## Local branch and patch audit

Baseline: `sync-filter-DoNotSync-and-Secret` at `8eca081`. The 42 existing importer
commits before the step-20 documentation pair were inspected for mixed feature/
test paths: none mixes `tests/` changes with feature files. Both new documentation
commits follow the same separation.

The entire importer patch history and current changed files were scanned for the
real feed host/query parameters. The sole allowed query-shaped fixture is the
synthetic `example.invalid` HTTP redaction test with a deliberately fake token.
No actual feed URL/token or downloaded feed was found. No profile JSON, generated
binary/spec, temporary plan or real-calendar fixture is tracked in the importer
diff. Synthetic `.ics` test fixtures are intentional.

The importer diff leaves `installer.iss`, `build.bat`, `requirements.txt`,
`mirror.py`, `tray.py`, `config.py`, `settings.py` and `sync_policy.py` unchanged.
Both independent installer commits (`72cacd3`, `727ca3a`) are absent from its
ancestry. The temporary plan remains local and untracked.

Useful read-only commands for repeating the branch/scope checks:

```powershell
git merge-base HEAD sync-filter-DoNotSync-and-Secret
git log --reverse --oneline sync-filter-DoNotSync-and-Secret..HEAD
git diff --name-only sync-filter-DoNotSync-and-Secret...HEAD
git diff --check sync-filter-DoNotSync-and-Secret...HEAD
git diff sync-filter-DoNotSync-and-Secret...HEAD -- installer.iss build.bat requirements.txt mirror.py tray.py config.py settings.py sync_policy.py
git merge-base --is-ancestor 72cacd3 HEAD
git merge-base --is-ancestor 727ca3a HEAD
git status --short
```

The two ancestry checks return **1**, meaning those commits are not ancestors;
the protected-file diff is empty. This is the intended result, not a test failure.

The local `main` (`3c7763e`) lacks the privacy dependency's four commits. Only the
fork's `origin` remote is configured. No external repository state was fetched,
no rebase performed and no remote MR/push created. The prepared
[MR draft](../../docs/ics-importer-merge-request.md) records the dependent target
and the need to refresh/review/test the upstream base before later publication.

Remaining acceptance limits are unchanged: actual server-side account free/busy,
Outlook-application restart, full concurrent cross-account Mirror synchronization,
and an independent Windows system without Python. See [ACCEPTANCE.md](ACCEPTANCE.md)
and [BUILD.md](BUILD.md) for evidence boundaries and reproduction instructions.

# Standalone importer build verification (step 19)

Recorded on 2026-10-26: Windows 11 x64, Python 3.12.10, PyInstaller 6.22.3,
hooks-contrib 2026.9, pywin32 312, icalendar 7.3.0, python-dateutil 3.0.0,
tzdata 2026.5. PyInstaller and dependencies were installed in `.venv` with
explicit download authorization. The build itself performs no package download.

## Reproduce

From the repository directory, using a normal, non-elevated Windows account:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-importer-build.txt
.\build-importer.bat
.\.venv\Scripts\python.exe tests/ics_importer/check_frozen_runtime.py
.\.venv\Scripts\python.exe -m unittest discover -s tests/ics_importer -p 'test_*.py' -q
.\.venv\Scripts\python.exe -m pip check
```

The build produced `dist/ics-importer/ics-importer.exe` plus `_internal` and
completed with exit code 0. The existing Mirror build and installer files were
unchanged. The generated EXE, spec and work directories are ignored, not committed.
Optional PyInstaller warnings concern platform-specific/optional imports, including
the unused pytz provider; the packaged zoneinfo provider passed the runtime check.

The opt-in smoke script copies the **entire** distribution to a temporary path
containing spaces, starts the real EXE with an unrelated working directory,
removes Python/Tcl/Tk environment overrides and virtual-environment activation,
and sets PATH to Windows directories only. `PYTHONTZPATH` is empty to require
the bundled timezone database. APPDATA/TEMP point to isolated test directories.
The child has a 45-second timeout. Temporary distribution/report/profile files
are removed afterward, leaving the original `dist/ics-importer` artifact intact.

The EXE's explicit `--check-runtime` path imports COM libraries without attaching
to Outlook, parses a synthetic timezone-aware event, checks winter/summer offsets,
saves/reloads a temporary profile and creates a withdrawn real importer GUI.
It checks that no job, preview or run has started and synchronization is disabled.
It does not call calendar selection, source transport or Outlook write operations.

## Results

- Real frozen EXE: exit code 0, `ok=true`, `frozen=true`, `elevated=false`.
- All four runtime checks passed: COM imports, timezone/parser, profile roundtrip,
  GUI initialized idle. No user APPDATA profile storage was created or accessed.
- **214 regression tests passed**, no skips reported; `pip check` passed.
- Existing user profile bytes were preserved by the runtime regression test.
  Failure injection produced only a static stage/type diagnostic.

This workstation still has Python installed. Removing Python from PATH and using
the frozen runtime verifies the portable bundle here; an independent Windows
machine with no Python installation has **not** been tested. Interactive EXE
usability review remains with the user. No Outlook integration test was repeated
for packaging: the offline smoke imports the COM runtime but does not attach to
Outlook. Existing real-COM acceptance evidence remains in [ACCEPTANCE.md](ACCEPTANCE.md).

The user separately reported that step 18 now imports at approximately one second
per appointment and that an import limited to a selected date range succeeded.
These are user acceptance observations, not measurements from this build smoke test.

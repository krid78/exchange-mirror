# ICS importer test foundation

Run from the repository root on Windows:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-importer.txt
.\.venv\Scripts\python.exe -m unittest discover -s tests/ics_importer -p "test_*.py" -v
```

The parser tests need the dependencies from `requirements-importer.txt`; package
installation requires network access once. The test run itself is entirely offline.
The tests import `ics_model` and `ics_source`, but do not import the tray/mirror application,
connect to Outlook, download a feed, or modify calendars. No credentials or further
test packages are required. This branch had no existing automated tests when step 1
was added; this command runs the importer suite.

## Model contracts (step 2)

`test_model.py` covers versioned identity encoding, distinct profiles/UIDs/slots,
opaque Unicode/escaped UIDs, equivalent UTC instants, date versus date-time,
microsecond precision, and stable keys after edits/rescheduling. It also checks
normalized fields, exclusive all-day ends, invalid/floating times, cancellation
tombstones, source revisions, immutable records, target flags and conflict evidence.
The operation records reject source/target identities that do not match their key.

These are model contracts only; parser behavior is covered separately below.

## Singleton parser (step 3)

`test_source.py` checks minimal and multiple events, UTF-8/folding/escapes,
categories and optional defaults, DTEND/DURATION, exclusive DATE ends, revisions,
tentative status and cancellation tombstones. Invalid or duplicate scalar fields,
bad times, partial snapshots, HTML, unexpected components and duplicate identities
must fail the entire parse instead of returning a partial event list.

UTC/date parsing is supported; floating/TZID timestamps, the timezone-bearing
timetable fixture and recurrence properties must currently fail explicitly.
Source alarms and attendees are not copied. LF/CRLF and UTF-8 BOM inputs are covered.
These tests do not prove HTTP, time-zone conversion, recurrence expansion, planning
or Outlook behavior; those belong to later steps.

## Fixtures

| File | Purpose |
| --- | --- |
| `fixtures/minimal.ics` | One synthetic UTC event with the minimum fields needed for the first parser tests. |
| `fixtures/timetable.ics` | Two independent synthetic events reflecting the observed source structure: publication feed, Europe/Berlin timezone definition, tentative status and Microsoft busy status, opaque transparency, text fields and sequence numbers. |

The timetable's recurrence rules describe daylight-saving timezone transitions;
its events are not recurring series. Subjects, rooms, dates and identifiers are
synthetic, not a copy of the actual calendar. The fixtures contain no source URL
or access token. Keep downloaded feeds out of version control.

The initial tests check that these assets can be read offline, have balanced
component boundaries and unique synthetic UIDs, and retain the source characteristics
needed by subsequent tests. They are **not** RFC validation or a production parser.
Parser behavior, synchronization planning and Outlook integration are not implemented
or tested in step 1. The source's tentative state is intentionally retained in the
fixture so later tests can verify that imported availability is forced to Busy.

Line endings may be LF or CRLF in the checkout. Future parser tests must additionally
exercise RFC-style CRLF input, folded lines, invalid inputs and time semantics.

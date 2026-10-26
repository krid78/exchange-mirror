# ICS importer test foundation

## GUI synchronization (step 17)

`test_gui_run.py` covers concrete confirmation/cancellation, non-overridable coverage
blockers, separate empty/large-removal confirmations, worker execution, double-click
protection, refreshed-preview review without automatic continuation, verified versus
uncertain row outcomes, cancellation/close reporting and coverage invalidation.
Unsaved identities cannot execute; saving matching reviewed inputs preserves the
preview, and execution verifies the persisted profile again before starting.
Coordinator tests verify monotonic verified progress and that observer errors cannot
abort successful writes or echo raw error data. The step-16 startup test now verifies
that execution starts disabled and is never invoked automatically. Picker tests also
check the visible Alt+Tab hint while the native dialog may be hidden.

Authorized real GUI-worker check on 2026-10-26 used a synthetic feed and real COM in
**Exchange Mirror Importer Test**, with explicit harness confirmations. It first
created two items, then displayed/applied one Create, one Update and one Delete.
The table showed each verified outcome and the report matched the receipts. The
final repeat also verified disabled synchronization until saving, then preservation
of the same reviewed preview. Profiles were saved only in a temporary test directory,
which was removed afterward; real user profile storage was untouched. Cleanup removed
the remaining two synthetic items and left the test calendar empty. Native dialogs
and interactive usability remain for user review.

## Preview GUI (step 16)

`test_gui.py` covers form defaults/validation, stable draft IDs, explicit mapping
changes, coverage reset, display-only destination renames, preview rows and static
error messages. Worker tests check background execution and rejection of overlapping
jobs. Withdrawn Tk windows exercise startup without automatic reads, validation,
cancelled selection, preview rendering, invalidation, temporary profile persistence,
read errors, event-loop responsiveness, late-result rejection, close-after-read and
corrupt-file preservation. All backends/dialogs are injected; no real Outlook, native
picker, network or user profile files are accessed. Tk cases skip only if no display
is available; the Windows run executed them successfully.

Separate authorized read-only GUI-worker check on 2026-10-26 used the real HTTPS
source and real Outlook inspection of the empty **Exchange Mirror Importer Test**
calendar. It rendered 87 preview rows and the additional-calendar free/busy notice.
No profile was saved, the execute API was never invoked and the test calendar
remained empty. The Tk root was withdrawn for this smoke test; native picker and
visual usability review remain for the user's interactive test.

`test_calendar_picker.py` tests plain metadata, cancellation with no fallback,
account-default versus additional calendars, unknown classification and invalid/
missing folders using fake Outlook sessions. No calendar picker is opened by tests.

Manual user review: start `.venv\Scripts\python.exe ics_importer.py`, paste the
desired public HTTPS source, choose the intended calendar, adjust the exclusive
date window and build a preview. Check layout, keyboard/paste behavior, folder
selection/cancellation, account/free-busy hints and preview notes. Stop after Build
preview for read-only checks; the current GUI can also apply explicitly confirmed
changes as described in step 17 above. Use `--profiles PATH` for a separate test profiles
file if desired. Interactive feedback remains with the user; the automatic checks
do not claim to replace native dialog or visual usability review.

## Run coordination and process locks (step 15)

`test_run.py` connects synthetic feeds to a stateful in-memory adapter. It covers
read-only preview, unchanged reruns, Create/Update-before-Delete ordering, source/
target/profile freshness, metadata-only changes, enumeration order, explicit empty-
snapshot approvals, pending deletions, conflict blocking, cancellation and partial
receipts. Failed/uncertain writes must not retry or proceed to deletion. Changed
source/target state or failed capture before deletion must stop that phase.

Restart tests simulate a persisted Create followed by an error or SystemExit before
the receipt: a fresh inventory finds the existing identity and does not create a
second copy. This is a stateful adapter interruption test, not a real COM-server
crash test. Other cases verify snapshot-local custom zones/BOM input, generator
materialization and source-URL redaction in failures/logs. Read-session tests now
also reject nested COM sessions before a second CoInitialize call.

`test_lock.py` uses actual OS locks and independent Python processes. A competing
process cannot acquire the same profile/destination, a different destination can,
and os._exit releases the lock without deleting its empty file. Children have a
15-second test timeout; no Outlook or network access is used by these tests.

Authorized real integration check on 2026-10-26 in **Exchange Mirror Importer Test**:
synthetic transport data plus real COM first created two items. Changing the source
after preview returned REVIEW_REQUIRED without writes. After renewed review, one
Create, one Update and one Delete completed; another identical run was Unchanged
for both remaining items. Cleanup removed the two remaining synthetic items, leaving
the test calendar empty. No production appointment was edited or message sent.
An initial harness incorrectly retained COM objects across nested sessions and failed
before writes; the corrected harness uses only plain destination IDs between sessions.
The product session API now rejects that unsupported nesting explicitly. The live
test did not kill Outlook or test server-side free/busy publication.

## Single-item deletion (step 14)

`test_outlook_delete.py` checks confirmation/operation shape, scoped deletion with
an unmarked neighbor, stale/missing targets, post-inventory parent/marker changes,
foreign-profile expectations, unknown schemas and duplicates. Meetings, recipients
and unsaved edits prevent deletion. Failures before/after removal, failed readback
and a reappearing identity must not trigger retries or false success. Save/Send are
never invoked by the delete path.

Authorized real Outlook check on 2026-10-26 in **Exchange Mirror Importer Test**:
two synthetic items were created. With confirmed synthetic coverage, the pure planner
produced one DELETE for the missing source and retained the sentinel. The adapter
rejected an unconfirmed call, deleted exactly the planned target, and verified the
remaining sentinel's complete state. Repeating the stale operation was rejected.
The sentinel was then removed by the isolated harness; the test calendar is empty.
No existing user calendar items were changed and no messages were sent.

## In-place updates (step 13)

`test_outlook_update.py` verifies retained EntryID/identity, text/location/time
updates, unchanged no-Save behavior, timed/all-day transitions, flag repairs, stale
snapshots, duplicate identities, unsaved edits, post-inventory ownership/parent
changes and recipient rejection. Save and readback failures must never create a
replacement or report success. The fakes prohibit Send/Delete during update.

Authorized real Outlook check on 2026-10-26 in **Exchange Mirror Importer Test**:
one synthetic item passed in-place text/time update, unchanged repeat, stale-snapshot
rejection, timed -> all-day -> timed transitions and Busy/reminder repair. Its
EntryID stayed unchanged and the folder contained exactly one item throughout.
The test item was then removed by the isolated harness; the folder is empty again.

## Single-create adapter (step 12)

`test_outlook_create.py` uses fake COM writes and independently reopened records.
It verifies all markers/content/flags on the first Save, UTC-tagged marshaling
inputs, exclusive all-day dates, long Unicode identities, LF/CRLF normalization,
repeat calls without another Add/Save, existing-state conflicts and preflight errors.
Save failures before/after persistence and readback/parent/marker/flag failures
must report uncertainty without retrying or deleting. Unexpected recipients discard
the unsaved object; no test sends a meeting response or invitation.

Authorized real Outlook check on 2026-10-26: a new isolated calendar **Exchange
Mirror Importer Test** was created below the standard calendar. An initial test
found that naive UTC inputs were shifted by pywin32 through the machine timezone;
the adapter now preserves explicit UTC tags. After that correction, a timed item
with multiline body, a multi-day all-day item, and a timed item with a 1,082-character
serialized Unicode identity all passed Save/reopen/content/marker/Busy checks.
Repeated create calls returned the same EntryIDs and did not increase the count.
All three synthetic items were removed by the isolated test harness; the test
calendar remains empty for later steps. The initial diagnostic items were also
removed. No existing calendar items were edited and no messages were sent.
This proves the tested client/store roundtrip, not a restart or server free/busy
verification; those remain later integration checks.

## Read-only Outlook adapter (step 11)

`test_outlook_read.py` uses fake namespaces, folders, collections, custom properties
and items. It covers concrete folder resolution with no Accounts dependency, full
unfiltered inventory, unmarked items without content access, UTC/all-day semantics,
canonical identity decoding, partial/corrupt markers, field/property/iterator errors,
limits/count inconsistencies, duplicate EntryIDs, preserved duplicate source keys,
unknown-schema conflicts, closed sessions and cross-thread rejection. Session fakes
verify COM cleanup after connection and context-body failures. Save/Delete fakes must
never be called. The repeatable test suite does not attach to Outlook.

Separate authorized read-only check on 2026-10-26 using pywin32 312 outside the
sandbox: the running Outlook instance's **Mein Stundenplan** calendar resolved by
its concrete IDs and yielded 87 items. Counts before, in the inventory and after
were all 87; all were unmanaged by this importer. No Add/Save/Delete/Send operation
was performed. This proves read connectivity and the unmarked inventory path, not
write permission, persistence of new markers or the real marked-item roundtrip;
those require the later isolated write tests.

## Deletion planning (step 10)

`test_deletions.py` checks confirmed versus missing/partial coverage, explicit
occurrence/UID cancellations, destination/profile/schema ownership, foreign and
unmanaged records, historical/narrowed windows, moved events, conflict handling,
and independent preview/empty-snapshot/large-removal confirmations. No test invokes
Outlook deletion; assertions inspect proposed operations and blocked candidates.
Profile tests also cover coverage persistence, old v1 compatibility and authority
reset after a source URL change.

Separate manual integration check on 2026-10-26: the authorized HTTPS source
returned 51,302 bytes; parsing produced 87 events and planning against a synthetic
empty inventory proposed 87 creates with no conflicts. No Outlook access, profile
write, source data commit or calendar mutation was involved. This one-time network
check is not part of the repeatable offline test command below and does not prove
the source's complete export coverage or future UID stability.

## Upsert planning (step 9)

`test_sync.py` covers create-to-unchanged idempotence, field and fixed-flag repairs,
comparison normalization, duplicate/schema conflicts, concrete destination/profile
checks, unmanaged/foreign records, moved-out events and absence of deletes. It
constructs normalized records directly; no COM or calendar write is simulated.

## Profiles (step 8)

`test_profiles.py` uses temporary directories under the test directory and never
touches real APPDATA. It covers roundtrip/schema, missing/corrupt/duplicate data,
stable UUIDs on confirmed mapping changes, rejection of implicit changes and a
simulated failed atomic replacement preserving the old file with no leftover temp.
No concurrent-write or Outlook-folder existence guarantee is claimed yet.

## HTTP transport (step 7)

`test_http.py` injects byte-stream responses; it never contacts a live host. Tests
cover success/metadata, no conditional headers, non-200/304/206 statuses, HTML,
truncation, size/deadline limits, invalid/compressed responses, origin checks,
same-origin redirects, HTTPS downgrade/cross-host rejection and safe diagnostics.
The default transport test verifies construction of the default TLS context; real
certificate verification belongs to the standard library, not a simulated certificate test.

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
bad times, partial snapshots, HTML, unexpected components and conflicting revisions
must fail the entire parse instead of returning a partial event list.

UTC/date parsing is supported; unconfigured floating timestamps, recurrence
without an explicit window, and unsupported RANGE/EXRULE properties fail explicitly.
Source alarms and attendees are not copied. LF/CRLF and UTF-8 BOM inputs are covered.
These tests do not prove HTTP or planning
or Outlook behavior; those belong to later steps.

## Timezones and all-day events (step 4)

`test_timezones.py` covers IANA summer/winter offsets, the embedded Europe/Berlin
definition in the timetable fixture, custom VTIMEZONE precedence, isolation between
successive snapshots, explicitly configured floating timestamps and unknown zones.
It checks elapsed time across DST, P1D versus PT24H, exclusive all-day ends and
invalid DATE/TZID combinations. Nonexistent/ambiguous local timestamps and broken
or duplicate zone definitions must fail instead of silently choosing an offset.

These tests use installed `tzdata` and synthetic calendar data only. No machine
timezone setting, network request or Outlook calendar is modified. The timetable
fixture is accepted; regular VEVENT recurrence expansion is covered below.

## Bounded recurrence expansion (step 5)

`test_recurrence.py` checks finite windows for unbounded daily rules, COUNT/UNTIL,
weekly BYDAY, RDATE unions and repeated values, EXDATE precedence (including DTSTART),
and all-day date sets. DST tests check stable local hours and per-occurrence duration
semantics; equivalent RDATE/EXDATE timestamps in different zones collapse correctly.
Window membership is based on local start dates, including starts on the previous
UTC day. Output identities retain the original occurrence slot.

Failure cases cover missing/invalid windows, work/output limits, old dense rules,
invalid RRULE termination, unsupported extensions, date/time mismatches, periods,
unsupported ranges and conflicting masters. No failed expansion is allowed to return
a partial success. Singleton behavior remains unchanged. Tests are entirely offline.

## Exceptions, cancellations and revisions (step 6)

`test_exceptions.py` verifies moved occurrences under their original keys, replacement
outside/inside the date window, occurrence tombstones, UID-wide cancellation and
standalone detached records. It checks all-day/type consistency and rejects RANGE,
duplicate RECURRENCE-ID and recurrence rules on detached exceptions.

Revision tests cover SEQUENCE before timestamps, LAST-MODIFIED/DTSTAMP ordering,
identical duplicate collapse, conflicting highest-ranked records, superseded conflicts
and independent master/exception ranks. Cancellation tests verify that an entire
cancelled series is not expanded into active imports and that contradictory newer
exceptions fail. These tests cover returned source records only; no deletion or
Outlook mutation is executed.

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

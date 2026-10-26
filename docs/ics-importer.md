# Manual ICS calendar importer

## Intended behavior

The importer will be a separate, manually started tool for synchronizing a public
ICS source into an explicitly selected Outlook calendar in one direction.
It will create new events, update existing imports, and remove owned imports when
the source's cancellation or confirmed export coverage establishes their removal.
Unrelated calendar items must remain unchanged. Stable source identities must
prevent duplicates when an event's subject or time changes.

Imported events will always use Outlook's **Busy** availability, including when
the source marks them as tentative. They will be ordinary appointments without
invitations or meeting responses. The selected destination must contribute to the
intended Exchange account's free/busy information for others to see that time as busy.

The tool will have its own entry point and source/destination settings. It will
not start with the existing tray application's synchronization loop and will not
require two Outlook accounts or privacy settings.

## Implementation status

Steps 1 and 2 establish synthetic ICS examples, an offline test foundation and
pure event/target/operation records in `ics_model.py`.
Step 3 adds offline parsing of simple VEVENTs in `ics_source.py`.
Step 4 adds snapshot-local TZID/VTIMEZONE resolution, explicit floating-time zones
and daylight-saving-aware duration handling.
Step 5 expands regular event recurrence sets in an explicit finite date window.
Step 6 resolves source revisions and overlays detached exceptions and cancellations.
Step 7 adds a bounded HTTPS download API in `ics_http.py`.
Step 8 adds validated JSON profiles and explicit source/destination changes.
Step 9 adds pure target projection and create/update/unchanged/conflict planning.
Step 10 adds coverage-aware deletion candidates and explicit confirmation gates.
Step 11 adds the read-only classic Outlook adapter in `ics_outlook.py`.
Step 12 adds an explicit single-appointment create and verification path.
Step 13 adds optimistic in-place updates of verified owned appointments.
Step 14 adds confirmed single-item deletes with ownership and readback checks.
Step 15 connects fetch, parse, preview and execution through a guarded run coordinator.
Step 16 adds a manually started, preview-only Tk GUI.
Step 17 adds confirmed GUI synchronization, progress, cancellation and result reporting.
The existing Exchange Mirror behavior is unchanged.

Later steps will complete integration acceptance, packaging and delivery documentation.

## Event identity contract

`EventIdentity` consists of a profile UUID, an opaque case-sensitive ICS UID and
the original recurrence slot. A single event uses an explicit `single` sentinel.
Recurring instances use their original date or timestamp, even when rescheduled.
The subject, current start time, source URL and Outlook EntryID are not part of
this identity. Changing a URL token must retain the profile UUID.

Serialization is a versioned JSON array with explicit date/date-time tags;
`key` is the SHA-256 hexadecimal digest of that UTF-8 encoding. A date is distinct
from midnight on that date. Aware timestamps normalize to UTC with microsecond
precision, so equivalent instants have the same identity. Naive timestamps are
rejected: a future parser must resolve floating times using the profile timezone.
The v1 encoding is a persistence contract; changing it later requires migration.

`EventFields` represents text, categories and a start/end pair (both dates or both
aware timestamps). Date ends are exclusive. Zero-duration timed events are allowed;
an end before the start is rejected. `SourceEvent` adds status and revision data;
cancelled events may be content-free tombstones. `TargetState` records observed
flags, including non-busy availability that a later planner must repair.
`SyncOperation` describes create/update/delete/unchanged/conflict without executing
anything and can retain multiple matching targets as conflict evidence.

These records themselves do not parse ICS, resolve revisions or execute actions.
The separate planner enforces the Busy projection; Outlook writing remains a later step.

## Offline parsing (step 3)

Install `requirements-importer.txt` into the development environment. The importer
uses the `icalendar` library, pinned independently of the existing application.
`parse_ics(data, profile_id)` accepts UTF-8 bytes or text and returns a tuple of
`SourceEvent` records. It performs no file, network or Outlook operations.

The current subset supports one complete VERSION:2.0 publication calendar with
PRODID, independent VEVENTs, required UID/DTSTAMP and DTSTART for active events.
UTC timestamps and DATE values are accepted. Source revisions and detached
exceptions follow the rules below. Folded lines
and escaped text are decoded by the library.

DTEND and positive DURATION are alternatives. A missing end defaults to the same
instant for timed events or the next exclusive date for all-day events. Missing
text/categories become empty, STATUS defaults to confirmed and SEQUENCE to zero.
LAST-MODIFIED, otherwise DTSTAMP, supplies revision time. CANCELLED events may be
identity-only tombstones with DTSTAMP. Source tentative status is preserved in the
model; forcing target availability to Busy belongs to later projection/writing.
Source alarms, attendees and organizer are not copied into the model.

Malformed components/properties, duplicate mapped scalar fields, conflicting end fields,
unsupported components and incomplete snapshots raise `IcsParseError` for the whole
snapshot. There is no partially successful result. A valid empty calendar returns
an empty tuple; this alone never authorizes deletion. This is a deliberately bounded
import subset, not a general RFC validator.

Library references: [usage](https://icalendar.readthedocs.io/en/stable/usage.html),
[parse errors](https://icalendar.readthedocs.io/en/stable/how-to/parse-errors.html).

## Timezone and all-day contract (step 4)

Event timestamps may use UTC or a TZID with an IANA timezone such as Europe/Berlin.
Embedded VTIMEZONE definitions take precedence over installed timezone data and
are rebuilt for each snapshot, including custom TZIDs. Missing/duplicate/broken
definitions and unknown zone references fail the complete snapshot. A prior parse
must not supply a cached custom zone to a later snapshot. `tzdata` provides IANA
data on Windows; `python-dateutil` supplies transition validation. Both direct
dependencies are recorded in `requirements-importer.txt`.

Floating timestamps require the caller to pass `floating_timezone="Europe/Berlin"`
(or a custom TZID defined in that snapshot) to `parse_ics`. The machine timezone
and X-WR-TIMEZONE display metadata are never implicit defaults. Unknown TZIDs do
not fall back to the profile zone. Ambiguous autumn times and nonexistent spring
times are rejected with `IcsParseError`; the source must provide an unambiguous
instant, for example UTC. This is an explicit conservative import policy rather
than silently selecting one side of a transition.

The model retains date-only all-day boundaries without timezone conversion; their
end is exclusive and must follow their start. DATE cannot carry TZID or a clock
duration. A missing all-day end means the following date. Timed explicit ends are
compared as UTC instants, including when start and end straddle a transition.
Durations apply nominal weeks/days in local calendar time, then exact hours,
minutes and seconds on the UTC timeline. Thus P1D can span 23 or 25 hours at a
daylight-saving transition, while PT24H always spans 24 hours. The parser preserves
the original duration units rather than inferring them from a normalized timedelta.

VTIMEZONE observance rules are used for timezone construction, independently of
VEVENT recurrence expansion. The synthetic timetable fixture parses successfully.

## Bounded recurrence sets (step 5)

For an event carrying RRULE, RDATE or EXDATE, `parse_ics` requires `window_start`
and `window_end` as `datetime.date` values, with an exclusive end. Membership uses
the occurrence's start date in the DTSTART timezone (or the date itself for all-day
events). This is a start-date expansion window, not an overlap query: an occurrence
starting earlier but ending inside the window is not included. Nonrecurring events
remain unfiltered at this parsing layer; final synchronization scope belongs to
the later planner. Window changes do not authorize deletion.

Example: `parse_ics(data, profile_id, window_start=date(2026, 10, 1),
window_end=date(2026, 11, 1))` expands starts in October. Each result is an
independent `SourceEvent` whose identity contains the original recurrence slot.
No Outlook series or current-time-dependent default is created.

The existing `python-dateutil` dependency supplies RRULE iteration. DTSTART and
RDATE form a union with rule-generated starts; EXDATE takes precedence, even over
DTSTART. Repeated dates collapse to one slot. Date/time types must match DTSTART;
equivalent timed RDATE/EXDATE values are compared in the series timezone. Daily
local times retain their wall-clock hour across DST, with the same rejection
policy for ambiguous/nonexistent output times as other local timestamps.

Explicit DTEND gives every occurrence the master's exact elapsed duration.
DURATION retains its nominal-day versus exact-hour semantics for each occurrence.
All-day starts and exclusive ends remain dates. EXRULE, RDATE periods and multiple
RRULEs remain unsupported and fail the whole snapshot. Revisions, detached exceptions
and cancellation records are handled as described below.

Defaults are 10,000 output events per snapshot and 100,000 iterated RRULE candidates
per series (`max_occurrences` / `max_candidates`). Explicit date-list entries also
count against the per-series candidate limit. Hitting a limit fails the entire parse,
never returns a truncated successful snapshot. Iteration counts candidates from
DTSTART, including those before the window; a very old dense series can therefore
fail instead of completing. These are work/output guards, not a hard elapsed-time
timeout inside the recurrence library. No additional package is introduced in step 5.

## Exceptions, cancellations and source revisions (step 6)

Revisions are grouped by profile, UID and original RECURRENCE-ID (or the single/master
sentinel). Highest SEQUENCE wins, then latest LAST-MODIFIED (otherwise DTSTAMP).
Selection is independent of input order. Identical highest-ranked duplicates collapse;
different highest-ranked payloads fail the entire snapshot, including differences in
recurrence definitions. Conflicts among superseded revisions do not invalidate a
newer unambiguous winner. Every input component must still be syntactically valid.

Detached exceptions use RECURRENCE-ID as their identity, never their replacement
DTSTART. After expanding a master, an exception suppresses that original slot and
supplies its replacement content or cancellation. Exception revisions are selected
independently of the master's revision number. Fields are interpreted as a complete
event with the usual defaults, not a partial patch of the master. Value types must
match (DATE versus DATE-TIME); an exception cannot define its own recurrence set.
RANGE parameters, including THISANDFUTURE, are explicitly rejected.

An exception is returned when either its original slot or its actual new start is
inside the expansion window. Thus moving out of the window still yields an update
under the old key, and moving into it can create the correct replacement. The master
timezone defines window membership when present; a standalone detached exception
uses its RECURRENCE-ID timezone. Standalone exceptions are accepted as explicit
records, without proving membership in an absent master's rule. An active standalone
exception still needs DTSTART; any exception needs a date window. An explicit override
can replace an EXDATE-excluded slot. A nonrecurring active master with detached
exceptions is rejected as contradictory.

STATUS:CANCELLED with RECURRENCE-ID is an occurrence tombstone and suppresses the
active generated slot. A cancelled master/no-RECURRENCE-ID record becomes one UID-wide
tombstone without expansion, even if no RRULE or DTSTART is supplied. This represents
a single-event or whole-series cancellation; the later planner must apply it only to
owned imports in the authorized target/scope. It is not itself an Outlook deletion.
Whole-series cancellation suppresses older/equal-ranked exceptions; a newer exception
without an active newer master is contradictory and fails rather than resurrecting
part of the cancelled series. A newer active master revision can supersede cancellation.

Only publication snapshots are accepted; this does not add an iTIP METHOD:CANCEL
message processor. Outlook changes and cross-run revision history remain later
concerns. Revisions are resolved within the supplied snapshot only.

## HTTPS downloads (step 7)

`fetch_ics(url)` returns an immutable `FeedSnapshot`: raw bytes, source host, UTC
fetch time, SHA-256 digest and optional ETag/Last-Modified metadata. It validates
one complete UTF-8 VCALENDAR envelope; callers must still parse every event before
using the snapshot for planning. No conditional requests or cache reuse are enabled.

Only HTTPS URLs without embedded credentials, fragments or whitespace are allowed.
TLS uses the system's verifying default context. Redirects must stay on the same
HTTPS host and port; the standard redirect handler limits loops. Cross-host redirects
require a future explicit policy change, never an implicit permission grant.

Default limits are 30 seconds and 5 MB. Socket operations use the timeout and elapsed
time is checked after each read; this is not a hard process-level deadline. Oversized,
compressed, truncated, HTML, malformed-envelope and non-200 responses (including 304)
fail with `FeedDownloadError`; they never become an empty successful snapshot.
text/calendar, text/plain and application/octet-stream are accepted only with a valid
calendar envelope. Diagnostics exclude URLs, query tokens and raw transport errors.
The optional opener injection is a trusted testing seam, not an alternate security policy.

## Profiles (step 8)

`ImportProfile` stores a stable UUID, display name, HTTPS source URL, concrete
`Destination` (store ID, folder ID, display name), floating-timezone name and an
exclusive date window. The timezone name is structurally validated here; resolving
it against IANA data or embedded VTIMEZONE definitions belongs to parsing.

`load_profiles()` / `save_profiles()` default to
`%APPDATA%\Exchange Mirror\ics-importer\profiles.json`; an explicit path is accepted
for tests. Missing files mean no profiles. Invalid JSON, unknown schema/fields,
duplicate UUIDs/JSON fields and malformed profile values fail without resetting or
rewriting the file. This is JSON data, never executable Python configuration.

Writes validate all records first, write/flush/fsync a temporary file in the same
directory and atomically replace the destination. A failed replace retains the old
file. This is atomic file replacement, not concurrent-writer coordination; locking
is a later step. Profile files contain the source URL and must not be logged or
committed. `update_profile` requires explicit confirmation for URL or destination
changes and preserves the UUID, including when renewing a query token. It does not
migrate existing Outlook items or infer that two URLs refer to the same calendar.

## Upsert planning (step 9)

`plan_upserts(profile, sources, inventory)` is a pure function over fully parsed
source records and a complete `TargetInventory`. Destination identity is checked
by store/folder IDs, not mutable display names. Source identities must belong to
the selected profile. Inventory items are marked `TargetState` records or explicit
`UnmanagedItem` records. A future adapter must abort on read/marker errors rather
than misclassify unreadable items as unmanaged.

`project_event` produces original public content with Busy=2, reminders off and
MeetingStatus=0; source tentative status does not change these flags. Comparison
ignores body line-ending differences/trailing whitespace and category ordering/case, while checking all
managed time/text/availability fields. Observed meeting status is now part of
`TargetState` (default 0 for existing callers).

The planner indexes the entire destination, so an existing item can be updated
when its new start moves outside the window. Eligibility uses actual source start,
original recurrence slot or existing target start in the profile window. Profile
timezone data defines date membership; callers can supply `window_timezone` for a
custom embedded timezone. Duplicate source/target identities or unsupported marker
schemas produce conflicts instead of another copy. Foreign-profile and unmanaged
items remain untouched. Cancelled or absent source records never cause deletion in
this API; step 10 supplies a separate deletion-aware planner. Plans perform no writes.

## Deletion planning and export authority (step 10)

`plan_sync(..., snapshot_complete=True)` adds deletion planning. The caller must
establish successful complete download and parsing before asserting completeness;
failed or partial snapshots must never be represented as an empty event list.
The function rejects a missing/false completeness assertion. No planner executes
Outlook operations, and a plan with any conflict must not be executed by later code.

Profiles can persist `ExportCoverage(start, end, confirmed)` as `source_coverage`.
Dates are exclusive at the upper bound. Missing/unconfirmed coverage grants no
absence-based deletion authority. Old v1 profile files without this optional field
load with no coverage. Updating a source URL clears coverage even for a token renewal,
so the caller must explicitly reconfirm export scope rather than silently transfer
deletion authority to a possibly different feed. Destination changes do not migrate items.

Absence candidates must be marked with this profile and known schema 1 in the exact
destination. Both the original recurrence slot (if any) and the target's actual start
must lie inside the current profile window and confirmed export coverage. This
protects historical items, narrowed windows and moved occurrences omitted by a partial
expansion. A source identity still present always prevents absence cleanup, even if
its new start moves outside the window. Unmarked and foreign-profile items are ignored.

Explicit occurrence/UID cancellations need no inference from absence and therefore
do not require export coverage, but remain limited to owned targets with original
slot or actual start in the synchronization window. Duplicate identities, unsupported
marker schemas or contradictory active/cancelled source data block all deletion
operations until resolved; unknown items are never silently treated as safe to remove.

By default every candidate stays in `SyncPlan.pending_deletions`, with a reason and
blocking conditions, rather than becoming a DELETE operation. `confirm_deletions=True`
represents acceptance of the exact preview. An empty snapshot additionally requires
`confirm_empty_snapshot=True`. Removing more than half of the owned in-scope items
requires `confirm_large_removal=True`; `maximum_removal_fraction` can configure this
threshold. These grants are per-call and not persisted. Confirmations cannot bypass
unknown coverage, scope or conflicts. `requires_confirmation` signals pending candidates;
the caller must inspect their blockers since some require corrected data/coverage.

Only confirmed, unblocked candidates enter `operations` as DELETE. No candidate is
selected by title, category or account display name, and nothing deletes from another
folder or profile. Future execution must revalidate ownership and freshness and honor
the same safety conditions; those adapter/coordinator steps are not implemented here.

## Read-only Outlook inventory (step 11)

`outlook_session()` attaches to an already running classic Outlook instance using
pywin32, initializes COM on the caller's thread and releases it on exit. It neither
starts the tray application nor loads Mirror configuration. No account-count check
is made: one Outlook account is sufficient. The context yields an `OutlookReader`
which must remain on that thread and cannot be used after closing. Nested sessions
on the same thread are rejected before COM initialization; pass plain destination
IDs between separate sessions, not retained COM folder/item references. The Windows
category separator is read from HKCU and parsed with the existing pure
`sync_policy.parse_categories` helper. pywin32 is an importer dependency on Windows;
the module's fake-adapter tests do not require a real COM server.

`describe_calendar(folder)` reads the concrete destination IDs and display path.
`read_inventory(destination)` resolves by both folder and store ID and validates
the result's identity and calendar item type. Missing/changed/non-calendar folders
fail. Resolution never uses a display-name fallback. This establishes readability,
not permission to write; write checks belong to subsequent adapter steps.

The inventory scans the entire folder without date filtering. The Items collection
view has IncludeRecurrences=False to avoid expanding infinite native series; this
does not modify an appointment. A default 100,000-item limit, iterator failures,
changed counts or duplicate EntryIDs abort the whole read. Count checks detect some
concurrent changes, not an atomic Outlook snapshot; later execution still needs
freshness checks. Duplicate importer source identities are retained as separate
records for the planner's conflict detection.

All four importer markers are custom text UserProperties:

| Name | Value |
| --- | --- |
| `IcsImportProfileId` | Canonical profile UUID |
| `IcsImportIdentity` | Complete canonical v1 `EventIdentity.serialize()` output |
| `IcsImportEventKey` | Matching SHA-256 identity key |
| `IcsImportSchemaVersion` | Positive canonical decimal text, currently `1` |

The full identity is necessary to recover UID and original occurrence slot for
updates/cancellations; its hash alone cannot provide them. Partial, malformed or
inconsistent markers fail the inventory. A readable identity with another positive
schema number remains visible as an unsupported-schema planner conflict. Completely
unmarked items become `UnmanagedItem` without reading their appointment content;
MirrorSourceID is never interpreted as importer ownership.

Marked objects must be nonrecurring appointments. Managed fields and flags are
read into `TargetState`; timed values use Outlook StartUTC/EndUTC as UTC wall-clock
COM dates, while all-day Start/End retain their exclusive date boundaries. Property
access errors are never treated as absent markers or skipped items. Errors omit
raw calendar content. `OutlookReader` contains no Add, Save, Delete or Send path;
the separate opt-in create writer is described below.

References: [folder IDs](https://learn.microsoft.com/en-us/office/vba/api/outlook.namespace.getfolderfromid),
[custom property lookup](https://learn.microsoft.com/en-us/office/vba/api/outlook.userproperties.find),
[UTC start](https://learn.microsoft.com/en-us/office/vba/api/outlook.appointmentitem.startutc),
[recurrence views](https://learn.microsoft.com/en-us/office/vba/api/outlook.items.includerecurrences).

## Single-appointment creation (step 12)

`outlook_session(write=True)` yields an `OutlookWriter`; the default session remains
read-only. `create_appointment(destination, source)` creates at most one ordinary
appointment in the concrete destination. It performs a new full inventory read
first. One matching, already-correct identity returns its observed state without
Add/Save. A different existing state, duplicate identity or unsupported matching
schema raises `OutlookCreateConflict` without writes; this API does not update.

The unsaved appointment receives source text, categories, dates and every identity
marker before its first Save. Marker fields are added to the item only, not the
folder's field registry. BusyStatus is always 2, ReminderSet is False and
MeetingStatus is 0. Unexpected recipients cause an abort before Save. The create
path never sends invitations/responses or invokes Delete.

Timed UTC values retain an explicit UTC timezone tag when passed to pywin32;
naive Python datetimes can be converted through the machine timezone. All-day
Start/End use UTC-tagged midnight COM DATE wall values to preserve their date-only,
exclusive boundaries. Those tags are a marshaling detail, not conversion of an
all-day event into a UTC-timed event. Subsecond times are rejected before Add because
this adapter does not promise their roundtrip precision. Categories containing the
Windows list separator or empty category names are also rejected rather than split.

After Save, the item is reopened by EntryID and StoreID. Parent folder, canonical
markers, identity, content and fixed flags are checked using the same public
`target_matches` comparison as planning, including LF/CRLF body normalization.
Only successful verification returns `TargetState`.

`OutlookCreateError` records whether Save was attempted and an EntryID if known.
A failed Save or failed readback may already have persisted the item. There is no
automatic retry or deletion/rollback; callers must read inventory again before
deciding what to do. Failure before explicit Save attempts to discard only the
unsaved object with Close(olDiscard). Run-wide locking and concurrent-writer races
remain step 15 concerns; the inventory preflight is not an atomic transaction.

References: [custom property creation](https://learn.microsoft.com/en-us/office/vba/api/outlook.userproperties.add),
[initial property values](https://learn.microsoft.com/en-us/office/vba/api/outlook.userproperty.value),
[Save](https://learn.microsoft.com/en-us/office/vba/api/outlook.appointmentitem.save).

## In-place updates (step 13)

`update_appointment(destination, source, expected)` requires a known-schema
`TargetState` from a recent inventory and an equal source identity. It rereads the
full inventory, requires one exact matching target snapshot, then reopens by ID
and checks destination, markers, content and flags again. Changed state, duplicate
identity, moved item or unsaved edits cause `OutlookWriteConflict` without mutation;
read/marker failures abort safely. It never falls back to Create or adopts an
unmarked item. Appointments with meeting recipients cannot be modified by this API.

If projected state already matches, no Save occurs. Otherwise the existing object
receives the same validated fields/flags and COM date handling as creation. Markers
and the original occurrence key remain unchanged. After Save, the item is reopened
under the same EntryID in the same store/folder and compared with the desired state.
Switching between timed and all-day content is supported; no replacement copy,
invitation, response or deletion is generated.

`OutlookUpdateError` carries `save_attempted` and the expected EntryID. Save/readback
failures are uncertain outcomes, not successful updates or implicit retries. A
failure before Save attempts to discard only unsaved edits. These optimistic checks
do not provide a transaction or eliminate all races with other clients; run-level
coordination remains step 15.

## Confirmed single-item deletion (step 14)

`delete_appointment(destination, operation, confirmed=True)` accepts exactly one
planned DELETE with a single expected target. The caller must provide an operation
emitted by the confirmed deletion planner; the low-level adapter does not infer
export authority or absence from the source. It rejects active-source operations,
other operation kinds and missing confirmation before any mutation.

As with updates, the full inventory is reread and the exact expected state must
still be unique. The item is reopened by EntryID/StoreID, then destination, canonical
profile/event markers, supported schema and all observed fields/flags are checked
again. Missing/stale/moved/unmarked targets, unsaved edits and marker errors abort.
Only ordinary MeetingStatus=0 appointments with no recipients can be deleted;
meeting objects are blocked rather than risking cancellation messages. There is
no fallback by subject/category or another calendar/account.

One Delete call is followed by a complete inventory read. Only if both the EntryID
and source identity are absent does the method return a `DeletedAppointment`
receipt. A remaining/reappearing identity or failed verification raises
`OutlookDeleteError` with `delete_attempted=True` and the expected EntryID. Errors
never cause a second Delete, replacement Create, Save or an assumed success.
Retrying an old operation after a successful deletion is a stale-state conflict.

Deletion uses normal Outlook behavior (typically moving to Deleted Items); it does
not purge Deleted Items or promise permanent/server-wide erasure. The method performs
no Send operation. Run-level sequencing, source freshness and cross-process locking
remain step 15 concerns; these optimistic checks are not a transaction with other clients.

## Manual run coordinator (step 15)

`ImportCoordinator` in `ics_run.py` provides the complete core flow for a future
GUI/CLI. `preview(profile)` locks the profile/destination, opens a read-only Outlook
session, downloads and fully parses the source, resolves its planning timezone and
reads the whole target inventory. It returns an immutable `ImportPreview` containing
plain records, source-byte digest and the unconfirmed plan. This dry-run invokes no
appointment mutations. Lock files are its only filesystem output; profiles are not
silently rewritten. `PreviewError` reports failure without exposing raw source/URL data.

`execute(profile, preview, approvals=RunApprovals(...))` takes the current profile
from the caller and acquires the same lock. It opens its own write-enabled COM
session and fetches/parses/reads again. Changed profile, resolved source content,
target state or proposed plan returns `REVIEW_REQUIRED` with a fresh preview before
any writes. Collection enumeration order does not matter. Raw-byte digests are
retained for traceability; changed DTSTAMP/SEQUENCE alone does not invalidate review
if revision resolution yields identical event identities, content, status and plan.
The profile timezone is resolved from each actual snapshot, including custom
VTIMEZONE definitions and BOM-bearing source data.

Conflicts block all operations. Otherwise all Create operations precede Update,
and all verified upserts precede Delete. Deletion approvals are explicit booleans
for the exact preview, empty snapshot and large removal; they are never persisted.
Unapproved/unsupported deletion candidates stay pending while safe upserts may
complete. Before any approved deletion, source and inventory are read again and
compared with the expected baseline including this run's verified upserts. New
changes, new upserts/conflicts or changed deletion decisions require fresh review;
old approval is never extended automatically to additional deletions.

Any write failure stops immediately, so a failed Create/Update cannot be followed
by deletion. There is no automatic retry or rollback of verified earlier operations.
`RunResult` carries status, verified per-operation receipts, counts, unchanged count
and a structured failure (phase, operation kind, identity hash and uncertain-outcome
flag). Counts describe verified completed operations, not a claim of an atomic
transaction or an exact count of underlying COM Save calls. Completion with pending
deletions is distinct from full completion; review/cancellation/failure after some
upserts preserves their receipts. A returned execution plan is a report, not a
reusable grant to skip a new preview. Logs use kinds/hashes and safe error types,
not subjects, raw errors or source query strings.

The optional `cancel_requested` callback is checked between stages and operations;
it does not interrupt an in-flight download/COM call or claim that a Save rolled
back. A new run always reads Outlook markers again. Thus a Save which persisted
before interruption can become Unchanged on restart without local checkpoint files.

`ics_lock.target_lock` uses nonblocking OS locks keyed by profile UUID plus concrete
store/folder IDs. Default files live beside profiles under `ics-importer/locks`;
tests can supply an explicit directory. All cooperating processes must use the
same directory. Windows locks the first byte (including beyond EOF); Unix uses
flock. Another owner causes `RunBusyError`/BUSY without waiting or starting COM.
Handles release on normal or abnormal process exit; empty files deliberately remain
to avoid unlink/recreate races. The lock does not coordinate other Windows users,
other tools or remote Outlook/Exchange clients and cannot provide a server transaction.

The coordinator retains no COM objects between preview and execution. Session,
transport and parser seams support offline harnesses; production defaults use the
existing HTTPS parser and guarded Outlook adapter. Preview and execute must run
outside any already-open `outlook_session` on the same thread.

## Preview GUI (step 16)

Start from the project directory with `.venv\Scripts\python.exe ics_importer.py`
after installing `requirements-importer.txt`. Classic Outlook must be running for
calendar selection and preview. `--profiles PATH` optionally selects another JSON
profiles file; this does not relocate the coordinator's per-user run locks.

The window offers saved profiles, New/Save profile, a HTTPS ICS URL, native Outlook
calendar selection, From/Until ISO dates, a floating-timezone setting and **Build
preview**. A new draft defaults to the current calendar year with an exclusive
Until date of January 1 of the following year. Its UUID remains stable through
previews and saving. Startup only loads the profiles file; it neither reads Outlook
nor downloads a feed. Profiles are saved only on the explicit Save action.

The selected folder is resolved by concrete IDs. Account/store and default-calendar
classification are shown; additional or unverified calendars carry a free/busy
notice. Cancelling Outlook's picker retains the previous selection. Missing or
non-calendar folders are errors, never a reason to guess another destination.
Display-name refreshes for unchanged store/folder IDs do not require a mapping
confirmation. Actual source/destination changes do; source changes clear export
coverage while retaining profile identity. Corrupt profiles disable saving to
preserve the original file; repair it and restart before saving again.

Read operations run on a worker which initializes its own COM sessions. Tk and
dialogs stay on the main thread; the worker returns only plain metadata and preview
records through a queue. Controls prevent a second simultaneous read. Late results
for changed inputs are discarded. Closing during a read waits responsively for it
to finish; cancel a still-open native Outlook folder dialog there. No unsafe thread
termination or in-flight COM interruption is attempted.

The scrollable table shows New/Update/Unchanged/Conflict entries and pending deletion
candidates with review notes and counts. Selecting a row shows its complete review
note below the table; labels wrap when the window is resized. Times are shown in the named IANA profile
zone, or explicitly UTC when a custom zone cannot be resolved locally for display.
All-day end dates stay exclusive. Input edits invalidate the old preview. Errors
show static actionable messages without raw URLs/query tokens. Preview itself makes
no appointment changes. The current GUI also offers the separate, explicitly
confirmed synchronization described below.

## GUI synchronization and results (step 17)

**Synchronize preview** is enabled only after a valid, conflict-free preview of a
saved profile. The currently displayed settings must match the persisted profile;
the file is checked again before execution. This prevents losing an unsaved import
identity on restart and creating a new profile's duplicate copies. Saving unchanged
reviewed inputs preserves the preview, so no redundant download is needed.
Input edits invalidate it. A concrete summary lists the destination, additions,
updates, eligible deletions and removals that remain pending before any writes.
Declining confirmation starts no run. Empty-source and unusually large deletions
retain separate confirmation gates. Unconfirmed coverage/scope/conflicts cannot
be overridden by answering a deletion dialog.

The checkbox **Source includes all events in this period — allow removal of missing imports**
explicitly asserts coverage for that exclusive date range. It enables absence-based
deletion planning, not deletion by itself. Source URL, period or timezone edits reset
the assertion. Saving persists it; loading checks it only for confirmed coverage
matching the displayed period. Wider/different saved coverage requires a new
assertion for the displayed range. Cancellation-based cleanup still follows the
existing planner rules. All actual deletion approvals remain per-run and unsaved.

Leave this unchecked when the source's coverage is unknown or is a rolling window.
New and changed appointments still synchronize. This is a statement about the
feed's completeness, not whether downloading or importing has finished. For example,
a complete semester export can support this assertion for its semester date range.
**About removals…** explains this distinction beside the checkbox without changing
the setting or granting deletion permission.

Execution invokes the existing coordinator on a worker thread. Tk confirmations,
progress and reports stay on the main thread. Immutable `RunProgress` updates carry
phase and verified-operation counts; observer failures cannot turn a saved operation
into a retry. Controls and a busy guard prevent double-click duplicate runs. The
table marks verified changes, failed/uncertain operations and operations not reached;
the final report counts only verified receipts and distinguishes pending, cancelled,
failed and complete runs. Unknown outcomes explicitly require fresh inventory.

Changed source/target data renders the coordinator's fresh preview and requires a
new user action/confirmation, never automatic continuation. Completed/failed/cancelled
runs require a fresh preview before another run. Unexpected worker interruption
does not claim that nothing changed. Defensive programmatic input changes during a
run request cancellation and preserve the eventual partial-result report.

**Cancel run** stops between operations. It cannot interrupt an in-flight Save and
does not undo earlier changes. Closing during execution requests cancellation,
waits for the current call and shows the final report before closing. Closing during
a read still waits for completion; a native Outlook picker must be cancelled there.

Starting calendar selection now shows a prominent inline notice that Outlook's
folder dialog may be behind the importer, with **Alt+Tab / taskbar** instructions.
The notice disappears when selection finishes. The app does not force foreground
focus or send keyboard shortcuts to Outlook.

## Acceptance scope (step 18)

The isolated Outlook acceptance run verified interrupted-process recovery,
idempotency, stable update identities, explicit cancellation and coverage-based
removal, persisted Busy=2, and direct Mirror create/update/delete interoperability.
The user reported successful GUI testing on 2026-10-26. Reproduction instructions
and exact evidence limits are in [the acceptance protocol](../tests/ics_importer/ACCEPTANCE.md).

Exchange server free/busy publication and the scheduling assistant's view of the
target account remain **unverified**. Persisted Busy=2 in an additional calendar
does not prove account availability publication. Restart recovery was verified with
a new importer process, not by restarting the user's running Outlook. Full parallel
account-to-account Mirror synchronization also remains outside this isolated test.

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
There is no importer command, UI, network client or Outlook write path yet.
The existing Exchange Mirror behavior is unchanged.

Later steps will extend time/recurrence parsing and introduce synchronization planning,
Outlook operations, user interface, and packaging in individually reviewable changes.

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

These records do not parse ICS, resolve source revisions, choose actions or enforce
the final Busy policy yet. Planning and Outlook writing remain separate later steps.

## Offline parsing (step 3)

Install `requirements-importer.txt` into the development environment. The importer
uses the `icalendar` library, pinned independently of the existing application.
`parse_ics(data, profile_id)` accepts UTF-8 bytes or text and returns a tuple of
`SourceEvent` records. It performs no file, network or Outlook operations.

The current subset supports one complete VERSION:2.0 publication calendar with
PRODID, independent VEVENTs, required UID/DTSTAMP and DTSTART for active events.
UTC timestamps and DATE values are accepted; floating times, TZID/VTIMEZONE,
recurrences and multiple revisions of one UID are explicitly rejected until their
planned steps. Consequently the timezone-bearing timetable fixture is not yet
accepted. Folded lines and escaped text are decoded by the library.

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

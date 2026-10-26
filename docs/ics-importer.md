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
UTC timestamps and DATE values are accepted. Recurrences and multiple revisions
of one UID remain explicitly rejected until their planned steps. Folded lines
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

VTIMEZONE observance rules are accepted for timezone construction; they are not
VEVENT recurrence support. The synthetic timetable fixture now parses successfully.

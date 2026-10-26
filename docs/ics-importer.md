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
There is no importer command, UI, network client, parser, or Outlook write path yet.
The existing Exchange Mirror behavior is unchanged.

Later steps will introduce parsing, synchronization planning,
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

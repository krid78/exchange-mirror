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

Step 1 establishes synthetic ICS examples and an offline test foundation only.
There is no importer command, UI, network client, parser, or Outlook write path yet.
The existing Exchange Mirror behavior is unchanged.

Later steps will introduce the event model, parsing, synchronization planning,
Outlook operations, user interface, and packaging in individually reviewable changes.

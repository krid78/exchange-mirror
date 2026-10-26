# Step 18 acceptance protocol

## Follow-up: change-key cache and live row progress (2026-10-26)

The opt-in command below was rerun successfully after introducing the session-local
change-key cache: crash/restart, idempotency, stable updates, cancellation/removal,
Busy=2 and Mirror isolation all passed. Cleanup verified the fixture empty. While
implementing the cache the real test caught differing short-term table EntryIDs
and long-term item EntryIDs; the final implementation verifies their equivalence
through Outlook. The test harness now prints sanitized read diagnostics on failure.

A separately authorized **read-only** measurement against the configured calendar
containing 1,702 items measured 23.857 seconds for the initial inventory and 0.118 /
0.141 seconds for two consecutive inventories in the same session. This measures
inventory reads, not complete per-appointment Save/verification time. No production
appointments or profiles were changed. The first inventory of each new session
still reads all items. Store/provider support determines whether the fast path is
available; unsupported providers use the complete scan.

Automatic regressions additionally cover concurrent changes between table
snapshots, two discarded reads followed by failure (three attempts maximum),
changed corrupt markers, and live per-row progress while the worker is paused.
Follow-up result: **212 tests passed**, `pip check` passed. After the final
table-lifetime adjustment, all eight cache regressions and the real Outlook
acceptance run passed again.

Run date: 2026-10-26, Windows, classic Outlook already running, Python 3.12 in
the project `.venv`, pywin32 312. The user separately reported successful GUI
testing. No messages or meeting invitations were sent. No production-calendar
appointments were used. No Outlook restart or Send/Receive was forced.

## Reproduce the isolated COM test

Use an already existing, empty **Exchange Mirror Importer Test** subcalendar
directly below Outlook's default calendar. Creation of that fixture and synthetic
write/delete testing were explicitly authorized for this session. A future operator
must obtain equivalent authorization before executing the opt-in command:

```powershell
.\.venv\Scripts\python.exe tests/ics_importer/outlook_acceptance.py --run
```

Without `--run` the script only displays help. It refuses a nonempty fixture and
never falls back to the default calendar. Synthetic UTC events in January 2030
and an offline `example.invalid` transport exercise the actual parser, planner,
locks, coordinator and COM adapter. The transport does not access the network.

The ignored `.venv/tmp/ics-acceptance-state.json` contains the test profile UUID
and concrete destination IDs for subprocess recovery. An existing state file
prevents a new run. On failure inspect it, then recover using:

```powershell
.\.venv\Scripts\python.exe tests/ics_importer/outlook_acceptance.py --cleanup
```

Cleanup revalidates the exact test destination and deletes only appointments with
this profile's importer marker or its unique synthetic Mirror source prefix. It
retains unexpected data and reports failure rather than deleting foreign items.
Deletion uses Outlook's ordinary Delete semantics; items may remain in Deleted
Items. The script never empties Deleted Items. State is retained if cleanup fails.

## Observed results

| Scenario | Result / assertion |
| --- | --- |
| Create interrupted after verified Save | Seed subprocess exits with code 23 inside the create adapter, before coordinator receipt and normal cleanup. |
| Fresh process and released OS lock | Resume subprocess reads the persisted item and completes with 0 creates, 1 unchanged; no duplicate. |
| Update plus additions | 1 update, 2 creates; original A retains its EntryID. |
| Explicit cancellation and absence removal | A cancelled, B absent within confirmed coverage: 2 deletes, C unchanged. Explicit deletion approvals supplied. |
| Source tentative/free availability | Source includes TENTATIVE and TRANSPARENT; all reopened originals have BusyStatus=2, ReminderSet=False, MeetingStatus=0. |
| Mirror create | Actual `mirror.create_mirror` accepts an imported original with no MirrorSourceID. |
| Update with Mirror present | C updates with stable EntryID, copy remains unmanaged; actual `mirror.update_mirror` updates its subject. |
| Delete isolation | Empty complete snapshot deletes only importer original C; Mirror copy survives with no importer profile marker. Actual `mirror.delete_mirror` then removes the copy. |
| Cleanup | Fresh session verifies fixture empty and removes temporary state. |

The run printed PASS for all these checks and `Cleanup verified: fixture empty`.
No adapter corrections were needed. Mirror configuration is overridden only in
the test subprocess; normal account configuration and synchronization are unused.
This proves direct entrypoint interoperability, not concurrent full-account sync.

## Still open: do not count as passed

- Exchange synchronization and server-side free/busy persistence. Reopening a COM
  item, even from a fresh importer process, can read Outlook's local cache.
- A second user's scheduling assistant must show the imported slot as **Busy**
  for the intended target account. An additional calendar cannot establish that
  account-level property. This needs a suitable approved test mailbox/default
  calendar and observer access; do not send an invitation.
- Restart of Outlook itself; only importer-process termination/restart was tested.
- Full concurrent bidirectional Mirror synchronization across separate accounts.

These limits prevent claiming complete Exchange availability acceptance. The
isolated tests and successful user GUI review are evidence for their stated scope.

## Offline regression

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests/ics_importer -p 'test_*.py' -q
.\.venv\Scripts\python.exe -m pip check
```

The new coverage-help regression confirms that explanation of missing-import
deletion does not enable coverage or invoke preview/execution. The live harness
is deliberately excluded from automatic unittest discovery.

Recorded result: **200 tests passed**, no skips reported; `pip check` reported
no broken requirements. `git diff --check` passed.

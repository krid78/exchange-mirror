# Exchange Mirror

<div align="center">

![Windows](https://img.shields.io/badge/Windows-10%2F11-0078D6?style=for-the-badge&logo=windows&logoColor=white)
![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Outlook](https://img.shields.io/badge/Outlook-COM%20Automation-0078D4?style=for-the-badge&logo=microsoftoutlook&logoColor=white)
![License](https://img.shields.io/badge/License-GPLv3-green?style=for-the-badge)

**A Windows tray app that bidirectionally mirrors calendar items between two Outlook/Exchange accounts**

[Features](#features) • [Quick Start](#quick-start) • [How It Works](#how-syncing-works) • [Configuration](#configuration-configpy) • [Building](#building-from-source)

</div>

---
<div align="center">
<img alt="exchange-mirror" src="https://github.com/user-attachments/assets/914f7cd3-325e-4ba6-812c-fbb6e2bc7180" />
</div>

## Features

- 🔄 **Bidirectional sync** — meetings in Account A appear as mirror copies in Account B, and vice versa
- 🧭 **System tray app** — runs quietly in the background with periodic automatic sync, no window to keep open
- 🪪 **Deduplication-safe** — every mirror carries a hidden `MirrorSourceID` marker, so mirrors are never re-mirrored into infinite copies
- 🙈 **Non-intrusive mirrors** — mirrored items are created as non-meeting entries with reminders disabled, so they never send invitations or trigger alerts
- 🏷️ **Clearly marked** — mirrors get a configurable subject prefix (`[Mirror] `) and Outlook category (`Mirrored`) so they're easy to spot and filter
- 🚫 **Exclude appointments** — the `DoNotSync` category prevents mirroring and removes existing copies within the sync window
- 🔒 **Hide appointment details** — use the `Secret` category for individual appointments or series, or anonymize all mirrors from Settings
- 🧪 **Dry-run mode** — preview exactly what would be created/updated/deleted before making any real changes
- 📅 **Flexible sync window** — relative (past/future days) or fixed date range, automatically split into chunks to avoid Outlook COM timeouts on large ranges
- 🧹 **Automatic cleanup** — deleted source meetings remove their mirror copies; orphaned mirrors are cleaned up safely, scoped per account direction
- 🖥️ **No admin rights needed** — settings are saved to a per-user config file, so the app never needs to run elevated (which would break Outlook COM automation)

## Quick Start

1. Run `Exchange_Mirror_Setup.exe` to install
2. Make sure Outlook is running with both accounts connected
3. Launch **Exchange Mirror** from the Start Menu (or it starts automatically on login if you chose that option during install)
4. The app appears as a blue sync icon in the system tray
5. Right-click the tray icon to start/stop sync, view settings, or exit

On first run, right-click the tray icon → **Settings...** to adjust your sync preferences — no need to hand-edit config files.

## How Syncing Works

Exchange Mirror performs a **bidirectional sync** between two Outlook accounts (Account A and Account B):

- **A → B**: Eligible appointments in A's calendar get mirror copies in B's calendar
- **B → A**: Eligible appointments in B's calendar get mirror copies in A's calendar
- **Updates**: When a source meeting changes (time, subject, location, etc.), the mirror copy is updated automatically
- **Deletes**: When a source meeting is deleted, the mirror copy is removed
- **Deduplication**: Each mirror is tagged with a hidden `MirrorSourceID` property linking it back to its source, so mirrors are never re-mirrored

Mirrored items are marked with a subject prefix (default: `[Mirror] `) and an Outlook category (default: `Mirrored`) so you can easily distinguish them. Mirrors are created as non-meeting items with reminders disabled, so they won't send invitations or trigger alerts.

The sync runs on a timer (default: every 15 minutes) and covers a configurable date window.

## Exclusion and Privacy

Apply these Outlook categories to the **original appointment**, or to the
**series** for recurring appointments. Names are matched in full, without
regard to case: `donotsync` matches `DoNotSync`, but `DoNotSyncLater` does not.
Changing the category of a mirror does not change the original's sync policy.

| Original or setting | Result in the other calendar |
|---|---|
| `DoNotSync` category | No mirror; existing matching mirrors are removed |
| `Secret` category | Anonymized mirror |
| **Anonymize all mirrored appointments** enabled | Anonymized mirrors for all eligible originals |
| None of the above | Normal appointment details |
| `DoNotSync` combined with either anonymization rule | Exclusion takes precedence |

An anonymized copy has the title **`[Mirror] Belegt`** by default. Its
description, including formatted content, location, original categories,
required/optional attendees and resource participants are cleared. Only the
configured mirror category remains. Start/end times, duration, all-day status,
availability and sensitivity are retained; reminders and invitations remain
disabled. The original is not edited.

To anonymize everything, open **Settings → General → Privacy** and enable
**Anonymize all mirrored appointments**, then save. The hint displays your
configured placeholder and secret category. The two category names and the
placeholder can be customized in `config.py`.

Changes apply on the **next sync cycle within the configured date window**:

- Adding `Secret` or enabling global anonymization also clears details from
  existing mirrors in that window.
- Removing `Secret` restores normal details unless global anonymization remains
  enabled. Turning global anonymization off restores details only for originals
  without `Secret`.
- Removing `DoNotSync` allows a mirror to be created again, with the applicable
  privacy policy.
- Copies outside the window are not retroactively cleared or removed. This also
  matters when an appointment moves outside the window.

Series markers apply to all occurrences in the window, including changed
exceptions. Where Outlook COM does not allow categories to be assigned to individual
occurrences or exceptions, set privacy categories on the entire series or on
ordinary non-recurring appointments. Per-occurrence category editing is not
a supported workflow for such clients.

Start with **Dry Run Mode** to preview creation, updates and deletion without
changing calendars. Dry-run logs for excluded or anonymized appointments omit
their original titles. New anonymized-action and deletion logs are also
redacted; old log entries are not retroactively cleared. If Outlook cannot read
the policy or complete anonymization, the app reports a warning and skips the
affected work. An existing copy may retain old details until a successful sync;
check the log rather than assuming that a failed cleanup succeeded.

Anonymization covers the fields described above. It is not a general scrubber
for arbitrary metadata or attachments manually added to a mirror.

## Account Setup

Exchange Mirror needs exactly two Outlook accounts to sync between.

**Auto-detect (default)**: If Outlook has exactly 2 accounts, they are used automatically. No configuration needed.

**Manual selection**: If Outlook has 3+ accounts, set `ACCOUNT_A` and `ACCOUNT_B` (in Settings, or directly in `config.py`) to specify which two to sync. Use either the SMTP email address or the display name:

```python
ACCOUNT_A = "user@company.com"
ACCOUNT_B = "user@personal.com"
```

You can verify your accounts by right-clicking the tray icon and choosing **Settings...** — it shows which accounts are configured and whether auto-detect is active, and includes a **Detect from Outlook** / **Test Connection** helper.

## Configuration (`config.py`)

Most settings can be changed from the tray icon's **Settings...** dialog. Advanced users can also edit `config.py` directly — changes take effect on the next sync cycle (no restart needed).

For an installed build, the active configuration is
`%APPDATA%\Exchange Mirror\config.py`; development runs use the file beside
`mirror.py`. Existing user values are preserved and missing settings are added
automatically on upgrade.

| Setting | Default | Description |
|---------|---------|-------------|
| `SYNC_INTERVAL_MINUTES` | `15` | How often to sync (minutes) |
| `SYNC_PAST_DAYS` | `7` | How many days back to sync |
| `SYNC_FUTURE_DAYS` | `60` | How many days ahead to sync |
| `SYNC_START_DATE` / `SYNC_END_DATE` | *(empty)* | Override sync window with specific dates (`YYYY-MM-DD`). Takes priority over days settings |
| `SYNC_CHUNK_DAYS` | `7` | Splits large sync windows into chunks of this many days to avoid Outlook COM slowdowns/timeouts. Set to `0` to disable chunking |
| `DRY_RUN` | `False` | Set `True` to preview changes without modifying anything |
| `LOG_FILE` | `mirror.log` | Log filename (created next to the exe) |
| `MIRROR_SUBJECT_PREFIX` | `[Mirror] ` | Prefix added to mirrored items (set to `""` to disable) |
| `MIRROR_CATEGORY` | `Mirrored` | Outlook category for mirrored items (set to `""` to disable) |
| `DO_NOT_SYNC_CATEGORY` | `DoNotSync` | Original category that excludes an appointment or series; config-file setting |
| `SECRET_CATEGORY` | `Secret` | Original category that anonymizes an appointment or series; config-file setting |
| `ANONYMIZE_ALL` | `False` | Anonymize all eligible mirrors; also available in Settings → General → Privacy |
| `ANONYMIZED_SUBJECT` | `Belegt` | Replacement title before applying the mirror prefix; config-file setting |
| `ACCOUNT_A` / `ACCOUNT_B` | *(empty)* | SMTP address or display name of accounts to sync (see Account Setup) |
| `START_ON_WINDOWS_LOGIN` | `False` | Start Exchange Mirror automatically when Windows starts |

The installer preserves your `config.py` edits during upgrades — it won't overwrite an existing config.

Privacy category names must be nonempty, distinct from each other and from the
mirror category, and contain only one category name. The placeholder must be
nonempty text. Category lists use the Windows regional list separator.

## Troubleshooting

**"Cannot connect to Outlook"**
- Make sure Outlook is running before launching Exchange Mirror
- The app connects via COM automation, so Outlook must be open (not just running in the background)

**"Need at least 2 accounts"**
- Outlook must have at least 2 email accounts configured. Check File → Account Settings in Outlook

**"ACCOUNT_A/B not found"**
- The value in `config.py` doesn't match any Outlook account. Use the exact SMTP address or display name. Right-click tray → **Settings...** to see detected accounts

**"Found N accounts — set ACCOUNT_A and ACCOUNT_B"**
- You have 3+ Outlook accounts and haven't told the app which two to sync. Set `ACCOUNT_A` and `ACCOUNT_B` in Settings or `config.py`

**Sync doesn't seem to run**
- Right-click the tray icon — if it says "Start Sync", sync is stopped. Click to start it
- Check `mirror.log` (Advanced tab → **Open Log**) for errors

**Mirrors aren't appearing**
- Enable `DRY_RUN = True` in Settings, check the log to see what would be synced, then set it back to `False`
- Make sure the sync window covers the dates you expect (`SYNC_PAST_DAYS` / `SYNC_FUTURE_DAYS` or fixed dates)

**Want to remove all mirrored items**
- Filter your calendar by the `Mirrored` category and delete them, or search for the `[Mirror]` subject prefix

## Manual ICS Importer

A separate, manually started tool synchronizes one public HTTPS ICS feed into one
explicitly selected **classic Outlook for Windows** calendar. The source is
authoritative: changed source events update imported appointments; local edits to
their imported fields can be overwritten on the next run. Nothing is written back
to the feed. Imported items are ordinary appointments, always **Busy**, with no
invitations, tentative responses or reminders. They have no privacy options.

### First import

As a normal user, start `ics-importer.exe` from the complete portable distribution
(see [Building from Source](#standalone-ics-importer)), or run the source checkout:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements-importer.txt
.\.venv\Scripts\python.exe .\ics_importer.py
```

Reuse an existing `.venv` instead of recreating it. No administrator rights are
needed. Starting the importer starts no preview or synchronization.

1. Start classic Outlook with the destination account connected. Enter a profile
   name and the source's HTTPS ICS URL (for example `https://example.invalid/calendar.ics`,
   a placeholder to replace with your actual feed).
2. Choose the destination calendar explicitly. If Outlook's folder dialog is
   behind the importer, use **Alt+Tab** or the taskbar. For account-level Exchange
   free/busy, prefer that account's default calendar; an additional calendar may
   not publish availability for the account.
3. Set **From**, **Until** and the floating-time timezone, usually `Europe/Berlin`.
   From is inclusive and Until exclusive: `2026-10-01` to `2026-11-01` covers October.
   All-day events also use an exclusive end date. Window/recurrence rules and
   pending-removal reasons are described in [the technical reference](docs/ics-importer.md).
4. Leave **Source includes all events in this period — allow removal of missing imports**
   unchecked unless the feed includes every event for that period. This describes
   source coverage, not whether downloading has finished. Unknown/rolling coverage:
   leave it off; additions and updates still work. **About removals…** explains it.
5. **Save profile**, then **Build preview**. Review the destination, new/updated/
   unchanged items, conflicts and pending removals. A preview changes no appointments.
   Edited inputs require a new preview; saving matching inputs preserves it.
6. Choose **Synchronize preview**, review the summary and confirm. The rows update
   when each change is verified. Empty-source and unusually large deletions require
   extra confirmation. **Cancel run** stops between operations; an in-flight Save
   may finish and earlier saved changes remain.

### Later runs and existing appointments

Select the same saved profile, build a fresh preview and confirm the changes.
Profiles are stored in `%APPDATA%\Exchange Mirror\ics-importer\profiles.json`.
Back up this file privately: it contains the source URL, destination IDs and stable
profile UUIDs. A feed's tokenized URL can grant access even when no login is needed.
Source and EXE use the same profile storage. `--profiles PATH` selects a separate
file when needed; restore the original file to recover the same import identity.

Avoid recreating an existing mapping with **New profile**: a new UUID owns a separate
set of imports and can create duplicates. A token refresh for the same feed should
edit the existing profile's URL; confirm the mapping change, save and rebuild the
preview. A changed URL clears the coverage assertion. Changing the destination
does not move or clean up imports in the previous calendar. Use a new profile for
an unrelated source; do not repurpose an old source identity.

Appointments previously imported manually or by another tool have no importer
ownership markers. They are left untouched, even if their subject/time matches a
source event. There is no automatic adoption or deduplication against those items.
For migration, first review a preview in a separate empty calendar. Before using
an existing calendar, back up and manually resolve only the known legacy copies.
Never delete unrelated appointments or hand-edit ownership markers.

### Removal and error handling

Only known-schema items owned by this profile in the selected destination may be
deleted, with per-run confirmation. Explicit source cancellations and absence from
a complete source are separate reasons. Absence-based removal additionally needs
confirmed coverage and both the original occurrence slot and current target start
inside the selected range. Shrinking a window does not authorize historical cleanup.
Unmarked items, other profiles and out-of-scope imports are protected. Conflicts
must be resolved before synchronization; a candidate with blockers remains pending.

| Result | Next action |
| --- | --- |
| Source/Outlook read error, invalid feed or marker | No incomplete snapshot is applied. Check the source/destination and static diagnostic, then build a fresh preview. |
| `inventory_changed` | The calendar changed during reading. Reads are retried at most twice; if it persists, finish parallel edits in Outlook and preview again. |
| Source or destination changed since preview | Review the refreshed preview and explicitly confirm again. |
| Save/Delete error or uncertain last operation | Earlier verified changes remain; the last operation may already exist. Build a fresh preview with the **same profile** before retrying. Do not assume rollback or create a replacement profile. |
| Another run is using this target | Wait for it to finish. Empty lock files may remain normally; do not remove files to bypass a live lock. |
| Profile cannot be loaded | Saving is disabled to preserve it. Keep the file and restore a known-good backup or correct its contents deliberately. |

The first destination read can be slow for large calendars; supported providers
reuse unchanged records within the session after checking Outlook change keys.
The importer is separate from the tray app and does not add autostart behavior.
Exchange server free/busy publication still needs validation for the chosen account;
Busy=2 in Outlook alone does not prove another user's scheduling-assistant view.

See [the technical reference](docs/ics-importer.md) and
[local merge-request preparation](docs/ics-importer-merge-request.md) for scope,
verification evidence and remaining integration limits.

## Building from Source

### Standalone ICS importer

Build as a normal Windows user from the project directory (no Inno Setup required):

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements-importer-build.txt
.\build-importer.bat
```

If `.venv` already exists, reuse it. The build script uses that environment and
does not download packages. Output: `dist\ics-importer\ics-importer.exe`.
Distribute the **entire `dist\ics-importer` folder**, including `_internal`.
The recipient does not need Python installed. Start `ics-importer.exe` manually;
it opens the importer without starting a preview or synchronization. Classic
Outlook must be running before selecting a calendar or synchronizing.

For offline runtime diagnostics, including the parser, bundled timezone rules,
temporary profile persistence and idle GUI initialization:

```powershell
.\dist\ics-importer\ics-importer.exe --check-runtime .\runtime-report.json
```

This writes a JSON report and exits; it does not connect to Outlook or the network
and does not load or change the user's profiles. Exit code 0 means the check
passed. The existing Mirror build and installer are independent of this artifact.

### Existing Mirror installer

Requires Python 3.10+, and [Inno Setup 6](https://jrsoftware.org/isinfo.php) for the installer.

**One-time setup** (run in an Administrator command prompt):

```bat
choco install innosetup -y
```

**Build:**

```bat
pip install -r requirements.txt
build.bat
```

`build.bat` does three things:
1. Installs Python dependencies
2. Runs PyInstaller to produce `dist\mirror\`
3. Runs Inno Setup to produce `dist\Exchange_Mirror_Setup.exe`

**Running in dev mode** (requires Outlook already running):

```bat
run_dev.bat
```

or directly:

```bash
python tray.py       # full tray app with background sync
python mirror.py     # one-shot sync (interactive account selection if 3+ accounts)
```

## Uninstalling

Use **Add or Remove Programs** in Windows Settings, or run the uninstaller from the Start Menu group. The auto-start registry entry is removed automatically.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the GNU General Public License v3.0 — see the [LICENSE](LICENSE) file for details.

```
Exchange Mirror
Copyright (C) 2026  Markus Begerow
```

## 🙋‍♂️ Get Involved

If you encounter any issues or have questions:
- 🐛 [Report bugs](https://github.com/markusbegerow/exchange-mirror/issues)
- 💡 [Request features](https://github.com/markusbegerow/exchange-mirror/issues)
- ⭐ Star the repo if you find it useful!

## ☕ Support the Project

If you like this project, support further development with a repost or coffee:

<a href="https://www.linkedin.com/sharing/share-offsite/?url=https://github.com/markusbegerow/exchange-mirror" target="_blank"> <img src="https://img.shields.io/badge/💼-Share%20on%20LinkedIn-blue" /> </a>

[![Buy Me a Coffee](https://img.shields.io/badge/☕-Buy%20me%20a%20coffee-yellow)](https://paypal.me/MarkusBegerow?country.x=DE&locale.x=de_DE)

## 📬 Contact

- 🧑‍💻 [Markus Begerow](https://linkedin.com/in/markusbegerow)
- 💾 [GitHub](https://github.com/markusbegerow)
- ✉️ [Twitter](https://x.com/markusbegerow)

---

**Privacy Notice**: Exchange Mirror runs locally through Outlook COM and does
not use a separate sync service. Mirrored data is saved to the configured
destination account and handled by that account's calendar provider. Use the
privacy controls above to choose which appointments and details are mirrored.

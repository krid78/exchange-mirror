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
- 🧪 **Dry-run mode** — preview exactly what would be created/updated/deleted before making any real changes
- 📅 **Flexible sync window** — relative (past/future days) or fixed date range, automatically split into chunks to avoid Outlook COM timeouts on large ranges
- 🧹 **Automatic cleanup** — deleted source meetings remove their mirror copies; orphaned mirrors are cleaned up safely, scoped per account direction
- 🖥️ **No admin rights needed** — settings are saved to a per-user config file, so the app never needs to run elevated (which would break Outlook COM automation)

## Quick Start

1. Run `Exchange_Mirror_Setup.exe` to install
2. Make sure Outlook is running with both accounts connected
3. Launch **Exchange Mirror** from the Start Menu
4. The app appears as a blue sync icon in the system tray
5. Right-click the tray icon to start/stop sync, view settings, or exit

On first run, right-click the tray icon → **Settings...** to adjust your sync preferences — no need to hand-edit config files.

### Installation scope and updates

For a new installation, Setup defaults to **Install for me only**, normally
under `%LOCALAPPDATA%\Programs\Exchange Mirror`, without requesting administrator
rights. You can instead select **Install for all users** to install under
`Program Files`; this requires administrator approval. Start Menu and optional
desktop shortcuts follow the selected scope. Run the installed application
normally as the Windows user whose Outlook profile should be synchronized.

Updates automatically reuse the detected installation scope and normally keep
the existing directory. An existing all-users installation therefore still
requires administrator approval when updated; it does not automatically become
a per-user installation. For scripted installations, `/CURRENTUSER` and
`/ALLUSERS` explicitly select the scope.

**Changing scope is a migration, not an in-place update.** Forcing a per-user
installation while an all-users installation exists can leave two independent
copies, duplicate shortcuts, and an autostart entry pointing to the old copy.
Running both copies may also cause concurrent synchronization. Before switching:
stop and exit Exchange Mirror, disable its autostart in Settings, back up your
configuration, and uninstall the old copy (administrator approval may be needed).
Then install in the desired scope and re-enable autostart for the new copy.
Existing runtime settings and logs are stored per user under
`%APPDATA%\Exchange Mirror`; preserve that directory and back up any legacy
`config.py` beside the old executable before uninstalling. Avoid keeping both
installation scopes installed at the same time.

### Start on Windows login

Autostart always belongs to an individual Windows user. A per-user installation
offers the **Start Exchange Mirror on Windows login** setup task. For an
all-users installation, each user enables **Start on Windows login** in the
application's **Settings** after launching it normally. The global installer
does not create a startup entry for the administrator account or enable startup
for everyone. Existing per-user startup entries remain in place during updates
in the same scope and directory. Disable autostart in Settings before uninstalling,
especially for an all-users installation, whose uninstaller cannot clean up every
user's startup entry.

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

Apply the `DoNotSync` category to an original appointment or to its recurring
series to exclude it from synchronization. Names are matched in full without
regard to case; `DoNotSyncLater` does not match `DoNotSync`.

Existing matching copies are removed on the next successful sync within the
configured date window. Removing the category allows a copy to be created
again. Copies outside the window are not retroactively removed.

Series markers apply to occurrences and changed exceptions. Where Outlook COM
does not support category assignment to individual occurrences or exceptions,
apply the category to the whole series or to a non-recurring appointment.

Use Dry Run Mode to preview exclusion and deletion without changing calendars.
If a marker cannot be read reliably, existing copies are preserved and the
affected appointment is skipped rather than treated as deleted.

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
| `ACCOUNT_A` / `ACCOUNT_B` | *(empty)* | SMTP address or display name of accounts to sync (see Account Setup) |
| `START_ON_WINDOWS_LOGIN` | `False` | Start Exchange Mirror automatically when Windows starts |

The installer preserves your `config.py` edits during upgrades — it won't overwrite an existing config.

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

## Building from Source

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

**Privacy Notice**: Exchange Mirror operates entirely locally via Outlook COM automation. No calendar data is sent to any external server — all sync happens directly between the two Outlook accounts already configured on your machine.

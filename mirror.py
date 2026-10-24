"""
Exchange Meeting Mirror
Bidirectionally syncs calendar items between two Outlook/Exchange accounts
using the Outlook COM interface.
"""

import importlib.util
import logging
import os
import re
import shutil
import sys
import tempfile
from datetime import datetime, timedelta

import win32com.client
import pywintypes

from sync_policy import build_target_fields, field_values_equal


def _get_app_dir():
    """Return the directory containing the executable (frozen) or script (dev)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = _get_app_dir()


def _get_data_dir():
    """Return a writable directory for logs and runtime data.

    Uses %APPDATA%\\Exchange Mirror for frozen builds (installed exe).
    Falls back to APP_DIR for dev runs.
    """
    if getattr(sys, "frozen", False):
        appdata = os.environ.get("APPDATA", "")
        if appdata:
            data_dir = os.path.join(appdata, "Exchange Mirror")
            os.makedirs(data_dir, exist_ok=True)
            return data_dir
    return APP_DIR


DATA_DIR = _get_data_dir()


def _get_config_path():
    """Return the path to the writable config.py, migrating it if needed.

    config.py lives in DATA_DIR (a user-writable folder) so that saving
    settings never requires admin rights. On first run after upgrading
    from a version that stored config.py next to the executable, an
    existing APP_DIR copy is migrated to DATA_DIR automatically.
    """
    data_config_path = os.path.join(DATA_DIR, "config.py")
    if not os.path.isfile(data_config_path):
        app_config_path = os.path.join(APP_DIR, "config.py")
        if os.path.isfile(app_config_path) and app_config_path != data_config_path:
            shutil.copyfile(app_config_path, data_config_path)
    return data_config_path


# Defaults for settings that may be missing from an existing user's
# config.py after an upgrade. config.py is migrated to DATA_DIR only once
# (see _get_config_path) and is never regenerated from the template
# afterwards, so a setting added in a later version would otherwise raise
# AttributeError for anyone who installed before it existed. New settings
# that need a default for pre-existing installs should be added here.
CONFIG_DEFAULTS = [
    ("SYNC_CHUNK_DAYS", 7, [
        "# Split large sync windows into chunks of this many days to avoid Outlook COM",
        "# slowdowns/timeouts when fetching a wide date range in one call. Set to 0 to",
        "# disable chunking and sync the whole window in a single pass.",
    ]),
    ("START_ON_WINDOWS_LOGIN", False, [
        "# Start Exchange Mirror automatically when Windows starts. Kept in sync",
        "# with the registry Run key by the Settings dialog when saved.",
    ]),
]


_STARTUP_REGISTRY_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
_STARTUP_VALUE_NAME = "ExchangeMirror"


def is_startup_enabled():
    """Return whether the per-user Run registry key for Exchange Mirror exists."""
    import winreg
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _STARTUP_REGISTRY_KEY) as key:
            winreg.QueryValueEx(key, _STARTUP_VALUE_NAME)
        return True
    except OSError:
        return False


def set_startup_enabled(enabled):
    """Create or remove the per-user Run registry key that starts Exchange
    Mirror on Windows login. No admin rights required (HKCU)."""
    import winreg
    if enabled and not getattr(sys, "frozen", False):
        raise RuntimeError("Startup toggle requires the installed application, not a dev run.")
    with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, _STARTUP_REGISTRY_KEY) as key:
        if enabled:
            winreg.SetValueEx(key, _STARTUP_VALUE_NAME, 0, winreg.REG_SZ, f'"{sys.executable}"')
        else:
            try:
                winreg.DeleteValue(key, _STARTUP_VALUE_NAME)
            except FileNotFoundError:
                pass


def _format_config_value(value):
    """Format a Python value as a config.py literal."""
    if isinstance(value, bool):
        return "True" if value else "False"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, str):
        escaped = value.replace("\\", "\\\\").replace('"', '\\"')
        return f'"{escaped}"'
    raise ValueError(f"Unsupported type for config value: {type(value)}")


def _backfill_missing_settings(config_path, content):
    """Append any CONFIG_DEFAULTS entries missing from a live config.py.

    Returns the (possibly updated) content. Existing keys, values, and
    comments are left untouched; missing settings are appended at the end,
    each with its explanatory comment, and written back atomically. If the
    write fails (e.g. permissions), the original content is returned
    unchanged and the caller falls back to in-memory defaults.
    """
    missing = [
        (key, default, comment_lines)
        for key, default, comment_lines in CONFIG_DEFAULTS
        if not re.search(rf"^{re.escape(key)}\s*=", content, re.MULTILINE)
    ]
    if not missing:
        return content

    blocks = [
        "\n".join(comment_lines + [f"{key} = {_format_config_value(default)}"])
        for key, default, comment_lines in missing
    ]
    new_content = content.rstrip("\n") + "\n\n" + "\n\n".join(blocks) + "\n"

    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w", encoding="utf-8", suffix=".py", dir=DATA_DIR, delete=False
        ) as tmp:
            tmp.write(new_content)
            tmp_path = tmp.name
        os.replace(tmp_path, config_path)
        return new_content
    except OSError:
        if tmp_path:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass
        return content


def _load_config():
    """Dynamically load config.py from the writable config location."""
    config_path = _get_config_path()
    if not os.path.isfile(config_path):
        raise RuntimeError(
            f"config.py not found at {config_path}\n"
            "Place config.py there and try again."
        )
    with open(config_path, "r", encoding="utf-8") as f:
        content = f.read()
    _backfill_missing_settings(config_path, content)

    spec = importlib.util.spec_from_file_location("config", config_path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    for key, default, _comment in CONFIG_DEFAULTS:
        if not hasattr(mod, key):
            setattr(mod, key, default)
    return mod


MIRROR_PROP_NAME = "MirrorSourceID"

SYNCED_FIELDS = [
    "Subject", "Body", "Location", "Start", "End", "Duration",
    "AllDayEvent", "BusyStatus", "ReminderSet", "Sensitivity",
    "Categories", "RequiredAttendees", "OptionalAttendees",
]

# Outlook constants
OL_FOLDER_CALENDAR = 9
OL_APPOINTMENT_ITEM = 1
OL_USER_PROPERTY_TEXT = 1  # olText

# Module-level state (lazy-initialized)
config = None
log = None
_initialized = False


def init():
    """Initialize module-level config and logging. Safe to call multiple times."""
    global config, log, _initialized
    if _initialized:
        return
    config = _load_config()
    log = setup_logging()
    _initialized = True


def setup_logging():
    cfg = _load_config() if config is None else config
    logger = logging.getLogger("mirror")
    logger.setLevel(logging.DEBUG)
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")

    if sys.stdout is not None:
        console = logging.StreamHandler(sys.stdout)
        console.setLevel(logging.INFO)
        console.setFormatter(fmt)
        logger.addHandler(console)

    log_path = os.path.join(DATA_DIR, cfg.LOG_FILE)
    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    return logger


def save_config(updates: dict) -> None:
    """Write changed config values back to config.py, preserving comments.

    Args:
        updates: dict of {key: new_value} to write. Only these fields are
                 touched; everything else (including comments) is preserved.
    Raises:
        RuntimeError: config.py not found or file write failed.
        ValueError:   a key was not found in config.py.
    """
    config_path = _get_config_path()
    if not os.path.isfile(config_path):
        raise RuntimeError(f"config.py not found at {config_path}")

    with open(config_path, "r", encoding="utf-8") as f:
        content = f.read()

    for key, new_value in updates.items():
        try:
            fv = _format_config_value(new_value)
        except ValueError:
            raise ValueError(f"Unsupported type for {key}: {type(new_value)}")

        pattern = rf"^({re.escape(key)}\s*=\s*)(.*)$"
        if not re.search(pattern, content, re.MULTILINE):
            raise ValueError(f"Key '{key}' not found in config.py")
        content = re.sub(pattern, rf"\g<1>{fv}", content, flags=re.MULTILINE)

    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w", encoding="utf-8", suffix=".py", dir=DATA_DIR, delete=False
        ) as tmp:
            tmp.write(content)
            tmp_path = tmp.name
        os.replace(tmp_path, config_path)
        if log is not None:
            log.info("Config saved: %s", ", ".join(updates.keys()))
    except Exception as exc:
        if tmp_path:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass
        raise RuntimeError(f"Failed to write config.py: {exc}") from exc


def get_outlook():
    """Get a reference to the running Outlook application."""
    init()
    try:
        return win32com.client.Dispatch("Outlook.Application")
    except Exception as exc:
        raise RuntimeError(f"Cannot connect to Outlook. Is it running? {exc}")


def get_outlook_accounts(outlook):
    """Return a list of (account_name, account_smtp, account_obj) for every account in Outlook."""
    namespace = outlook.GetNamespace("MAPI")
    accounts = []
    for i in range(1, namespace.Accounts.Count + 1):
        acct = namespace.Accounts.Item(i)
        accounts.append((acct.DisplayName, acct.SmtpAddress, acct))
    return accounts


def resolve_accounts(accounts, account_a_hint="", account_b_hint=""):
    """Pick two accounts from the list using hints (SMTP or display name).

    Returns (acct_a, acct_b) tuples of (name, smtp, account_obj).
    Raises RuntimeError if resolution fails.
    """
    if account_a_hint and account_b_hint:
        acct_a = _find_account(accounts, account_a_hint)
        acct_b = _find_account(accounts, account_b_hint)
        if acct_a is None:
            raise RuntimeError(f"ACCOUNT_A '{account_a_hint}' not found in Outlook accounts.")
        if acct_b is None:
            raise RuntimeError(f"ACCOUNT_B '{account_b_hint}' not found in Outlook accounts.")
        return acct_a, acct_b

    if len(accounts) == 2:
        return accounts[0], accounts[1]

    names = ", ".join(f"{name} ({smtp})" for name, smtp, _ in accounts)
    raise RuntimeError(
        f"Found {len(accounts)} accounts: {names}\n"
        "Set ACCOUNT_A and ACCOUNT_B in config.py to select which two to sync."
    )


def _find_account(accounts, hint):
    """Match an account by SMTP address or display name (case-insensitive)."""
    hint_lower = hint.lower()
    for name, smtp, acct in accounts:
        if smtp.lower() == hint_lower or name.lower() == hint_lower:
            return (name, smtp, acct)
    return None


def get_calendar_folder(account):
    """Get the default calendar folder for an account."""
    store = account.DeliveryStore
    return store.GetDefaultFolder(OL_FOLDER_CALENDAR)


def get_calendar_items(calendar_folder, start_date, end_date):
    """Fetch appointments from a calendar folder within the given date range.

    Returns a list of appointment COM objects. Recurring items are expanded.
    """
    items = calendar_folder.Items
    # IncludeRecurrences must be set before Sort(); doing it in the other
    # order can make Outlook re-expand the recurrence view inconsistently
    # and yield the same occurrence more than once from GetFirst()/GetNext().
    items.IncludeRecurrences = True
    items.Sort("[Start]")

    # Use DASL syntax with an unambiguous YYYY/MM/DD date literal instead of
    # the classic "[Start] >= 'MM/DD/YYYY ...'" syntax. The classic syntax is
    # parsed according to the OS regional date format, so on non-US locales
    # (e.g. de-DE, which expects DD/MM/YYYY) any date where both day and
    # month are <= 12 gets silently misread as the wrong date, making
    # Restrict() return zero results with no error.
    start_str = start_date.strftime("%Y/%m/%d") + " 12:00 AM"
    end_str = end_date.strftime("%Y/%m/%d") + " 11:59 PM"
    restriction = (
        f'@SQL="urn:schemas:calendar:dtstart" >= \'{start_str}\' AND '
        f'"urn:schemas:calendar:dtstart" <= \'{end_str}\''
    )
    restricted = items.Restrict(restriction)

    results = []
    item = restricted.GetFirst()
    while item is not None:
        try:
            # Only sync appointments (class 26 = olAppointment)
            if item.Class == 26:
                results.append(item)
        except pywintypes.com_error:
            pass
        item = restricted.GetNext()
    return results


def get_mirror_source_id(item):
    """Read the MirrorSourceID UserProperty from an item, or return None."""
    try:
        props = item.UserProperties
        prop = props.Find(MIRROR_PROP_NAME)
        if prop is not None:
            return prop.Value
    except (pywintypes.com_error, AttributeError):
        pass
    return None


def set_mirror_source_id(item, source_id):
    """Set the MirrorSourceID UserProperty on an item."""
    props = item.UserProperties
    prop = props.Find(MIRROR_PROP_NAME)
    if prop is None:
        prop = props.Add(MIRROR_PROP_NAME, OL_USER_PROPERTY_TEXT)
    prop.Value = source_id


def build_source_id(source_email, entry_id, start=None):
    """Build the deduplication key: email::EntryID[::StartISO].

    Recurring appointment occurrences fetched with IncludeRecurrences all
    report the same EntryID (the master series' EntryID, an Outlook COM
    limitation), so the occurrence's Start time is appended to give each
    occurrence its own stable, unique key.
    """
    if start is not None:
        return f"{source_email}::{entry_id}::{start.strftime('%Y-%m-%dT%H:%M:%S')}"
    return f"{source_email}::{entry_id}"


def find_mirror(dest_items, source_id):
    """Find a mirrored item in the destination by its MirrorSourceID."""
    for item in dest_items:
        if get_mirror_source_id(item) == source_id:
            return item
    return None


def get_category_separator():
    """Read the same regional list separator Outlook uses for Categories."""
    import winreg
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Control Panel\International") as key:
        separator, _ = winreg.QueryValueEx(key, "sList")
    if not isinstance(separator, str) or not separator:
        raise ValueError("Invalid Windows category list separator")
    return separator


def get_target_fields(source, separator):
    """Snapshot source fields and calculate normal mirror values."""
    init()
    source_fields = {}
    for field in SYNCED_FIELDS:
        try:
            source_fields[field] = getattr(source, field)
        except (pywintypes.com_error, AttributeError) as exc:
            log.debug("Could not read field %s: %s", field, exc)
    return build_target_fields(
        source_fields, subject_prefix=config.MIRROR_SUBJECT_PREFIX,
        mirror_category=config.MIRROR_CATEGORY, separator=separator,
    )


def copy_fields(source, dest, *, target_fields=None, separator=None):
    """Write calculated target values, including markers and fixed flags."""
    init()
    if target_fields is None:
        separator = separator or get_category_separator()
        target_fields = get_target_fields(source, separator)
    for field, value in target_fields.items():
        try:
            setattr(dest, field, value)
        except (pywintypes.com_error, AttributeError) as exc:
            if field in ("MeetingStatus", "ReminderSet"):
                raise
            log.debug("Could not copy field %s: %s", field, exc)


def apply_mirror_markers(item):
    """Apply marker values using the same calculation as the sync path."""
    init()
    target = build_target_fields(
        {"Subject": item.Subject, "Categories": item.Categories},
        subject_prefix=config.MIRROR_SUBJECT_PREFIX,
        mirror_category=config.MIRROR_CATEGORY, separator=get_category_separator(),
    )
    item.Subject = target["Subject"]
    item.Categories = target["Categories"]


def needs_update(source, mirror, *, target_fields=None, separator=None):
    """Compare the exact desired values that copy_fields will write."""
    init()
    separator = separator or get_category_separator()
    if target_fields is None:
        target_fields = get_target_fields(source, separator)
    for field, expected in target_fields.items():
        try:
            if not field_values_equal(field, expected, getattr(mirror, field), separator):
                return True
        except (pywintypes.com_error, AttributeError):
            pass
    return False


def create_mirror(source_item, dest_calendar, source_email):
    """Create a new mirrored appointment in the destination calendar."""
    init()
    source_id = build_source_id(source_email, source_item.EntryID, source_item.Start)
    subject = source_item.Subject

    if config.DRY_RUN:
        log.info("  [DRY RUN] Would CREATE mirror: %s", subject)
        return

    separator = get_category_separator()
    target_fields = get_target_fields(source_item, separator)
    new_appt = dest_calendar.Items.Add(OL_APPOINTMENT_ITEM)
    copy_fields(source_item, new_appt, target_fields=target_fields)
    set_mirror_source_id(new_appt, source_id)
    new_appt.Save()
    log.info("  Created mirror: %s", subject)


def update_mirror(source_item, mirror_item):
    """Update an existing mirror if fields have changed."""
    init()
    subject = source_item.Subject

    separator = get_category_separator()
    target_fields = get_target_fields(source_item, separator)
    if not needs_update(source_item, mirror_item, target_fields=target_fields,
                        separator=separator):
        log.debug("  No changes for: %s", subject)
        return

    if config.DRY_RUN:
        log.info("  [DRY RUN] Would UPDATE mirror: %s", subject)
        return

    copy_fields(source_item, mirror_item, target_fields=target_fields)
    mirror_item.Save()
    log.info("  Updated mirror: %s", subject)


def delete_mirror(mirror_item):
    """Delete a mirrored item whose source no longer exists."""
    init()
    subject = mirror_item.Subject

    if config.DRY_RUN:
        log.info("  [DRY RUN] Would DELETE mirror: %s", subject)
        return

    mirror_item.Delete()
    log.info("  Deleted mirror: %s", subject)


def get_sync_window():
    """Return (start_date, end_date) based on config.

    Uses SYNC_START_DATE / SYNC_END_DATE if set, otherwise falls back
    to the relative SYNC_PAST_DAYS / SYNC_FUTURE_DAYS.
    """
    init()
    if config.SYNC_START_DATE:
        start_date = datetime.strptime(config.SYNC_START_DATE, "%Y-%m-%d")
    else:
        start_date = datetime.now() - timedelta(days=config.SYNC_PAST_DAYS)

    if config.SYNC_END_DATE:
        end_date = datetime.strptime(config.SYNC_END_DATE, "%Y-%m-%d")
    else:
        end_date = datetime.now() + timedelta(days=config.SYNC_FUTURE_DAYS)

    return start_date, end_date


def _chunk_date_range(start_date, end_date, chunk_days):
    """Split [start_date, end_date] into consecutive (chunk_start, chunk_end) tuples.

    If chunk_days is <= 0 or the window is no larger than chunk_days, yields the
    whole window as a single chunk (no behavior change for small windows).
    """
    if chunk_days <= 0 or (end_date - start_date).days <= chunk_days:
        yield start_date, end_date
        return

    chunk_start = start_date
    while chunk_start <= end_date:
        chunk_end = min(chunk_start + timedelta(days=chunk_days - 1), end_date)
        yield chunk_start, chunk_end
        chunk_start = chunk_end + timedelta(days=1)


def sync_one_direction(source_account, dest_account, source_email):
    """Sync meetings from source_account's calendar into dest_account's calendar."""
    init()
    log.info(
        "Syncing: %s -> %s",
        source_account.DisplayName,
        dest_account.DisplayName,
    )

    start_date, end_date = get_sync_window()

    source_calendar = get_calendar_folder(source_account)
    dest_calendar = get_calendar_folder(dest_account)

    chunks = list(_chunk_date_range(start_date, end_date, config.SYNC_CHUNK_DAYS))
    total = len(chunks)
    for i, (chunk_start, chunk_end) in enumerate(chunks, start=1):
        if total > 1:
            log.info(
                "  Chunk %d/%d: %s -> %s",
                i, total, chunk_start.strftime("%Y-%m-%d"), chunk_end.strftime("%Y-%m-%d"),
            )
        _sync_window(source_calendar, dest_calendar, source_email, chunk_start, chunk_end)


def _sync_window(source_calendar, dest_calendar, source_email, start_date, end_date):
    """Create/update/delete mirrors for a single date window."""
    init()
    source_items = get_calendar_items(source_calendar, start_date, end_date)
    dest_items = get_calendar_items(dest_calendar, start_date, end_date)

    log.info(
        "  Source items: %d, Dest items: %d", len(source_items), len(dest_items)
    )

    # Track which source IDs are still valid (for orphan cleanup)
    valid_source_ids = set()

    # Create or update mirrors
    for src_item in source_items:
        # Skip items that are themselves mirrors (created by a previous sync)
        try:
            if get_mirror_source_id(src_item) is not None:
                continue
        except (pywintypes.com_error, AttributeError) as exc:
            log.warning("  Skipping item due to error: %s", exc)
            continue

        # Mark the source ID valid as soon as it's known (before any risky
        # create/update work) so a transient error below can never make an
        # already-mirrored, still-existing item look like an orphan.
        try:
            source_id = build_source_id(source_email, src_item.EntryID, src_item.Start)
        except (pywintypes.com_error, AttributeError) as exc:
            log.warning("  Skipping item due to error: %s", exc)
            continue
        valid_source_ids.add(source_id)

        try:
            mirror = find_mirror(dest_items, source_id)
            if mirror is None:
                create_mirror(src_item, dest_calendar, source_email)
            else:
                update_mirror(src_item, mirror)
        except (pywintypes.com_error, AttributeError) as exc:
            log.warning("  Skipping item due to error: %s", exc)
            continue

    # Delete orphaned mirrors (source meeting was deleted)
    for dest_item in dest_items:
        try:
            mirror_id = get_mirror_source_id(dest_item)
            if mirror_id is None:
                continue
            # Only clean up mirrors that originated from this source account
            if not mirror_id.startswith(source_email + "::"):
                continue
            if mirror_id not in valid_source_ids:
                delete_mirror(dest_item)
        except (pywintypes.com_error, AttributeError) as exc:
            log.warning("  Skipping item during orphan cleanup due to error: %s", exc)
            continue


def main():
    init()
    outlook = get_outlook()
    accounts = get_outlook_accounts(outlook)

    if len(accounts) < 2:
        log.error(
            "Need at least 2 accounts in Outlook, found %d.", len(accounts)
        )
        sys.exit(1)

    log.info("Found %d accounts:", len(accounts))
    for i, (name, smtp, _acct) in enumerate(accounts):
        log.info("  [%d] %s (%s)", i, name, smtp)

    if len(accounts) == 2:
        acct_a = accounts[0]
        acct_b = accounts[1]
    else:
        print("\nEnter the index numbers of the two accounts to sync.")
        idx_a = int(input("Account A index: "))
        idx_b = int(input("Account B index: "))
        acct_a = accounts[idx_a]
        acct_b = accounts[idx_b]

    name_a, smtp_a, account_a = acct_a
    name_b, smtp_b, account_b = acct_b

    log.info("Syncing between: %s <-> %s", name_a, name_b)
    if config.DRY_RUN:
        log.info("*** DRY RUN MODE - no changes will be made ***")

    # A -> B
    sync_one_direction(account_a, account_b, smtp_a)
    # B -> A
    sync_one_direction(account_b, account_a, smtp_b)

    log.info("Sync complete.")


if __name__ == "__main__":
    main()

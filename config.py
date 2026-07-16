# How far back to sync (days)
SYNC_PAST_DAYS = 7

# How far ahead to sync (days)
SYNC_FUTURE_DAYS = 60

# Optional: override the sync window with specific dates (format: "YYYY-MM-DD")
# When set, these take priority over SYNC_PAST_DAYS / SYNC_FUTURE_DAYS
SYNC_START_DATE = ""
SYNC_END_DATE = ""

# Split large sync windows into chunks of this many days to avoid Outlook COM
# slowdowns/timeouts when fetching a wide date range in one call. Set to 0 to
# disable chunking and sync the whole window in a single pass.
SYNC_CHUNK_DAYS = 7

# Preview mode - set True to see what would be synced without making changes
DRY_RUN = False

# Start Exchange Mirror automatically when Windows starts
START_ON_WINDOWS_LOGIN = False

# Log file path
LOG_FILE = "mirror.log"

# Prefix added to the subject of mirrored copies (set to "" to disable)
MIRROR_SUBJECT_PREFIX = "[Mirror] "

# Outlook category assigned to mirrored copies (set to "" to disable)
MIRROR_CATEGORY = "Mirrored"

# How often to sync in tray mode (minutes)
SYNC_INTERVAL_MINUTES = 15

# Account selection for tray mode (SMTP address or display name)
# Leave empty to auto-detect (works when exactly 2 accounts exist)
ACCOUNT_A = ""
ACCOUNT_B = ""

# =============================================================================# --- XLSX Column Layout ---
# Single definition shared by step1 and step2. Edit here only.
READWISE_COLS = [
    "highlight_id", "title", "author", "date_added", "date_modified",
    "tags", "date", "quote", "note", "location", "location_type",
    "source_url", "readwise_url", "category", "book_id",
]
CURATION_COLS = ["CommonBook", "updated"]
ALL_COLS      = READWISE_COLS + CURATION_COLS

# TagOS Readwise → Commonplace Book Configuration
# =============================================================================

# --- Readwise API ---
READWISE_TOKEN = ""  # https://readwise.io/access_token

# --- Paths ---
XLSX_PATH            = ""
OBSIDIAN_QUOTES_DIR  = ""
OBSIDIAN_STAGING_DIR = ""
QUOTES_TEMPLATE_PATH = ""

# --- XLSX Sheet ---
SHEET_NAME = "Highlights"

# --- Staging delay ---
# Save the XLSX every N rows processed (new or updated). Prevents data loss
# if the script crashes mid-run. Set to 0 to disable periodic saves.
XLSX_SAVE_INTERVAL = 1
# Seconds to wait after moving each new note into the vault.
# Gives Obsidian/Templater time to process each file before the next arrives.
STAGING_DELAY = 1.0
INCLUDE_VALUE = "Y"       # Value in CommonBook column to include a quote
DEFAULT_VALUE = ""        # Default (empty = No)

# --- Curation ---
INCLUDE_VALUE = "Y"       # Value in CommonBook column to include a quote
DEFAULT_VALUE = ""        # Default (empty = No)

# --- CommonBook Auto-inclusion ---
# New highlights are auto-marked Y in CommonBook only if their category is in
# INCLUDE_CATEGORIES, AND they do not carry any tag in EXCLUDE_TAGS.
# Tags are matched against the merged highlight+book tag string.
# Readwise categories: books, articles, podcasts, tweets, supplementals
INCLUDE_CATEGORIES = ["books", "supplementals"]
EXCLUDE_TAGS       = ["study", "sy"]


# --- Note Title ---
# Format: YYYY-MM-DD -- Q{counter}
# Counter is zero-padded to this many digits (2 → Q00, Q01 ... Q99)
COUNTER_PADDING = 2

# --- YAML: extra static tags added to every note ---
EXTRA_TAGS = ["readwise"]

# --- Obsidian Database Type tag (shown in your example notes) ---
OBS_DATABASE_TYPE = "QuotesV0_3"

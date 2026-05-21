import os
from dotenv import load_dotenv

load_dotenv()  # reads .env from the project root if present

# =============================================================================
# --- XLSX Column Layout ---
# Single definition shared by step1 and step2. Edit here only.
READWISE_COLS = [
    "highlight_id", "title", "author", "date_added", "date_modified",
    "tags", "date", "quote", "note", "location", "location_type",
    "source_url", "readwise_url", "category", "book_id",
]

# Optional display-name overrides for XLSX column headers.
# Use "" to keep the READWISE_COLS name at that position.
# Must match the length of READWISE_COLS (shorter lists are padded with "").
COL_NAMES = [
    "highlight id", "", "", "date added", "date modified",
    "", "", "", "", "", "location type",
    "source", "RW url", "", "book id",
]

CURATION_COLS = ["CommonBook", "updated"]

# Derived column values (XLSX_COLS, ALL_COLS, KEY_TO_HEADER, HEADER_TO_KEY)
# are computed in columns.py — import them from there, not here.

# TagOS Readwise → Commonplace Book Configuration
# =============================================================================

# --- Personal variables (set in .env, never committed) ---
READWISE_TOKEN       = os.getenv("READWISE_TOKEN", "")   # https://readwise.io/access_token
XLSX_PATH            = os.getenv("XLSX_PATH", "")
OBSIDIAN_QUOTES_DIR  = os.getenv("OBSIDIAN_QUOTES_DIR", "")
OBSIDIAN_STAGING_DIR = os.getenv("OBSIDIAN_STAGING_DIR", "")
QUOTES_TEMPLATE_PATH = os.getenv("QUOTES_TEMPLATE_PATH", "")

# --- XLSX Sheet ---
SHEET_NAME = "Highlights"

# --- Timing ---
# Save the XLSX every N rows processed (new or updated). Prevents data loss
# if the script crashes mid-run. Set to 0 to disable periodic saves.
XLSX_SAVE_INTERVAL = 1
# Seconds to wait after moving each new note into the vault.
# Gives Obsidian/Templater time to process each file before the next arrives.
STAGING_DELAY = 1.0

# --- Curation ---
INCLUDE_VALUE = "Y"       # Value in CommonBook column to include a quote
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
EXTRA_TAGS = []

# --- Obsidian Database Type tag (shown in your example notes) ---
OBS_DATABASE_TYPE = "QuotesV0_3"

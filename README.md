# Readwise → Obsidian CommonPlace Book

A Python pipeline that pulls highlights from Readwise into individual, atomic Obsidian notes — one note per highlight — forming a personal commonplace book that integrates with your daily notes workflow.

---

## Overview

Most Readwise → Obsidian integrations produce one file per book or article. This pipeline goes further: each highlight becomes its own standalone note with full YAML frontmatter, making highlights individually linkable, queryable via Dataview, and embeddable into daily notes.

The pipeline is designed around deliberate curation. Not every highlight you make is worth keeping permanently — this system gives you a review layer (an Excel spreadsheet) between Readwise and your vault, so you decide what enters your knowledge bank. 

Take note: At this time, the script mostly decides what you want to keep based on tags. Changes YOU make to the excel sheet may only last long enough for a single run!

Also, current update processes are poor. If you change something in the config, best approach is to delete the xlsx and rebuild all files.

---

## How It Works

```
Readwise API → XLSX (review) → Python → Staging folder → Obsidian vault
                                                              ↓
                                                         Templater fires
                                                              ↓
                                                      Formatted quote note
```

1. **`step1_fetch.py`** pulls all highlights from the Readwise API and writes them to an Excel spreadsheet. Each highlight is a row. Two columns are managed by you: `CommonBook` (include this highlight?) and `updated` (has Readwise data changed?).

2. **You review the spreadsheet.** Filter by category, author, tags, or any other column. `CommonBook` is auto-populated based on configurable rules (e.g. auto-include all books, exclude highlights tagged `study`). Override by adjusting tags in Readwise — not by editing the spreadsheet.

3. **`step2_atomize.py`** reads the spreadsheet and for each `CommonBook = Y` row:
   - Builds a YAML frontmatter block from the highlight's metadata
   - Appends your Obsidian Templater quote template as the note body
   - Writes the complete file to a staging folder **outside** the vault
   - Moves it into the vault — Templater fires automatically on the complete file

4. **Templater resolves** the dynamic expressions in the note body (author, quote text, tags, date) against the YAML that's already present. The result is a fully formatted quote note ready to embed in daily notes, link from project notes, or surface via Dataview.

---

## Requirements

- Python 3.10+
- A [Readwise](https://readwise.io) account and API token
- [Obsidian](https://obsidian.md) with the following plugins:
  - [Templater](https://github.com/SilentVoid13/Templater)
  - [Dataview](https://github.com/blacksmithgu/obsidian-dataview) (optional, for queries)

### Python dependencies

```bash
pip install requests openpyxl pyyaml
```

---

## Setup

### 1. Configure `config.py`

Copy `config.example.py` to `config.py` and fill in your values:

```python
# Readwise
READWISE_TOKEN = "your_token_here"    # https://readwise.io/access_token

# Paths
XLSX_PATH            = "/path/to/readwise_highlights.xlsx"
OBSIDIAN_QUOTES_DIR  = "/path/to/vault/Quotes/"         # inside vault
OBSIDIAN_STAGING_DIR = "/path/to/staging/"              # outside vault
QUOTES_TEMPLATE_PATH = "/path/to/vault/Templates/Quotes Template.md"

# Auto-inclusion rules
INCLUDE_CATEGORIES = ["books"]          # which Readwise categories to auto-include
EXCLUDE_TAGS       = ["study"]          # highlights with these tags are excluded

# Behaviour
STAGING_DELAY      = 2.0    # seconds between file moves (gives Templater time)
XLSX_SAVE_INTERVAL = 50     # save XLSX every N rows (0 = only save at end)
```

### 2. Create your Quotes Templater template

Create a file at `QUOTES_TEMPLATE_PATH`. This is the note body that Templater will render after Python writes the YAML. It reads from the frontmatter Python has already written.

A minimal example:

```markdown
[[Quotes]] || [[<% tp.file.title.slice(0, 10) %>]]
***
![[Quotes Information#^yourid]]

> [!quote] <% tp.frontmatter.author %>
> <% tp.frontmatter.Quote %>
`$= dv.current().file.tags.join(' ')`
^quotes
```

**How this works:**
- `tp.frontmatter.author` / `tp.frontmatter.Quote` — pulls from the YAML Python wrote
- The `> [!quote]` callout renders the highlight for display
- The Dataview inline expression surfaces the note's tags
- The `^quotes` block ID makes the callout individually embeddable (e.g. `![[2025-03-01 - RW -- Q00#^quotes]]` in a daily note)

> **Important:** The template must contain **no YAML frontmatter of its own**. Python writes the frontmatter first; Templater appends the body. If the template has its own `---` block, the file will have duplicate frontmatter.

### 3. Configure Templater folder trigger

In Obsidian: **Settings → Templater → Folder Templates**

Map your `OBSIDIAN_QUOTES_DIR` to your quotes template. Templater will then automatically apply the template to every new file that appears in that folder — which is exactly what happens when `step2_atomize.py` moves files in.

---

## Usage

### Step 1 — Sync highlights from Readwise

```bash
python3 step1_fetch.py
```

- Fetches all highlights and writes/updates the XLSX
- New highlights get `CommonBook` set automatically based on `INCLUDE_CATEGORIES` / `EXCLUDE_TAGS`
- Existing highlights are re-evaluated; changed rows get `updated = Y`
- Safe to run repeatedly — unchanged rows are skipped

### Step 2 — Atomize to Obsidian

```bash
python3 step2_atomize.py
```

- Reads rows where `CommonBook = Y`
- New notes are written to staging, moved into vault, Templater fires
- Existing notes with `updated = Y` are auto-updated if tracked fields changed
- `updated` flag is cleared after each processed row

**Graceful shutdown:** press `Ctrl+C` once. The script finishes the current row, saves the XLSX, and exits cleanly.

---

## XLSX Structure

The spreadsheet is formatted as an Excel Table with filter dropdowns on every column.

### Readwise columns
| Column | Description |
|---|---|
| `highlight_id` | Readwise highlight ID |
| `title` | Book / article title |
| `author` | Author |
| `date_added` | When the highlight was made |
| `date_modified` | When Readwise last modified it |
| `tags` | Merged highlight + book tags |
| `date` | Highlight date (YYYY-MM-DD) |
| `quote` | Highlight text |
| `note` | Your annotation from Readwise |
| `location` | Location in source (page, offset, etc.) |
| `location_type` | Location format |
| `source_url` | URL of source document |
| `readwise_url` | Direct link back to highlight in Readwise |
| `category` | Readwise category (books, articles, podcasts, etc.) |
| `book_id` | Readwise book/document ID |

### Curation columns
| Column | Description |
|---|---|
| `CommonBook` | `Y` = include in vault; blank = exclude. Auto-set by step1; do not edit manually. |
| `updated` | `Y` = Readwise data has changed; step2 will auto-update the vault note. Cleared after processing. |

> **Tip:** Use Excel's filter dropdowns to review highlights by category or author before running step2. You can also filter to `CommonBook = Y` to see exactly what will be written to your vault.

---

## Note Format

### Filename
```
YYYY-MM-DD - RW -- Q00
YYYY-MM-DD - RW -- Q01
...
```

The date is the highlight date. The `QXX` counter is global across all highlights on that date and continues from the highest existing counter in the vault, so re-runs never produce duplicate filenames.

For highlights with no date (some supplementals), the fallback order is: `highlighted_at` → `updated_at` → today's date.

### YAML frontmatter

```yaml
title: 2025-03-01 - RW -- Q00
book title: The Innovator's Dilemma
author:
  - Clayton Christensen
tags:
  - readwise
obs note type: Quotes
obs version: v0_5_0
Quote: The reason why it is so difficult for existing firms to capitalize on disruptive
  innovations is that their processes and values are optimized for sustaining innovations.
summary: Key thesis of the book
source: ''
RW source: https://readwise.io/open/123456789
highlight id: '123456789'
```

---

## Update Behaviour

Step2 never blindly overwrites existing notes. It compares only the fields that Readwise controls:

**Tracked fields:** `book title`, `author`, `tags`, `Quote`, `summary`, `source`, `RW source`, `highlight id`

**Never compared:** `title`, `obs note type`, `obs version`, note body

This means edits you make inside Obsidian (annotations, links, body content) are never overwritten by a re-run. Only a genuine change in Readwise (e.g. you edited the highlight text or added a note) triggers an update, and only when `updated = Y` is set by step1.

---

## Daily Notes Integration

Once notes are in your vault, the `^quotes` block ID in the template body makes each quote individually embeddable:

```markdown
![[2025-03-01 - RW -- Q00#^quotes]]
```

You can also query your quotes vault with Dataview. For example, all quotes from a specific book:

```dataview
TABLE author, date FROM "Quotes"
WHERE book-title = "The Innovator's Dilemma"
SORT date ASC
```

Or surface recent quotes in your daily note template:

```dataview
LIST FROM "Quotes"
WHERE date = date(today)
```

---

## File Structure

```
project/
├── config.py              # your local config (not committed)
├── config.example.py      # template — copy and fill in
├── step1_fetch.py         # Readwise → XLSX
├── step2_atomize.py       # XLSX → Obsidian notes
├── step1_fetch.log        # generated on run
├── step2_atomize.log      # generated on run
└── README.md
```

---

## Design Decisions

**One note per highlight, not one per book.**
Atomic notes are individually linkable and queryable. A highlight buried in a book-level note cannot be embedded in a daily note or surfaced by Dataview topic queries.

**XLSX as review layer, not a database.**
A spreadsheet is the most accessible format for human curation — filtering, sorting, spot-checking — without requiring a separate UI. The Excel Table format adds filtering at no cost.

**`CommonBook` is script-managed.**
Auto-inclusion rules live in `config.py`; exclusions are expressed as Readwise tags on the source material. This keeps curation decisions close to the content, not in the spreadsheet. Re-tagging a book in Readwise and re-running step1 automatically clears it from the pipeline.

**Staging → vault move, not direct write.**
Writing the complete file outside the vault first, then moving it in, ensures Templater fires on a file that already has its YAML frontmatter. Writing directly to the vault risks Templater executing before the file is fully written.

**Field-scoped comparison.**
Comparing full file content would flag every note on every run (due to `date modified` changing). Only the fields Readwise controls are compared, so user edits inside Obsidian are never at risk.

---

## License

MIT

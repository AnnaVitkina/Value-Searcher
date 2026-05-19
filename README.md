## What is it?

**Value Finder (File Search Utility)** is an internal tool designed to locate a **specific value** across many shared-drive folders without opening each workbook manually. The user enters a search term and one or more root folder paths; the tool walks the folder tree, finds “leaf” locations that hold the latest reference files, searches the **most recently modified** supported document in each location, and reports **exact word matches** (case-insensitive) with sheet/cell or row details—so analysts can see **where** a name or code appears (e.g. account coding logic, Packard logic) instead of checking every versioned file by hand.

## **What the automation actually does**

### **One run across many folders instead of manual file-by-file search**

With **Value Finder**, a user runs `final.py`, enters the **value to search for** and **folder path(s)**, and the script then automatically:

1. **Resolves paths** — Combines optional `BASE_PATH` (e.g. Google Colab Shared drives) with relative folder names, or accepts full paths. Converts Windows **Google Drive for Desktop** paths (e.g. `G:\Shared drives\...`) into the Colab/Drive layout; rejects pasted Google Drive **URLs** and shows path help instead.
2. **Discovers “final” folders** — Recursively scans each root path and collects folders that have **no subfolders** except optionally one named **`Old`** (case-insensitive). Folders named `Old` and anything inside them are **skipped**; paths under an `Old` ancestor are ignored.
3. **Lists supported files per final folder** — In each final folder only (not recursive into subfolders), collects files with extensions **`.xlsx`**, **`.xls`**, **`.docx`**, **`.csv`**.
4. **Selects one file to search** — Picks the file with the **latest modification time** (`st_mtime`) among those supported types. **Version numbers and “draft” in the filename are not used**—only the filesystem date.
5. **Applies exact-match rules** — Normalizes text (Unicode NFKC, Unicode spaces → ASCII space), then matches the search term as a **whole word** (regex word boundaries, case-insensitive). Examples: `Anna` matches `anna` in `Hello Anna!` but not `Joanna` or `annabel`; `Turner` matches inside `Davies Turner`.
6. **Searches Excel (.xlsx)** — Opens the workbook read-only with **calculated values** (`data_only=True`), walks every sheet and cell, extracts plain text from **rich-text / italic** cells via `CellRichText`, and records up to five match locations (sheet + cell coordinate).
7. **Searches legacy Excel (.xls)** — Uses `xlrd` over all sheets and cells with the same exact-match logic.
8. **Searches Word (.docx)** — Scans paragraph text and table cell text; reports paragraph previews or table positions.
9. **Searches CSV** — Tries multiple encodings (UTF-8, UTF-8 BOM, Latin-1, CP1252, etc.), line-by-line and per-column exact match, with Excel-style column letters where possible.
10. **Builds structured results per folder** — For each final folder: folder name/path, full list of supported filenames (sorted), which file was searched, its modified timestamp, whether a match was found, and match details (locations).
11. **Prints a clear summary** — Match run: folders with hits show all files in the folder (★ marks the searched file), cell/row details, and full path to the matching file. **No matches**: summary still lists each folder and **which file was actually searched** (so users can see if a newer file was chosen instead of the one they expected).

The workflow is **path- and folder-structure-driven**: point at shared-drive roots, enter a term, re-run when folders or files change—no manual “open each xlsx and Ctrl+F” across dozens of versioned logic files.

## **What you still decide (human in the loop)**

- **Search value** — The exact word or phrase to find (e.g. `Turner`, `Davies Turner`). Matching is whole-word only, not substring inside another word.
- **Which root folder(s)** to scan — One or more paths (relative to `BASE_PATH` or full paths). You can add multiple roots in one run.
- **`BASE_PATH` (in code)** — Typically set for Colab (`/content/drive/Shared drives/`) or local Windows; leave empty to type full paths only.
- **Which file gets searched in each final folder** — Only the **newest by modified date** among supported types. If several logic files sit in the same folder, ensure the intended file is the most recent, or adjust dates / file placement.
- **Interpreting results** — Confirm the reported “Searched” file when there is no hit; verify the folder is a “final” folder (no extra subfolders besides `Old`). Analysts still validate edge cases (merged cells, formulas vs displayed values, special characters).

### **How you run it**

The tool is a **command-line script** (`python final.py`), not a web upload UI. It prompts interactively for the search value and folder path(s), shows progress while scanning final folders, and prints formatted results to the console. It is suited to **Google Colab** (mounted Shared drive + `BASE_PATH`) or **local** use with paths from Google Drive for Desktop. Dependencies: `openpyxl`, `xlrd`, `python-docx` (installed via `pip` if missing).

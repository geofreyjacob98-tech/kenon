# CLAUDE.md

This repo holds the **SALMA Fleet Accounting System** — an Excel/Google Sheets workbook for JAFA TRADE, a cross-border trucking company (Dar es Salaam ↔ Zambia/DRC). There is no application code here; the "codebase" is the workbook's sheets and formulas.

**Before doing anything in this repo, read [`docs/PROJECT_STATE.md`](docs/PROJECT_STATE.md) in full.** It is the authoritative handoff and covers:

- Two divergent copies of this workbook exist — a local `.xlsx` and a live Google Sheets file the user actually uses — and they are **not interchangeable**. §0 of that document explains which is authoritative and why.
- The Google Sheets copy has a **bound Apps Script project with a live `onEdit` trigger** managing Trip Status by code, not formula. Do not assume any column there is a plain formula without checking whether that script owns it.
- A near-miss happened last session editing that script's code editor via keyboard shortcut — see §7.3 for the specific shortcut to avoid and the safer editing discipline to follow.
- A prioritized, exact list of what's left to do is in §5 and §9.

## Working conventions established so far

- **Never** edit the `.xlsx` via a full `openpyxl` load+save round-trip — it has corrupted charts/drawings before. Edit via raw OOXML XML surgery: unzip → targeted string/regex replace on the specific `sheetN.xml`/`styles.xml`/`workbook.xml` → rezip → verify (fresh `openpyxl` load + zip listing + chart/drawing/sheetProtection counts) → promote.
- **Always** back up the live `.xlsx` to `backups/` (timestamped) before promoting a change.
- The workbook must stay Excel/Google-Sheets-formula-compatible: no `XLOOKUP`, no dynamic arrays, no Excel Tables.
- `.gitignore` excludes `*.xlsx` — neither workbook copy is git-tracked. Don't assume git history reflects the workbook's state; `docs/PROJECT_STATE.md` is the source of truth for that.
- When editing the live Google Sheets file via the Apps Script editor (script.google.com), take one action at a time with a screenshot check after each — do not batch keyboard shortcuts blindly. `Cmd+G` does not mean "go to line" there; it typed text into the file last time.

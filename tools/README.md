# tools/ — Module builders

## STRUCTIFY_Build_All_Modules.gs

Google Apps Script that generates the P-003…P-007 module Sheets in the correct
STRUCTIFY house layout, directly in the `STRUCTIFY_WIP` Drive folder.

### Run it

1. Go to [script.google.com](https://script.google.com) → **New project**
2. Paste `STRUCTIFY_Build_All_Modules.gs`, **Save**
3. Run **`buildAllModules`** and approve the Drive permission once
   → ~23 Sheets appear in `STRUCTIFY_WIP`

Helpers: `buildOne('M-021')` builds a single module (for testing);
`deleteExistingModuleCopies()` clears old copies before a rebuild.

The target folder is set at the top: `WIP_FOLDER_ID`.

### Module layout spec (matches M-000_START_HERE_SYSTEM)

The house layout, derived from the reference module:

- **Multi-tab** — one section per tab. Tab names in `UPPERCASE_WITH_UNDERSCORES`.
- **First tab = `OVERVIEW`** — a plain `SECTION | CONTENT` table (module, code,
  purpose, how to use, rule, status). No separate index/navigation tab.
- **Each tab starts at row 1 = the header row.** No title row, no subtitle, no
  tagline above the table.
- Header row 1: **bold, white text on black fill, frozen.** Data from row 2.
- Font: **Montserrat 10** throughout.
- **No zebra banding, no cell borders** — clean rows.
- Typically a narrow label column A + wider content column B; data modules
  (trackers, budgets) use as many columns as they need.
- Formulas use Apps Script's comma syntax (locale-independent).
- Data/tracker tabs get a block of blank working rows (`GROW_ROWS`) below the
  seeded rows: per-row formulas (NET, REMAINING, DIFFERENCE, DONE, %) are guarded
  with `IF(...="","",…)` and pre-filled across that range, and `SUM` totals cover
  it — so values a user adds later still calculate and roll into the totals.
- Column widths in the config are char-like and converted to pixels (`px_`), since
  `setColumnWidth` takes pixels; money columns are formatted through the total row.

### Modules generated

| Pack | Modules |
|------|---------|
| P-003 All You Need | M-001, M-002, M-003, M-004, M-005, M-045, M-046, M-037, M-038 |
| P-004 Money Control | M-021, M-022, M-023, M-024, M-025 |
| P-005 STRUCTIFY Kids | M-029, M-030, M-031, M-032 |
| P-006 Family OS | M-026, M-027, M-028 (+ shared M-029, M-031) |
| P-007 SHE-O OS | M-048, M-047 (+ shared M-046, M-021, M-037) |

Module content and names come from `STRUCTIFY_MODULE_REGISTRY`. Example rows are
placeholders to show structure — replace with real content per module.

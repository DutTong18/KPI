# checkStageProgress.ts — README

## Overview

This Office Script is a single merged workflow that runs inside Excel and does four things on every execution:

1. Builds and maintains a **KPI table** on the `SchedulerData` sheet, populated with live formula references to the `Stope Cadence` source sheet
2. Automatically adds any new stopes that hit **BLACK** or **RED** Zone Status to the KPI table
3. Adds a **date-stamped check column** showing whether each stope has progressed at least 2 stages forward since the last run (Y / N / ?)
4. Maintains a **summary block** at the top of the sheet with total stope counts by status

The script enforces a maximum of **7 date check columns** in the live table. When the limit is reached, the entire table is archived to a new sheet and the live table is reset to begin a fresh cycle.

---

## How It Works 

### 1. Configuration

All key behaviour is controlled by constants at the top of the script:

| Constant | Default | Purpose |
|---|---|---|
| `SourceSheet` | `"Stope Cadence"` | Sheet where source data lives |
| `TargetSheet` | `"SchedulerData"` | Sheet where the KPI table is built |
| `TargetTableName` | `"KPI"` | Name of the live KPI table |
| `StateSheetName` | `"_StageStateCache"` | Name of the hidden cache sheet |
| `TableAnchor` | `"A25"` | Top-left cell of the KPI table |
| `SummaryAnchor` | `"A1"` | Top-left cell of the summary block |
| `IndexOfIdStope` | `2` | Column C in source — stope ID |
| `IndexOfUser` | `7` | Column H in source — user |
| `IndexOfDesignStage` | `21` | Column V in source — design stage |
| `IndexOfSubProcess` | `22` | Column W in source — sub process |
| `IndexOfZoneStatus` | `34` | Column AI in source — zone status |
| `FilterValues` | `["BLACK", "RED"]` | Zone statuses that get tracked |
| `MIN_STEPS_FORWARD` | `2` | Minimum stage steps for a pass (Y) |
| `MAX_CHECKS` | `7` | Period checks per table before archive |

---

### 2. Source Data Read

The source is read directly from the `Stope Cadence` sheet using `getUsedRange()` — no named table is required. Row 0 is treated as the header and rows 1 onward as data. The script scans each row, picks out unique stope IDs whose Zone Status matches `FilterValues`, and tallies BLACK and RED counts.

---

### 3. KPI Table Structure

The KPI table has 5 base columns followed by alternating date check and notes columns:

| ID Stope | User | Design Stage | Sub Process | Zone Status | 12-05-2026 | 12-05-2026 Notes | 19-05-2026 | 19-05-2026 Notes | ... |
|---|---|---|---|---|---|---|---|---|---|
| ST-101 | DUT | Draft_Design | 50% | BLACK | Y | | N | ... | |
| ST-102 | JSM | Final_Design | 0% | RED | ? | | Y | ... | |

The **ID Stope** column holds plain values. The other four base columns use live **INDEX/MATCH** formulas:

```
=IFERROR(INDEX('Stope Cadence'!H:H, MATCH([@[ID Stope]], 'Stope Cadence'!C:C, 0)), "")
```

This means the User, Design Stage, Sub Process, and Zone Status always reflect the current source values without re-running the script. New rows automatically inherit the same formula pattern via the structured reference `[@[ID Stope]]`.

---

### 4. Adding New Stopes Automatically

On every run the script:

1. Builds a list of all BLACK / RED stope IDs from the source sheet
2. Reads the existing ID Stope column from the KPI table
3. Compares the two lists and appends any IDs that are missing

The 4 lookup formulas are applied to the new rows immediately, so they populate as soon as the worksheet recalculates.

Once a stope is added it stays in the table even if its status later changes — the lookup formula will reflect the new status in the Zone Status column.

---

### 5. Stage Order and Progress Detection

A **stage key** is built for each row as `"DesignStage::SubProcess"`. The script holds an ordered list of all known stage keys, and the position in this list defines progression:

```
Re-Sequenced::      → index 0
Not_Started::       → index 1
Geology_Review::    → index 2
Draft_Design::0%    → index 4
...
COMPLETE::          → index 34
```

The list is seeded on first run from a hardcoded `SEED_ORDER` array and stored in column E of the hidden cache sheet. After that, the persisted version is the source of truth, and any new stage keys discovered in the data are automatically appended to the end of the order.

For each row, the script:
1. Looks up the current stage index
2. Reads the previously saved stage index from the cache
3. Computes `delta = current - previous`
4. Assigns the result:

| Condition | Cell Value | Cell Colour |
|---|---|---|
| `delta >= 2` | **Y** | Green |
| `delta < 2` | **N** | Red |
| No previous state recorded | **?** | Yellow |
| Stage key unrecognised | **?** | Yellow |

---

### 6. Date-Stamped Check Columns

The new check column is named with today's date in `DD-MM-YYYY` format (e.g. `12-05-2026`). A paired blank `12-05-2026 Notes` column is added immediately after for manual comments.

If a column with the same date already exists (the script was run twice in one day), the suffix `(2)`, `(3)` etc. is appended automatically to keep names unique.

---

### 7. Rotation at 7 Checks

Before adding a new date column, the script counts existing period columns. If the count is already at `MAX_CHECKS` (7):

1. The entire KPI table (values, not formulas) is copied to a new archive sheet named `KPI_Archive_DD-MM-YYYY`
2. All date and notes columns are removed from the live KPI table (the 5 base columns remain)
3. The new date column is then added as the first of a new cycle

The hidden state cache **persists across rotation**, so the first check on a new cycle still measures real progress against where each stope was at the end of the previous cycle. If multiple archives happen on the same day, an incrementing suffix `_2`, `_3` is added to the archive sheet name.

---

### 8. Summary Block

A 3-row block is written above the KPI table at `SummaryAnchor` (`A1` by default):

| Total Stopes | BLACK | RED |
|---|---|---|
| 24 | 6 | 9 |
| *Last updated: 12/05/2026 09:15:00* | | |

The values overwrite in place each run so the block does not grow over time.

---

### 9. State Persistence — The Hidden Cache Sheet

The script maintains a sheet named `_StageStateCache` set to `veryHidden`. Layout:

| Columns A–C | Column E |
|---|---|
| Stope state — `StopeID / DesignStage / SubProcess` | Stage order list (one key per row) |

Both halves are read at the start of each run and re-written at the end.

---

---

## Code Walkthrough — Section by Section

The script is organised into a `main()` function containing 13 numbered steps, followed by helper functions. The two standalone helpers at the bottom of the file (`columnLetterFromIndex` and `rowFromAnchor`) live outside `main()` because they take no script state and could be reused independently.

### Top of `main()` — Configuration Block (lines 3–48)

Three categories of constants are defined upfront so all behaviour is controllable from a single location:

- **Sheet and table names** (`SourceSheet`, `TargetSheet`, `TargetTableName`, `StateSheetName`, `TableAnchor`, `SummaryAnchor`) — what the script reads from and writes to
- **Source column indices** (`IndexOfIdStope`, `IndexOfUser`, `IndexOfDesignStage`, `IndexOfSubProcess`, `IndexOfZoneStatus`) — 0-based positions of the fields the script needs from the source sheet
- **Behaviour switches** (`BASE_HEADERS`, `FilterValues`, `MIN_STEPS_FORWARD`, `MAX_CHECKS`, `SEED_ORDER`) — control which statuses are tracked, what counts as progress, when to rotate, and the initial stage progression

The block ends by building `baseHeader`, today's date as `DD-MM-YYYY`, which will be used as the new period column name.

---

### STEP 1 — Sheet Sanity (lines 50–55)

Verifies that the source and target sheets actually exist. If either is missing, the script logs a message and exits early. This is a fast fail-safe before any expensive work begins.

---

### STEP 2 — Source Read and Filter (lines 57–84)

Reads the entire used range from the `Stope Cadence` sheet. Row 0 is treated as headers and skipped. For each subsequent row the script:

- Reads the stope ID from column C and the zone status from column AI
- Skips the row if the ID is blank, the zone is not BLACK or RED, or the ID has already been seen
- Adds qualifying IDs to `blackRedIds` and increments the `countBlack` / `countRed` tallies

A `Set` is used to deduplicate IDs efficiently. If no qualifying stopes are found, the script exits — there is nothing to track.

---

### STEP 3 — Get or Create the KPI Table (lines 86–109)

Attempts to fetch the existing KPI table by name. If it does not exist, the script:

1. Calls `T_sheet.addTable()` with a 5-column header range starting at `TableAnchor` — the end column letter is computed dynamically using `columnLetterFromIndex(BASE_HEADERS.length)` so the table width matches the number of base headers
2. Sets the table name to `TargetTableName`
3. Writes the headers from `BASE_HEADERS`
4. Seeds the first data row with the first stope ID via `buildBaseRow()`, then applies the INDEX/MATCH lookup formulas to columns 2-5 via `applyLookupFormulasToRow()`
5. Adds any remaining stopes via `addRows()`, then applies the formulas to each of the new rows

The placeholder row exists because every `addTable` call requires at least one row — it cannot be a header-only table.

---

### STEP 4 — Rotation Check (lines 111–120)

Counts the period columns in the existing table using `countPeriodColumns()` (which excludes the 5 base columns and any column ending in `" Notes"`). If the count is already at `MAX_CHECKS`:

1. `archivePeriodColumns()` copies the table's values to a new sheet named `KPI_Archive_DD-MM-YYYY`
2. `removePeriodColumns()` deletes every column except the 5 base ones
3. The headers list is re-read so subsequent steps see the cleaned table

---

### STEP 5 — Auto-Add New Stopes (lines 122–146)

This step is what makes the script continuous across runs:

1. Reads the current ID Stope column and collects all existing IDs into a `Set`
2. Filters `blackRedIds` to find IDs that exist in the source but not yet in the table
3. If any are new, `addRows()` appends them with placeholder values, then `applyLookupFormulasToRow()` writes the INDEX/MATCH formulas into the User / Stage / Sub / Zone cells for each new row

Stopes that were previously tracked but no longer match BLACK or RED status are left in place — their lookup formulas will simply update to reflect their new status.

---

### STEP 6 — Load Cache (lines 148–155)

Calls three cache helpers:

- `getStateSheet()` returns the hidden `_StageStateCache` sheet (creates it if needed)
- `readSavedState()` returns a dictionary of `StopeID → {designStage, subProcess}` from columns A-C
- `readStageOrder()` returns the persisted stage progression list from column E, falling back to `SEED_ORDER` if empty

A `stageIdx` dictionary is then built for fast `O(1)` lookup of any stage key's position in the order.

---

### STEP 7 — Read KPI Table and Discover New Stages (lines 157–175)

Reads all values from the live KPI table. Because the User / Stage / Sub / Zone columns contain INDEX/MATCH formulas, this captures the **calculated values** — i.e. the current source-sheet state for each stope.

Line 159 is a no-op nudge that forces Excel to recalculate before the read — it sets the row height to its current value, which triggers a recalculation without changing the visible layout.

The script then iterates every row, builds a stage key for the current Design Stage / Sub Process combination, and registers any keys not already in the order list. New keys are appended to the end and `orderChanged` is flagged so they get persisted at the end of the run.

---

### STEP 8 — Evaluate Y / N / ? (lines 177–203)

For each row in the table:

1. Build the current stage key and look up its index
2. Fetch the previously saved state for this stope ID from the cache
3. If no previous state exists, the cell value is `?` (first time we have seen this stope)
4. Otherwise, compute the previous stage index, calculate `delta = currI - prevI`, and assign:
   - `Y` if `delta >= MIN_STEPS_FORWARD` (default: 2)
   - `N` otherwise
5. Push the result to `cellValues` and update `passCount` for the final log message

The current state for every stope is collected in `currentStateList` so it can be persisted at the end.

---

### STEP 9 — Generate Unique Date Header (lines 205–212)

Uses `baseHeader` (today's date) and checks if a column with that name already exists. If so, appends `(2)`, `(3)`, etc. until a unique name is found. This handles the case where the script is run more than once on the same day.

---

### STEP 10 — Append Date Column with Y / N / ? Values (lines 214–242)

1. Adds a new column to the right of the table via `addColumn(-1)`
2. Sets its name to the unique date header
3. Re-reads the actual row count from the live range (in case rows changed during the run) and pads `cellValues` with `?` if necessary
4. Writes all values in one `setValues()` call
5. Loops through every cell to apply colour formatting based on its value:
   - **Y** → light green fill (`#C6EFCE`), dark green text (`#276221`)
   - **N** → light red fill (`#FFC7CE`), dark red text (`#9C0006`)
   - **?** → light yellow fill (`#FFEB9C`), dark orange text (`#9C5700`)
6. Centre-aligns each cell horizontally

---

### STEP 11 — Append Paired Notes Column (lines 244–248)

Adds a second new column immediately after the date column, named `<DateHeader> Notes`. Fills it with empty strings so the cells are explicitly blank and ready for user comments.

---

### STEP 12 — Update Summary Block (line 251)

Calls `writeSummaryBlock()` to overwrite the 3-row block at `SummaryAnchor` with the current total, BLACK, and RED counts plus a "Last updated" timestamp.

---

### STEP 13 — Persist State (lines 253–257)

1. `writeSavedState()` overwrites columns A-C of the cache sheet with the current state of every stope so it becomes the previous state on the next run
2. `writeStageOrder()` updates column E only if new stage keys were discovered

A final log message summarises the run for the Office Scripts console.

---

## Helper Functions Inside `main()`

These are nested inside `main()` so they can directly reference the configuration constants without needing parameters for every value.

### `buildBaseRow(id)` (lines 261–264)
Returns an array `[id, "", "", "", ""]` — the ID is set as a real value, and the other four cells are placeholders that will be overwritten with lookup formulas. The empty strings exist because `addRows()` requires a value array and rejects formulas directly.

### `applyLookupFormulasToRow(rowRange)` (lines 266–277)
Loops through columns 2-5 of a single table row and sets each one to an INDEX/MATCH formula that looks up the appropriate source column by stope ID.

### `buildLookupFormula(sourceColIdx)` (lines 279–283)
Constructs the actual formula string for a given source column. The formula uses a structured reference `[@[ID Stope]]` so it works identically on every row, looking up the value in the source sheet's ID column and returning the corresponding cell from the requested column. Wrapped in `IFERROR(..., "")` so unmatched IDs show as blank rather than `#N/A`.

Example output for the User column:
```
=IFERROR(INDEX('Stope Cadence'!H:H,MATCH([@[ID Stope]],'Stope Cadence'!C:C,0)),"")
```

### `countPeriodColumns(headers)` (lines 285–287)
Returns the number of columns whose name is neither in `BASE_HEADERS` nor ends in `" Notes"`. This is how the rotation check decides whether the 7-check limit has been reached.

### `removePeriodColumns(table)` (lines 289–298)
Iterates the table's columns from right to left (so indices stay stable as items are deleted) and removes any column whose name is not in `BASE_HEADERS`. This wipes both date columns and their paired Notes columns in one pass.

### `archivePeriodColumns(workbook, table, runDate)` (lines 300–318)
1. Builds an archive sheet name in the form `KPI_Archive_DD-MM-YYYY`
2. If that name is taken, appends `_2`, `_3`, etc. until a free name is found
3. Creates the archive sheet
4. Reads all values from the live table (headers + data) and writes them to the archive sheet starting at A1

Values are captured rather than formulas so the archive remains accurate even if the source sheet is later modified.

### `stageKey(designStage, subProcess)` (lines 320–322)
Joins a Design Stage and Sub Process into a single key string using `::` as the separator, trimming whitespace from both. Used everywhere the script needs to identify or compare a stage.

### `getStateSheet()` (lines 324–333)
Returns the hidden cache sheet. If it does not exist, creates it, sets it to `veryHidden`, and writes the two header rows (`StopeID / DesignStage / SubProcess` in A1:C1 and `StageOrder` in E1).

### `readStageOrder(sheet)` (lines 335–347)
Reads column E (index 4) row by row, skipping the header. Returns a list of stage keys, or a copy of `SEED_ORDER` if the column is empty.

### `writeStageOrder(sheet, order)` (lines 349–353)
Clears column E and writes the supplied order list back to it. The +20 buffer in the clear range guarantees no stale rows remain even if the new list is shorter than the old one.

### `readSavedState(sheet)` (lines 355–367)
Reads columns A-C of the cache sheet and returns a dictionary keyed by stope ID, where each value contains the saved Design Stage and Sub Process. This is the data structure used to compare against the current state in Step 8.

### `writeSavedState(sheet, list)` (lines 369–374)
Clears the A-C data range and writes the current state list back to it. Always overwrites; never appends.

### `writeSummaryBlock(sheet, anchor, total, black, red, runDate)` (lines 376–410)
Writes three rows starting at `anchor`:
- Row 1: header labels with coloured fills (blue / black / red backgrounds, contrasting font colours, bold, centre-aligned)
- Row 2: numeric counts with matching coloured fills
- Row 3: italic timestamp showing when the script last ran

The helper uses a nested arrow function `cell(r, c)` to translate row / column offsets into absolute cell references relative to the anchor.

---

## Helper Functions Outside `main()`

These two functions live at the file level because they are pure utilities that don't depend on any script state.

### `columnLetterFromIndex(colNum)` (lines 414–422)
Converts a 1-based column number to its Excel column letter (1 → A, 2 → B, 27 → AA, 52 → AZ, 53 → BA, etc.) using the standard base-26 algorithm with the wrinkle that Excel columns are 1-indexed (no "column zero"). Used when building INDEX/MATCH formulas to translate source column indices into letter references.

### `rowFromAnchor(anchor)` (lines 425–428)
Extracts the row number from a cell reference like `"A25"` using a simple regex, returning `25` as a number. Used when constructing the initial table range in Step 3.

---

## Potential Errors and How to Fix Them

### Sheet "Stope Cadence" or "SchedulerData" not found
**Cause:** A sheet has been renamed or deleted.
**Fix:** Update `SourceSheet` or `TargetSheet` at the top of the script to match the current sheet names. Names are case-sensitive.

---

### All stopes show `?` on first run
**Cause:** Expected behaviour. The cache has no previous state to compare against on the very first run. The current state is saved on this run, so the second run will produce Y / N results.

---

### Stope shows `?` even though progress was made
**Cause:** The Design Stage or Sub Process value in the source cell does not exactly match any entry in the stage order. Common causes include leading or trailing spaces, different capitalisation, or a brand new stage value.
**Fix:** Run the script once — new stage keys are auto-registered at the end of the order list. If the new stage should sit elsewhere in the progression rather than at the end, manually re-order the keys in column E of the `_StageStateCache` sheet.

---

### New stage appears at the end of the order but logically belongs in the middle
**Cause:** Auto-registration always appends new keys to the end of the order. A stope that moves into this new stage will only show a delta of +1 if the previous stage was the second-to-last known key.
**Fix:** Open the hidden `_StageStateCache` sheet (Format → Hide/Unhide → Unhide Sheet), find column E, and drag the new key into its correct position in the progression.

---

### Lookup columns show blank or `#N/A` after script runs
**Cause:** The stope ID in the KPI table does not exist in the source sheet's ID column, or the source sheet has been renamed.
**Fix:** Check that the stope ID is present in column C of `Stope Cadence`. If the sheet was renamed, the formulas need rebuilding — the easiest path is to delete the KPI table and rerun the script to regenerate it.

---

### Period column not added — script exits with "No BLACK or RED stopes found in source"
**Cause:** No rows in the source sheet currently have a Zone Status of BLACK or RED.
**Fix:** This is expected when there is nothing to track. No action needed — the script will run normally as soon as a stope hits BLACK or RED status.

---

### Same-day re-runs produce columns like `12-05-2026 (2)`, `12-05-2026 (3)`
**Cause:** Expected behaviour. To prevent name collisions, a numeric suffix is automatically added when the script runs more than once on the same date.
**Fix:** None required. If the duplicate runs were unintentional, delete the extra columns manually before the next run.

---

### Rotation happens unexpectedly
**Cause:** The number of period columns has reached `MAX_CHECKS` (7) and the script archived the table.
**Fix:** This is expected behaviour. Find the new sheet named `KPI_Archive_DD-MM-YYYY` for the previous cycle's data. To change the limit, edit `MAX_CHECKS` at the top of the script.

---

### Archive sheet has formula errors instead of values
**Cause:** Should not happen — the archive function reads `getValues()` which captures calculated results. If it does, the source sheet may have been renamed before the archive was created.
**Fix:** Manually re-create the archive by copying the values from the live table before running the next rotation.

---

### `_StageStateCache` sheet is accidentally deleted or visible
**Cause:** A user manually deleted or unhid the cache sheet.
**Fix:** The script will recreate the sheet on the next run, but all previous state will be lost — all stopes will show `?` for that run only. The sheet is automatically reset to `veryHidden`.

---

### Duplicate stope IDs in the source
**Cause:** Two or more rows in the source share the same stope ID.
**Fix:** The lookup formulas return the value from the **first** matching row. If this is wrong, fix the duplicate IDs in the source.

---

### Period columns accumulating with wrong numbering after a column is manually deleted
**Cause:** The script auto-generates unique date headers and only checks for direct name conflicts. If a date column is manually deleted, the next run on that same date will reuse the date header (without a suffix), which may cause unexpected gaps.
**Fix:** Avoid manually deleting period columns. If a column must be removed, do not run the script again on the same date.

---

## Manual Operations

### Reset the cache without affecting the KPI table
1. Unhide `_StageStateCache` via Format → Hide/Unhide → Unhide Sheet
2. Delete all rows from A2 downward (keep the header row)
3. Optionally clear column E to reset the stage order back to the seed
4. Hide the sheet again

The next run will treat every stope as "first seen" and show `?` for all.

---

### Insert a new stage in the middle of the order
1. Unhide `_StageStateCache`
2. In column E, right-click the row where the new stage should sit and choose Insert
3. Type the new stage key in the format `"DesignStage::SubProcess"` (or `"DesignStage::"` if no sub process)
4. Hide the sheet again

The new stage takes effect on the next run.

---

### Force a manual rotation before hitting the limit
1. Run the script normally until it has the desired number of checks
2. Manually duplicate the `SchedulerData` sheet to keep an archive
3. Delete all period and notes columns from the live KPI table
4. The next run will start a fresh cycle

---

## Summary of Files Created and Maintained

| Sheet / Table | Visibility | Purpose |
|---|---|---|
| `KPI` table on `SchedulerData` | Visible | Live working table |
| `_StageStateCache` sheet | Very hidden | Stope state + stage order |
| `KPI_Archive_DD-MM-YYYY` sheet | Visible | Historical archive (created on rotation) |

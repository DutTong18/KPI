# KPI
# Current Stage: TESTING
# checkStageProgress.ts — README

## Overview

This Office Script runs inside Excel and tracks whether each stope in the **KPI** table has progressed far enough through the design workflow between script runs. Each time it is run it adds two new columns to the KPI table — a **Period** result column (`Y` / `N` / `?`) and a blank **Notes** column for manual comments. The previous state of every stope is saved in a hidden sheet so that progress can be measured on the next run.

---

## How It Works — Step by Step

### 1. Configuration
At the top of the script, a set of constants controls all key behaviour. These are the only values that should ever need changing:

| Constant | Default | Purpose |
|---|---|---|
| `TargetTableName` | `"KPI"` | Name of the Excel table being written to |
| `StateSheetName` | `"_StageStateCache"` | Name of the hidden sheet used to persist state |
| `BaseColumnHeader` | `"Period"` | Prefix for new columns — produces *Period 1*, *Period 2*, etc. |
| `COL_ID_STOPE` | `0` | Column index (0-based) of the Stope ID in the KPI table |
| `COL_DESIGN_STAGE` | `2` | Column index of the Design Stage |
| `COL_SUB_PROCESS` | `3` | Column index of the Sub Process |
| `MIN_STEPS_FORWARD` | `2` | Minimum number of stage steps forward required for a **Y** (pass) |

---

### 2. Stage Ordering
Every valid design stage and sub-process combination is listed in `STAGE_ORDER` as a flat ordered array. Each entry uses the format:

```
"DesignStage::SubProcess"
```

If a stage has no sub-process it is stored as `"DesignStage::"`. The position in the array is the progression index — the higher the index, the further along in the workflow.

**Current order (index 0 → 34):**

| Index | Stage Key |
|---|---|
| 0 | Re-Sequenced:: |
| 1 | Not_Started:: |
| 2 | Geology_Review:: |
| 3 | Draft_Design:: |
| 4 | Draft_Design::0% |
| 5 | Draft_Design::25% |
| 6 | Draft_Design::50% |
| 7 | Draft_Design::75% |
| 8 | External_Review:: |
| 9 | External_Review::0% |
| 10 | External_Review::25% |
| 11 | External_Review::50% |
| 12 | External_Review::75% |
| 13 | Concept:: |
| 14 | Shape_Review:: |
| 15 | Shape_Review::Wait on Meeting |
| 16 | Shape_Review::Updates Post ISR |
| 17 | Final_Design:: |
| 18 | Final_Design::0% |
| 19 | Final_Design::25% |
| 20 | Final_Design::50% |
| 21 | Final_Design::75% |
| 22 | Peer_Review:: |
| 23 | Peer_Review::Wait on Meeting |
| 24 | Peer_Review::Updates Post PR |
| 25 | Direction_Meeting:: |
| 26 | Direction_Meeting::Wait on Meeting |
| 27 | Direction_Meeting::Updates Post meeting |
| 28 | IFR:: |
| 29 | 02_Technical:: |
| 30 | 03_Operations:: |
| 31 | 04_Superintendent:: |
| 32 | 05_Manager:: |
| 33 | 06_Upload:: |
| 34 | COMPLETE:: |

---

### 3. State Persistence — The Hidden Sheet
The script creates and maintains a sheet named `_StageStateCache` which is set to `veryHidden` (not visible in the Excel sheet tab bar). This sheet stores one row per stope with three columns:

```
| StopeID | DesignStage | SubProcess |
```

On every run the script reads this sheet to get the *previous* state, compares it to the *current* state in the KPI table, and then overwrites the sheet with the current state so it becomes the previous state on the next run.

---

### 4. Progress Evaluation
For each row in the KPI table the script:

1. Reads the current `DesignStage` and `SubProcess` and builds a stage key
2. Looks up its index in `STAGE_ORDER`
3. Retrieves the previous state for the same Stope ID from `_StageStateCache`
4. Calculates `delta = currentIndex - previousIndex`
5. Assigns a result:

| Condition | Cell Value | Cell Colour |
|---|---|---|
| `delta >= 2` | **Y** | Green |
| `delta < 2` | **N** | Red |
| No previous state found | **NEWLY RED/BLACK** | Yellow |
| Stage key not in `STAGE_ORDER` | **?** | Yellow |

---

### 5. Column Output
Two columns are appended to the KPI table on each run:

- **Period N** — contains Y / N / ? / NEWLY RED/BLACK with colour formatting
- **Period N Notes** — blank, for manual comments

The counter increments automatically so existing columns are never overwritten.

---

## Potential Errors and How to Fix Them

### Table "KPI" not found
**Cause:** The KPI table has been renamed or does not exist on the target sheet.  
**Fix:** Check the table name in Excel (*Table Design* tab → *Table Name*) and update `TargetTableName` at the top of the script to match exactly (case-sensitive).

---

### Stage shows `?` instead of Y or N
**Cause:** The value in the Design Stage or Sub Process cell does not exactly match any entry in `STAGE_ORDER`. Common causes include:
- Extra leading or trailing spaces in the cell
- Different capitalisation (e.g. `"draft_design"` vs `"Draft_Design"`)
- A new sub-process value that has not been added to the list yet

**Fix:** Open the console log output in Office Scripts to see the unrecognised key. Add the missing key to `STAGE_ORDER` in the correct position, or correct the data in the source sheet.

---

### All stopes show `NEWLY RED/BLACK` on first run
**Cause:** This is expected behaviour. The first time the script runs there is no previous state in `_StageStateCache` to compare against. The current state is saved on this first run, so the second run will produce Y / N results.

---

### Wrong column indices (Y / N assigned to wrong stopes)
**Cause:** The KPI table columns have been reordered, or the `moveColumns.ts` script was changed so the 5 columns are in a different order.  
**Fix:** Update `COL_ID_STOPE`, `COL_DESIGN_STAGE`, and `COL_SUB_PROCESS` at the top of the script to match the current 0-based column positions in the KPI table.

---

### `_StageStateCache` sheet is accidentally deleted or visible
**Cause:** A user manually deleted or unhid the cache sheet.  
**Fix:** The script will automatically recreate the sheet on the next run, but all previous state will be lost — all stopes will show `?` for that run. The sheet will be set back to `veryHidden` automatically.

---

### Stope ID is blank or duplicated in the KPI table
**Cause:** If a row has an empty Stope ID, it will be saved to the cache under an empty string key and may overwrite other blank-ID rows. If two rows share the same Stope ID, the second row will overwrite the first in the cache.  
**Fix:** Ensure every row in the KPI table has a unique, non-empty Stope ID before running the script.

---

### Script runs but nothing is added to the table
**Cause:** The KPI table has no data rows (only a header row).  
**Fix:** The script exits early with the message `"KPI table has no data rows."` Check that the `moveColumns.ts` script has been run first to populate the table.

---

### Period columns accumulating with wrong numbering after a column is manually deleted
**Cause:** The script finds the next available `Period N` number by scanning existing column headers. If *Period 2* is manually deleted but *Period 3* exists, the next run will add *Period 2* again (filling the gap) rather than *Period 4*.  
**Fix:** Do not manually delete Period columns. If a column must be removed, also rename any remaining Period columns to maintain a clean sequence.

---

## Adding a New Stage to the List
If a new design stage or sub-process is introduced:

1. Open `checkStageProgress.ts` in Office Scripts
2. Find the `STAGE_ORDER` array
3. Insert the new entry `"DesignStage::SubProcess"` at the correct position in the progression
4. Save and run — all future comparisons will include the new stage

> NOTE: Inserting a new stage shifts the index of every entry below it. Any stopes that were at those stages when the cache was last written will have their delta recalculated based on the new indices on the next run, which may produce unexpected Y / N results for a single run after the change.

---

## Future Changes
Changes will be made to;
1. correct errors in sub process checking (append process and sub process -> check with list of stages as strings)
3. reference or add new stope status automatically
4. correct errors with Draft_Design and sub processes
5. correction to header row colour
6. generate new sheet of historic stage
7. manage completed stopes (marked as IFR Need to ask )
8. changes made to script to allow new stopes to be added withought column error
9. **change logic to check individual stage, to prevent table growing laterally indefinitely**
10. changes will be made to save completed stope data / Graphs and how are they presented
11. User percentages and presentation (leaderboard?? [may be polarising])
12. Deployment!
13. continued updates 

## Nice to haves
1.  standardise naming conventions
2.  connecto to powerBI
3.  


> NOTE: 

## If there are any issues found that have not been adressed here please feel free to reach out!


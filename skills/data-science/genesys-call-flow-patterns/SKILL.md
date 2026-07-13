---
name: genesys-call-flow-patterns
description: "Common patterns and gotchas when processing Genesys call center exports for Switchboard and Non-Atty flows."
version: 1.0.0
author: Hermes Agent
license: MIT
platforms: [linux, macos, windows]
metadata:
  hermes:
    tags: [genesys, call-center, data-validation]
    category: data-science
---

# Genesys Call Flow Patterns

This skill documents the common patterns and gotchas when processing Genesys call center exports, particularly for Switchboard and Non-Atty call flows.

## Outcome Classification Logic

## Unified Classification

All flow types (Switchboard, Non-ATTY, ATTY, LRS Spanish, Bar Examiners) use the same formula:

| Category | Condition |
|----------|-----------|
| **Successful** | `Outcome Success == 1` AND `Abandoned != 'YES'` |
| **Failed** | `Outcome Success != 1` AND `Abandoned != 'YES'` |
| **Abandoned** | `Abandoned == 'YES'` (regardless of Outcome Success value) |

**PITFALL:** Outcome Success can be `2` for non-abandoned calls. Simply counting `Outcome Success == 1` overcounts if abandoned=YES rows also have Outcome Success = 1. Always exclude abandoned rows first, then split by Outcome Success value.

**PITFALL:** `Failed = Total - Successful - Abandoned` does NOT work — some rows appear in both Successful AND Abandoned categories.

## Duration Parsing
- Format: ` 00:04:12.710` (note leading space)
- Trim whitespace before parsing with `pd.to_timedelta()`
- Filter outliers: durations > 24 hours (86,400 seconds) are data errors

### Call Paths
- Queue column contains semicolon-separated values: `SWITCHBOARD; LRS`
- Parse by splitting on `;`, stripping whitespace, joining with ` → `
- Empty queue values indicate calls that never reached a queue (voicemail, callback, etc.)

## Verification Checklist
- [ ] Total calls in Summary matches row count
- [ ] Success + failed + abandon == total (exact match for conditional logic)
- [ ] Duration is realistic (< 24 hours for average)
- [ ] Call Paths sheet has at least one entry

## Consolidating Monthly Files

When combining monthly call flow exports into a single report for delivery:

1. **Load each month's file** and extract Summary + Call Paths sheets
2. **Build consolidated Summary** with columns: `Metric | April | May | June | Total`
   - Count-based metrics (Total Calls, Successful, Failed, Abandoned) are **summed**
   - Percentage metrics (Success %, Failure %, Abandon %) are **weighted averages** by Total Calls
   - Duration metrics use **weighted average** by call count (not simple average)
   - Duration metric name is **Average Call Time (MM:SS)** (not "Average Duration")
   - Duration format is `mm:ss` (no milliseconds, no milliseconds.ms)
3. **Build Call Paths** by concatenating all months and grouping by Call Path + summing counts
4. **Deliver as single Excel** with:
   - `Summary` sheet — consolidated monthly + totals
   - Individual monthly sheets named `April`, `May`, `June` with separate Call Path tables
   - Each monthly sheet has a **column bar chart** showing call counts by path
5. **Format**: blue header row (`4472C4`), thin borders, Total column highlighted with `D9E2F3` fill
6. **Never include** merged Call Paths sheet in final deliverable — only individual monthly sheets

### Chart Requirements for Monthly Sheets
- Chart type: `BarChart()` (horizontal 2D bar chart, NOT columns)
- Style: `chart.style = 1` (plain 2D, no 3D shading)
- Data references must start from **row 2** (skip header row) — if you include row 1, the first bar shows the column header (e.g., "Queue Path") instead of the first actual call path (e.g., "Switchboard")
- Labels: `Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)`
- Data: `Reference(ws, min_col=2, min_row=2, max_row=ws.max_row)`
- `chart.add_data(data, titles_from_data=False)` then `chart.set_categories(labels)`
- Position chart below the data table (e.g., at `A{max_row+5}`)
- Chart dimensions: ~15cm wide × 12cm tall

### Reference Process
See `scripts/consolidate_call_flow_reports.py` for the complete implementation pattern (reads from pre-consolidated Excel files).

### Auto-discovering Categories from CSV Files (v2 pattern)
When source files are raw CSVs (not pre-consolidated Excel), use a filename-based category discovery approach instead of hard-coding:

1. **Scan CSV_DIR** for all `.csv` files.
2. **Extract month name** (case-insensitive "april|may|june") from each filename to determine the month.
3. **Derive category** from the prefix before the month name, stripping `doc_[hex]_` prefixes and trailing hyphens/spaces.
4. **Normalize typos** — common Genesys export typos include "Switchbaord" (b-a-o vs b-o-a-r-d). Use `re.sub(r'(?i)switchboa?r[dh]|switchba?o?r?d', 'Switchboard', category)` to normalize.
5. **Group by category**, then process all 5 categories (Bar Examiners, ATTY, Switchboard, Non-ATTY, LRS Spanish) in a single run.
6. **Output**: one `.xlsx` per category with Summary + April/May/June tabs (each monthly tab has a bar chart).

**PITFALL:** Always strip trailing hyphens and spaces after normalization — "Switchbaord-" becomes "Switchboard" (not "Switchboard ") which would create a duplicate category entry.

**PITFALL:** The regex `(?i)switchboa?r[dh]|switchba?o?r?d` is fragile — if Genesys changes the export format or adds new categories, this pattern will miss them. Test against filenames before running.

**Example filenames and their extracted categories:**
```
doc_b696083cdb96_Switchbaord-April.csv  →  Switchboard (April)
doc_a7d02a33e3f6_Bar Examiners June.csv  →  Bar Examiners (June)
doc_74e69e181d95_ATTY June.csv           →  ATTY (June)
```

See `scripts/consolidate_all_from_csv.py` for the full implementation.

## Project Working Directory

The canonical project directory is `/Users/trey/.hermes/projects/Genesys Call Analysis/`. All monthly source files, consolidated outputs, and scripts live here.

## Templates
- `templates/validate_call_flow_output.py` - validation script template

## References
- `references/lrs-spanish-logic.md` - Historical rationale (all flows now use unified approach)
- `references/atty-logic.md` - Historical rationale (all flows now use unified approach)

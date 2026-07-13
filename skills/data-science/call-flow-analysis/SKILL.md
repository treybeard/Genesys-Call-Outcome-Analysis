---
name: call-flow-analysis
description: "Analyze Genesys call center CSV exports to generate summary metrics, queue flow paths, and bar charts."
version: 4.0.0
author: Hermes Agent
license: MIT
platforms: [linux, macos, windows]
metadata:
  hermes:
    tags: [genesys, call-center, data-analysis, excel]
    category: data-science
---

# Call Flow Analysis

Analyze Genesys call center CSV exports to generate summary metrics, queue flow paths, and bar charts. Produces Excel workbooks with Summary (metrics), Call Paths (monthly data), and bar charts per monthly tab.

## When to Use

- User provides a Genesys CSV with outcome columns (`Outcome Success`, `Outcome Failure`, `Abandoned`)
- User wants summary metrics: total calls, success/failure/abandon counts and rates, average duration
- User wants ordered queue paths preserving Queue column sequence
- Output must be in Excel format with Summary tab + monthly tabs with bar charts

**Don't use for**: Raw CSV inspection without metric extraction or non-Genesys data.

### Variant: LRS Spanish Exports

LRS Spanish flows use a different outcome classification where:
- **Successful** = `Outcome Success == 1` AND `Abandoned != 'YES'`
- **Failed** = `Outcome Success != 1` AND `Abandoned != 'YES'`
- **Abandoned** = `Abandoned == 'YES'` (any outcome)

See `genesys-call-flow-patterns/references/lrs-spanish-logic.md` for details.

## Core Workflow

### Direct CSV → Consolidated XLSX (preferred)

Use `csv_to_consolidated.py` at `/Users/trey/.hermes/projects/Genesys Call Analysis/csv_to_consolidated.py` to process raw CSVs directly into a consolidated XLSX per batch.

Consolidated XLSX structure:
- **Summary tab**: Columns = Metric, Apr, May, Jun, Total. Rows = 7 metrics (Total Calls, Successful, Failed, Abandoned, Success %, Abandon %, Average Duration).
  - **Percentages stored as raw decimals** (e.g., `0.8961`) with Excel number format `0.00%` — not formatted strings like `\"89.61%\"`
  - **Average Duration stored as Excel time fraction** (seconds / 86400) with number format `h:mm:ss` and **right alignment**
  - Numeric metrics (Total Calls, Successful, Failed, Abandoned) summed for Total column
  - Percentages averaged for Total column
  - Average Duration averaged for Total column
  - **No Call Path or Month rows in Summary**
- **Individual month tabs** (April, May, June): Each contains:
  - Call paths data: "Path" and "Count" columns, sorted by count descending
  - **2D horizontal bar chart** titled "Call Count / Call Path" anchored at D2, spanning columns D→R, with height matching the data table's row range
  - Blue header row (`#4472C4` background, white bold text)
  - All cells have thin borders

#### Chart Generation Method (CRITICAL — Do NOT use openpyxl BarChart)

**openpyxl's `BarChart` class produces Excel-incompatible chart XML that Excel silently rejects** (charts render as invisible). This has been verified through multiple debugging sessions — EMU post-processing, anchor fixes, and `strRef`/`numRef` patches all fail because the root XML structure is wrong.

**Correct approach: hand-craft chart XML and drawing XML, then rebuild the xlsx as a ZIP.**

Chart XML (`xl/charts/chart*.xml`) must contain:
- `c:strRef` with `c:f` formula + `c:strCache` containing `<c:ptCount>` and one `<c:pt idx="N"><c:v>path_label</c:v></c:pt>` for every row (all categories embedded inline, not as cell references)
- `c:numRef` with `c:f` formula + `c:numCache` containing `<c:formatCode>General</c:formatCode>`, `<c:ptCount>`, and one `<c:pt idx="N"><c:v>count_value</c:v></c:pt>` for every row
- Namespace declarations on `chartSpace`: `xmlns:c`, `xmlns:a`, `xmlns:r`, `xmlns:c16r2`
- Required elements: `c:style val="2"`, `c:date1904 val="0"`, `c:lang val="en-US"`, `c:roundedCorners val="1"`, `c:varyColors val="1"`, `c:extLst` with `c16:uniqueId`
- Axes: `c:catAx` (with title "Call Path") + `c:valAx` (with title "Count" + `c:majorGridlines`)
- Layout: `c:legend` with `legendPos val="r"`, `c:plotVisOnly val="1"`, `c:dispBlanksAs val="gap"`

Drawing XML (`xl/drawings/drawing*.xml`) must use:
- `xdr:` namespace prefix (NOT default namespace) — the three keys are `xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"`, `xmlns:a`, `xmlns:c`, `xmlns:r`, `xmlns:a16`
- `xdr:oneCellAnchor` with `xdr:col>3</xdr:col>` (D column) and `xdr:row>1</xdr:row>` (row 2)
- `xdr:ext cx="540000000000" cy="{num_paths × 68580000000}"` — plain integers, NEVER scientific notation
- `a:extLst` → `a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"` → `a16:creationId` (Office 2016+ compatibility)
- `a:graphic` → `a:graphicData uri="chart"` → `c:chart r:id="rId1"`

Relationships:
- Drawing rels: `xl/drawings/_rels/drawing*.xml.rels` linking to `/xl/charts/chart*.xml`
- Worksheet rels: `xl/worksheets/_rels/sheet{i}.xml.rels` linking to `/xl/drawings/drawing*.xml`

Rebuild workflow:
1. Save workbook with openpyxl (for data cells only, no charts)
2. Re-open as ZIP archive
3. Remove all `xl/charts/chart*.xml`, `xl/drawings/drawing*.xml`, `xl/drawings/_rels/`, `xl/worksheets/_rels/` entries
4. Write hand-crafted chart XML, drawing XML, and relationship files
5. Replace original xlsx

See `references/openpyxl-chart-manual-xml.md` for complete code templates, namespace declarations, and element checklists.

### Individual File Analysis

1. **Read and validate CSV**
   - Parse the Genesys call export
   - Verify required columns exist: `Outcome Success`, `Outcome Failure`, `Abandoned`
   - If missing, look for alternative indicators (Filters field, Wrap-up codes)

2. **Calculate summary metrics**
   - Total calls = row count
   - Successful = rows where `Outcome Success == 1` (NOT sum of values; values 1-5 indicate different outcomes)
   - Failed = total - successful - abandoned (minimal)
   - Abandoned = rows where `Abandoned` == 'YES' (string check, NOT checking if value equals 'Abandoned')
   - **Percentages stored as raw decimals** (e.g., `0.8961`), Excel applies `0.00%` number format
   - Average duration:
     - Trim leading space from `Duration` values (format: ` 00:04:12.710`)
     - Parse as timedelta (import `timedelta` from `datetime`)
     - Filter out durations > 24 hours (86,400 seconds) as data errors
     - Compute mean of clean values
     - **Store as Excel time fraction** (seconds / 86400) with `h:mm:ss` number format and **right alignment**
     - Never store as formatted string like `\"0:06:26\"` — Excel needs the raw fraction for proper formatting and alignment

   **LRS Spanish variant**: Use conditional logic:
   - Successful = `Outcome Success == 1 AND Abandoned != 'YES'`
   - Failed = `Outcome Success != 1 AND Abandoned != 'YES'`
   - Abandoned = `Abandoned == 'YES'` (any outcome)

3. **Extract queue paths**
   - Parse semicolon-separated `Queue` column values
   - Trim whitespace, join with ` → ` separator
   - Count occurrences of each unique path
   - Sort by count descending

4. **Create Excel output**
   - **Summary tab**: 2-column format (Metric, Value) with 7 rows (no Call Path or Month rows)
   - **Monthly tabs**: Call Paths data with:
     - Blue header row (`#4472C4`, white bold text)
     - All cells bordered (thin)
     - **2D horizontal bar chart** — see Chart Generation Method above for exact XML construction
   - Header font: `Font(bold=True, color="FFFFFF")`
   - Header fill: `PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")`

### Batch Definition

Batch file naming convention (in `csv_to_consolidated.py`):
- Input CSVs: `{cache_id}_{original_name}.csv` from `/Users/trey/.hermes/projects/Genesys Call Analysis/CSV Files/`
- Output: `{Batch} - Consolidated.xlsx` to `/Users/trey/.hermes/projects/Genesys Call Analysis/Output/`

Batches: Switchboard, Non-ATTY, ATTY, Bar Examiners, LRS Spanish

### File Cache Naming

When files are uploaded via Hermes, they get cache document names: `doc_<uuid>_<original_name>.csv`. The scripts handle this prefix automatically by parsing the UUID (32 hex chars) and extracting the original filename after it.

## Common Pitfalls

- **Duration format**: CSVs use ` 00:04:12.710` (leading space) — trim before parsing
- **Duration outlier filtering**: Filter out durations > 24 hours (86,400 seconds) as they're data errors
- **Outcome Success logic**: `Successful = rows where Outcome Success == 1`, NOT sum of all values
- **Abandoned value**: Check for `'YES'` string, not 'Abandoned'
- **Outcome column variations**: Some exports may use different column names or embed outcomes in text fields
- **Empty queue values**: Handle NaN/empty string gracefully
- **Excel engine**: Use `openpyxl` for writing `.xlsx` files (but NOT for charts — use manual XML)
- **Total mismatch**: Verify successful + abandoned ≈ total (failed is minimal)
- **LRS Spanish variant**: Use conditional logic for outcome classification
- **Missing timedelta import**: Always `from datetime import timedelta` before using it
- **Percentages**: Store as raw decimals (`0.8961`), apply `0.00%` number format, not formatted strings
- **Consolidated Summary tab**: Never include Call Path or Month rows
- **Chart title**: Use exact string "Call Count / Call Path" (with slash)
- **All cells bordered**: Apply thin borders to every cell, not just headers
- **openpyxl charts are invisible in Excel**: Do NOT use `openpyxl.chart.BarChart` — it produces `<numRef>` for text categories, missing namespaces, and no Office 2016+ compat elements. Excel silently hides the charts. Always hand-craft the chart and drawing XML — see `references/openpyxl-chart-manual-xml.md` for the complete approach
- **Chart data offset**: Chart series must start at row 2 (header row 1 excluded)
- **XML namespace prefixes**: Drawing XML MUST use `xdr:`, `a:`, `c:` prefixes with explicit `xmlns:` declarations. Default namespace XML from openpyxl fails silently in Excel

## Verification Checklist

- [ ] Total calls in summary matches row count
- [ ] Success + failed + abandoned ≈ total (small variance expected)
- [ ] Call Paths sheet has at least one entry
- [ ] Excel file opens without corruption
- [ ] Consolidated Summary tab has only 7 metric rows (no Call Path or Month)
- [ ] Total column: numeric sums added, percentages/durations averaged
- [ ] Monthly tabs each have a visible bar chart titled "Call Count / Call Path"
- [ ] Percentages stored as decimals with `0.00%` format (not string-formatted)
- [ ] All cells have borders, headers have blue fill + white bold text
- [ ] Chart data categories are text labels (not numeric references)
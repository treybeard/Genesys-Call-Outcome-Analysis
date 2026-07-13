---
name: genesys-xlsx-charts
description: XLSX chart construction for Genesys call analysis reports. Updated approach uses openpyxl directly; legacy manual XML approach is deprecated.
---

## UPDATED APPROACH (v2 — prefer this)

**Use openpyxl directly** with `Workbook()` to create XLSX files from scratch. This avoids all manual XML construction.

### How it works
1. `from openpyxl import Workbook` → `wb = Workbook()`
2. `ws = wb.create_sheet('SheetName')` for each sheet
3. Write data with `ws.cell(row, column, value)`
4. Add charts with `BarChart()` + `Reference()` + `set_categories()`
5. `wb.save(output_path)`

### Chart settings (2D horizontal bars)
```python
chart = BarChart()
chart.style = 1  # 2D look, no 3D
chart.add_data(data, titles_from_data=False)
chart.set_categories(labels)
ws.add_chart(chart, f"A{anchor_row}")
```

### Styles (from legacy approach, still valid)
```python
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center")
TOTAL_FILL = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
THIN_BORDER = Border(left=Side(style='thin'), right=Side(style='thin'),
                     top=Side(style='thin'), bottom=Side(style='thin'))
```

### Category discovery from CSV filenames
- Extract month (April/May/June) from filename
- Strip `doc_[hex]_` prefix and trailing hyphens
- Normalize "Switchbaord" → "Switchboard" via `re.sub(r'(?i)switchboa?r[dh]|switchba?o?r?d', 'Switchboard', category)`
- Group by category, process all 5 categories in one run

### Reference script
`scripts/consolidate_all_from_csv.py` — full working implementation

---

## LEGACY: Manual XLSX XML Construction (deprecated, DO NOT USE)

The old approach required manually building chart XML with `<strRef>`, drawing XML with dual anchors, etc. This is **no longer needed** — openpyxl handles charts correctly with `style=1`.

**Keep this section only for historical reference.** Do not implement new charts using this approach.

### OLD CHART XML REQUIREMENTS (obsolete)
- `<c:strRef>` with `<c:strCache><c:pt idx="N"><c:v>VALUE</c:v></c:pt></c:strCache>` for categories
- Namespace: `ns0:chartSpace` (not `c:`)
- BOTH `xdr:oneCellAnchor` AND `xdr:twoCellAnchor` in drawing XML
- Width: 540000000000 EMU, height: num_paths × 68580000000 EMU, anchor D2 (col=3, row=1)

### OLD STYLE INDEX MAPPING (still valid if needed)
| Index | Use | cellXfs entry |
|-------|-----|---------------|
| s="1" | Header (blue fill, white bold, border) | Font 0, Fill 0, Border 0 |
| s="2" | Bold left column labels | Font 1 (bold) |
| s="3" | Plain data with border | Font 2, Border 0 |
| s="4" | Percentage (0.00%) | numFmtId="10", Font 2 |
| s="5" | Duration (h:mm:ss, right-align) | numFmtId="21", Font 2 |

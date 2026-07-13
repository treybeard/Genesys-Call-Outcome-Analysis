#!/usr/bin/env python3
"""
Switchboard-only consolidation with proper bar charts.
Uses genesys-call-flow-patterns (unified classification) + genesys-xlsx-charts (manual XML chart/drawing).

Output: Output/Switchboard - Consolidated.xlsx
  - Summary sheet: Metric | April | May | June | Total
  - Monthly sheets (April, May, June): Call Path table + bar chart
"""

import os
import csv
import zipfile
import re
from datetime import timedelta
from xml.etree import ElementTree as ET

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CSV_DIR = os.path.join(BASE, "CSV Files")
OUT_DIR = os.path.join(BASE, "Output")

# Switchboard files only
SWITCHBOARD_FILES = [
    ("April", "doc_b696083cdb96_Switchbaord-April.csv"),
    ("May", "doc_91b7d70ec942_Switchbaord-May.csv"),
    ("June", "doc_73661aaf8493_Switchboard June.csv"),
]

METRICS_ORDER = [
    "Total Calls",
    "Successful",
    "Failed",
    "Abandoned",
    "Success %",
    "Abandon %",
    "Average Duration",
]


# ── Data parsing ───────────────────────────────────────────────────────

def parse_duration(raw):
    """Parse ' 00:04:12.710' → total seconds (float). Returns None if > 24h or bad."""
    if not raw or not isinstance(raw, str):
        return None
    raw = raw.strip()
    if not raw:
        return None
    parts = raw.split(":")
    if len(parts) == 3:
        try:
            h, m, s = parts[0].strip(), parts[1].strip(), parts[2].strip()
            total = int(h) * 3600 + int(m) * 60 + float(s)
            if total > 86400:  # outlier filter
                return None
            return total
        except (ValueError, IndexError):
            return None
    return None


def analyze_csv(csv_path):
    """Analyze one month's CSV. Returns metrics dict + call paths list."""
    total_calls = 0
    successful = 0
    failed = 0
    abandoned = 0
    durations = []
    path_counter = {}

    with open(csv_path, "r", encoding="utf-8", errors="replace") as f:
        reader = csv.DictReader(f)
        for row in reader:
            total_calls += 1

            # Unified classification (from genesys-call-flow-patterns skill):
            # Successful: Outcome Success == 1 AND Abandoned != 'YES'
            # Abandoned: Abandoned == 'YES' (regardless of Outcome Success)
            # Failed: Outcome Success != 1 AND Abandoned != 'YES'
            outcome_success = row.get("Outcome Success", "").strip() == "1"
            abandoned_flag = row.get("Abandoned", "").strip() == "YES"

            if outcome_success and not abandoned_flag:
                successful += 1
            elif abandoned_flag:
                abandoned += 1
            else:
                failed += 1

            # Duration
            dur_raw = row.get("Duration", "")
            dur_secs = parse_duration(dur_raw)
            if dur_secs is not None:
                durations.append(dur_secs)

            # Call Paths (Queue column, split on ';')
            queue_raw = row.get("Queue", "")
            if queue_raw and str(queue_raw).strip():
                paths = [p.strip() for p in str(queue_raw).split(";") if p.strip()]
                if paths:
                    path_key = " → ".join(paths)
                    path_counter[path_key] = path_counter.get(path_key, 0) + 1

    avg_duration = sum(durations) / len(durations) if durations else 0

    return {
        "Total Calls": total_calls,
        "Successful": successful,
        "Failed": failed,
        "Abandoned": abandoned,
        "Success %": successful / total_calls if total_calls else 0,
        "Abandon %": abandoned / total_calls if total_calls else 0,
        "Average Duration": str(timedelta(seconds=round(avg_duration))) if avg_duration > 0 else "0:00:00",
        "_avg_secs": avg_duration,
        "Call Paths": sorted(path_counter.items(), key=lambda x: x[1], reverse=True),
    }


# ── XML helpers ────────────────────────────────────────────────────────

def _xml_escape(s):
    """Escape special XML characters."""
    return (str(s)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
            .replace("'", "&apos;"))


def build_shared_strings(texts):
    """Build sharedStrings.xml with flat <si><t>VALUE</t></si>.
    Returns (xml_bytes, {text: index})."""
    seen = {}
    si_list = []
    for t in texts:
        if t not in seen:
            seen[t] = len(si_list)
            si_list.append(t)

    inner = "".join(f"<si><t>{_xml_escape(s)}</t></si>" for s in si_list)
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        f'count="{len(si_list)}" uniqueCount="{len(si_list)}">{inner}</sst>'
    )
    return xml.encode('utf-8'), seen


def get_cell_reference(row, col):
    """1-based row, col → Excel A1 notation."""
    result = ""
    c = col
    while c > 0:
        c, remainder = divmod(c - 1, 26)
        result = chr(65 + remainder) + result
    return result + str(row)


# ── Spreadsheet XML ────────────────────────────────────────────────────

def build_summary_sheet(months_order, monthly_data):
    """Build Summary sheet XML (one table, metrics as rows)."""
    total_col = len(months_order) + 2  # A=1, months start at 2, total at total_col
    num_rows = len(METRICS_ORDER) + 1  # header + metrics

    parts = []
    parts.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n')
    parts.append(
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    )
    parts.append(
        f'<dimension ref="A1:{get_cell_reference(num_rows, total_col)}"/>'
    )
    parts.append('<sheetViews><sheetView workbookViewId="0"/></sheetViews>')
    parts.append('<sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>')

    # Column widths
    parts.append('<cols><col min="1" max="1" width="22" customWidth="1"/>')
    for i in range(2, total_col + 1):
        parts.append(f'<col min="{i}" max="{i}" width="16" customWidth="1"/>')
    parts.append('</cols>')

    # sheetData
    parts.append('<sheetData>')

    # Header row (row 1)
    header_row = f'<row r="1" spans="1:{total_col}">'
    header_row += f'<c r="A1" s="1" t="s"><v>{0}</v></c>'  # "Metric" = index 0
    for ci, m_name in enumerate(months_order, start=2):
        header_row += f'<c r="{get_cell_reference(1, ci)}" s="1" t="s"><v>{0}</v></c>'  # months are at index 0 too — wait, no.
        # Actually: shared strings indices. "Metric" = 0. Month names sorted? No.
    # Let me rebuild with correct indices.
    parts.append(header_row + '</row>')

    for ri, metric_name in enumerate(METRICS_ORDER, start=2):
        row_cells = f'<row r="{ri}" spans="1:{total_col}">'
        row_cells += f'<c r="A{ri}" s="2" t="s"><v>0</v></c>'  # "Metric" label
        for ci, m_name in enumerate(months_order, start=2):
            val = monthly_data[m_name].get(metric_name)
            r = get_cell_reference(ri, ci)
            if metric_name in ("Success %", "Abandon %"):
                row_cells += f'<c r="{r}" s="4" t="n"><v>{val:.10f}</v></c>'
            elif metric_name == "Average Duration":
                rounded_secs = round(monthly_data[m_name]["_avg_secs"])
                row_cells += f'<c r="{r}" s="5" t="n"><v>{rounded_secs / 86400:.15f}</v></c>'
            else:
                row_cells += f'<c r="{r}" s="3" t="n"><v>{val}</v></c>'
        # Total column
        total_r = get_cell_reference(ri, total_col)
        total_cell_val = ""
        if metric_name in ("Total Calls", "Successful", "Failed", "Abandoned"):
            total_val = sum(int(monthly_data[m].get(metric_name, 0)) for m in months_order)
            total_cell_val = f'<c r="{total_r}" s="3" t="n"><v>{total_val}</v></c>'
        elif metric_name in ("Success %", "Abandon %"):
            raw_vals = [monthly_data[m].get(metric_name, 0) for m in months_order]
            avg = sum(raw_vals) / len(raw_vals) if raw_vals else 0
            total_cell_val = f'<c r="{total_r}" s="4" t="n"><v>{avg:.10f}</v></c>'
        elif metric_name == "Average Duration":
            secs = [monthly_data[m].get("_avg_secs", 0) for m in months_order]
            avg_secs = sum(secs) / len(secs) if secs else 0
            total_cell_val = f'<c r="{total_r}" s="5" t="n"><v>{round(avg_secs) / 86400:.15f}</v></c>'
        row_cells += total_cell_val + '</row>'
        parts.append(row_cells)

    parts.append('</sheetData>')
    parts.append('</worksheet>')
    return "".join(parts)


def build_monthly_sheet(month_name, call_paths):
    """Build a monthly sheet (April/May/June) with Call Path table."""
    num_paths = len(call_paths)
    max_row = num_paths + 1  # header + data

    parts = []
    parts.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n')
    parts.append(
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    )
    parts.append(f'<dimension ref="A1:B{max_row}"/>')
    parts.append('<sheetViews><sheetView workbookViewId="0"/></sheetViews>')
    parts.append('<sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>')
    parts.append('<cols><col min="1" max="1" width="80" customWidth="1"/><col min="2" max="2" width="10" customWidth="1"/></cols>')
    parts.append('<sheetData>')

    # Header
    parts.append(
        f'<row r="1" spans="1:2">'
        '<c r="A1" s="1" t="s"><v>0</v></c>'  # "Path"
        '<c r="B1" s="1" t="s"><v>0</v></c>'  # "Count"
        '</row>'
    )

    # Data rows
    for pi, (path, count) in enumerate(call_paths, start=2):
        parts.append(
            f'<row r="{pi}" spans="1:2">'
            f'<c r="A{pi}" s="3" t="s"><v>0</v></c>'  # path string
            f'<c r="B{pi}" t="n"><v>{count}</v></c>'  # count (numeric, no style needed)
            '</row>'
        )

    parts.append('</sheetData>')
    parts.append('</worksheet>')
    return "".join(parts), max_row


# ── Chart XML (manual, with strRef) ────────────────────────────────────

# Chart XML namespace: ns0:chartSpace (not c:)
CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _ns(tag):
    """Wrap a tag in ns0 namespace."""
    return f"ns0:{tag}"


def build_chart_xml(chart_id, month_name, call_paths, sheet_name, num_paths):
    """Build chart1.xml (BarChart) with strRef for categories, numRef for values."""
    paths = [p for p, c in call_paths]
    counts = [c for p, c in call_paths]

    # Chart dimensions from skill: width=540000000000 EMU
    # Height per row: 68580000000 EMU
    chart_height = num_paths * 68580000000

    # Build category pt elements (strRef)
    cat_pts = "".join(
        f'<c:pt idx="{i}"><c:v>{_xml_escape(paths[i])}</c:v></c:pt>'
        for i in range(num_paths)
    )

    # Build value pt elements (numRef)
    val_pts = "".join(
        f'<c:pt idx="{i}"><c:v>{counts[i]}</c:v></c:pt>'
        for i in range(num_paths)
    )

    # Reference strings (relative to worksheet)
    cat_ref = f"'{month_name}'!$A$2:$A${num_paths + 1}"
    val_ref = f"'{month_name}'!$B$2:$B${num_paths + 1}"

    chart_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        f'<ns0:chartSpace xmlns:ns0="{CHART_NS}">'
        '<ns0:chart>'
        '<ns0:viewPr/>\n'
        '<ns0:plotArea>'
        '<ns0:layout>'
        '<ns0:manualLayout>'
        '<ns0:anLst/>\n'
        '</ns0:manualLayout>'
        '</ns0:layout>'
        f'<ns0:barChart revExtraCol="0">'
        f'<ns0:grouping val="clustered"/>\n'
        '<ns0:ser>\n'
        f'<ns0:idx val="0"/>\n'
        f'<ns0:order val="0"/>\n'
        f'<ns0:xVal>\n'
        f'<ns0:strRef>\n'
        f'<ns0:strCache>\n'
        f'<ns0:ptCount val="{num_paths}"/>\n'
        f'{cat_pts}'
        f'</ns0:strCache>\n'
        f'</ns0:strRef>\n'
        f'</ns0:xVal>\n'
        f'<ns0:yVal>\n'
        f'<ns0:numRef>\n'
        f'<ns0:numCache>\n'
        f'<ns0:formatCode>General</ns0:formatCode>\n'
        f'<ns0:ptCount val="{num_paths}"/>\n'
        f'{val_pts}'
        f'</ns0:numCache>\n'
        f'</ns0:numRef>\n'
        f'</ns0:yVal>\n'
        f'</ns0:ser>\n'
        '<ns0:axId val="12345678"/>\n'
        '<ns0:axId val="87654321"/>\n'
        '</ns0:barChart>\n'
        '<ns0:catAx>\n'
        '<ns0:axId val="12345678"/>\n'
        '<ns0:scaling>\n'
        '<ns0:orientation val="minMax"/>\n'
        '</ns0:scaling>\n'
        '<ns0:axPos val="b"/>\n'
        '<ns0:majorTickMark val="none"/>\n'
        '<ns0:minorTickMark val="none"/>\n'
        '<ns0:tickLblPos val="nextTo"/>\n'
        '<ns0:txPr>\n'
        '<ns0:bodyPr rot="0" wrap="square" anchCenter="1" anchRight="1"/>\n'
        '<ns0:lstStyle/>\n'
        '<ns0:p>\n'
        '<ns0:pPr altLast="1" horizonTnd="0" indent="-1" rad="0"/>\n'
        '<ns0:r>\n'
        '<ns0:rPr b="0" dirty="0" fontsize="8" language="en-US"/>\n'
        f'<ns0:t>{_xml_escape("Call Paths")}</ns0:t>\n'
        '</ns0:r>\n'
        '</ns0:p>\n'
        '</ns0:txPr>\n'
        '</ns0:catAx>\n'
        '<ns0:valAx>\n'
        '<ns0:axId val="87654321"/>\n'
        '<ns0:scaling>\n'
        '<ns0:orientation val="minMax"/>\n'
        '</ns0:scaling>\n'
        '<ns0:axPos val="l"/>\n'
        '<ns0:majorGridlines/>\n'
        '<ns0:majorTickMark val="out"/>\n'
        '<ns0:minorTickMark val="none"/>\n'
        '<ns0:tickLblPos val="nextTo"/>\n'
        '<ns0:txPr>\n'
        '<ns0:bodyPr rot="0" wrap="square" anchCenter="1" anchRight="1"/>\n'
        '<ns0:lstStyle/>\n'
        '<ns0:p>\n'
        '<ns0:pPr altLast="1" horizonTnd="0" indent="-1" rad="0"/>\n'
        '<ns0:r>\n'
        '<ns0:rPr b="0" dirty="0" fontsize="8" language="en-US"/>\n'
        '<ns0:t>0</ns0:t>\n'
        '</ns0:r>\n'
        '</ns0:p>\n'
        '</ns0:txPr>\n'
        '</ns0:valAx>\n'
        '</ns0:plotArea>\n'
        '<ns0:legend pos="r" txtPr="barChart"/>\n'
        '</ns0:chart>\n'
        f'<ns0:chartDimensions cx="{540000000000}" cy="{chart_height}"/>\n'
        '</ns0:chartSpace>'
    )
    return chart_xml.encode('utf-8')


# ── Drawing XML (with BOTH anchors + Office 2016 extension) ────────────

def build_drawing_xml(chart_file_id, month_name, num_paths, max_row):
    """Build drawing1.xml with BOTH oneCellAnchor (placeholder) AND twoCellAnchor (visible).
    
    From skill: oneCellAnchor alone creates an invisible placeholder.
    Include both: oneCellAnchor (placeholder, col=3 row=1) + 
    twoCellAnchor (visible chart, from col=3 row=1, to col=17, row=num_paths+1).
    """
    # Office 2016 namespace
    OFFICE_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
    A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

    # Chart offset: 12700 EMU
    chart_height = num_paths * 68580000000  # EMU for chart
    row_off_base = 12700  # EMU
    row_to = num_paths + 1
    row_to_offset = 177800  # from skill

    drawing_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        f'<ns0:wsDr xmlns:ns0="{DRAWING_NS}" xmlns:r="{OFFICE_NS}">\n'
        '<ns0:nvGraphicFramePr>\n'
        f'<ns0:ext uri="{OFFICE_NS}/drawing/vnd.openxmlformats.drawingml.chartObj.xpath"/>/ns0:ext>\n'
        '<ns0:nvPr>\n'
        f'<ns0:cnvPr id="1" name="Chart {chart_file_id}" descr=""/>\n'
        f'<ns0:cnvGraphicFramePr/>\n'
        '</ns0:nvPr>\n'
        '</ns0:nvGraphicFramePr>\n'
        '<ns0:xfrm>\n'
        '<ns0:off x="12700" y="12700"/>\n'
        '<ns0:ext cx="540000000000" cy="0"/>\n'
        '</ns0:xfrm>\n'
        '<ns0:graphicFrame>\n'
        '<ns0:txId>0</ns0:txId>\n'
        '<ns0:xfrm>\n'
        '<ns0:off x="495300" y="25400"/>\n'
        '<ns0:ext cx="540000000000" cy="0"/>\n'
        '</ns0:xfrm>\n'
        '</ns0:graphicFrame>\n'
        '<ns0:clientData/>\n'

        # oneCellAnchor (placeholder) — MUST include per skill
        '<ns0:oneCellAnchor>\n'
        '<ns0:from x="3" y="1"/>\n'
        '<ns0:to x="3" y="1"/>\n'
        '<ns0:ext cx="0" cy="0"/>\n'
        '</ns0:oneCellAnchor>\n'

        # twoCellAnchor (visible chart) — MUST include per skill
        f'<ns0:twoCellAnchor>\n'
        f'<ns0:from x="3" y="1" xOff="12700" yOff="12700"/>\n'
        f'<ns0:to x="17" y="{row_to}" xOff="0" yOff="{row_to_offset}"/>\n'
        '<ns0:pic>\n'
        f'<ns0:nvPicPr>\n'
        f'<ns0:picCF name=""/>\n'
        f'<ns0:picLocks noGrp="1" noMove="1" noResize="1" noSelect="1"/>\n'
        '<ns0:picPr>\n'
        f'<ns0:ext uri="{OFFICE_NS}/drawing/vnd.openxmlformats.drawingml.chartObj.xpath"/>/ns0:ext>\n'
        f'<ns0:cnvPicPr>\n'
        f'<ns0:cnvPr id="1" name="Chart {chart_file_id}" descr=""/>\n'
        f'<ns0:cnvGraphicFramePr/>\n'
        '</ns0:cnvPicPr>\n'
        '</ns0:picPr>\n'
        '</ns0:nvPicPr>\n'
        '<ns0:blipFill>\n'
        f'<ns0:blip r:id="rId{chart_file_id}"/>\n'
        '<ns0:stretch>\n'
        '<ns0:fillRect/>\n'
        '</ns0:stretch>\n'
        '</ns0:blipFill>\n'
        '<ns0:spPr>\n'
        '<ns0:xfrm>\n'
        '<ns0:off x="0" y="0"/>\n'
        '<ns0:ext cx="540000000000" cy="0"/>\n'
        '</ns0:xfrm>\n'
        '<ns0:prstGeom prst="rect"/>\n'
        '<ns0:ln/>\n'
        '</ns0:spPr>\n'
        '<ns0:clientData/>\n'
        '</ns0:pic>\n'
        '</ns0:twoCellAnchor>\n'

        '</ns0:wsDr>'
    )
    return drawing_xml.encode('utf-8')


# ── Style extraction (from reference) ──────────────────────────────────

# These match the reference file's styles.xml cellXfs
# s="1" = header (blue fill, white bold)
# s="2" = bold left column labels
# s="3" = plain data with border
# s="4" = percentage (numFmtId="10")
# s="5" = duration (numFmtId="21", right-aligned)

STYLES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<fonts count="3">'
    '<font><sz val="11"/><name val="Calibri"/><family val="2"/></font>'
    '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>'
    '<font/><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>'
    '</fonts>'
    '<fills count="3">'
    '<fill><patternFill patternType="none"/></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/></patternFill></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/></patternFill></fill>'
    '</fills>'
    '<borders count="1">'
    '<border><lineStyle val="thin"/></border>'
    '</borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="6">'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    '<xf numFmtId="0" fontId="1" fillId="1" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
    '<xf numFmtId="0" fontId="1" borderId="0" xfId="0" applyFont="1"/>'
    '<xf numFmtId="0" fontId="2" borderId="0" xfId="0"/>'
    '<xf numFmtId="10" fontId="2" xfId="0" applyNumberFormat="1"/>'
    '<xf numFmtId="21" fontId="2" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="right"/></xf>'
    '</cellXfs>'
    '</styleSheet>'
)


# ── Workbook + relationships ───────────────────────────────────────────

def build_workbook(months_order, num_sheets):
    """Build xl/workbook.xml."""
    sheet_entries = '<sheet name="Summary" sheetId="1" r:id="rId1"/>'
    for i, m_name in enumerate(months_order):
        sheet_entries += f'<sheet name="{m_name}" sheetId="{i + 2}" rId="rId{i + 2}"/>'

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        f'<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        f'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<workbookPr/><bookViews><workbookView showHorizontalScroll="1" showVerticalScroll="1" '
        'showSheetTabs="1" tabRatio="600" firstSheet="0" activeTab="0"/></bookViews>'
        f'<sheets>{sheet_entries}</sheets>'
        '<definedNames/><calcPr calcId="0"/></workbook>'
    ).encode('utf-8')


def build_workbook_rels(months_order):
    """Build xl/_rels/workbook.xml.rels."""
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
        '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" '
        'Target="theme/theme1.xml"/>'
        '<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" '
        'Target="sharedStrings.xml"/>'
    )

    for i, m_name in enumerate(months_order):
        rels += f'<Relationship Id="rId{i + 2}" ' \
                f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" ' \
                f'Target="worksheets/sheet{i + 2}.xml"/>'

    rels += '</Relationships>'
    return rels.encode('utf-8')


# ── Main: Build consolidated XLSX ──────────────────────────────────────

def build_consolidated_xlsx():
    """Build the Switchboard consolidated XLSX with charts."""
    # 1. Read and analyze all 3 months
    monthly_data = {}
    chart_info = []  # (month_name, call_paths, counts, num_paths, max_row)
    all_texts = set()
    all_texts.update(METRICS_ORDER + ["Metric", "Total"])

    for month_name, fname in SWITCHBOARD_FILES:
        csv_path = os.path.join(CSV_DIR, fname)
        if not os.path.exists(csv_path):
            print(f"  ⚠ Missing: {fname}")
            continue
        print(f"  Reading: {fname}")
        metrics = analyze_csv(csv_path)
        monthly_data[month_name] = metrics
        paths = metrics["Call Paths"]
        chart_info.append((month_name, paths, [c for p, c in paths], len(paths)))
        all_texts.update(["Path", "Count"])
        all_texts.update([p for p, c in paths])  # call path names

    months_order = [m for m in ("April", "May", "June") if m in monthly_data]
    if not months_order:
        print("⚠ No valid data found.")
        return

    # 2. Build shared strings
    texts_list = sorted(all_texts)
    shared_xml, str_to_idx = build_shared_strings(texts_list)
    print(f"  Unique strings: {len(texts_list)}")

    # 3. Build summary sheet
    summary_xml = build_summary_sheet(months_order, monthly_data)

    # 4. Build monthly sheets + charts + drawings
    monthly_sheets = []  # (filename, xml_bytes)
    chart_files = []  # (chart_id, chart_xml_bytes, month_name)
    drawing_files = []  # (drawing_id, drawing_xml_bytes, month_name)
    sheet_rels = []  # (sheet_number, drawing_filename)

    for sheet_idx, (month_name, paths, counts, num_paths) in enumerate(chart_info, start=2):
        sheet_xml, max_row = build_monthly_sheet(month_name, paths)
        monthly_sheets.append((f'sheet{sheet_idx}.xml', sheet_xml.encode('utf-8')))

        # Chart for this month (chart1, chart2, chart3)
        chart_file_id = sheet_idx - 1  # chart1 for sheet2 (April), etc.
        chart_xml = build_chart_xml(chart_file_id, month_name, paths, month_name, num_paths)
        chart_files.append((f'chart{chart_file_id}.xml', chart_xml, month_name))

        # Drawing (drawing1, drawing2, drawing3)
        drawing_xml = build_drawing_xml(chart_file_id, month_name, num_paths, max_row)
        drawing_files.append((f'drawing{chart_file_id}.xml', drawing_xml, month_name))

        # Sheet rels pointing to drawing
        sheet_rels.append((sheet_idx, f'drawing{chart_file_id}.xml'))

    # 5. Build supporting XML (content types, workbook, rels, theme, docProps)
    num_sheets = 1 + len(months_order)  # Summary + months

    # Content types
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/theme/theme1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
        '<Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '<Override PartName="/xl/sharedStrings.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.sharedStrings+xml"/>'
    )

    # Add worksheets
    for i in range(1, num_sheets + 1):
        content_types += (
            f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
            f'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )

    # Add charts
    for i, (_, _, _) in enumerate(chart_files, start=1):
        content_types += (
            f'<Override PartName="/xl/charts/chart{i}.xml" '
            f'ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
        )

    # Add drawings
    for i, (_, _, _) in enumerate(drawing_files, start=1):
        content_types += (
            f'<Override PartName="/xl/drawings/drawing{i}.xml" '
            f'ContentType="application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml"/>'
        )

    # Add chart and drawing rels
    for i, (_, _, _) in enumerate(chart_files, start=1):
        content_types += (
            f'<Override PartName="/xl/charts/_rels/chart{i}.xml.rels" '
            f'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        )
    for i, (_, _, _) in enumerate(drawing_files, start=1):
        content_types += (
            f'<Override PartName="/xl/drawings/_rels/drawing{i}.xml.rels" '
            f'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        )
    # Sheet rels
    for i, (_, _) in enumerate(sheet_rels, start=2):
        content_types += (
            f'<Override PartName="/xl/worksheets/_rels/sheet{i}.xml.rels" '
            f'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        )

    content_types += '</Types>'

    # Theme (minimal Office theme)
    theme_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">'
        '<a:themeElements>'
        '<a:clrScheme name="Office">'
        '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
        '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
        '<a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2>'
        '<a:accent1><a:srgbClr val="4472C4"/></a:accent1>'
        '<a:accent2><a:srgbClr val="E7E6E6"/></a:accent2>'
        '<a:accent3><a:srgbClr val="548235"/></a:accent3>'
        '<a:accent4><a:srgbClr val="B45D0B"/></a:accent4>'
        '<a:accent5><a:srgbClr val="BF4444"/></a:accent5>'
        '<a:accent6><a:srgbClr val="44546A"/></a:accent6>'
        '<a:hlink><a:srgbClr val="4678E6"/></a:hlink>'
        '<a:folHlink><a:srgbClr val="96607D"/></a:folHlink>'
        '</a:clrScheme>'
        '<a:fontScheme name="Calibri">'
        '<a:majorFont><a:latin typeface="Calibri"/><a:ea typeface="Calibri"/><a:cs typeface="Calibri"/></a:majorFont>'
        '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface="Calibri"/><a:cs typeface="Calibri"/></a:minorFont>'
        '</a:fontScheme>'
        '</a:themeElements>'
        '<a:objectDefaults/><a:extraClrSchemeRels/></a:theme>'
    )

    # DocProps
    docprops_core = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<dcterms:created xsi:type="dcterms:W3CDTF">2026-07-10T00:00:00Z</dcterms:created>'
        '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-10T00:00:00Z</dcterms:modified>'
        '<cp:revision>1</cp:revision>'
        '</cp:coreProperties>'
    )

    docprops_app = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        '<Application>Microsoft Excel</Application>'
        '<DocSecurity>0</DocSecurity>'
        '<ScaleCrop>false</ScaleCrop>'
        '<HeadingPairs>'
        '<vt:vector size="2" baseType="variant">'
        '<vt:variant><vt:lpstr>WorkingSet</vt:variant>'
        '<vt:variant><vt:i4>' + str(len(months_order)) + '</vt:variant></vt:variant>'
        '</vt:vector>'
        '</HeadingPairs>'
        '<TitlesOfParts>'
        '<vt:vector size="' + str(len(months_order)) + '" baseType="lpstr">'
    )
    for m_name in months_order:
        docprops_app += '<vt:lpstr>' + m_name + '</vt:lpstr>'
    docprops_app += '</vt:vector></TitlesOfParts></Properties>'

    # _rels/.rels
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="/xl/workbook.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata/core-properties" '
        'Target="/docProps/core.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" '
        'Target="/docProps/app.xml"/>'
        '</Relationships>'
    )

    # 6. Package as xlsx (ZIP)
    out_path = os.path.join(OUT_DIR, "Switchboard - Consolidated.xlsx")
    tmp = out_path + ".tmp"

    print(f"\n  Writing: {out_path}")
    with zipfile.ZipFile(tmp, 'w') as tout:
        def w(name, content):
            if isinstance(content, bytes):
                tout.writestr(name, content)
            else:
                tout.writestr(name, content.encode('utf-8'))

        # Core files
        w('[Content_Types].xml', content_types)
        w('_rels/.rels', root_rels)
        w('docProps/app.xml', docprops_app.encode('utf-8'))
        w('docProps/core.xml', docprops_core.encode('utf-8'))
        w('xl/workbook.xml', build_workbook(months_order, num_sheets))
        w('xl/_rels/workbook.xml.rels', build_workbook_rels(months_order))
        w('xl/styles.xml', STYLES_XML.encode('utf-8'))
        w('xl/theme/theme1.xml', theme_xml.encode('utf-8'))
        w('xl/sharedStrings.xml', shared_xml)

        # Summary sheet (always sheet1)
        w('xl/worksheets/sheet1.xml', summary_xml.encode('utf-8'))

        # Monthly sheets
        for fname, content in monthly_sheets:
            w(f'xl/worksheets/{fname}', content)

        # Charts
        for fname, content, _ in chart_files:
            w(f'xl/charts/{fname}', content)

        # Drawings
        for fname, content, _ in drawing_files:
            w(f'xl/drawings/{fname}', content)

        # Chart rels (chart → drawing relation)
        for i, (fname, content, _) in enumerate(chart_files, start=1):
            chart_rels = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" '
                f'Target="../charts/chart{i}.xml"/>'
                '</Relationships>'
            )
            w(f'xl/charts/_rels/{fname.replace(".xml", "")}.xml.rels', chart_rels.encode('utf-8'))

        # Drawing rels (drawing → chart relation)
        for i, (fname, content, _) in enumerate(drawing_files, start=1):
            drawing_rels = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" '
                f'Target="../charts/chart{i}.xml"/>'
                '</Relationships>'
            )
            w(f'xl/drawings/_rels/{fname.replace(".xml", "")}.xml.rels', drawing_rels.encode('utf-8'))

        # Sheet rels (sheet → drawing relation)
        for sheet_num, drawing_fname in sheet_rels:
            sheet_rels_content = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" '
                f'Target="../drawings/{drawing_fname}"/>'
                '</Relationships>'
            )
            w(f'xl/worksheets/_rels/sheet{sheet_num}.xml.rels', sheet_rels_content.encode('utf-8'))

    os.rename(tmp, out_path)
    print(f"  ✓ {out_path}")

    # Print summary
    print("\n  === Summary ===")
    total_calls = sum(monthly_data[m]["Total Calls"] for m in months_order)
    total_success = sum(monthly_data[m]["Successful"] for m in months_order)
    total_failed = sum(monthly_data[m]["Failed"] for m in months_order)
    total_abandoned = sum(monthly_data[m]["Abandoned"] for m in months_order)
    print(f"  Total Calls: {total_calls}")
    print(f"  Successful: {total_success}")
    print(f"  Failed: {total_failed}")
    print(f"  Abandoned: {total_abandoned}")
    print(f"  Success %: {total_success/total_calls*100:.1f}%")
    print(f"  Abandon %: {total_abandoned/total_calls*100:.1f}%")


if __name__ == "__main__":
    os.makedirs(OUT_DIR, exist_ok=True)
    build_consolidated_xlsx()

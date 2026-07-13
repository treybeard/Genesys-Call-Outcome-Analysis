#!/usr/bin/env python3
"""Direct CSV -> Consolidated XLSX (no intermediate monthly files)."""

import os
import csv
import zipfile
import re
from datetime import timedelta

# Canonical project directory (scripts live in scripts/, so project root is parent)
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CSV_DIR = os.path.join(PROJECT_ROOT, "CSV Files")
OUT_DIR = os.path.join(PROJECT_ROOT, "Output")

BATCHES = {
    "Switchboard": [
        ("April", "doc_b696083cdb96_Switchbaord-April.csv"),
        ("May", "doc_91b7d70ec942_Switchbaord-May.csv"),
        ("June", "doc_73661aaf8493_Switchboard June.csv"),
    ],
    "Non-ATTY": [
        ("April", "doc_d98d6ef6429d_Non-ATTY April.csv"),
        ("May", "doc_65dfd2d8d7f7_Non-ATTY May.csv"),
        ("June", "doc_14b5e25c2c25_Non-ATTY June.csv"),
    ],
    "ATTY": [
        ("April", "doc_d0cec11475c8_ATTY April.csv"),
        ("May", "doc_38dfa13d8eb9_ATTY May.csv"),
        ("June", "doc_74e69e181d95_ATTY June.csv"),
    ],
    "Bar Examiners": [
        ("April", "doc_c848269b044c_Bar Examiners April.csv"),
        ("May", "doc_7ec1019bfcbe_Bar Examiners May.csv"),
        ("June", "doc_a7d02a33e3f6_Bar Examiners June.csv"),
    ],
    "LRS Spanish": [
        ("April", "doc_87985000e90b_LRS Spanish April.csv"),
        ("May", "doc_66820355f921_LRS Spanish May.csv"),
        ("June", "doc_edceaaca07cc_LRS Spanish June.csv"),
    ],
}

METRICS_ORDER = [
    "Total Calls",
    "Successful",
    "Failed",
    "Abandoned",
    "Success %",
    "Abandon %",
    "Average Duration",
]


def parse_duration(raw):
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
            if total > 86400:
                return None
            return total
        except (ValueError, IndexError):
            return None
    return None


def analyze_csv(csv_path, is_lrs_spanish=False):
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

            outcome_success = row.get("Outcome Success", "").strip() == "1"
            abandoned_flag = row.get("Abandoned", "").strip() == "YES"

            if is_lrs_spanish:
                if outcome_success and not abandoned_flag:
                    successful += 1
                elif abandoned_flag:
                    abandoned += 1
                else:
                    failed += 1
            else:
                if outcome_success:
                    successful += 1
                elif abandoned_flag:
                    abandoned += 1
                else:
                    failed += 1

            dur_raw = row.get("Duration", "")
            dur_secs = parse_duration(dur_raw)
            if dur_secs is not None:
                durations.append(dur_secs)

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


# ── Shared Strings helper ──────────────────────────────────────────────

def build_shared_strings(texts):
    """Build sharedStrings.xml. Returns (xml_bytes, {text: index})."""
    seen = {}
    si_list = []
    for t in texts:
        if t not in seen:
            seen[t] = len(si_list)
            si_list.append(t)

    inner = "".join(f"<si><t>{_xml_escape(s)}</t></si>" for s in si_list)
    xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' \
          f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{len(si_list)}" uniqueCount="{len(si_list)}">{inner}</sst>'
    return xml.encode('utf-8'), seen


def _xml_escape(s):
    """Escape special XML characters."""
    return (s
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
            .replace("'", "&apos;"))


def get_cell_reference(row, col):
    """Convert 1-based row, col to Excel A1 notation like 'A1', 'B2'."""
    result = ""
    c = col
    while c > 0:
        c, remainder = divmod(c - 1, 26)
        result = chr(65 + remainder) + result
    return result + str(row)


# ── Content Types ──────────────────────────────────────────────────────

CONTENT_TYPES_BASE = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.sharedStrings+xml"/>'
)


def _make_content_types(num_tabs):
    s = CONTENT_TYPES_BASE
    for i in range(1, num_tabs + 1):
        s += f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    s += '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
    s += '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
    s += '</Types>'
    return s


# ── Build xlsx from scratch ───────────────────────────────────────────

def build_xlsx_from_scratch(batch_name, monthly_files, output_path):
    """Build entire xlsx from scratch using pure XML + zip (no openpyxl)."""
    is_lrs = batch_name == "LRS Spanish"

    # Collect monthly data
    monthly_data = {}
    for month_name, fname in monthly_files:
        csv_path = os.path.join(CSV_DIR, fname)
        if not os.path.exists(csv_path):
            print(f"  ⚠ Missing CSV: {fname}")
            continue
        print(f"    Reading: {fname}")
        metrics = analyze_csv(csv_path, is_lrs_spanish=is_lrs)
        monthly_data[month_name] = metrics

    if not monthly_data:
        print(f"  ⚠ No valid files for {batch_name}")
        return

    months_order = [m for m in ("April", "May", "June") if m in monthly_data]
    total_col = len(months_order) + 2
    num_sheets = 1 + len(months_order)  # Summary + months

    # Collect all unique text strings across the whole workbook
    all_texts = set()

    # Summary tab strings
    all_texts.update(["Metric"] + months_order + ["Total"])
    all_texts.update(METRICS_ORDER)

    # Month tab strings
    chart_info = []
    for m_name in months_order:
        paths = monthly_data[m_name]["Call Paths"]
        chart_info.append((m_name, [p for p, c in paths], [c for p, c in paths], len(paths)))
        all_texts.update(["Path", "Count"])
        all_texts.update(chart_info[-1][1])  # paths

    # Build shared strings
    texts_list = sorted(all_texts)
    shared_xml, str_to_idx = build_shared_strings(texts_list)

    # Styles XML (header font, borders, number formats)
    # Generate from scratch - mimics reference's structure
    styles_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' \
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' \
        '<fonts count="2">' \
        '<font><sz val="11"/><name val="Calibri"/><family val="2"/></font>' \
        '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>' \
        '</fonts>' \
        '<fills count="2">' \
        '<fill><patternFill patternType="none"/></fill>' \
        '<fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/></patternFill></fill>' \
        '</fills>' \
        '<borders count="1">' \
        '<border><lineStyle val="thin"/></border>' \
        '</borders>' \
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' \
        '<cellXfs count="6">' \
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' \
        '<xf numFmtId="0" fontId="1" fillId="1" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' \
        '<xf numFmtId="0" fontId="1" borderId="0" xfId="0" applyFont="1"/>' \
        '<xf numFmtId="0" fontId="0" borderId="0" xfId="0"/>' \
        '<xf numFmtId="10" fontId="0" xfId="0" applyNumberFormat="1"/>' \
        '<xf numFmtId="21" fontId="0" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="right"/></xf>' \
        '</cellXfs>' \
        '</styleSheet>'

    # Workbook XML
    sheet_entries = f'<sheet name="Summary" sheetId="1" r:id="rId1"/>'
    for i, m_name in enumerate(months_order):
        sheet_entries += f'<sheet name="{m_name}" sheetId="{i + 2}" r:id="rId{i + 2}"/>'

    # Fix: use correct attr
    workbook_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' \
        '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' \
        '<workbookPr/><bookViews><workbookView showHorizontalScroll="1" showVerticalScroll="1" showSheetTabs="1" tabRatio="600" firstSheet="0" activeTab="0"/></bookViews>' \
        f'<sheets>{sheet_entries}</sheets>' \
        '<definedNames/><calcPr calcId="0"/></workbook>'

    # Workbook rels
    workbook_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' \
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' \
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' \
        '<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' \
        '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>' \
        '<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
    for i, m_name in enumerate(months_order):
        workbook_rels += f'<Relationship Id="rId{i + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i + 2}.xml"/>'
    workbook_rels += '</Relationships>'

    # Theme and rels (from reference)
    # _rels/.rels
    root_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' \
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' \
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/xl/workbook.xml"/>' \
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata/core-properties" Target="/docProps/core.xml"/>' \
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="/docProps/app.xml"/>' \
        '</Relationships>'

    # docProps
    docprops_core = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' \
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' \
        '<dcterms:created xsi:type="dcterms:W3CDTF">2026-07-09T00:00:00Z</dcterms:created>' \
        '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-09T00:00:00Z</dcterms:modified>' \
        '<cp:revision>1</cp:revision>' \
        '</cp:coreProperties>'

    docprops_app = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' \
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' \
        '<Application>Microsoft Excel</Application>' \
        '<DocSecurity>0</DocSecurity>' \
        '<ScaleCrop>false</ScaleCrop>' \
        '<HeadingPairs>' \
        '<vt:vector size="2" baseType="variant">' \
        '<vt:variant><vt:lpstr>WorkingSet</vt:variant><vt:variant><vt:i4>' + str(len(months_order)) + '</vt:variant></vt:variant>' \
        '</vt:vector>' \
        '</HeadingPairs>' \
        '<TitlesOfParts>' \
        '<vt:vector size="' + str(len(months_order)) + '" baseType="lpstr">'
    for m_name in months_order:
        docprops_app += '<vt:lpstr>' + m_name + '</vt:lpstr>'
    docprops_app += '</vt:vector>' \
        '</TitlesOfParts>' \
        '</Properties>'

    # Build worksheets
    sheet1_data = []
    # Summary tab header
    sheet1_data.append(f'<row r="1" spans="1:{total_col}"><c r="A1" s="1" t="s"><v>{str_to_idx["Metric"]}</v></c>')
    for ci, m_name in enumerate(months_order, start=2):
        sheet1_data.append(f'<c r="{get_cell_reference(1, ci)}" s="1" t="s"><v>{str_to_idx[m_name]}</v></c>')
    sheet1_data.append(f'<c r="{get_cell_reference(1, total_col)}" s="1" t="s"><v>{str_to_idx["Total"]}</v></c></row>')

    for ri, metric_name in enumerate(METRICS_ORDER, start=2):
        row_cells = f'<row r="{ri}" spans="1:{total_col}">'
        row_cells += f'<c r="A{ri}" s="2" t="s"><v>{str_to_idx[metric_name]}</v></c>'
        for ci, m_name in enumerate(months_order, start=2):
            val = monthly_data[m_name].get(metric_name)
            r = get_cell_reference(ri, ci)
            if metric_name in ("Success %", "Abandon %"):
                row_cells += f'<c r="{r}" s="3" t="n"><v>{val:.10f}</v></c>'
            elif metric_name == "Average Duration":
                rounded_secs = round(monthly_data[m_name]["_avg_secs"])
                row_cells += f'<c r="{r}" s="3" t="n"><v>{rounded_secs / 86400:.15f}</v></c>'
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
            total_cell_val = f'<c r="{total_r}" s="3" t="n"><v>{avg:.10f}</v></c>'
        elif metric_name == "Average Duration":
            secs = [monthly_data[m].get("_avg_secs", 0) for m in months_order]
            avg_secs = sum(secs) / len(secs) if secs else 0
            total_cell_val = f'<c r="{total_r}" s="3" t="n"><v>{round(avg_secs) / 86400:.15f}</v></c>'
        row_cells += total_cell_val + '</row>'
        sheet1_data.append(row_cells)

    sheet1_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' \
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' \
        '<dimension ref="A1:' + get_cell_reference(len(METRICS_ORDER) + 1, total_col) + '"/>' \
        '<sheetViews><sheetView workbookViewId="0"/></sheetViews>' \
        '<sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>' \
        '<cols><col min="1" max="1" width="22" customWidth="1"/>'
    for i in range(2, total_col + 1):
        sheet1_xml += f'<col min="{i}" max="{i}" width="16" customWidth="1"/>'
    sheet1_xml += '</cols><sheetData>'
    sheet1_xml += ''.join(sheet1_data)
    sheet1_xml += '</sheetData></worksheet>'

    # Month tabs
    month_sheets = []
    for sheet_idx, (sheet_name, paths, counts, num_paths) in enumerate(chart_info, start=2):
        rows = []
        # Header
        rows.append(f'<row r="1" spans="1:2"><c r="A1" s="1" t="s"><v>{str_to_idx["Path"]}</v></c><c r="B1" s="1" t="s"><v>{str_to_idx["Count"]}</v></c></row>')
        # Data rows
        for pi, (path, count) in enumerate(zip(paths, counts), start=2):
            rows.append(f'<row r="{pi}" spans="1:2"><c r="A{pi}" s="3" t="s"><v>{str_to_idx[path]}</v></c><c r="B{pi}" s="3" t="n"><v>{count}</v></c></row>')

        sheet_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' \
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' \
            f'<dimension ref="A1:B{num_paths + 1}"/>' \
            '<sheetViews><sheetView workbookViewId="0"/></sheetViews>' \
            '<sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>' \
            '<cols><col min="1" max="1" width="80" customWidth="1"/><col min="2" max="2" width="10" customWidth="1"/></cols>' \
            '<sheetData>' + ''.join(rows) + '</sheetData></worksheet>'
        month_sheets.append((f'sheet{sheet_idx}.xml', sheet_xml))

    # Write everything to xlsx
    tmp = output_path + ".tmp"
    with zipfile.ZipFile(tmp, 'w') as tout:
        def w(name, content):
            if isinstance(content, bytes):
                tout.writestr(name, content)
            else:
                tout.writestr(name, content.encode('utf-8'))

        w('[Content_Types].xml', _make_content_types(num_sheets))
        w('_rels/.rels', root_rels)
        w('docProps/app.xml', docprops_app)
        w('docProps/core.xml', docprops_core)
        w('xl/workbook.xml', workbook_xml)
        w('xl/_rels/workbook.xml.rels', workbook_rels)
        w('xl/styles.xml', styles_xml)

        # Minimal theme
        theme_content = (
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
            '<a:fmtScheme name="Office">'
            '<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            '<a:gradFill rotWithShape="1"><a:gsLst>'
            '<a:gs pos="0"><a:schemeClr val="phClr"><a:alpha val="50000"/></a:schemeClr></a:gs>'
            '<a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="300000"/><a:alpha val="5000"/></a:schemeClr></a:gs>'
            '<a:gs pos="100000"><a:schemeClr val="phClr"><a:alpha val="0"/></a:schemeClr></a:gs>'
            '</a:gsLst><a:lin ang="16000000" scaled="1"/></a:gradFill></a:fillStyleLst>'
            '<a:lnStyleLst>'
            '<a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
            '<a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
            '<a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
            '<a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
            '<a:ln w="31750"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
            '<a:ln w="38100"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
            '<a:ln w="44450"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
            '<a:ln w="50800"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
            '<a:ln w="63500"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
            '</a:lnStyleLst>'
            '<a:effectStyleLst>'
            '<a:effectStyle><a:effectLst/><a:spPr/><a:lnSpPr/></a:effectStyle>'
            '<a:effectStyle><a:effectLst/><a:spPr/><a:lnSpPr/></a:effectStyle>'
            '<a:effectStyle><a:effectLst/><a:spPr/><a:lnSpPr/></a:effectStyle>'
            '</a:effectStyleLst>'
            '<a:bgFillStyleLst>'
            '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            '<a:solidFill><a:schemeClr val="phClr"/><a:tint val="35000"/><a:satMod val="150000"/></a:solidFill>'
            '<a:gradFill rotWithShape="1"><a:gsLst>'
            '<a:gs pos="0"><a:schemeClr val="phClr"/><a:tint val="100000"/><a:satMod val="300000"/></a:schemeClr></a:gs>'
            '<a:gs pos="50000"><a:schemeClr val="phClr"/><a:tint val="95000"/><a:satMod val="350000"/></a:schemeClr></a:gs>'
            '<a:gs pos="100000"><a:schemeClr val="phClr"/><a:tint val="90000"/><a:satMod val="400000"/><a:alpha val="30000"/></a:schemeClr></a:gs>'
            '</a:gsLst><a:lin ang="16000000" scaled="0"/></a:gradFill></a:bgFillStyleLst>'
            '</a:fmtScheme>'
            '</a:themeElements>'
            '<a:objectDefaults/><a:extraClrSchemeRels/></a:theme>'
        )
        w('xl/theme/theme1.xml', theme_content.encode('utf-8'))

        # Shared strings
        w('xl/sharedStrings.xml', shared_xml)

        # Worksheets
        w('xl/worksheets/sheet1.xml', sheet1_xml)
        for fname, content in month_sheets:
            w(f'xl/worksheets/{fname}', content)

    os.rename(tmp, output_path)
    print(f"  ✓ {output_path}")


def main():
    for batch_name, files in BATCHES.items():
        print(f"\nConsolidating: {batch_name}")
        out_name = f"{batch_name} - Consolidated.xlsx"
        out_path = os.path.join(OUT_DIR, out_name)
        build_xlsx_from_scratch(batch_name, files, out_path)

    print(f"\nDone! All consolidated files in: {OUT_DIR}")


if __name__ == "__main__":
    main()

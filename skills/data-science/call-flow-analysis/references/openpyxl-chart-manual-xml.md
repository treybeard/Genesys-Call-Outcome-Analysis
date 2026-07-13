# Hand-Crafted Chart XML for openpyxl Workbooks

## Problem

openpyxl's `BarChart` class produces Excel-incompatible chart XML:
1. Uses `<numRef>` for text categories instead of `<strRef>` with `<strCache>`
2. Uses default XML namespaces instead of prefixed `c:`, `a:`, `r:` namespaces
3. Omits required `a16:creationId`, `a:extLst`, `c:style`, `c:varyColors` elements
4. The resulting drawing XML lacks `xdr:` prefix namespaces and Office 2016+ compat elements

Excel silently rejects these charts — they don't render at all.

## Solution: Manual XML Generation

Generate chart XML by hand using the pattern below. The template builds every required element with correct namespace prefixes.

### Step 1: Build chart XML (write to `xl/charts/chart*.xml`)

```python
from xml.etree import ElementTree as ET
import uuid

C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A16_NS = "http://schemas.microsoft.com/office/drawing/2014/main"
C16_NS = "http://schemas.microsoft.com/office/drawing/2014/chart"
C16R2_NS = "http://schemas.microsoft.com/office/drawing/2015/06/chart"

def make_chart_xml(sheet_name, paths, counts):
    """Build a complete chart1.xml matching the reference format."""
    num = len(paths)
    root = ET.Element(f"{{{C_NS}}}chartSpace")
    root.set("xmlns:c", C_NS)
    root.set("xmlns:a", A_NS)
    root.set("xmlns:r", R_NS)
    root.set("xmlns:c16r2", C16R2_NS)

    ET.SubElement(root, f"{{{C_NS}}}style").set("val", "2")
    ET.SubElement(root, f"{{{C_NS}}}date1904").set("val", "0")
    ET.SubElement(root, f"{{{C_NS}}}lang").set("val", "en-US")
    ET.SubElement(root, f"{{{C_NS}}}roundedCorners").set("val", "1")

    chart = ET.SubElement(root, f"{{{C_NS}}}chart")

    # Title
    title_el = ET.SubElement(chart, f"{{{C_NS}}}title")
    tx = ET.SubElement(title_el, f"{{{C_NS}}}tx")
    rich = ET.SubElement(tx, f"{{{C_NS}}}rich")
    ET.SubElement(rich, f"{{{A_NS}}}bodyPr")
    ET.SubElement(rich, f"{{{A_NS}}}lstStyle")
    p = ET.SubElement(rich, f"{{{A_NS}}}p")
    p_pr = ET.SubElement(p, f"{{{A_NS}}}pPr")
    ET.SubElement(p_pr, f"{{{A_NS}}}defRPr")
    r = ET.SubElement(p, f"{{{A_NS}}}r")
    t = ET.SubElement(r, f"{{{A_NS}}}t")
    t.text = "Call Count / Call Path"
    ET.SubElement(title_el, f"{{{C_NS}}}overlay").set("val", "1")
    ET.SubElement(chart, f"{{{C_NS}}}autoTitleDeleted").set("val", "0")

    # Plot area
    plot_area = ET.SubElement(chart, f"{{{C_NS}}}plotArea")
    ET.SubElement(plot_area, f"{{{C_NS}}}layout")

    # Bar chart
    bar_chart = ET.SubElement(plot_area, f"{{{C_NS}}}barChart")
    ET.SubElement(bar_chart, f"{{{C_NS}}}barDir").set("val", "bar")
    ET.SubElement(bar_chart, f"{{{C_NS}}}grouping").set("val", "clustered")

    # Series
    ser = ET.SubElement(bar_chart, f"{{{C_NS}}}ser")
    ET.SubElement(ser, f"{{{C_NS}}}idx").set("val", "0")
    ET.SubElement(ser, f"{{{C_NS}}}order").set("val", "0")

    # spPr
    sp_pr = ET.SubElement(ser, f"{{{C_NS}}}spPr")
    ln = ET.SubElement(sp_pr, f"{{{A_NS}}}ln")
    ET.SubElement(ln, f"{{{A_NS}}}prstDash").set("val", "solid")
    ET.SubElement(ser, f"{{{C_NS}}}invertIfNegative").set("val", "1")

    # Category (strRef with embedded strCache)
    cat = ET.SubElement(ser, f"{{{C_NS}}}cat")
    str_ref = ET.SubElement(cat, f"{{{C_NS}}}strRef")
    f_el = ET.SubElement(str_ref, f"{{{C_NS}}}f")
    f_el.text = f"{sheet_name}!$A$2:$A${num + 1}"
    str_cache = ET.SubElement(str_ref, f"{{{C_NS}}}strCache")
    ET.SubElement(str_cache, f"{{{C_NS}}}ptCount").set("val", str(num))
    for idx, path in enumerate(paths):
        pt = ET.SubElement(str_cache, f"{{{C_NS}}}pt")
        pt.set("idx", str(idx))
        v = ET.SubElement(pt, f"{{{C_NS}}}v")
        v.text = path

    # Value (numRef with embedded numCache)
    val = ET.SubElement(ser, f"{{{C_NS}}}val")
    num_ref = ET.SubElement(val, f"{{{C_NS}}}numRef")
    f_el2 = ET.SubElement(num_ref, f"{{{C_NS}}}f")
    f_el2.text = f"{sheet_name}!$B$2:$B${num + 1}"
    num_cache = ET.SubElement(num_ref, f"{{{C_NS}}}numCache")
    ET.SubElement(num_cache, f"{{{C_NS}}}formatCode").text = "General"
    ET.SubElement(num_cache, f"{{{C_NS}}}ptCount").set("val", str(num))
    for idx, count in enumerate(counts):
        pt = ET.SubElement(num_cache, f"{{{C_NS}}}pt")
        pt.set("idx", str(idx))
        v = ET.SubElement(pt, f"{{{C_NS}}}v")
        v.text = str(count)

    # extLst with uniqueId
    ext_lst = ET.SubElement(ser, f"{{{C_NS}}}extLst")
    ext_el = ET.SubElement(ext_lst, f"{{{C_NS}}}ext")
    ext_el.set("uri", "{C3380CC4-5D6E-409C-BE32-E72D297353CC}")
    ET.SubElement(ext_el, f"{{{C16_NS}}}uniqueId").set("val", "{" + str(uuid.uuid4()).upper() + "}")

    # gapWidth
    gap = ET.SubElement(bar_chart, f"{{{C_NS}}}gapWidth")
    gap.set("val", "150")

    # Axes (catAx + valAx) - omit axes element
    # Build catAx and valAx similar to reference XML
    # ... (full implementation in csv_to_consolidated.py)
```

### Step 2: Build drawing XML (write to `xl/drawings/drawing*.xml`)

```python
drawing_namespaces = {
    'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'a16': 'http://schemas.microsoft.com/office/drawing/2014/main',
}

# D column = col index 3, row 2 = row index 1 (0-based)
cx = 540000000000       # Width D->R
cy = num_paths * 68580000000  # Height = rows x row-height

nsmap = ' '.join(f'xmlns:{k}="{v}"' for k, v in drawing_namespaces.items())
drawing = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr {nsmap}>
<xdr:oneCellAnchor>
<xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
<xdr:ext cx="{cx}" cy="{cy}"/>
<xdr:graphicFrame macro="">
<xdr:nvGraphicFramePr>
<xdr:cNvPr id="2" name="Chart 1">
<a:extLst>
<a:ext uri="{{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}}">
<a16:creationId xmlns:a16="{A16_NS}" id="{{00000000-0008-0000-0{chart_idx}0-000002000000}}"/>
</a:ext>
</a:extLst>
</xdr:cNvPr>
<xdr:cNvGraphicFramePr/>
</xdr:nvGraphicFramePr>
<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
<a:graphic>
<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
<c:chart xmlns:c="{C_NS}" xmlns:r="{R_NS}" r:id="rId1"/>
</a:graphicData>
</a:graphic>
</xdr:graphicFrame>
<xdr:clientData/>
</xdr:oneCellAnchor>
</xdr:wsDr>'''
```

### Step 3: Build relationship files

**Drawing rels** (`xl/drawings/_rels/drawing*.xml.rels`):
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
  Target="/xl/charts/chart{idx}.xml" Id="rId1"/>
</Relationships>
```

**Worksheet rels** (`xl/worksheets/_rels/sheet{i}.xml.rels`):
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
  Target="/xl/drawings/drawing{idx}.xml" Id="rId1"/>
</Relationships>
```

### Step 4: Rebuild the xlsx

```python
# 1. Save with openpyxl (for data cells only)
wb.save(output_path)

# 2. Re-open as ZIP, remove old chart/drawing/relation XMLs, write new ones
tmp = output_path + ".tmp"
with zipfile.ZipFile(output_path, 'r') as zin:
    with zipfile.ZipFile(tmp, 'w') as tout:
        # Copy everything except old charts/drawings/rels
        for item in zin.infolist():
            name = item.filename
            if (name.startswith('xl/charts/chart') and name.endswith('.xml')):
                continue
            if (name.startswith('xl/drawings/drawing') and name.endswith('.xml')):
                continue
            if name.startswith('xl/drawings/_rels/'):
                continue
            if name.startswith('xl/worksheets/_rels/'):
                continue
            tout.writestr(item, zin.read(item))

        # Write hand-crafted XMLs
        for idx, (sheet_name, paths, counts, num_paths) in enumerate(chart_info, start=1):
            chart_xml = make_chart_xml(sheet_name, paths, counts)
            tout.writestr(f'xl/charts/chart{idx}.xml', chart_xml)

            drawing = build_drawing_xml(idx, num_paths, A16_NS, C_NS, R_NS)
            tout.writestr(f'xl/drawings/drawing{idx}.xml', drawing)

            drawing_rels = '...rels XML...'
            tout.writestr(f'xl/drawings/_rels/drawing{idx}.xml.rels', drawing_rels)

            sheet_num = idx + 1  # Summary=1, April=2, etc.
            ws_rels = '...rels XML...'
            tout.writestr(f'xl/worksheets/_rels/sheet{sheet_num}.xml.rels', ws_rels)

os.unlink(output_path)
os.rename(tmp, output_path)
```

## Key Namespace/Element Checklist

Every working chart needs:

- **chartSpace root**: `xmlns:c`, `xmlns:a`, `xmlns:r`, `xmlns:c16r2`
- **chartSpace children**: `c:style val="2"`, `c:date1904 val="0"`, `c:lang val="en-US"`, `c:roundedCorners val="1"`
- **Series**: `c:spPr`, `c:invertIfNegative val="1"`, `c:extLst` with `c16:uniqueId`
- **Categories**: `c:strRef` -> `c:f` + `c:strCache` -> `c:ptCount` + `c:pt idx/N/v` (one per row)
- **Values**: `c:numRef` -> `c:f` + `c:numCache` -> `c:formatCode General` + `c:ptCount` + `c:pt idx/N/v` (one per row)
- **Axes**: `c:catAx` with `c:title` for "Call Path", `c:valAx` with `c:title` for "Count" + `c:majorGridlines`
- **Chart layout**: `c:legend` with `legendPos val="r"`, `c:plotVisOnly val="1"`, `c:dispBlanksAs val="gap"`

Drawing XML needs:
- `xdr:` namespace prefix (not default namespace)
- `xdr:oneCellAnchor` with `xdr:col>3</xdr:col>` (D column) and `xdr:row>1</xdr:row>` (row 2)
- `xdr:ext cx` and `xdr:ext cy` as plain integers (no scientific notation)
- `a:extLst` -> `a:ext uri="{FF2B5EF4-...}"` -> `a16:creationId`
- `a:graphic` -> `a:graphicData uri="chart"` -> `c:chart r:id="rId1"`

## EMU Dimensions

| Property | Value | Notes |
|----------|-------|-------|
| Width | 540000000000 EMU | Columns D->R (15 columns) |
| Height per row | 68580000000 EMU | One data row height |
| Total height | num_paths x 68580000000 | Varies per tab |
| Anchor | D2 | col=3, row=1 (0-based) |

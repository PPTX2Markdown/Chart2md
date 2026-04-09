# chart2md

Convert OOXML charts (DrawingML and chartEx) to Markdown tables. Supports `.pptx`, `.xlsx`, and `.docx` files with no external dependencies.

## Installation

```bash
pip install chart2md
```

## Quick Start

```python
from chart2md import convert_chart, load_chart_parts

for root, ctx in load_chart_parts("presentation.pptx"):
    print(convert_chart(root, ctx))
```

Output:

```
**Chart type:** Column Chart

|        | Series 1 | Series 2 | Series 3 |
| ------ | -------- | -------- | -------- |
| Item 1 | 4.3      | 2.4      | 2        |
| Item 2 | 2.5      | 4.4      | 2        |
```

## CLI

```bash
chart2md input.pptx                        # print all charts to stdout
chart2md input.pptx -o output.md           # save to file
chart2md input.pptx --chart-source excel   # use embedded Excel data for DrawingML charts
chart2md input.pptx --chartex-source xml   # use XML data for chartEx charts
```

## API

### `load_chart_parts(path)`

Scans an OOXML file and returns a list of `(root, ctx)` pairs, one per chart.
`root` is an `ET.Element` (the chart XML root) and `ctx` is a `ZipContext`
that the converter uses to access resources inside the archive.

For `.pptx` files, slide order is preserved. For `.xlsx` and `.docx`, charts
are returned in filename sort order.

```python
from chart2md import load_chart_parts, convert_chart

for root, ctx in load_chart_parts("presentation.pptx"):
    print(convert_chart(root, ctx))
```

### `convert_chart(root, ctx, chart_source="xml", chartex_source="excel")`

Converts a chart XML root element to a Markdown table string.
Automatically detects whether the chart is a traditional DrawingML chart
(`c:chartSpace`) or a modern chartEx chart (`cx:chartSpace`).

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `root` | `ET.Element` | — | Chart XML root returned by `load_chart_parts` or resolved from a slide |
| `ctx` | `ZipContext \| None` | — | Context object for the archive. Pass `None` to skip embedded resource lookups |
| `chart_source` | `str` | `"xml"` | Data source for DrawingML charts: `"xml"` or `"excel"` |
| `chartex_source` | `str` | `"excel"` | Data source for chartEx charts: `"excel"` or `"xml"` |

### `ZipContext`

OOXML files (`.pptx`, `.xlsx`, `.docx`) are ZIP archives that contain many
XML files inside. `ZipContext` pairs an open `zipfile.ZipFile` with the path
of a specific XML file within the archive, so the converter can follow
internal references (e.g. a chart linking to an embedded Excel workbook).

When you use `load_chart_parts()`, `ZipContext` objects are created and
returned automatically. You only need to construct one manually when building
a custom pipeline (see below).

```python
import zipfile
from chart2md import ZipContext

zf = zipfile.ZipFile("presentation.pptx")
ctx = ZipContext(zf, "ppt/charts/chart1.xml")
```

## Advanced: Full Pipeline Integration

`load_chart_parts()` is convenient but returns charts without slide context.
When you need to convert an entire PPTX in slide order — preserving reading
order and position — iterate the slides manually:

```python
import posixpath
import zipfile
import xml.etree.ElementTree as ET
from chart2md import convert_chart, ZipContext

PML_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
DML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _read_rels(zf, xml_path):
    """Read the .rels file for a given XML part and return {rId: resolved_path}."""
    directory = posixpath.dirname(xml_path)
    filename = posixpath.basename(xml_path)
    rels_path = posixpath.join(directory, "_rels", filename + ".rels")
    result = {}
    try:
        for rel in ET.fromstring(zf.read(rels_path)):
            tag = rel.tag.split("}")[-1] if "}" in rel.tag else rel.tag
            if tag != "Relationship":
                continue
            rid = rel.get("Id", "")
            target = rel.get("Target", "")
            if rel.get("TargetMode") == "External" or not rid:
                continue
            if target.startswith("/"):
                resolved = target.lstrip("/")
            else:
                resolved = posixpath.normpath(
                    posixpath.join(directory, target)
                ).lstrip("/")
            result[rid] = resolved
    except KeyError:
        pass
    return result


with zipfile.ZipFile("presentation.pptx") as zf:
    # 1. Read slide order from presentation.xml
    prs = ET.fromstring(zf.read("ppt/presentation.xml"))
    prs_rels = _read_rels(zf, "ppt/presentation.xml")

    for sld_id_el in prs.findall(f".//{{{PML_NS}}}sldIdLst/{{{PML_NS}}}sldId"):
        rid = sld_id_el.get(f"{{{REL_NS}}}id")
        slide_path = prs_rels.get(rid or "")
        if not slide_path:
            continue

        slide = ET.fromstring(zf.read(slide_path))
        slide_rels = _read_rels(zf, slide_path)

        # 2. Find graphicFrame shapes that contain charts
        for gf in slide.iter():
            if gf.tag.split("}")[-1] != "graphicFrame":
                continue

            graphic = gf.find(f".//{{{DML_NS}}}graphic")
            if graphic is None:
                continue
            graphic_data = graphic.find(f"{{{DML_NS}}}graphicData")
            if graphic_data is None:
                continue

            # 3. Chart is identified by "chart" in the graphicData uri
            if "chart" not in graphic_data.get("uri", ""):
                continue

            # 4. Find the chart element (c:chart or cx:chart) and extract r:id
            chart_el = None
            for child in graphic_data.iter():
                if child.tag.split("}")[-1] == "chart":
                    chart_el = child
                    break
            if chart_el is None:
                continue

            chart_rid = None
            for attr, val in chart_el.attrib.items():
                if attr.endswith("}id"):
                    chart_rid = val
                    break
            if not chart_rid:
                continue

            # 5. Resolve the chart file path and convert
            chart_path = slide_rels.get(chart_rid)
            if not chart_path:
                continue

            chart_root = ET.fromstring(zf.read(chart_path))
            ctx = ZipContext(zf, chart_path)
            print(convert_chart(chart_root, ctx))
```

## Supported Chart Types

**DrawingML (`c:chartSpace`):** bar, line, pie, area, scatter, bubble, radar, stock, surface, and more

**chartEx (`cx:chartSpace`):** waterfall, funnel, treemap, sunburst, histogram, box & whisker, region map, pareto

## License

Apache 2.0 — Copyright 2026 INSEONG LEE

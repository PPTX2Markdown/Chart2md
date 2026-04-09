"""chart2md — Convert OOXML charts to Markdown tables.

Basic usage:

    from chart2md import convert_chart, load_chart_parts

    for root, ctx in load_chart_parts("file.pptx"):
        print(convert_chart(root, ctx))

For type-specific converters:

    from chart2md import convert_chartex   # modern chartEx charts (cx:chartSpace)
"""

from __future__ import annotations

import argparse
import sys
import xml.etree.ElementTree as ET
from typing import Optional

from .chart2md import convert_chart as _convert_drawingml, load_chart_parts
from .chartex2md import convert_chartex
from .excel2md import convert_excel_to_table
from .ooxml_context import ZipContext, iter_parts_matching

# cx:chartSpace namespace identifier
_CHARTEX_NS = "http://schemas.microsoft.com/office/drawing/2014/chartex"


def convert_chart(
    root: ET.Element,
    ctx: Optional[ZipContext] = None,
    chart_source: str = "xml",
    chartex_source: str = "excel",
) -> str:
    """Convert a chart XML root element to a Markdown table string.

    Automatically detects whether the root is a traditional DrawingML chart
    (``c:chartSpace``) or a modern chartEx chart (``cx:chartSpace``) and
    dispatches to the appropriate converter.

    Args:
        root: ``c:chartSpace`` or ``cx:chartSpace`` root ``ET.Element``.
        ctx: ``ZipContext`` wrapping the source OOXML archive. Pass ``None``
            to skip relationship resolution (e.g. linked Excel workbooks).
        chart_source: Data source for traditional DrawingML charts.
            ``"xml"`` (default) — parse the embedded DrawingML XML cache.
            ``"excel"`` — prefer the embedded ``.xlsx`` workbook, fall back to XML.
        chartex_source: Data source for modern chartEx charts.
            ``"excel"`` (default) — prefer the embedded ``.xlsx`` workbook, fall back to XML.
            ``"xml"`` — parse the XML cache directly.

    Returns:
        Markdown table string.
    """
    if root.tag.startswith(f"{{{_CHARTEX_NS}}}"):
        return convert_chartex(root, ctx, chartex_source)
    return _convert_drawingml(root, ctx, chart_source)


def main() -> None:
    """CLI 진입점 — chart2md [options] input."""
    parser = argparse.ArgumentParser(
        description="Convert ECMA-376 chart XML to Markdown tables.")
    parser.add_argument('input', help="Input file (.xml, .pptx, .xlsx, .docx)")
    parser.add_argument('-o', '--output', help="Output file (default: stdout)")
    parser.add_argument(
        '--chart-source', choices=('xml', 'excel'), default='xml',
        help="전통 DrawingML 차트 데이터 소스 (기본값: xml).",
    )
    parser.add_argument(
        '--chartex-source', choices=('xml', 'excel'), default='excel',
        help="chartEx 차트 데이터 소스 (기본값: excel).",
    )
    args = parser.parse_args()

    contexts = load_chart_parts(args.input)
    if not contexts:
        print("No chart data found.", file=sys.stderr)
        sys.exit(1)

    parts = []
    for root, ctx in contexts:
        parts.append(convert_chart(root, ctx,
                                   chart_source=args.chart_source,
                                   chartex_source=args.chartex_source))

    # 공유 ZipFile 닫기
    seen_zf: set[int] = set()
    for _, ctx in contexts:
        if ctx is not None and id(ctx.zf) not in seen_zf:
            seen_zf.add(id(ctx.zf))
            try:
                ctx.zf.close()
            except Exception:
                pass

    output = "\n---\n\n".join(parts)
    if args.output:
        with open(args.output, 'w', encoding='utf-8') as f:
            f.write(output)
    else:
        print(output, end='')


__all__ = [
    # CLI
    "main",
    # main entry point
    "convert_chart",
    "load_chart_parts",
    # type-specific converters
    "convert_chartex",
    # utilities
    "convert_excel_to_table",
    "ZipContext",
    "iter_parts_matching",
]

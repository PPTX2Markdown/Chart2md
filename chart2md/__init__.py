"""chart2md — OOXML DrawingML 차트를 Markdown 테이블로 변환하는 패키지.

일반적인 사용:

    from chart2md import convert, load_chart_parts

    for root, ctx in load_chart_parts("file.pptx"):
        print(convert(root, ctx))

타입별 직접 호출이 필요할 때:

    from chart2md import convert_chart     # 전통 DrawingML 차트 (c:chartSpace)
    from chart2md import convert_chartex   # 현대 chartEx 차트 (cx:chartSpace)
"""

from __future__ import annotations

import argparse
import sys
import xml.etree.ElementTree as ET
from typing import Optional

from .chart2md import convert_chart, load_chart_parts
from .chartex2md import convert_chartex
from .excel2md import convert_excel_to_table
from .ooxml_context import ZipContext, iter_parts_matching

# cx:chartSpace 네임스페이스 식별자
_CHARTEX_NS = "http://schemas.microsoft.com/office/drawing/2014/chartex"


def convert(
    root: ET.Element,
    ctx: Optional[ZipContext] = None,
    chart_source: str = "xml",
    chartex_source: str = "excel",
) -> str:
    """차트 XML 루트 Element를 Markdown 테이블 문자열로 변환한다.

    루트 태그의 네임스페이스를 보고 전통 DrawingML 차트(c:chartSpace)와
    현대 chartEx 차트(cx:chartSpace)를 자동 판별해 적절한 변환기로 디스패치한다.

    Args:
        root:           c:chartSpace 또는 cx:chartSpace 루트 ET.Element.
        ctx:            ZipContext 인스턴스. None이면 관계 참조를 건너뛴다.
        chart_source:   전통 DrawingML 차트의 데이터 소스.
                        "xml" (기본값) — DrawingML XML 캐시 파싱.
                        "excel"        — 내장 .xlsx 우선, 없으면 xml 폴백.
        chartex_source: 현대 chartEx 차트의 데이터 소스.
                        "excel" (기본값) — 내장 .xlsx 우선, 없으면 xml 폴백.
                        "xml"           — XML 캐시(lvl/pt) 직접 파싱.

    Returns:
        Markdown 테이블 문자열.
    """
    if root.tag.startswith(f"{{{_CHARTEX_NS}}}"):
        return convert_chartex(root, ctx, chartex_source)
    return convert_chart(root, ctx, chart_source)


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
        parts.append(convert(root, ctx,
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
    # 통합 entry point (권장)
    "convert",
    "load_chart_parts",
    # 타입별 직접 호출
    "convert_chart",
    "convert_chartex",
    # 유틸리티
    "convert_excel_to_table",
    "ZipContext",
    "iter_parts_matching",
]

"""chart2md_example.py — chart.pptx에 포함된 차트를 슬라이드 순서대로 Markdown으로 변환하는 예제.

슬라이드 순서 보장:
    presentation.xml → sldIdLst(슬라이드 순서)
    → 각 슬라이드 XML → p:graphicFrame → a:graphicData[@uri]
    → 슬라이드 .rels → r:id resolve → chart XML 경로
    → convert(root, ctx)

실행:
    python chart2md_example.py              # 결과를 stdout에 출력
    python chart2md_example.py -o out.md    # 결과를 파일로 저장
"""

from __future__ import annotations

import argparse
import posixpath
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from chart2md import convert, ZipContext

PPTX_PATH = Path(__file__).parent / "chart.pptx"

# ---------------------------------------------------------------------------
# 네임스페이스
# ---------------------------------------------------------------------------
NS_R   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_P   = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A   = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_C   = "http://schemas.openxmlformats.org/drawingml/2006/chart"

CHART_REL     = f"{NS_R}/chart"
CHARTEX_URIS  = {
    "http://schemas.microsoft.com/office/drawing/2014/chartex",
    "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex",
}


# ---------------------------------------------------------------------------
# 헬퍼
# ---------------------------------------------------------------------------

def _rels_path(part_path: str) -> str:
    """파트 경로로부터 .rels 경로를 반환한다."""
    d, f = posixpath.split(part_path)
    return posixpath.join(d, "_rels", f + ".rels")


def _load_rels(zf: zipfile.ZipFile, part_path: str) -> dict[str, str]:
    """파트의 .rels를 읽어 {rId: 절대 경로} 매핑을 반환한다."""
    rels_path = _rels_path(part_path)
    try:
        root = ET.fromstring(zf.read(rels_path))
    except KeyError:
        return {}
    result: dict[str, str] = {}
    for rel in root:
        rid    = rel.get("Id", "")
        target = rel.get("Target", "")
        if rel.get("TargetMode") == "External" or not rid:
            continue
        if target.startswith("/"):
            resolved = target.lstrip("/")
        else:
            base_dir = posixpath.dirname(part_path)
            resolved = posixpath.normpath(posixpath.join(base_dir, target))
        result[rid] = resolved
    return result


def _slide_paths(zf: zipfile.ZipFile) -> list[str]:
    """presentation.xml의 sldIdLst 순서대로 슬라이드 경로 목록을 반환한다."""
    prs_path = "ppt/presentation.xml"
    prs_root = ET.fromstring(zf.read(prs_path))
    prs_rels = _load_rels(zf, prs_path)

    sld_id_lst = prs_root.find(f"{{{NS_P}}}sldIdLst")
    if sld_id_lst is None:
        return []
    return [
        prs_rels[sld.get(f"{{{NS_R}}}id")]
        for sld in sld_id_lst
        if sld.get(f"{{{NS_R}}}id") in prs_rels
    ]


# ---------------------------------------------------------------------------
# 슬라이드 순서 기반 차트 변환
# ---------------------------------------------------------------------------

def iter_charts(pptx_path: str):
    """pptx 내 차트를 슬라이드 순서대로 순회하며 (markdown, slide_path) 를 yield한다."""
    zf = zipfile.ZipFile(pptx_path)

    for slide_path in _slide_paths(zf):
        try:
            slide_root = ET.fromstring(zf.read(slide_path))
        except (KeyError, ET.ParseError):
            continue

        slide_rels = _load_rels(zf, slide_path)

        for gf in slide_root.iter(f"{{{NS_P}}}graphicFrame"):
            graphic_data = gf.find(
                f"{{{NS_A}}}graphic/{{{NS_A}}}graphicData"
            )
            if graphic_data is None:
                continue

            uri = graphic_data.get("uri", "")
            if "chart" not in uri:
                continue

            is_chartex = uri in CHARTEX_URIS

            # c:chart 또는 cx:chart 요소에서 r:id 추출
            rid = None
            for el in graphic_data.iter():
                local = el.tag.split("}")[-1] if "}" in el.tag else el.tag
                if local == "chart":
                    for attr, val in el.attrib.items():
                        if attr.endswith("}id") or attr == "r:id":
                            rid = val
                            break
                if rid:
                    break

            if not rid or rid not in slide_rels:
                continue

            chart_path = slide_rels[rid]
            try:
                chart_root = ET.fromstring(zf.read(chart_path))
            except (KeyError, ET.ParseError):
                continue

            ctx = ZipContext(zf, chart_path)
            try:
                md = convert(chart_root, ctx)
                yield md.strip(), slide_path
            except Exception as exc:
                yield f"> [Error: {chart_path}: {exc}]", slide_path


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main() -> int:
    parser = argparse.ArgumentParser(
        description="Convert all charts in chart.pptx to Markdown tables (slide order)."
    )
    parser.add_argument("-o", "--output", metavar="FILE",
                        help="Output file (default: stdout)")
    args = parser.parse_args()

    parts = []
    for md, slide_path in iter_charts(str(PPTX_PATH)):
        slide_name = posixpath.basename(slide_path)
        parts.append(f"<!-- {slide_name} -->\n\n{md}")

    output = "\n\n---\n\n".join(parts) + "\n"

    if args.output:
        Path(args.output).write_text(output, encoding="utf-8")
        print(f"Saved to {args.output}")
    else:
        print(output, end="")

    return 0


if __name__ == "__main__":
    sys.exit(main())

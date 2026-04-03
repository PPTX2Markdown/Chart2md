"""Convert embedded Excel worksheets (xlsx) to Markdown tables.

Used by the chart pipeline when ``chart_source='excel'``: the xlsx workbook
that PowerPoint embeds alongside a chart shape is parsed and its first
data-sheet is rendered as a GFM pipe table.

This module has **no external dependencies** — only the Python standard library
is used (zipfile + xml.etree.ElementTree).
"""
from __future__ import annotations
import io
import re
import zipfile
import xml.etree.ElementTree as ET
from datetime import date, timedelta
from typing import Optional

__all__ = ["convert_excel_to_table"]

# ── Namespace constants ───────────────────────────────────────────────────────
# Transitional (Office 2007~) 네임스페이스 — 가장 일반적
_XL_NS_TRANSITIONAL = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
# Strict (ISO/IEC 29500 Strict) 네임스페이스
_XL_NS_STRICT = "http://purl.oclc.org/ooxml/spreadsheetml/main"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _detect_xl_ns(xml_bytes: bytes) -> str:
    """XML 선언 또는 루트 태그에서 SpreadsheetML 네임스페이스를 감지한다.

    Strict 네임스페이스가 발견되면 _XL_NS_STRICT, 아니면 _XL_NS_TRANSITIONAL 반환.
    """
    if b"purl.oclc.org/ooxml/spreadsheetml" in xml_bytes:
        return _XL_NS_STRICT
    return _XL_NS_TRANSITIONAL


# 기본값 (하위 호환 — 함수들은 _XL_NS 변수를 직접 참조하지 않음)
_XL_NS = _XL_NS_TRANSITIONAL


# ── Date helpers ─────────────────────────────────────────────────────────────

# ECMA-376 §18.8.30 기준 내장 날짜 numFmtId 집합
_BUILTIN_DATE_NUM_FMT_IDS = frozenset({14, 15, 16, 17, 22})

# Excel epoch: 1900 윤년 버그를 감안한 기준일
_EXCEL_EPOCH = date(1899, 12, 30)


def _is_date_format_code(fmt_code: str) -> bool:
    """커스텀 서식 코드가 날짜 형식인지 판별한다."""
    # 조건부 서식 [...] 및 문자열 리터럴 "..." 제거
    fmt_code = re.sub(r'\[.*?\]', '', fmt_code)
    fmt_code = re.sub(r'"[^"]*"', '', fmt_code)
    fmt_lower = fmt_code.lower()
    # 연도(y)와 일(d)이 함께 있으면 날짜 서식으로 판단
    return 'y' in fmt_lower and 'd' in fmt_lower


def _read_date_style_indices(xl: zipfile.ZipFile, ns: str) -> frozenset[int]:
    """date 서식이 적용된 cellXfs 인덱스 집합을 반환한다."""
    if "xl/styles.xml" not in xl.namelist():
        return frozenset()
    try:
        data = xl.read("xl/styles.xml")
        sty_ns = _detect_xl_ns(data)
        root = ET.fromstring(data)
    except (KeyError, ET.ParseError):
        return frozenset()

    # 커스텀 numFmt에서 날짜 서식 ID 수집
    custom_date_ids: set[int] = set()
    num_fmts = root.find(f"{{{sty_ns}}}numFmts")
    if num_fmts is not None:
        for nf in num_fmts.findall(f"{{{sty_ns}}}numFmt"):
            try:
                fmt_id = int(nf.get("numFmtId", -1))
                if _is_date_format_code(nf.get("formatCode", "")):
                    custom_date_ids.add(fmt_id)
            except (ValueError, TypeError):
                pass

    # cellXfs에서 날짜 서식 인덱스 수집
    date_indices: set[int] = set()
    cell_xfs = root.find(f"{{{sty_ns}}}cellXfs")
    if cell_xfs is not None:
        for i, xf in enumerate(cell_xfs.findall(f"{{{sty_ns}}}xf")):
            try:
                fmt_id = int(xf.get("numFmtId", 0))
            except (ValueError, TypeError):
                continue
            if fmt_id in _BUILTIN_DATE_NUM_FMT_IDS or fmt_id in custom_date_ids:
                date_indices.add(i)

    return frozenset(date_indices)


def _serial_to_date_str(serial: float) -> str:
    """Excel 날짜 시리얼을 'YYYY-MM-DD' 문자열로 변환한다."""
    try:
        return (_EXCEL_EPOCH + timedelta(days=int(serial))).isoformat()
    except (ValueError, OverflowError):
        return str(int(serial))


# ── Coordinate helpers ────────────────────────────────────────────────────────

def _col_str_to_idx(col_str: str) -> int:
    """Convert column letters to 0-based index.  A→0, B→1, AA→26, …"""
    idx = 0
    for ch in col_str.upper():
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


_CELL_RE = re.compile(r"([A-Za-z]+)(\d+)")


def _cell_ref_to_rc(ref: str) -> tuple[int, int] | tuple[None, None]:
    m = _CELL_RE.match(ref)
    if not m:
        return None, None
    return int(m.group(2)) - 1, _col_str_to_idx(m.group(1))


# ── xlsx sub-parsers ──────────────────────────────────────────────────────────

def _read_shared_strings(xl: zipfile.ZipFile, ns: str) -> list[str]:
    """Parse xl/sharedStrings.xml → list of strings (index = string ID)."""
    if "xl/sharedStrings.xml" not in xl.namelist():
        return []
    data = xl.read("xl/sharedStrings.xml")
    # 실제 파일의 네임스페이스로 재감지 (sharedStrings가 다른 ns일 수 있음)
    ns = _detect_xl_ns(data)
    root = ET.fromstring(data)
    result: list[str] = []
    for si in root.findall(f"{{{ns}}}si"):
        t_el = si.find(f"{{{ns}}}t")
        if t_el is not None:
            result.append(t_el.text or "")
        else:
            # Rich-text: concatenate r/t children
            parts = []
            for r in si.findall(f"{{{ns}}}r"):
                t = r.find(f"{{{ns}}}t")
                if t is not None:
                    parts.append(t.text or "")
            result.append("".join(parts))
    return result


def _find_data_sheet_path(xl: zipfile.ZipFile, ns: str) -> Optional[str]:
    """Return the xl/ path of the first non-chart worksheet, or None."""
    try:
        wb_data = xl.read("xl/workbook.xml")
        wb_root = ET.fromstring(wb_data)
        wb_rels = ET.fromstring(xl.read("xl/_rels/workbook.xml.rels"))
    except (KeyError, ET.ParseError):
        return None

    # workbook.xml 네임스페이스 재감지
    wb_ns = _detect_xl_ns(wb_data)

    rid_to_target: dict[str, str] = {
        rel.get("Id", ""): rel.get("Target", "")
        for rel in wb_rels
    }

    for sheet in wb_root.findall(f".//{{{wb_ns}}}sheet"):
        if sheet.get("type", "") == "chart":
            continue
        rid = sheet.get(f"{{{_REL_NS}}}id") or sheet.get("r:id", "")
        target = rid_to_target.get(rid, "")
        if not target:
            continue
        # Target is relative to xl/
        if not target.startswith("xl/"):
            target = "xl/" + target.lstrip("/")
        if target in xl.namelist():
            return target
    return None


def _parse_worksheet(
    xl: zipfile.ZipFile,
    sheet_path: str,
    shared_strings: list[str],
    date_style_indices: frozenset[int] = frozenset(),
) -> list[list[str]]:
    """Parse a worksheet XML part → 2-D list of string values."""
    try:
        ws_data = xl.read(sheet_path)
        root = ET.fromstring(ws_data)
    except (KeyError, ET.ParseError):
        return []

    # 워크시트별 네임스페이스 감지 (Strict/Transitional 혼재 가능)
    ns = _detect_xl_ns(ws_data)

    cells: dict[tuple[int, int], str] = {}

    for row_el in root.findall(f".//{{{ns}}}row"):
        for c_el in row_el.findall(f"{{{ns}}}c"):
            ref = c_el.get("r", "")
            row_i, col_i = _cell_ref_to_rc(ref)
            if row_i is None:
                continue

            cell_type = c_el.get("t", "n")
            v_el = c_el.find(f"{{{ns}}}v")
            # Inline string (t="inlineStr"): is/t 우선, 없으면 is/r/t rich-text 연결
            is_node = c_el.find(f"{{{ns}}}is")
            if is_node is not None:
                is_el = is_node.find(f"{{{ns}}}t")
                if is_el is not None:
                    value = is_el.text or ""
                else:
                    # Rich inline string: is/r/t 연결 (ECMA-376 §18.3.1.4)
                    value = "".join(
                        t.text or ""
                        for t in is_node.findall(f".//{{{ns}}}t")
                    )
                cells[(row_i, col_i)] = value
                continue

            if v_el is None or not v_el.text:
                continue
            elif cell_type == "s":          # shared string index
                idx = int(v_el.text)
                value = shared_strings[idx] if idx < len(shared_strings) else ""
            elif cell_type in ("str", "e"): # formula string / error
                value = v_el.text
            elif cell_type == "b":          # boolean
                value = "TRUE" if v_el.text == "1" else "FALSE"
            else:                           # number (includes dates as serial)
                try:
                    f = float(v_el.text)
                    style_idx = c_el.get("s")
                    if style_idx is not None and int(style_idx) in date_style_indices:
                        value = _serial_to_date_str(f)
                    elif f == int(f):
                        value = str(int(f))
                    else:
                        value = str(f)  # Python 3 최단 표현 — 부동소수점 노이즈 제거
                except (ValueError, TypeError):
                    value = v_el.text

            cells[(row_i, col_i)] = value

    if not cells:
        return []

    min_row = min(r for r, _ in cells)
    max_row = max(r for r, _ in cells)
    min_col = min(c for _, c in cells)
    max_col = max(c for _, c in cells)

    grid: list[list[str]] = []
    for r in range(min_row, max_row + 1):
        row_vals = [cells.get((r, c), "") for c in range(min_col, max_col + 1)]
        grid.append(row_vals)

    # Drop trailing fully-empty rows
    while grid and all(v == "" for v in grid[-1]):
        grid.pop()

    # Drop trailing fully-empty columns
    if grid:
        max_used_col = max(
            next((len(row) - 1 - i for i, v in enumerate(reversed(row)) if v), 0)
            for row in grid
        )
        grid = [row[: max_used_col + 1] for row in grid]

    return grid


# ── Markdown rendering ────────────────────────────────────────────────────────

def _grid_to_markdown(grid: list[list[str]]) -> str:
    """Render 2-D string grid as a GFM pipe table."""
    if not grid:
        return ""

    ncols = max(len(row) for row in grid)
    for row in grid:
        while len(row) < ncols:
            row.append("")

    def _esc(s: str) -> str:
        return s.replace("|", "\\|").replace("\n", " ")

    lines = [
        "| " + " | ".join(_esc(v) for v in grid[0]) + " |",
        "| " + " | ".join("---" for _ in range(ncols)) + " |",
    ]
    for row in grid[1:]:
        lines.append("| " + " | ".join(_esc(v) for v in row) + " |")

    return "\n".join(lines)


# ── Public API ────────────────────────────────────────────────────────────────

def convert_excel_to_table(excel_bytes: bytes) -> str:
    """Parse an embedded Excel workbook and return a GFM markdown table.

    Reads the first non-chart worksheet found in the workbook.
    Returns an empty string when the workbook cannot be parsed or contains
    no usable data.
    """
    try:
        with zipfile.ZipFile(io.BytesIO(excel_bytes)) as xl:
            # workbook.xml 기준 네임스페이스 감지 (Strict / Transitional)
            try:
                wb_data = xl.read("xl/workbook.xml")
                ns = _detect_xl_ns(wb_data)
            except KeyError:
                ns = _XL_NS_TRANSITIONAL
            shared_strings = _read_shared_strings(xl, ns)
            sheet_path = _find_data_sheet_path(xl, ns)
            if sheet_path is None:
                return ""
            date_style_indices = _read_date_style_indices(xl, ns)
            grid = _parse_worksheet(xl, sheet_path, shared_strings, date_style_indices)
            return _grid_to_markdown(grid)
    except Exception:
        return ""

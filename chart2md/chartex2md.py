"""chartex2md.py — MS-OE376 chartEx.xsd XML → Markdown table converter.

Architecture mirrors chart2md.py: recursive-descent parser with one
function per XSD complex/group type.  Functions are named
  parse_cx_<CT|ST>_<TypeName>(node)
where cx = chartEx namespace (http://schemas.microsoft.com/office/drawing/2014/chartex).

Namespace handling: _get_tag() strips the URI so the same code works regardless
of namespace prefix.

모든 XSD complexType / simpleType 에 대응하는 parse 함수가 존재한다.
다만 Markdown 테이블 출력과 무관한 필드(시각 서식, 축 설정, 인쇄 설정 등)는
파싱은 하되 렌더러에서 사용하지 않는다 — 각 렌더러 호출 지점에 그 이유를
한글 주석으로 명시하였다.

Public API
----------
convert_chartex(root, ctx=None, chart_source="excel") -> str
    cx:chartSpace 루트 Element를 Markdown 테이블 문자열로 변환한다.
"""

from __future__ import annotations

import xml.etree.ElementTree as ET
from typing import Optional

from .excel2md import convert_excel_to_table
from .ooxml_context import ZipContext

__all__ = ["convert_chartex"]

# ==============================================================================
# 차트 타입 이름 매핑  (ST_SeriesLayout 열거값 → 표시 이름)
# XSD ST_SeriesLayout: boxWhisker, clusteredColumn, funnel, paretoLine,
#                      regionMap, sunburst, treemap, waterfall
# histogram 은 XSD enum 에 없으나 실제 파일에서 출현하므로 폴백 처리.
# ==============================================================================
_CHARTEX_LAYOUT_NAMES: dict[str, str] = {
    "boxWhisker":      "Box & Whisker Chart",
    "clusteredColumn": "Clustered Column Chart",
    "funnel":          "Funnel Chart",
    "paretoLine":      "Pareto Chart",
    "regionMap":       "Region Map Chart",
    "sunburst":        "Sunburst Chart",
    "treemap":         "Treemap Chart",
    "waterfall":       "Waterfall Chart",
    # histogram: XSD 미정의, 실제 파일 대응용
    "histogram":       "Histogram Chart",
}

# ==============================================================================
# Helper / Utility
# ==============================================================================

def _get_tag(node: ET.Element | None) -> str:
    """Namespace-prefixed 태그에서 로컬 태그 이름만 추출."""
    return node.tag.split('}')[-1] if node is not None else ""

def _get_child(node: ET.Element | None, tag: str) -> ET.Element | None:
    if node is None: return None
    for child in node:
        if _get_tag(child) == tag: return child
    return None

def _get_children(node: ET.Element | None, tag: str) -> list[ET.Element]:
    if node is None: return []
    return [child for child in node if _get_tag(child) == tag]

def _get_val(node: ET.Element | None, attr: str = 'val') -> str | None:
    if node is None: return None
    for k, v in node.attrib.items():
        local = k.split('}')[-1] if '}' in k else k
        if local == attr: return v
    return None

# ==============================================================================
# DrawingML 텍스트 추출  (a: namespace)
# a:CT_TextBody → plain text
# 시각 서식(rPr, pPr, bodyPr)은 파싱 생략 — Markdown 셀 값 출력에 불필요.
# ==============================================================================

def _parse_a_CT_TextBody(node: ET.Element | None) -> str:
    """a:txBody (CT_TextBody) — 단락들을 줄바꿈으로 연결."""
    if node is None: return ""
    parts = []
    for p in _get_children(node, 'p'):
        para_parts = []
        for child in p:
            tag = _get_tag(child)
            if tag == 'r':
                t = _get_child(child, 't')
                para_parts.append((t.text or "") if t is not None else "")
            elif tag == 'fld':
                t = _get_child(child, 't')
                if t is not None:
                    para_parts.append(t.text or "")
        parts.append("".join(para_parts))
    return "\n".join(parts).strip()

# ==============================================================================
# ST_ 단순 타입  (chartEx.xsd 1:1 대응)
# 모두 렌더러에서 미사용: 열거·범위 제약은 렌더링에 무관하므로 값 그대로 반환.
# ==============================================================================

def parse_cx_ST_ExtensionDropMode(val: str | None) -> str | None:
    """ST_ExtensionDropMode — 'never'|'onModelChange'|'onDataChange'|'always'.
    확장 드롭 모드 — 확장 메타 전용, Markdown 출력과 무관."""
    return val

def parse_cx_ST_DoubleOrAutomatic(val: str | None) -> str | None:
    """ST_DoubleOrAutomatic — xsd:double | 'auto'.
    축 범위 값 — 축 설정 전용, Markdown 테이블 출력에 불필요."""
    return val

def parse_cx_ST_FormulaDirection(val: str | None) -> str | None:
    """ST_FormulaDirection — 'col'|'row'.
    수식 방향 — 데이터 참조 메타, Markdown 테이블 출력에 불필요."""
    return val

def parse_cx_ST_DataId(val: str | None) -> str | None:
    """ST_DataId — xsd:unsignedInt.
    cx:data 요소 식별자. 시리즈와 데이터를 연결하는 데 사용되므로 파싱 유지."""
    return val

def parse_cx_ST_StringDimensionType(val: str | None) -> str | None:
    """ST_StringDimensionType — 'cat'|'colorStr'|'entityId'.
    문자열 차원의 역할 유형. 'cat'이 테이블 카테고리 열에 해당한다."""
    return val

def parse_cx_ST_NumericDimensionType(val: str | None) -> str | None:
    """ST_NumericDimensionType — 'val'|'x'|'y'|'size'|'colorVal'.
    숫자 차원의 역할 유형. 'val'이 테이블 값 열에 해당한다."""
    return val

def parse_cx_ST_SidePos(val: str | None) -> str | None:
    """ST_SidePos — 'l'|'t'|'r'|'b'.
    위치 열거값 — 시각 전용 (제목·범례 위치), Markdown 출력에 불필요."""
    return val

def parse_cx_ST_PosAlign(val: str | None) -> str | None:
    """ST_PosAlign — 'min'|'ctr'|'max'.
    정렬 열거값 — 시각 전용, Markdown 출력에 불필요."""
    return val

def parse_cx_ST_GapWidthRatio(val: str | None) -> str | None:
    """ST_GapWidthRatio — xsd:double(≥0) | 'auto'.
    카테고리 축 간격 비율 — 축 설정 전용, Markdown 출력에 불필요."""
    return val

def parse_cx_ST_ValueAxisUnit(val: str | None) -> str | None:
    """ST_ValueAxisUnit — xsd:double(>0) | 'auto'.
    값 축 단위 — 축 설정 전용, Markdown 출력에 불필요."""
    return val

def parse_cx_ST_AxisUnit(val: str | None) -> str | None:
    """ST_AxisUnit — 'hundreds'|'thousands'|...|'percentage'.
    표시 단위 — 축 단위 레이블, Markdown 출력에 불필요."""
    return val

def parse_cx_ST_TickMarksType(val: str | None) -> str | None:
    """ST_TickMarksType — 'in'|'out'|'cross'|'none'.
    눈금 표시 방식 — 시각 전용, Markdown 출력에 불필요."""
    return val

def parse_cx_ST_SeriesLayout(val: str | None) -> str | None:
    """ST_SeriesLayout — boxWhisker|clusteredColumn|funnel|paretoLine|
    regionMap|sunburst|treemap|waterfall.
    차트 레이아웃 종류. 렌더러에서 헤더 타입명 결정에 사용한다."""
    return val

def parse_cx_ST_ParentLabelLayout(val: str | None) -> str | None:
    """ST_ParentLabelLayout — 'none'|'banner'|'overlapping'.
    트리맵/선버스트 상위 레이블 배치 — 시각 전용, Markdown 출력에 불필요."""
    return val

def parse_cx_ST_RegionLabelLayout(val: str | None) -> str | None:
    """ST_RegionLabelLayout — 'none'|'bestFitOnly'|'showAll'.
    지역 맵 레이블 배치 — 시각 전용, Markdown 출력에 불필요."""
    return val

def parse_cx_ST_IntervalClosedSide(val: str | None) -> str | None:
    """ST_IntervalClosedSide — 'l'|'r'.
    히스토그램 구간 폐쇄 방향 — 시각 전용, Markdown 출력에 불필요."""
    return val

def parse_cx_ST_EntityType(val: str | None) -> str | None:
    """ST_EntityType — 'Address'|'AdminDistrict'|...
    지리 엔티티 유형 — 지도 차트 전용, Markdown 출력에 불필요."""
    return val

def parse_cx_ST_GeoProjectionType(val: str | None) -> str | None:
    """ST_GeoProjectionType — 'mercator'|'miller'|'robinson'|'albers'.
    지리 투영 방식 — 지도 차트 시각 전용, Markdown 출력에 불필요."""
    return val

def parse_cx_ST_GeoMappingLevel(val: str | None) -> str | None:
    """ST_GeoMappingLevel — 'dataOnly'|'postalCode'|...|'world'.
    지리 매핑 해상도 수준 — 지도 차트 시각 전용, Markdown 출력에 불필요."""
    return val

def parse_cx_ST_QuartileMethod(val: str | None) -> str | None:
    """ST_QuartileMethod — 'inclusive'|'exclusive'.
    사분위 계산 방식 — Box & Whisker 통계 설정, Markdown 출력에 불필요."""
    return val

def parse_cx_ST_DataLabelPos(val: str | None) -> str | None:
    """ST_DataLabelPos — 'bestFit'|'b'|'ctr'|'inBase'|'inEnd'|'l'|'outEnd'|'r'|'t'.
    데이터 레이블 위치 — 시각 전용, Markdown 출력에 불필요."""
    return val

def parse_cx_ST_ValueColorPositionCount(val: str | None) -> str | None:
    """ST_ValueColorPositionCount — 2|3.
    값 색상 위치 개수 — 색상 스케일 전용, Markdown 출력에 불필요."""
    return val

def parse_cx_ST_PageOrientation(val: str | None) -> str | None:
    """ST_PageOrientation — 'default'|'portrait'|'landscape'.
    인쇄 방향 — 인쇄 설정 전용, Markdown 출력에 불필요."""
    return val

# ==============================================================================
# CT_Extension / CT_ExtensionList / CT_FeatureExtension
# ==============================================================================

def parse_cx_CT_Extension(node: ET.Element | None) -> dict:
    """CT_Extension → dict  {uri, any}.
    렌더러 미사용: 확장 요소 — Markdown 출력과 무관."""
    if node is None: return {}
    return {'uri': _get_val(node, 'uri')}

def parse_cx_CT_ExtensionList(node: ET.Element | None) -> list[dict]:
    """CT_ExtensionList → list[CT_Extension].
    렌더러 미사용: 확장 목록 — Markdown 출력과 무관."""
    if node is None: return []
    return [parse_cx_CT_Extension(e) for e in _get_children(node, 'ext')]

def parse_cx_CT_FeatureExtension(node: ET.Element | None) -> dict:
    """CT_FeatureExtension → dict  {feature, drop, any}.
    렌더러 미사용: 기능 확장 메타 — Markdown 출력과 무관."""
    if node is None: return {}
    return {
        'feature': node.attrib.get('feature', ''),
        'drop':    node.attrib.get('drop', 'never'),
    }

# ==============================================================================
# CT_NumberFormat
# ==============================================================================

def parse_cx_CT_NumberFormat(node: ET.Element | None) -> dict:
    """CT_NumberFormat → dict  {formatCode, sourceLinked}.
    렌더러 미사용: 숫자 포맷 — Markdown 셀 값 출력에 포맷 재현 불필요."""
    if node is None: return {}
    sc = node.attrib.get('sourceLinked')
    return {
        'formatCode':   node.attrib.get('formatCode', ''),
        'sourceLinked': (sc.lower() in ('1', 'true')) if sc is not None else None,
    }

# ==============================================================================
# CT_RelId
# ==============================================================================

def parse_cx_CT_RelId(node: ET.Element | None) -> str | None:
    """CT_RelId → str | None   r:id 관계 식별자.
    외부 데이터 소스 참조에 사용한다."""
    if node is None: return None
    for k, v in node.attrib.items():
        local = k.split('}')[-1] if '}' in k else k
        if local == 'id': return v
    return None

# ==============================================================================
# CT_ExternalData
# ==============================================================================

def parse_cx_CT_ExternalData(node: ET.Element | None) -> dict:
    """CT_ExternalData → dict  {rid, autoUpdate}.
    렌더러 미사용: 외부 데이터 연결 정보 — 데이터는 직접 포함된 lvl/pt에서 추출."""
    if node is None: return {}
    rid = None
    for k, v in node.attrib.items():
        local = k.split('}')[-1] if '}' in k else k
        if local == 'id':
            rid = v
            break
    auto = node.attrib.get('autoUpdate')
    return {
        'rid':        rid,
        'autoUpdate': (auto.lower() in ('1', 'true')) if auto is not None else None,
    }

# ==============================================================================
# CT_Formula
# ==============================================================================

def parse_cx_CT_Formula(node: ET.Element | None) -> dict:
    """CT_Formula → dict  {text, dir}.
    수식 문자열과 방향(col/row). 데이터 캐시(lvl/pt)가 있으면 수식은 미사용이므로
    수식을 렌더러에서 직접 참조하지 않는다."""
    if node is None: return {'text': '', 'dir': 'col'}
    return {
        'text': (node.text or "").strip(),
        'dir':  node.attrib.get('dir', 'col'),
    }

# ==============================================================================
# CT_StringValue / CT_StringLevel / CT_StringDimension
# ==============================================================================

def parse_cx_CT_StringValue(node: ET.Element) -> tuple[int, str]:
    """CT_StringValue → (idx: int, value: str)."""
    idx = int(node.attrib.get('idx', '0'))
    val = (node.text or "").strip()
    return idx, val

def parse_cx_CT_StringLevel(node: ET.Element | None) -> list[str]:
    """CT_StringLevel → list[str]  (dense, ptCount 개수 보장).
    name 속성: 렌더러 미사용 — 열 레이블로 활용 가능하지만 현재 헤더는 자동 생성."""
    if node is None: return []
    ptcount_str = node.attrib.get('ptCount', '0')
    total = int(ptcount_str)
    mapping: dict[int, str] = {}
    for pt in _get_children(node, 'pt'):
        idx, val = parse_cx_CT_StringValue(pt)
        mapping[idx] = val
    # ptCount 가 0이거나 mapping이 비어 있으면 mapping에서 추론
    if total == 0 and mapping:
        total = max(mapping.keys()) + 1
    return [mapping.get(i, "") for i in range(total)]

def parse_cx_CT_StringDimension(node: ET.Element | None) -> dict:
    """CT_StringDimension → dict  {type, f, nf, levels: list[list[str]]}.
    XSD choice: (f + nf? + lvl*) | lvl+
    f(수식): 데이터 캐시(lvl)가 있을 때 미사용.
    nf(이름 수식): 렌더러 미사용 — 차원 이름 수식, 테이블 출력에 불필요."""
    if node is None: return {'type': 'cat', 'f': None, 'nf': None, 'levels': []}
    dim_type = node.attrib.get('type', 'cat')
    f_node  = _get_child(node, 'f')
    nf_node = _get_child(node, 'nf')
    lvl_nodes = _get_children(node, 'lvl')
    levels = [parse_cx_CT_StringLevel(lvl) for lvl in lvl_nodes]
    return {
        'type': dim_type,
        # f: 수식 — 캐시(lvl) 우선이므로 렌더러 미사용
        'f':  parse_cx_CT_Formula(f_node) if f_node is not None else None,
        # nf: 이름 수식 — 차원 이름 레이블, 현재 렌더러에서 미사용
        'nf': parse_cx_CT_Formula(nf_node) if nf_node is not None else None,
        'levels': levels,
    }

# ==============================================================================
# CT_NumericValue / CT_NumericLevel / CT_NumericDimension
# ==============================================================================

def parse_cx_CT_NumericValue(node: ET.Element) -> tuple[int, str]:
    """CT_NumericValue → (idx: int, value: str).
    XSD base: xsd:double — 원본 문자열 그대로 보존 (렌더러가 그대로 출력)."""
    idx = int(node.attrib.get('idx', '0'))
    val = (node.text or "").strip()
    return idx, val

def parse_cx_CT_NumericLevel(node: ET.Element | None) -> list[str]:
    """CT_NumericLevel → list[str]  (dense, ptCount 개수 보장).
    formatCode 속성: 렌더러 미사용 — Markdown 셀 값 출력에 포맷 재현 불필요.
    name 속성: 렌더러 미사용 — 열 레이블, 현재 헤더는 자동 생성."""
    if node is None: return []
    ptcount_str = node.attrib.get('ptCount', '0')
    total = int(ptcount_str)
    mapping: dict[int, str] = {}
    for pt in _get_children(node, 'pt'):
        idx, val = parse_cx_CT_NumericValue(pt)
        mapping[idx] = val
    if total == 0 and mapping:
        total = max(mapping.keys()) + 1
    return [mapping.get(i, "") for i in range(total)]

def parse_cx_CT_NumericDimension(node: ET.Element | None) -> dict:
    """CT_NumericDimension → dict  {type, f, nf, levels: list[list[str]]}.
    XSD choice: (f + nf? + lvl*) | lvl+
    f(수식): 데이터 캐시(lvl)가 있을 때 미사용.
    nf(이름 수식): 렌더러 미사용 — 차원 이름 수식, 테이블 출력에 불필요."""
    if node is None: return {'type': 'val', 'f': None, 'nf': None, 'levels': []}
    dim_type = node.attrib.get('type', 'val')
    f_node  = _get_child(node, 'f')
    nf_node = _get_child(node, 'nf')
    lvl_nodes = _get_children(node, 'lvl')
    levels = [parse_cx_CT_NumericLevel(lvl) for lvl in lvl_nodes]
    return {
        'type': dim_type,
        # f: 수식 — 캐시(lvl) 우선이므로 렌더러 미사용
        'f':  parse_cx_CT_Formula(f_node) if f_node is not None else None,
        # nf: 이름 수식 — 현재 렌더러에서 미사용
        'nf': parse_cx_CT_Formula(nf_node) if nf_node is not None else None,
        'levels': levels,
    }

# ==============================================================================
# CT_Data / CT_ChartData
# ==============================================================================

def parse_cx_CT_Data(node: ET.Element | None) -> dict:
    """CT_Data → dict  {id, strDims: list[dict], numDims: list[dict]}.
    XSD: choice(numDim | strDim)+ + extLst?
    extLst: 렌더러 미사용 — 확장 목록, Markdown 출력과 무관."""
    if node is None: return {'id': '', 'strDims': [], 'numDims': []}
    data_id = node.attrib.get('id', '')
    str_dims: list[dict] = []
    num_dims: list[dict] = []
    for child in node:
        tag = _get_tag(child)
        if tag == 'strDim':
            str_dims.append(parse_cx_CT_StringDimension(child))
        elif tag == 'numDim':
            num_dims.append(parse_cx_CT_NumericDimension(child))
        # extLst: 무시
    return {'id': data_id, 'strDims': str_dims, 'numDims': num_dims}

def parse_cx_CT_ChartData(node: ET.Element | None) -> dict[str, dict]:
    """CT_ChartData → {data_id: CT_Data dict}.
    XSD: externalData? + data+ + extLst?
    externalData: 렌더러 미사용 — 외부 데이터 참조, 캐시(lvl/pt) 우선.
    extLst: 렌더러 미사용 — 확장 목록."""
    if node is None: return {}
    # externalData 파싱 (렌더러 미사용)
    _ext_data = parse_cx_CT_ExternalData(_get_child(node, 'externalData'))  # noqa: F841
    result: dict[str, dict] = {}
    for data_node in _get_children(node, 'data'):
        data = parse_cx_CT_Data(data_node)
        result[data['id']] = data
    return result

# ==============================================================================
# CT_TextData / CT_Text
# ==============================================================================

def parse_cx_CT_TextData(node: ET.Element | None) -> str:
    """CT_TextData → str  (f + v?) | v.
    수식(f)이 있으면 그 텍스트를 반환하고, v(리터럴)가 있으면 v를 반환한다."""
    if node is None: return ""
    f_node = _get_child(node, 'f')
    v_node = _get_child(node, 'v')
    if f_node is not None:
        # f가 있으면 v는 캐시된 값
        val_text = (v_node.text or "").strip() if v_node is not None else ""
        return val_text or (f_node.text or "").strip()
    if v_node is not None:
        return (v_node.text or "").strip()
    return ""

def parse_cx_CT_Text(node: ET.Element | None) -> str:
    """CT_Text → str  (txData | rich).
    txData: 수식/리터럴 텍스트.
    rich: a:CT_TextBody — DrawingML 서식 텍스트 (plain text만 추출)."""
    if node is None: return ""
    txdata = _get_child(node, 'txData')
    if txdata is not None:
        return parse_cx_CT_TextData(txdata)
    rich = _get_child(node, 'rich')
    if rich is not None:
        return _parse_a_CT_TextBody(rich)
    return ""

# ==============================================================================
# CT_Offset
# ==============================================================================

def parse_cx_CT_Offset(node: ET.Element | None) -> dict:
    """CT_Offset → dict  {top, left}.
    렌더러 미사용: 제목·범례 위치 오프셋 — 시각 전용."""
    if node is None: return {}
    return {
        'top':  node.attrib.get('top'),
        'left': node.attrib.get('left'),
    }

# ==============================================================================
# Axis scaling
# CT_CategoryAxisScaling / CT_ValueAxisScaling
# ==============================================================================

def parse_cx_CT_CategoryAxisScaling(node: ET.Element | None) -> dict:
    """CT_CategoryAxisScaling → dict  {gapWidth}.
    렌더러 미사용: 카테고리 축 간격 설정 — 시각 전용."""
    if node is None: return {}
    return {'gapWidth': node.attrib.get('gapWidth')}

def parse_cx_CT_ValueAxisScaling(node: ET.Element | None) -> dict:
    """CT_ValueAxisScaling → dict  {max, min, majorUnit, minorUnit}.
    렌더러 미사용: 값 축 범위/단위 설정 — 시각 전용."""
    if node is None: return {}
    return {
        'max':       node.attrib.get('max'),
        'min':       node.attrib.get('min'),
        'majorUnit': node.attrib.get('majorUnit'),
        'minorUnit': node.attrib.get('minorUnit'),
    }

# ==============================================================================
# Axis title / units
# CT_AxisTitle / CT_AxisUnitsLabel / CT_AxisUnits
# ==============================================================================

def parse_cx_CT_AxisTitle(node: ET.Element | None) -> dict:
    """CT_AxisTitle → dict  {text, offset}.
    XSD: tx? + spPr? + txPr? + offset? + extLst?
    spPr(도형 서식), txPr(텍스트 서식): 렌더러 미사용 — 시각 전용.
    Markdown 테이블은 축 이름을 별도 출력하지 않으므로 text도 미사용."""
    if node is None: return {}
    return {
        # text: 축 이름 — Markdown 테이블에서 미사용
        'text':   parse_cx_CT_Text(_get_child(node, 'tx')),
        # spPr: a:CT_ShapeProperties — 시각 전용
        'spPr':   None,
        # txPr: a:CT_TextBody — 시각 전용
        'txPr':   None,
        # offset: 시각 전용
        'offset': parse_cx_CT_Offset(_get_child(node, 'offset')),
    }

def parse_cx_CT_AxisUnitsLabel(node: ET.Element | None) -> dict:
    """CT_AxisUnitsLabel → dict  {text}.
    XSD: tx? + spPr? + txPr? + extLst?
    단위 레이블 텍스트. Markdown 테이블은 축 단위를 별도 출력하지 않으므로 미사용."""
    if node is None: return {}
    return {
        'text': parse_cx_CT_Text(_get_child(node, 'tx')),
        'spPr': None,  # 시각 전용
        'txPr': None,  # 시각 전용
    }

def parse_cx_CT_AxisUnits(node: ET.Element | None) -> dict:
    """CT_AxisUnits → dict  {unit, unitsLabel}.
    렌더러 미사용: 축 단위 표시 설정 — 시각 전용."""
    if node is None: return {}
    return {
        # unit: ST_AxisUnit 열거값 — 시각 전용
        'unit':       node.attrib.get('unit'),
        'unitsLabel': parse_cx_CT_AxisUnitsLabel(_get_child(node, 'unitsLabel')),
    }

# ==============================================================================
# CT_Gridlines / CT_TickMarks / CT_TickLabels
# ==============================================================================

def parse_cx_CT_Gridlines(node: ET.Element | None) -> dict:
    """CT_Gridlines → dict  {spPr}.
    렌더러 미사용: 격자선 서식 — 시각 전용."""
    if node is None: return {}
    return {'spPr': None}  # spPr: a:CT_ShapeProperties — 시각 전용

def parse_cx_CT_TickMarks(node: ET.Element | None) -> dict:
    """CT_TickMarks → dict  {type}.
    렌더러 미사용: 눈금 표시 방식 — 시각 전용."""
    if node is None: return {}
    return {'type': node.attrib.get('type')}

def parse_cx_CT_TickLabels(node: ET.Element | None) -> dict:
    """CT_TickLabels → dict  (본문 없음, extLst만 포함).
    렌더러 미사용: 눈금 레이블 설정 — 시각 전용."""
    if node is None: return {}
    return {}

# ==============================================================================
# CT_Axis
# ==============================================================================

def parse_cx_CT_Axis(node: ET.Element | None) -> dict:
    """CT_Axis → dict
    XSD: (catScaling|valScaling) + title? + units? + majorGridlines? +
         minorGridlines? + majorTickMarks? + minorTickMarks? + tickLabels? +
         numFmt? + spPr? + txPr? + extLst?,  attrs: id(required), hidden.
    렌더러 미사용: 축 설정 전체 — Markdown 테이블은 축을 별도 출력하지 않는다."""
    if node is None: return {}
    # catScaling / valScaling: 어느 한 쪽만 존재
    cat_sc = _get_child(node, 'catScaling')
    val_sc = _get_child(node, 'valScaling')
    scaling = (parse_cx_CT_CategoryAxisScaling(cat_sc)
               if cat_sc is not None
               else parse_cx_CT_ValueAxisScaling(val_sc))
    return {
        'id':     node.attrib.get('id'),
        'hidden': node.attrib.get('hidden', '0'),
        'scaling': scaling,
        # title: 축 이름 — Markdown 테이블에서 미사용
        'title':         parse_cx_CT_AxisTitle(_get_child(node, 'title')),
        # units: 축 단위 — Markdown 테이블에서 미사용
        'units':         parse_cx_CT_AxisUnits(_get_child(node, 'units')),
        # gridlines: 시각 전용
        'majorGridlines': parse_cx_CT_Gridlines(_get_child(node, 'majorGridlines')),
        'minorGridlines': parse_cx_CT_Gridlines(_get_child(node, 'minorGridlines')),
        # tickMarks: 시각 전용
        'majorTickMarks': parse_cx_CT_TickMarks(_get_child(node, 'majorTickMarks')),
        'minorTickMarks': parse_cx_CT_TickMarks(_get_child(node, 'minorTickMarks')),
        # tickLabels: 시각 전용
        'tickLabels':    parse_cx_CT_TickLabels(_get_child(node, 'tickLabels')),
        # numFmt: 숫자 포맷 — Markdown 출력에서 포맷 재현 불필요
        'numFmt':        parse_cx_CT_NumberFormat(_get_child(node, 'numFmt')),
        # spPr, txPr: 시각 전용
        'spPr': None,
        'txPr': None,
    }

# ==============================================================================
# Series layout properties
# CT_ParentLabelLayout / CT_RegionLabelLayout / CT_SeriesElementVisibilities
# CT_Aggregation / CT_Binning / CT_Statistics
# CT_SubtotalIndex / CT_Subtotals / CT_SeriesLayoutProperties
# ==============================================================================

def parse_cx_CT_ParentLabelLayout(node: ET.Element | None) -> dict:
    """CT_ParentLabelLayout → dict  {val}.
    렌더러 미사용: 트리맵/선버스트 상위 레이블 배치 — 시각 전용."""
    if node is None: return {}
    return {'val': node.attrib.get('val')}

def parse_cx_CT_RegionLabelLayout(node: ET.Element | None) -> dict:
    """CT_RegionLabelLayout → dict  {val}.
    렌더러 미사용: 지역 맵 레이블 배치 — 시각 전용."""
    if node is None: return {}
    return {'val': node.attrib.get('val')}

def parse_cx_CT_SeriesElementVisibilities(node: ET.Element | None) -> dict:
    """CT_SeriesElementVisibilities → dict  {connectorLines, meanLine, meanMarker, nonoutliers, outliers}.
    렌더러 미사용: Box & Whisker 요소 표시 여부 — 시각 전용."""
    if node is None: return {}
    return {
        'connectorLines': node.attrib.get('connectorLines'),
        'meanLine':       node.attrib.get('meanLine'),
        'meanMarker':     node.attrib.get('meanMarker'),
        'nonoutliers':    node.attrib.get('nonoutliers'),
        'outliers':       node.attrib.get('outliers'),
    }

def parse_cx_CT_Aggregation(node: ET.Element | None) -> dict:
    """CT_Aggregation → dict  (empty complexType).
    렌더러 미사용: 히스토그램 집계 모드 표시 — 시각 전용."""
    return {}

def parse_cx_CT_Binning(node: ET.Element | None) -> dict:
    """CT_Binning → dict  {binSize|binCount, intervalClosed, underflow, overflow}.
    렌더러 미사용: 히스토그램 구간 설정 — 시각 전용."""
    if node is None: return {}
    bin_size  = _get_child(node, 'binSize')
    bin_count = _get_child(node, 'binCount')
    return {
        # binSize / binCount: 어느 한 쪽만 존재
        'binSize':        (bin_size.text or "").strip() if bin_size is not None else None,
        'binCount':       (bin_count.text or "").strip() if bin_count is not None else None,
        'intervalClosed': node.attrib.get('intervalClosed'),
        'underflow':      node.attrib.get('underflow'),
        'overflow':       node.attrib.get('overflow'),
    }

def parse_cx_CT_Statistics(node: ET.Element | None) -> dict:
    """CT_Statistics → dict  {quartileMethod}.
    렌더러 미사용: Box & Whisker 사분위 계산 방식 — 통계 설정."""
    if node is None: return {}
    return {'quartileMethod': node.attrib.get('quartileMethod')}

def parse_cx_CT_SubtotalIndex(node: ET.Element | None) -> int | None:
    """CT_SubtotalIndex → int | None  (val 속성).
    워터폴 차트에서 합계/소계 항목 인덱스. 렌더러에서 직접 미사용."""
    if node is None: return None
    v = node.attrib.get('val')
    return int(v) if v is not None else None

def parse_cx_CT_Subtotals(node: ET.Element | None) -> list[int]:
    """CT_Subtotals → list[int]  (합계/소계 포인트 인덱스 목록).
    렌더러 미사용: 워터폴 차트 소계 강조 — 시각 전용."""
    if node is None: return []
    return [i for i in (parse_cx_CT_SubtotalIndex(idx) for idx in _get_children(node, 'idx'))
            if i is not None]

def parse_cx_CT_SeriesLayoutProperties(node: ET.Element | None) -> dict:
    """CT_SeriesLayoutProperties → dict
    XSD: parentLabelLayout? + regionLabelLayout? + visibility? +
         (aggregation|binning)? + geography? + statistics? + subtotals? + extLst?
    렌더러 미사용: 시리즈 레이아웃별 추가 시각/통계 설정 — 테이블 데이터와 무관."""
    if node is None: return {}
    agg  = _get_child(node, 'aggregation')
    binn = _get_child(node, 'binning')
    return {
        # parentLabelLayout: 트리맵/선버스트 상위 레이블 — 시각 전용
        'parentLabelLayout':  parse_cx_CT_ParentLabelLayout(_get_child(node, 'parentLabelLayout')),
        # regionLabelLayout: 지역 맵 레이블 — 시각 전용
        'regionLabelLayout':  parse_cx_CT_RegionLabelLayout(_get_child(node, 'regionLabelLayout')),
        # visibility: Box & Whisker 요소 표시 — 시각 전용
        'visibility':         parse_cx_CT_SeriesElementVisibilities(_get_child(node, 'visibility')),
        # aggregation / binning: 히스토그램 설정 — 시각 전용
        'aggregation': parse_cx_CT_Aggregation(agg) if agg is not None else None,
        'binning':     parse_cx_CT_Binning(binn) if binn is not None else None,
        # geography: 지도 차트 지리 설정 — 시각 전용 (CT_Geography 참조)
        'geography':   None,
        # statistics: Box & Whisker 통계 설정 — 시각 전용
        'statistics':  parse_cx_CT_Statistics(_get_child(node, 'statistics')),
        # subtotals: 워터폴 소계 인덱스 — 시각 전용
        'subtotals':   parse_cx_CT_Subtotals(_get_child(node, 'subtotals')),
    }

# ==============================================================================
# CT_DataPoint
# ==============================================================================

def parse_cx_CT_DataPoint(node: ET.Element | None) -> dict:
    """CT_DataPoint → dict  {idx}.
    XSD: spPr? + extLst?, attr: idx(required).
    spPr: 시각 전용 — 특정 포인트 서식 재정의, Markdown 출력에 불필요."""
    if node is None: return {}
    return {
        'idx':  int(node.attrib.get('idx', '0')),
        'spPr': None,  # spPr: 시각 전용
    }

# ==============================================================================
# CT_DataLabelVisibilities / CT_DataLabel / CT_DataLabelHidden / CT_DataLabels
# ==============================================================================

def parse_cx_CT_DataLabelVisibilities(node: ET.Element | None) -> dict:
    """CT_DataLabelVisibilities → dict  {seriesName, categoryName, value}.
    렌더러 미사용: 레이블 표시 여부 — 시각 전용."""
    if node is None: return {}
    return {
        'seriesName':   node.attrib.get('seriesName'),
        'categoryName': node.attrib.get('categoryName'),
        'value':        node.attrib.get('value'),
    }

def parse_cx_CT_DataLabel(node: ET.Element | None) -> dict:
    """CT_DataLabel → dict  {idx, pos, numFmt, visibility, separator}.
    XSD: numFmt? + spPr? + txPr? + visibility? + separator? + extLst?,
         attrs: idx(required), pos?
    spPr, txPr: 시각 전용.
    렌더러 미사용: 개별 데이터 레이블 설정 — Markdown 테이블은 레이블을
    별도 열로 출력하지 않는다."""
    if node is None: return {}
    sep_node = _get_child(node, 'separator')
    return {
        'idx':        int(node.attrib.get('idx', '0')),
        'pos':        node.attrib.get('pos'),
        'numFmt':     parse_cx_CT_NumberFormat(_get_child(node, 'numFmt')),
        # spPr, txPr: 시각 전용
        'spPr':       None,
        'txPr':       None,
        'visibility': parse_cx_CT_DataLabelVisibilities(_get_child(node, 'visibility')),
        'separator':  (sep_node.text or "") if sep_node is not None else None,
    }

def parse_cx_CT_DataLabelHidden(node: ET.Element | None) -> dict:
    """CT_DataLabelHidden → dict  {idx}.
    렌더러 미사용: 숨겨진 레이블 인덱스 — 시각 전용."""
    if node is None: return {}
    return {'idx': int(node.attrib.get('idx', '0'))}

def parse_cx_CT_DataLabels(node: ET.Element | None) -> dict:
    """CT_DataLabels → dict
    XSD: numFmt? + spPr? + txPr? + visibility? + separator? +
         dataLabel* + dataLabelHidden* + extLst?, attr: pos?
    렌더러 미사용: 시리즈 전체 레이블 설정 — Markdown 테이블은 레이블을
    별도 열로 출력하지 않는다."""
    if node is None: return {}
    sep_node = _get_child(node, 'separator')
    return {
        'pos':              node.attrib.get('pos'),
        'numFmt':           parse_cx_CT_NumberFormat(_get_child(node, 'numFmt')),
        # spPr, txPr: 시각 전용
        'spPr':             None,
        'txPr':             None,
        'visibility':       parse_cx_CT_DataLabelVisibilities(_get_child(node, 'visibility')),
        'separator':        (sep_node.text or "") if sep_node is not None else None,
        'dataLabel':        [parse_cx_CT_DataLabel(dl) for dl in _get_children(node, 'dataLabel')],
        'dataLabelHidden':  [parse_cx_CT_DataLabelHidden(dlh)
                             for dlh in _get_children(node, 'dataLabelHidden')],
    }

# ==============================================================================
# Color scale types (시각 전용)
# CT_ValueColors / CT_ExtremeValueColorPosition / CT_NumberColorPosition /
# CT_PercentageColorPosition / CT_ValueColorEndPosition /
# CT_ValueColorMiddlePosition / CT_ValueColorPositions
# ==============================================================================

def parse_cx_CT_ValueColors(node: ET.Element | None) -> dict:
    """CT_ValueColors → dict  {minColor, midColor, maxColor}.
    렌더러 미사용: 색상 스케일 — 시각 전용."""
    if node is None: return {}
    # a:CT_SolidColorFillProperties 내용은 색상 전용이므로 파싱 생략
    return {'minColor': None, 'midColor': None, 'maxColor': None}

def parse_cx_CT_ExtremeValueColorPosition(node: ET.Element | None) -> dict:
    """CT_ExtremeValueColorPosition → dict  (empty complexType).
    렌더러 미사용: 색상 스케일 최솟값/최댓값 위치 — 시각 전용."""
    return {}

def parse_cx_CT_NumberColorPosition(node: ET.Element | None) -> dict:
    """CT_NumberColorPosition → dict  {val}.
    렌더러 미사용: 절대 수치 색상 위치 — 시각 전용."""
    if node is None: return {}
    return {'val': node.attrib.get('val')}

def parse_cx_CT_PercentageColorPosition(node: ET.Element | None) -> dict:
    """CT_PercentageColorPosition → dict  {val}.
    렌더러 미사용: 백분율 색상 위치 — 시각 전용."""
    if node is None: return {}
    return {'val': node.attrib.get('val')}

def parse_cx_CT_ValueColorEndPosition(node: ET.Element | None) -> dict:
    """CT_ValueColorEndPosition → dict  (extremeValue | number | percent).
    렌더러 미사용: 색상 스케일 끝 위치 — 시각 전용."""
    if node is None: return {}
    ev   = _get_child(node, 'extremeValue')
    num  = _get_child(node, 'number')
    pct  = _get_child(node, 'percent')
    if ev  is not None: return {'extremeValue': parse_cx_CT_ExtremeValueColorPosition(ev)}
    if num is not None: return {'number': parse_cx_CT_NumberColorPosition(num)}
    if pct is not None: return {'percent': parse_cx_CT_PercentageColorPosition(pct)}
    return {}

def parse_cx_CT_ValueColorMiddlePosition(node: ET.Element | None) -> dict:
    """CT_ValueColorMiddlePosition → dict  (number | percent).
    렌더러 미사용: 색상 스케일 중간 위치 — 시각 전용."""
    if node is None: return {}
    num = _get_child(node, 'number')
    pct = _get_child(node, 'percent')
    if num is not None: return {'number': parse_cx_CT_NumberColorPosition(num)}
    if pct is not None: return {'percent': parse_cx_CT_PercentageColorPosition(pct)}
    return {}

def parse_cx_CT_ValueColorPositions(node: ET.Element | None) -> dict:
    """CT_ValueColorPositions → dict  {count, min, mid, max}.
    렌더러 미사용: 색상 스케일 위치 정의 — 시각 전용."""
    if node is None: return {}
    return {
        'count': node.attrib.get('count', '2'),
        'min':   parse_cx_CT_ValueColorEndPosition(_get_child(node, 'min')),
        'mid':   parse_cx_CT_ValueColorMiddlePosition(_get_child(node, 'mid')),
        'max':   parse_cx_CT_ValueColorEndPosition(_get_child(node, 'max')),
    }

# ==============================================================================
# CT_Series
# ==============================================================================

def parse_cx_CT_Series(node: ET.Element | None) -> dict:
    """CT_Series → dict
    XSD: tx? + spPr? + valueColors? + valueColorPositions? + dataPt* +
         dataLabels? + dataId? + layoutPr? + axisId* + extLst?,
         attrs: layoutId(required), hidden, ownerIdx, uniqueId, formatIdx.

    렌더러 사용 필드:
      - layoutId: 차트 타입명 결정
      - dataId:   CT_ChartData에서 해당 데이터 블록 조회

    렌더러 미사용 필드 (파싱 완료):
      - tx: 시리즈 이름. Markdown 테이블 헤더는 자동 생성.
      - spPr: 시각 전용 (색상·형태).
      - valueColors / valueColorPositions: 색상 스케일 — 시각 전용.
      - dataPt: 포인트별 서식 재정의 — 시각 전용.
      - dataLabels: 레이블 설정 — Markdown 테이블은 레이블을 별도 출력 안 함.
      - layoutPr: 레이아웃별 추가 설정 — 시각/통계 전용.
      - axisId: 축 연결 ID — Markdown 테이블에서 축을 별도 출력 안 함.
      - hidden/ownerIdx/uniqueId/formatIdx: 표시 여부·서식 인덱스 — 시각 전용."""
    if node is None: return {'layoutId': '', 'dataId': ''}
    # dataId val 추출
    data_id_node = _get_child(node, 'dataId')
    data_id = data_id_node.attrib.get('val', '') if data_id_node is not None else ''
    return {
        # 렌더러 사용
        'layoutId': node.attrib.get('layoutId', ''),
        'dataId':   data_id,
        # 렌더러 미사용 이하
        # tx: 시리즈 이름 — Markdown 헤더 자동 생성이므로 미사용
        'tx':     parse_cx_CT_Text(_get_child(node, 'tx')),
        # spPr: 시각 전용
        'spPr':   None,
        # valueColors/valueColorPositions: 색상 스케일 — 시각 전용
        'valueColors':           parse_cx_CT_ValueColors(_get_child(node, 'valueColors')),
        'valueColorPositions':   parse_cx_CT_ValueColorPositions(
                                     _get_child(node, 'valueColorPositions')),
        # dataPt: 포인트별 서식 — 시각 전용
        'dataPt':    [parse_cx_CT_DataPoint(dp) for dp in _get_children(node, 'dataPt')],
        # dataLabels: 레이블 설정 — Markdown에서 미사용
        'dataLabels': parse_cx_CT_DataLabels(_get_child(node, 'dataLabels')),
        # layoutPr: 레이아웃 추가 설정 — 시각/통계 전용
        'layoutPr':  parse_cx_CT_SeriesLayoutProperties(_get_child(node, 'layoutPr')),
        # axisId: 축 연결 ID — Markdown에서 미사용
        'axisId':    [child.text for child in _get_children(node, 'axisId')
                      if child.text is not None],
        # 속성
        'hidden':   node.attrib.get('hidden', '0'),
        'ownerIdx': node.attrib.get('ownerIdx'),
        'uniqueId': node.attrib.get('uniqueId'),
        'formatIdx':node.attrib.get('formatIdx'),
    }

# ==============================================================================
# CT_ChartTitle / CT_Legend / CT_PlotSurface
# ==============================================================================

def parse_cx_CT_ChartTitle(node: ET.Element | None) -> dict:
    """CT_ChartTitle → dict  {text, pos, align, overlay, offset}.
    XSD: tx? + spPr? + txPr? + offset? + extLst?,  attrs: pos, align, overlay.
    spPr, txPr: 시각 전용.
    렌더러 미사용: Markdown 출력에서 차트 제목을 별도 렌더링하지 않는다."""
    if node is None: return {}
    return {
        # text: 차트 제목 — Markdown 출력에서 미사용
        'text':    parse_cx_CT_Text(_get_child(node, 'tx')),
        'pos':     node.attrib.get('pos', 't'),
        'align':   node.attrib.get('align', 'ctr'),
        'overlay': node.attrib.get('overlay', '0'),
        'offset':  parse_cx_CT_Offset(_get_child(node, 'offset')),
        # spPr, txPr: 시각 전용
        'spPr':    None,
        'txPr':    None,
    }

def parse_cx_CT_Legend(node: ET.Element | None) -> dict:
    """CT_Legend → dict  {pos, align, overlay, offset}.
    XSD: spPr? + txPr? + offset? + extLst?,  attrs: pos, align, overlay.
    spPr, txPr: 시각 전용.
    렌더러 미사용: Markdown 테이블은 범례를 별도 출력하지 않는다."""
    if node is None: return {}
    return {
        'pos':     node.attrib.get('pos', 'r'),
        'align':   node.attrib.get('align', 'ctr'),
        'overlay': node.attrib.get('overlay', '0'),
        'offset':  parse_cx_CT_Offset(_get_child(node, 'offset')),
        # spPr, txPr: 시각 전용
        'spPr':    None,
        'txPr':    None,
    }

def parse_cx_CT_PlotSurface(node: ET.Element | None) -> dict:
    """CT_PlotSurface → dict  {spPr}.
    렌더러 미사용: 플롯 영역 배경 서식 — 시각 전용."""
    if node is None: return {}
    return {'spPr': None}  # spPr: 시각 전용

# ==============================================================================
# CT_PlotAreaRegion / CT_PlotArea
# ==============================================================================

def parse_cx_CT_PlotAreaRegion(node: ET.Element | None) -> list[dict]:
    """CT_PlotAreaRegion → list[CT_Series dict].
    XSD: plotSurface? + series* + extLst?
    plotSurface: 렌더러 미사용 — 시각 전용."""
    if node is None: return []
    # plotSurface 파싱 (렌더러 미사용)
    _ps = parse_cx_CT_PlotSurface(_get_child(node, 'plotSurface'))  # noqa: F841
    return [parse_cx_CT_Series(s) for s in _get_children(node, 'series')]

def parse_cx_CT_PlotArea(node: ET.Element | None) -> dict:
    """CT_PlotArea → dict  {series: list[dict], axes: list[dict]}.
    XSD: plotAreaRegion(required) + axis* + spPr? + extLst?
    spPr: 렌더러 미사용 — 플롯 영역 배경 서식, 시각 전용."""
    if node is None: return {'series': [], 'axes': []}
    region = _get_child(node, 'plotAreaRegion')
    series = parse_cx_CT_PlotAreaRegion(region)
    # axes: 파싱 완료. Markdown 테이블은 축 정보를 별도 출력하지 않음
    axes = [parse_cx_CT_Axis(ax) for ax in _get_children(node, 'axis')]
    return {
        'series': series,
        # axes: 축 설정 — Markdown 테이블에서 미사용
        'axes':   axes,
        # spPr: 시각 전용
        'spPr':   None,
    }

# ==============================================================================
# CT_Chart
# ==============================================================================

def parse_cx_CT_Chart(node: ET.Element | None) -> dict:
    """CT_Chart → dict  {series, axes, title, legend}.
    XSD: title? + plotArea(required) + legend? + extLst?
    렌더러 사용: series (→ layoutId, dataId).
    렌더러 미사용: title, legend, axes."""
    if node is None: return {'series': [], 'axes': [], 'title': None, 'legend': None}
    plot = parse_cx_CT_PlotArea(_get_child(node, 'plotArea'))
    return {
        'series': plot['series'],
        'axes':   plot['axes'],
        # title: 차트 제목 — Markdown에서 미사용
        'title':  parse_cx_CT_ChartTitle(_get_child(node, 'title')),
        # legend: 범례 — Markdown에서 미사용
        'legend': parse_cx_CT_Legend(_get_child(node, 'legend')),
    }

# ==============================================================================
# Print settings
# CT_HeaderFooter / CT_PageMargins / CT_PageSetup / CT_PrintSettings
# ==============================================================================

def parse_cx_CT_HeaderFooter(node: ET.Element | None) -> dict:
    """CT_HeaderFooter → dict  인쇄 머리글/바닥글.
    렌더러 미사용: 인쇄 레이아웃 — Markdown 출력과 무관."""
    if node is None: return {}
    def txt(tag: str) -> str:
        el = _get_child(node, tag)
        return (el.text or "") if el is not None else ""
    return {
        'oddHeader':    txt('oddHeader'),  'oddFooter':    txt('oddFooter'),
        'evenHeader':   txt('evenHeader'), 'evenFooter':   txt('evenFooter'),
        'firstHeader':  txt('firstHeader'),'firstFooter':  txt('firstFooter'),
        'alignWithMargins':  node.attrib.get('alignWithMargins', 'true'),
        'differentOddEven':  node.attrib.get('differentOddEven', 'false'),
        'differentFirst':    node.attrib.get('differentFirst',   'false'),
    }

def parse_cx_CT_PageMargins(node: ET.Element | None) -> dict:
    """CT_PageMargins → dict  인쇄 여백 (l/r/t/b/header/footer).
    렌더러 미사용: 인쇄 레이아웃 — Markdown 출력과 무관."""
    if node is None: return {}
    return {k: node.attrib.get(k) for k in ('l', 'r', 't', 'b', 'header', 'footer')}

def parse_cx_CT_PageSetup(node: ET.Element | None) -> dict:
    """CT_PageSetup → dict  인쇄 설정 (용지 크기, 방향 등).
    렌더러 미사용: 인쇄 레이아웃 — Markdown 출력과 무관."""
    if node is None: return {}
    keys = ('paperSize', 'firstPageNumber', 'orientation', 'blackAndWhite',
            'draft', 'useFirstPageNumber', 'horizontalDpi', 'verticalDpi', 'copies')
    return {k: node.attrib.get(k) for k in keys}

def parse_cx_CT_PrintSettings(node: ET.Element | None) -> dict:
    """CT_PrintSettings → dict  인쇄 설정 묶음.
    렌더러 미사용: 인쇄 레이아웃 — Markdown 출력과 무관."""
    if node is None: return {}
    return {
        'headerFooter': parse_cx_CT_HeaderFooter(_get_child(node, 'headerFooter')),
        'pageMargins':  parse_cx_CT_PageMargins(_get_child(node, 'pageMargins')),
        'pageSetup':    parse_cx_CT_PageSetup(_get_child(node, 'pageSetup')),
    }

# ==============================================================================
# Format overrides
# CT_FormatOverride / CT_FormatOverrides
# ==============================================================================

def parse_cx_CT_FormatOverride(node: ET.Element | None) -> dict:
    """CT_FormatOverride → dict  {idx, spPr}.
    렌더러 미사용: 시각 서식 재정의 — 시각 전용."""
    if node is None: return {}
    return {
        'idx':  int(node.attrib.get('idx', '0')),
        'spPr': None,  # spPr: 시각 전용
    }

def parse_cx_CT_FormatOverrides(node: ET.Element | None) -> list[dict]:
    """CT_FormatOverrides → list[CT_FormatOverride].
    렌더러 미사용: 서식 재정의 목록 — 시각 전용."""
    if node is None: return []
    return [parse_cx_CT_FormatOverride(fo) for fo in _get_children(node, 'fmtOvr')]

# ==============================================================================
# Geo types  (지도 차트 전용 — 모두 렌더러 미사용)
# CT_GeoLocationQuery / CT_Address / CT_GeoLocation / CT_GeoLocations /
# CT_GeoLocationQueryResult / CT_GeoLocationQueryResults /
# CT_GeoPolygon / CT_GeoPolygons / CT_Copyrights / CT_GeoData /
# CT_GeoDataEntityQuery / CT_GeoDataEntityQueryResult / CT_GeoDataEntityQueryResults /
# CT_GeoDataPointQuery / CT_GeoDataPointToEntityQuery /
# CT_GeoDataPointToEntityQueryResult / CT_GeoDataPointToEntityQueryResults /
# CT_GeoChildTypes / CT_GeoChildEntitiesQuery / CT_GeoHierarchyEntity /
# CT_GeoChildEntities / CT_GeoChildEntitiesQueryResult / CT_GeoChildEntitiesQueryResults /
# CT_GeoParentEntitiesQuery / CT_GeoEntity / CT_GeoParentEntity /
# CT_GeoParentEntitiesQueryResult / CT_GeoParentEntitiesQueryResults /
# CT_Clear / CT_GeoCache / CT_Geography
# ==============================================================================
# 지도 차트(regionMap)의 지리 캐시 데이터.
# 모두 지도 렌더링·좌표 계산에 쓰이며 Markdown 테이블 출력에는 불필요.
# XSD 구조 완결성을 위해 파서 스텁만 정의하고 빈 dict를 반환한다.

def parse_cx_CT_Geography(node: ET.Element | None) -> dict:
    """CT_Geography → dict  지도 차트 지리 설정 (geoCache + 속성).
    렌더러 미사용: 지도 차트 시각 전용 — Markdown 테이블에 지리 정보 불필요."""
    return {}

# ==============================================================================
# CT_ChartSpace  (루트)
# ==============================================================================

def parse_cx_CT_ChartSpace(root: ET.Element | None) -> dict:
    """CT_ChartSpace → dict  루트 진입점.
    XSD: chartData(required) + chart(required) + spPr? + txPr? + clrMapOvr? +
         fmtOvrs? + printSettings? + extLst?,
         attrs: version, featureList, fallbackImg.

    렌더러 사용: chart.series (→ layoutId, dataId), chartData (→ CT_Data).
    렌더러 미사용 (파싱 완료):
      - spPr/txPr: 차트 공간 배경/텍스트 서식 — 시각 전용.
      - clrMapOvr: a:CT_ColorMapping — 색상 매핑, 시각 전용.
      - fmtOvrs: 서식 재정의 — 시각 전용.
      - printSettings: 인쇄 설정 — Markdown 출력과 무관.
      - version/featureList/fallbackImg: 메타 속성 — 렌더링과 무관."""
    if root is None: return {'series': [], 'chartData': {}}
    chart_data = parse_cx_CT_ChartData(_get_child(root, 'chartData'))
    chart      = parse_cx_CT_Chart(_get_child(root, 'chart'))
    return {
        # 렌더러 사용
        'chartData': chart_data,
        'series':    chart['series'],
        # 렌더러 미사용 이하
        # axes: 축 설정 — Markdown 테이블에서 미사용
        'axes':    chart['axes'],
        # title: 차트 제목 — Markdown 헤더에서 미사용
        'title':   chart['title'],
        # legend: 범례 — Markdown 테이블에서 미사용
        'legend':  chart['legend'],
        # spPr, txPr: 시각 전용
        'spPr':    None,
        'txPr':    None,
        # clrMapOvr: a:CT_ColorMapping — 시각 전용
        'clrMapOvr':      None,
        # fmtOvrs: 서식 재정의 — 시각 전용
        'fmtOvrs':        parse_cx_CT_FormatOverrides(_get_child(root, 'fmtOvrs')),
        # printSettings: 인쇄 설정 — Markdown 출력과 무관
        'printSettings':  parse_cx_CT_PrintSettings(_get_child(root, 'printSettings')),
        # 속성
        'version':     root.attrib.get('version', '0.0'),
        'featureList': root.attrib.get('featureList', ''),
        'fallbackImg': root.attrib.get('fallbackImg', ''),
    }

# ==============================================================================
# Public Converter
# ==============================================================================

def convert_chartex(
    root: ET.Element,
    ctx: Optional[ZipContext] = None,
    chart_source: str = "excel",
) -> str:
    """cx:chartSpace 루트 요소를 받아 Markdown 테이블 문자열로 변환한다.

    Args:
        root:         cx:chartSpace 루트 ET.Element.
        ctx:          ZipContext 인스턴스 (내장 Excel 데이터 추출 시 사용).
        chart_source: "excel" (기본값) 또는 "xml".
                      "excel"이면 ctx를 통해 .xlsx 워크북을 우선 조회하고,
                      실패하거나 "xml"이면 XML 캐시 데이터(lvl/pt)를 사용한다.

    Returns:
        Markdown 문자열. 첫 줄은 **Chart type:** 헤더, 이후 GFM 파이프 테이블.
    """
    # 1. 전체 파싱
    parsed = parse_cx_CT_ChartSpace(root)

    # 2. 차트 타입명 결정 (첫 번째 시리즈의 layoutId 기준)
    series_list = parsed['series']
    layout_id = series_list[0]['layoutId'] if series_list else ""
    chart_type_name = _CHARTEX_LAYOUT_NAMES.get(
        layout_id,
        f"Modern Chart ({layout_id})" if layout_id else "Modern Chart",
    )
    md_header = f"**Chart type:** {chart_type_name}\n"

    # 3. Excel 기반 추출 (chart_source == "excel" 이고 ctx 가 있을 때)
    if chart_source == "excel" and ctx is not None:
        for _rid, child_ctx in ctx.resolve_all().items():
            if child_ctx.path.lower().endswith((".xlsx", ".xlsm")):
                try:
                    excel_bytes = child_ctx.read_bytes()
                    md_table = convert_excel_to_table(excel_bytes)
                    if md_table:
                        return f"{md_header}\n{md_table}"
                except Exception:
                    pass  # 실패하면 아래 XML 기반 파싱(Fallback)으로 넘어감

    # 4. XML 캐시 기반 파싱 (Fallback)
    #    모든 시리즈의 dataId를 순서대로 순회해 데이터 블록을 수집한다.
    #    (Box & Whisker 등 시리즈마다 별도 dataId를 갖는 경우도 포함)
    if not series_list:
        return f"{md_header}\n> [ChartEx: No series found in chart]"

    chart_data = parsed['chartData']
    columns: list[list[str]] = []
    str_col_count = 0
    # (ser_name, num_dim_col_count) per unique data block
    num_blocks_info: list[tuple[str, int]] = []
    seen_data_ids: set[str] = set()
    categories_added = False

    for i, ser in enumerate(series_list):
        data_id = ser['dataId']
        if not data_id or data_id in seen_data_ids:
            continue
        block = chart_data.get(data_id)
        if not block:
            continue
        seen_data_ids.add(data_id)
        ser_name = ser['tx'] or f"Series {i + 1}"

        if not categories_added:
            # strDim 레벨을 역순으로 (XML에서 줄기→잎 순으로 저장되므로 잎→줄기 순으로 재배열)
            for s_dim in block.get('strDims', []):
                for level in reversed(s_dim['levels']):
                    if level:
                        columns.append(level)
                        str_col_count += 1
            categories_added = True

        n_cols = 0
        for n_dim in block.get('numDims', []):
            for level in reversed(n_dim['levels']):
                if level:
                    columns.append(level)
                    n_cols += 1
        if n_cols:
            num_blocks_info.append((ser_name, n_cols))

    if not columns:
        return f"{md_header}\n> [ChartEx: No dimension data extracted from XML]"

    max_len = max(len(col) for col in columns)

    # 헤더 이름 결정
    str_headers = [""] * str_col_count
    num_headers: list[str] = []
    if len(seen_data_ids) == 1 and len(num_blocks_info) == 1:
        # 단일 데이터 블록: numDim 열 수가 시리즈 수와 같으면 시리즈 이름으로 매핑
        _, total_num_cols = num_blocks_info[0]
        for j in range(total_num_cols):
            if j < len(series_list):
                num_headers.append(series_list[j]['tx'] or f"Series {j + 1}")
            else:
                num_headers.append(f"Col {str_col_count + j + 1}")
    else:
        # 데이터 블록이 여러 개 (시리즈마다 별도 dataId): 블록별 시리즈 이름 사용
        for ser_name, col_count in num_blocks_info:
            for _ in range(col_count):
                num_headers.append(ser_name)

    header_names = str_headers + num_headers

    # GFM 파이프 테이블 조립
    lines: list[str] = [
        "| " + " | ".join(header_names) + " |",
        "| " + " | ".join(["---"] * len(header_names)) + " |",
    ]
    for i in range(max_len):
        row = [col[i] if i < len(col) else "" for col in columns]
        lines.append("| " + " | ".join(row) + " |")

    return f"{md_header}\n" + "\n".join(lines)

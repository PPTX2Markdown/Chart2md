"""chart2md.py — ECMA-376 dml-chart.xsd XML → Markdown table converter.

Architecture mirrors omml2latex.py: recursive-descent parser with one
function per XSD complex/group type.  Functions are named
  parse_<ns>_<CT|EG|Group>_<TypeName>(node)
where ns is 'c' (drawingml/chart) or 'a' (drawingml/main).

Namespace handling: get_tag() strips the URI, so the same code works for
both strict  (http://purl.oclc.org/ooxml/drawingml/chart)
and transitional (http://schemas.openxmlformats.org/drawingml/2006/chart).

모든 XSD complexType / group 에 대응하는 parse 함수가 존재한다.
다만 Markdown 테이블 출력과 무관한 필드(시각 서식, 3D 뷰, 축 설정 등)는
파싱은 하되 렌더러에서 사용하지 않는다 — 각 렌더러 호출 지점에 그 이유를
한글 주석으로 명시하였다.

Public API
----------
convert_chart(root, ctx=None) -> str
    차트 XML 루트 Element를 Markdown 테이블 문자열로 변환한다.
load_chart_parts(path) -> list[tuple[ET.Element, ZipContext | None]]
    OOXML 파일에서 차트 파트를 로드한다.
"""

from __future__ import annotations

import argparse
import sys
import zipfile
import xml.etree.ElementTree as ET
from typing import Optional

from .ooxml_context import ZipContext, iter_parts_matching

__all__ = ["convert_chart", "load_chart_parts"]

# ==============================================================================
# 모듈 수준 날짜 에포크 플래그
# parse_c_CT_ChartSpace 에서 date1904 값을 읽어 설정하고,
# parse_c_CT_NumData 의 날짜 직렬 변환에서 참조한다.
# 단일 스레드 사용 전제 (Office 파일 변환은 단일 스레드).
# ==============================================================================
_date1904: bool = False

# ==============================================================================
# Helper / Utility  (verbatim from omml2latex.py)
# ==============================================================================

def _get_tag(node):
    return node.tag.split('}')[-1] if node is not None else ""

def _get_child(node, tag):
    if node is None: return None
    for child in node:
        if _get_tag(child) == tag: return child
    return None

def _get_children(node, tag):
    if node is None: return []
    return [child for child in node if _get_tag(child) == tag]

def _get_val(node, attr='val'):
    if node is None: return None
    for k, v in node.attrib.items():
        if k == attr or k.endswith(f"}}{attr}"):
            return v
    return None

# ==============================================================================
# DrawingML 텍스트 추출  (a: namespace)
# CT_RegularTextRun → CT_TextParagraph → CT_TextBody
# ==============================================================================

def parse_a_CT_RegularTextRun(node):
    """a:r (CT_RegularTextRun) — 텍스트 런에서 순수 문자열 추출."""
    if node is None: return ""
    t = _get_child(node, 't')
    return (t.text or "") if t is not None else ""

def parse_a_CT_TextParagraph(node):
    """a:p (CT_TextParagraph) — 단락 내 모든 런을 연결."""
    if node is None: return ""
    parts = []
    for child in node:
        tag = _get_tag(child)
        if tag == 'r':
            parts.append(parse_a_CT_RegularTextRun(child))
        elif tag == 'fld':
            t = _get_child(child, 't')
            if t is not None:
                parts.append(t.text or "")
    return "".join(parts)

def parse_a_CT_TextBody(node):
    """a:txBody (CT_TextBody) — 단락들을 줄바꿈으로 연결."""
    if node is None: return ""
    paras = _get_children(node, 'p')
    return "\n".join(parse_a_CT_TextParagraph(p) for p in paras).strip()

# ==============================================================================
# 공통 단순 타입 파서
# CT_Boolean, CT_Double, CT_UnsignedInt
# (속성만 있는 단순 wrapper — get_val로 직접 읽어도 되지만 명명 일관성을 위해 정의)
# ==============================================================================

def parse_c_CT_Boolean(node):
    """CT_Boolean → bool  (val 속성 없으면 XSD default=true)"""
    if node is None: return None
    v = _get_val(node)
    if v is None: return True
    return v.lower() in ('1', 'true')

def parse_c_CT_Double(node):
    """CT_Double → float"""
    if node is None: return None
    v = _get_val(node)
    return float(v) if v is not None else None

def parse_c_CT_UnsignedInt(node):
    """CT_UnsignedInt → int"""
    if node is None: return None
    v = _get_val(node)
    return int(v) if v is not None else None

# ==============================================================================
# ST_ 단순 타입 파서  (dml-chart.rnc 1:1 대응)
# 모두 렌더러에서 미사용: 열거형·범위 제약은 렌더링에 무관하므로 val 그대로 반환
# ==============================================================================

def parse_c_ST_AxisUnit(val):
    """ST_AxisUnit — xsd:double > 0. 축 단위 — Markdown 테이블 출력에 불필요."""
    return val

def parse_c_ST_AxPos(val):
    """ST_AxPos — 'b'|'l'|'r'|'t'. 축 위치 — 축 설정 — Markdown 테이블 출력에 불필요."""
    return val

def parse_c_ST_BarDir(val):
    """ST_BarDir — 'bar'|'col'. 막대 방향 — 막대 차트 시각 설정 — Markdown에서 미사용."""
    return val

def parse_c_ST_BarGrouping(val):
    """ST_BarGrouping — 'percentStacked'|'clustered'|'standard'|'stacked'. 막대 차트 시각 설정 — Markdown에서 미사용."""
    return val

def parse_c_ST_BubbleScale(val):
    """ST_BubbleScale — '0%'-'300%'. 버블 크기 배율 — 시각 전용 — Markdown 테이블에 표현 불가."""
    return val

def parse_c_ST_BubbleScalePercent(val):
    """ST_BubbleScalePercent — 내부 타입 (ST_BubbleScale 동의어). 시각 전용 — Markdown 테이블에 표현 불가."""
    return val

def parse_c_ST_BuiltInUnit(val):
    """ST_BuiltInUnit — 'hundreds'|'thousands'|...|'trillions'. 축 단위 표시 — 축 설정 — Markdown 테이블 출력에 불필요."""
    return val

def parse_c_ST_CrossBetween(val):
    """ST_CrossBetween — 'between'|'midCat'. 축 교차 위치 — 축 설정 — Markdown 테이블 출력에 불필요."""
    return val

def parse_c_ST_Crosses(val):
    """ST_Crosses — 'autoZero'|'max'|'min'. 축 교차 방식 — 축 설정 — Markdown 테이블 출력에 불필요."""
    return val

def parse_c_ST_DepthPercent(val):
    """ST_DepthPercent — '20%'-'2000%'. 3D 깊이 — 3D 시점 — Markdown 2D 테이블과 무관."""
    return val

def parse_c_ST_DepthPercentWithSymbol(val):
    """ST_DepthPercentWithSymbol — 내부 타입 (ST_DepthPercent 동의어). 3D 시점 — Markdown 2D 테이블과 무관."""
    return val

def parse_c_ST_DispBlanksAs(val):
    """ST_DispBlanksAs — 'span'|'gap'|'zero'. 빈 값 표시 방식 — 시각 전용 — Markdown 테이블에 표현 불가."""
    return val

def parse_c_ST_DLblPos(val):
    """ST_DLblPos — 'bestFit'|'b'|'ctr'|'inBase'|'inEnd'|'l'|'outEnd'|'r'|'t'. 데이터 레이블 위치 — 시각 전용 — Markdown 테이블에 표현 불가."""
    return val

def parse_c_ST_ErrBarType(val):
    """ST_ErrBarType — 'both'|'minus'|'plus'. 오차 막대 유형 — 추세선/오차 설정 — Markdown에서 미사용."""
    return val

def parse_c_ST_ErrDir(val):
    """ST_ErrDir — 'x'|'y'. 오차 막대 방향 — 추세선/오차 설정 — Markdown에서 미사용."""
    return val

def parse_c_ST_ErrValType(val):
    """ST_ErrValType — 'cust'|'fixedVal'|'percentage'|'stdDev'|'stdErr'. 오차 값 유형 — 추세선/오차 설정 — Markdown에서 미사용."""
    return val

def parse_c_ST_FirstSliceAng(val):
    """ST_FirstSliceAng — 0-360 unsigned short. 파이 첫 슬라이스 각도 — 파이 차트 시각 설정 — Markdown에서 미사용."""
    return val

def parse_c_ST_GapAmount(val):
    """ST_GapAmount — '0%'-'500%'. 막대 간격 — 막대 차트 시각 설정 — Markdown에서 미사용."""
    return val

def parse_c_ST_GapAmountPercent(val):
    """ST_GapAmountPercent — 내부 타입 (ST_GapAmount 동의어). 막대 차트 시각 설정 — Markdown에서 미사용."""
    return val

def parse_c_ST_Grouping(val):
    """ST_Grouping — 'percentStacked'|'standard'|'stacked'. 꺾은선/영역 차트 그룹화 — 시각 전용 — Markdown 테이블에 표현 불가."""
    return val

def parse_c_ST_HoleSize(val):
    """ST_HoleSize — '1%'-'90%'. 도넛 구멍 크기 — 파이 차트 시각 설정 — Markdown에서 미사용."""
    return val

def parse_c_ST_HoleSizePercent(val):
    """ST_HoleSizePercent — 내부 타입 (ST_HoleSize 동의어). 파이 차트 시각 설정 — Markdown에서 미사용."""
    return val

def parse_c_ST_HPercent(val):
    """ST_HPercent — 동의어 ST_HPercentWithSymbol. 차트 높이 비율 — 시각 전용 — Markdown 테이블에 표현 불가."""
    return val

def parse_c_ST_HPercentWithSymbol(val):
    """ST_HPercentWithSymbol — '5%'-'500%'. 차트 높이 비율 — 시각 전용 — Markdown 테이블에 표현 불가."""
    return val

def parse_c_ST_LayoutMode(val):
    """ST_LayoutMode — 'edge'|'factor'. 레이아웃 모드 — 시각 전용 — Markdown 테이블에 표현 불가."""
    return val

def parse_c_ST_LayoutTarget(val):
    """ST_LayoutTarget — 'inner'|'outer'. 레이아웃 대상 — 시각 전용 — Markdown 테이블에 표현 불가."""
    return val

def parse_c_ST_LblAlgn(val):
    """ST_LblAlgn — 'ctr'|'l'|'r'. 레이블 정렬 — 시각 전용 — Markdown 테이블에 표현 불가."""
    return val

def parse_c_ST_LblOffset(val):
    """ST_LblOffset — '0%'-'1000%'. 레이블 오프셋 — 축 설정 — Markdown 테이블 출력에 불필요."""
    return val

def parse_c_ST_LblOffsetPercent(val):
    """ST_LblOffsetPercent — 내부 타입 (ST_LblOffset 동의어). 축 설정 — Markdown 테이블 출력에 불필요."""
    return val

def parse_c_ST_LegendPos(val):
    """ST_LegendPos — 'b'|'tr'|'l'|'r'|'t'. 범례 위치 — 시각 전용 — Markdown 테이블에 표현 불가."""
    return val

def parse_c_ST_LogBase(val):
    """ST_LogBase — 2.0-1000.0 double. 로그 축 밑 — 축 설정 — Markdown 테이블 출력에 불필요."""
    return val

def parse_c_ST_MarkerSize(val):
    """ST_MarkerSize — 2-72 unsigned byte. 마커 크기 — 시각 전용 (색상·크기·모양) — Markdown 테이블에 표현 불가."""
    return val

def parse_c_ST_MarkerStyle(val):
    """ST_MarkerStyle — 'circle'|'dash'|...|'auto'. 마커 모양 — 시각 전용 (색상·크기·모양) — Markdown 테이블에 표현 불가."""
    return val

def parse_c_ST_OfPieType(val):
    """ST_OfPieType — 'pie'|'bar'. 보조 파이/막대 유형 — 파이 차트 시각 설정 — Markdown에서 미사용."""
    return val

def parse_c_ST_Order(val):
    """ST_Order — 2-6 unsigned byte. 추세선 차수 — 추세선/오차 설정 — Markdown에서 미사용."""
    return val

def parse_c_ST_Orientation(val):
    """ST_Orientation — 'maxMin'|'minMax'. 축 방향 — 축 설정 — Markdown 테이블 출력에 불필요."""
    return val

def parse_c_ST_Overlap(val):
    """ST_Overlap — '-100%'-'100%'. 막대 겹침 — 막대 차트 시각 설정 — Markdown에서 미사용."""
    return val

def parse_c_ST_OverlapPercent(val):
    """ST_OverlapPercent — 내부 타입 (ST_Overlap 동의어). 막대 차트 시각 설정 — Markdown에서 미사용."""
    return val

def parse_c_ST_PageSetupOrientation(val):
    """ST_PageSetupOrientation — 'default'|'portrait'|'landscape'. 인쇄 방향 — 인쇄 서식 — Markdown 출력과 무관."""
    return val

def parse_c_ST_Period(val):
    """ST_Period — unsigned int >= 2. 이동 평균 기간 — 추세선/오차 설정 — Markdown에서 미사용."""
    return val

def parse_c_ST_Perspective(val):
    """ST_Perspective — 0-240 unsigned byte. 3D 원근감 — 3D 시점 — Markdown 2D 테이블과 무관."""
    return val

def parse_c_ST_PictureFormat(val):
    """ST_PictureFormat — 'stretch'|'stack'|'stackScale'. 그림 채우기 방식 — 시각 전용 (색상·크기·모양) — Markdown 테이블에 표현 불가."""
    return val

def parse_c_ST_PictureStackUnit(val):
    """ST_PictureStackUnit — double > 0. 그림 스택 단위 — 시각 전용 (색상·크기·모양) — Markdown 테이블에 표현 불가."""
    return val

def parse_c_ST_RadarStyle(val):
    """ST_RadarStyle — 'standard'|'marker'|'filled'. 레이더 차트 스타일 — 시각 전용 — Markdown 테이블에 표현 불가."""
    return val

def parse_c_ST_RotX(val):
    """ST_RotX — -90 to 90 byte. X축 회전 — 3D 시점 — Markdown 2D 테이블과 무관."""
    return val

def parse_c_ST_RotY(val):
    """ST_RotY — 0-360 unsigned short. Y축 회전 — 3D 시점 — Markdown 2D 테이블과 무관."""
    return val

def parse_c_ST_ScatterStyle(val):
    """ST_ScatterStyle — 'none'|'line'|'lineMarker'|'marker'|'smooth'|'smoothMarker'. 분산형 스타일 — 시각 전용 — Markdown 테이블에 표현 불가."""
    return val

def parse_c_ST_SecondPieSize(val):
    """ST_SecondPieSize — '5%'-'200%'. 보조 파이 크기 — 파이 차트 시각 설정 — Markdown에서 미사용."""
    return val

def parse_c_ST_SecondPieSizePercent(val):
    """ST_SecondPieSizePercent — 내부 타입 (ST_SecondPieSize 동의어). 파이 차트 시각 설정 — Markdown에서 미사용."""
    return val

def parse_c_ST_Shape(val):
    """ST_Shape — 'cone'|'coneToMax'|'box'|'cylinder'|'pyramid'|'pyramidToMax'. 3D 막대 모양 — 시각 전용 (색상·크기·모양) — Markdown 테이블에 표현 불가."""
    return val

def parse_c_ST_SizeRepresents(val):
    """ST_SizeRepresents — 'area'|'w'. 버블 크기 표현 방식 — 시각 전용 — Markdown 테이블에 표현 불가."""
    return val

def parse_c_ST_Skip(val):
    """ST_Skip — unsigned int >= 1. 눈금 건너뜀 — 축 설정 — Markdown 테이블 출력에 불필요."""
    return val

def parse_c_ST_SplitType(val):
    """ST_SplitType — 'auto'|'cust'|'percent'|'pos'|'val'. 보조 파이 분할 방식 — 파이 차트 시각 설정 — Markdown에서 미사용."""
    return val

def parse_c_ST_Style(val):
    """ST_Style — 1-48 unsigned byte. 차트 내장 스타일 번호 — 시각 전용 (색상·크기·모양) — Markdown 테이블에 표현 불가."""
    return val

def parse_c_ST_Thickness(val):
    """ST_Thickness — 동의어 ST_ThicknessPercent. 표면 차트 두께 — 시각 전용 — Markdown 테이블에 표현 불가."""
    return val

def parse_c_ST_ThicknessPercent(val):
    """ST_ThicknessPercent — '[0-9]+%' 형식. 표면 차트 두께 백분율 — 시각 전용 — Markdown 테이블에 표현 불가."""
    return val

def parse_c_ST_TickLblPos(val):
    """ST_TickLblPos — 'high'|'low'|'nextTo'|'none'. 눈금 레이블 위치 — 축 설정 — Markdown 테이블 출력에 불필요."""
    return val

def parse_c_ST_TickMark(val):
    """ST_TickMark — 'cross'|'in'|'none'|'out'. 눈금 표시 방식 — 축 설정 — Markdown 테이블 출력에 불필요."""
    return val

def parse_c_ST_TimeUnit(val):
    """ST_TimeUnit — 'days'|'months'|'years'. 시간 축 단위 — 축 설정 — Markdown 테이블 출력에 불필요."""
    return val

def parse_c_ST_TrendlineType(val):
    """ST_TrendlineType — 'exp'|'linear'|'log'|'movingAvg'|'poly'|'power'. 추세선 유형 — 추세선/오차 설정 — Markdown에서 미사용."""
    return val

# ==============================================================================
# CT_ 단순 래퍼 타입 파서  (dml-chart.rnc 1:1 대응)
# 대부분 val 단일 속성 컨테이너 — Markdown 테이블 출력에 직접 사용하지 않음
# ==============================================================================

def parse_c_CT_AxisUnit(node):
    """CT_AxisUnit → val. 렌더러 미사용: 축 설정 — Markdown 테이블 출력에 불필요."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_AxPos(node):
    """CT_AxPos → val. 렌더러 미사용: 축 설정 — Markdown 테이블 출력에 불필요."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_BarDir(node):
    """CT_BarDir → val (default 'col'). 렌더러 미사용: 막대 차트 시각 설정 — Markdown에서 미사용."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_BarGrouping(node):
    """CT_BarGrouping → val (default 'clustered'). 렌더러 미사용: 막대 차트 시각 설정 — Markdown에서 미사용."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_BubbleScale(node):
    """CT_BubbleScale → val (default '100%'). 렌더러 미사용: 시각 전용 — Markdown 테이블에 표현 불가."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_BuiltInUnit(node):
    """CT_BuiltInUnit → val (default 'thousands'). 렌더러 미사용: 축 설정 — Markdown 테이블 출력에 불필요."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_CrossBetween(node):
    """CT_CrossBetween → val. 렌더러 미사용: 축 설정 — Markdown 테이블 출력에 불필요."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_Crosses(node):
    """CT_Crosses → val. 렌더러 미사용: 축 설정 — Markdown 테이블 출력에 불필요."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_CustSplit(node):
    """CT_CustSplit → list[int]. 렌더러 미사용: 파이 차트 시각 설정 — Markdown에서 미사용."""
    if node is None: return None
    return [parse_c_CT_UnsignedInt(pt) for pt in _get_children(node, 'secondPiePt')]

def parse_c_CT_DepthPercent(node):
    """CT_DepthPercent → val (default '100%'). 렌더러 미사용: 3D 시점 — Markdown 2D 테이블과 무관."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_DispBlanksAs(node):
    """CT_DispBlanksAs → val (default 'zero'). 렌더러 미사용: 시각 전용 — Markdown 테이블에 표현 불가."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_DLblPos(node):
    """CT_DLblPos → val. 렌더러 미사용: 시각 전용 — Markdown 테이블에 표현 불가."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_ErrBarType(node):
    """CT_ErrBarType → val (default 'both'). 렌더러 미사용: 추세선/오차 설정 — Markdown에서 미사용."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_ErrDir(node):
    """CT_ErrDir → val. 렌더러 미사용: 추세선/오차 설정 — Markdown에서 미사용."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_ErrValType(node):
    """CT_ErrValType → val (default 'fixedVal'). 렌더러 미사용: 추세선/오차 설정 — Markdown에서 미사용."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_Extension(node):
    """CT_Extension → {uri, content}. 렌더러 미사용: 확장 요소 — Markdown 출력과 무관."""
    if node is None: return None
    uri = _get_val(node, 'uri')
    return {'uri': uri}

def parse_c_CT_Extension_any(node):
    """CT_Extension_any → wildcard 자식. 렌더러 미사용: 확장 요소 와일드카드 — Markdown 출력과 무관."""
    if node is None: return None
    return node

def parse_c_CT_ExtensionList(node):
    """CT_ExtensionList → list[CT_Extension]. 렌더러 미사용: 확장 요소 목록 — Markdown 출력과 무관."""
    if node is None: return None
    return [parse_c_CT_Extension(ext) for ext in _get_children(node, 'ext')]

def parse_c_CT_FirstSliceAng(node):
    """CT_FirstSliceAng → val (default 0). 렌더러 미사용: 파이 차트 시각 설정 — Markdown에서 미사용."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_GapAmount(node):
    """CT_GapAmount → val (default '150%'). 렌더러 미사용: 막대 차트 시각 설정 — Markdown에서 미사용."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_Grouping(node):
    """CT_Grouping → val (default 'standard'). 렌더러 미사용: 시각 전용 — Markdown 테이블에 표현 불가."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_HoleSize(node):
    """CT_HoleSize → val (default '10%'). 렌더러 미사용: 파이 차트 시각 설정 — Markdown에서 미사용."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_HPercent(node):
    """CT_HPercent → val (default '100%'). 렌더러 미사용: 시각 전용 — Markdown 테이블에 표현 불가."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_LayoutMode(node):
    """CT_LayoutMode → val (default 'factor'). 렌더러 미사용: 시각 전용 — Markdown 테이블에 표현 불가."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_LayoutTarget(node):
    """CT_LayoutTarget → val (default 'outer'). 렌더러 미사용: 시각 전용 — Markdown 테이블에 표현 불가."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_LblAlgn(node):
    """CT_LblAlgn → val. 렌더러 미사용: 시각 전용 — Markdown 테이블에 표현 불가."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_LblOffset(node):
    """CT_LblOffset → val (default '100%'). 렌더러 미사용: 축 설정 — Markdown 테이블 출력에 불필요."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_LegendPos(node):
    """CT_LegendPos → val (default 'r'). 렌더러 미사용: 시각 전용 — Markdown 테이블에 표현 불가."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_LogBase(node):
    """CT_LogBase → val. 렌더러 미사용: 축 설정 — Markdown 테이블 출력에 불필요."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_MarkerSize(node):
    """CT_MarkerSize → val (default 5). 렌더러 미사용: 시각 전용 (색상·크기·모양) — Markdown 테이블에 표현 불가."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_MarkerStyle(node):
    """CT_MarkerStyle → val. 렌더러 미사용: 시각 전용 (색상·크기·모양) — Markdown 테이블에 표현 불가."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_OfPieType(node):
    """CT_OfPieType → val (default 'pie'). 렌더러 미사용: 파이 차트 시각 설정 — Markdown에서 미사용."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_Order(node):
    """CT_Order → val (default 2). 렌더러 미사용: 추세선/오차 설정 — Markdown에서 미사용."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_Orientation(node):
    """CT_Orientation → val (default 'minMax'). 렌더러 미사용: 축 설정 — Markdown 테이블 출력에 불필요."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_Overlap(node):
    """CT_Overlap → val (default '0%'). 렌더러 미사용: 막대 차트 시각 설정 — Markdown에서 미사용."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_Period(node):
    """CT_Period → val (default 2). 렌더러 미사용: 추세선/오차 설정 — Markdown에서 미사용."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_Perspective(node):
    """CT_Perspective → val (default 30). 렌더러 미사용: 3D 시점 — Markdown 2D 테이블과 무관."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_PictureFormat(node):
    """CT_PictureFormat → val. 렌더러 미사용: 시각 전용 (색상·크기·모양) — Markdown 테이블에 표현 불가."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_PictureStackUnit(node):
    """CT_PictureStackUnit → val. 렌더러 미사용: 시각 전용 (색상·크기·모양) — Markdown 테이블에 표현 불가."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_RadarStyle(node):
    """CT_RadarStyle → val (default 'standard'). 렌더러 미사용: 시각 전용 — Markdown 테이블에 표현 불가."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_RotX(node):
    """CT_RotX → val (default 0). 렌더러 미사용: 3D 시점 — Markdown 2D 테이블과 무관."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_RotY(node):
    """CT_RotY → val (default 0). 렌더러 미사용: 3D 시점 — Markdown 2D 테이블과 무관."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_ScatterStyle(node):
    """CT_ScatterStyle → val (default 'marker'). 렌더러 미사용: 시각 전용 — Markdown 테이블에 표현 불가."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_SecondPieSize(node):
    """CT_SecondPieSize → val (default '75%'). 렌더러 미사용: 파이 차트 시각 설정 — Markdown에서 미사용."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_Shape(node):
    """CT_Shape → val (default 'box'). 렌더러 미사용: 시각 전용 (색상·크기·모양) — Markdown 테이블에 표현 불가."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_SizeRepresents(node):
    """CT_SizeRepresents → val (default 'area'). 렌더러 미사용: 시각 전용 — Markdown 테이블에 표현 불가."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_Skip(node):
    """CT_Skip → val. 렌더러 미사용: 축 설정 — Markdown 테이블 출력에 불필요."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_SplitType(node):
    """CT_SplitType → val (default 'auto'). 렌더러 미사용: 파이 차트 시각 설정 — Markdown에서 미사용."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_Style(node):
    """CT_Style → val. 렌더러 미사용: 시각 전용 (색상·크기·모양) — Markdown 테이블에 표현 불가."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_TextLanguageID(node):
    """CT_TextLanguageID → val (s:ST_Lang). 렌더러 미사용: 언어 태그 — Markdown 출력과 무관."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_Thickness(node):
    """CT_Thickness → val. 렌더러 미사용: 시각 전용 — Markdown 테이블에 표현 불가."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_TickLblPos(node):
    """CT_TickLblPos → val (default 'nextTo'). 렌더러 미사용: 축 설정 — Markdown 테이블 출력에 불필요."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_TickMark(node):
    """CT_TickMark → val (default 'cross'). 렌더러 미사용: 축 설정 — Markdown 테이블 출력에 불필요."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_TimeUnit(node):
    """CT_TimeUnit → val (default 'days'). 렌더러 미사용: 축 설정 — Markdown 테이블 출력에 불필요."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_TrendlineType(node):
    """CT_TrendlineType → val (default 'linear'). 렌더러 미사용: 추세선/오차 설정 — Markdown에서 미사용."""
    if node is None: return None
    return _get_val(node)

def parse_c_CT_BandFmt(node):
    """CT_BandFmt → {idx, spPr}. 렌더러 미사용: 시각 전용 (색상·크기·모양) — Markdown 테이블에 표현 불가."""
    if node is None: return None
    idx_node = _get_child(node, 'idx')
    idx = parse_c_CT_UnsignedInt(idx_node)
    spPr = _get_child(node, 'spPr')
    return {'idx': idx, 'spPr': spPr}

def parse_c_CT_BandFmts(node):
    """CT_BandFmts → list[CT_BandFmt]. 렌더러 미사용: 시각 전용 (색상·크기·모양) — Markdown 테이블에 표현 불가."""
    if node is None: return None
    return [parse_c_CT_BandFmt(bf) for bf in _get_children(node, 'bandFmt')]

# ==============================================================================
# 리프 데이터 타입
# CT_NumVal, CT_StrVal, CT_NumData, CT_StrData
# ==============================================================================

def parse_c_CT_NumVal(node):
    """CT_NumVal → (idx:int, value:str)
    formatCode 속성도 XSD에 정의되어 있지만 Markdown 셀 값 출력에는
    숫자 포맷을 재현할 필요가 없으므로 무시한다."""
    if node is None: return (0, "")
    idx = int(_get_val(node, 'idx') or 0)
    v = _get_child(node, 'v')
    value = (v.text or "") if v is not None else ""
    return (idx, value)

def parse_c_CT_StrVal(node):
    """CT_StrVal → (idx:int, value:str)"""
    if node is None: return (0, "")
    idx = int(_get_val(node, 'idx') or 0)
    v = _get_child(node, 'v')
    value = (v.text or "") if v is not None else ""
    return (idx, value)

def _pts_to_list(pts_fn, pt_nodes, ptcount_node):
    """sparse pt 요소들을 dense list[str]로 변환하는 공유 헬퍼."""
    total = int(_get_val(ptcount_node) or 0) if ptcount_node is not None else None
    mapping = {}
    for pt in pt_nodes:
        idx, val = pts_fn(pt)
        mapping[idx] = val
    if total is None and mapping:
        total = max(mapping.keys()) + 1
    elif total is None:
        total = 0
    return [mapping.get(i, "") for i in range(total)]

def parse_c_CT_NumData(node):
    """CT_NumData → list[str]  (sparse → dense, 빈 슬롯은 "")
    formatCode 자식 요소를 검사하여 날짜 형식인 경우 Serial 반환값을 변환한다."""
    if node is None: return []
    
    fc_node = _get_child(node, 'formatCode')
    format_code = fc_node.text if fc_node is not None else ""
    
    is_date = False
    if format_code:
        fc = format_code.lower()
        # 날짜 패턴 감지 (ECMA-376 §22.2.2.40 formatCode + Excel 관행)
        # yy/yyyy: 연도, dd/ddd/dddd: 일/요일, mmm/mmmm: 월 이름,
        # m/d: 슬래시·하이픈 구분 날짜, [$-NNN]: 로케일 접두사
        if ('yy' in fc or 'mmm' in fc or 'm/d' in fc or 'm-d' in fc
                or fc.startswith('[$-')
                or ('d' in fc and any(p in fc for p in ('dd', 'ddd', 'dddd')))):
            is_date = True
        # 단독 'd' (단일 문자 day 포맷)
        elif fc.strip('"').strip() == 'd':
            is_date = True

    raw_list = _pts_to_list(parse_c_CT_NumVal,
                            _get_children(node, 'pt'),
                            _get_child(node, 'ptCount'))
                            
    if is_date:
        import datetime
        res = []
        for val in raw_list:
            if not val:
                res.append(val)
                continue
            try:
                serial = float(val)
                if _date1904:
                    # Mac 1904 날짜 체계: 1904-01-01 = 0
                    dt = datetime.datetime(1904, 1, 1) + datetime.timedelta(days=int(serial))
                else:
                    # Windows 1900 날짜 체계: 엑셀 날짜 버그(1900년 윤년 가정) 보정
                    delta = serial - 1 if serial < 60 else serial - 2
                    dt = datetime.datetime(1900, 1, 1) + datetime.timedelta(days=int(delta))
                res.append(dt.strftime("%Y-%m-%d"))
            except Exception:
                res.append(val)
        return res
    return raw_list

def parse_c_CT_StrData(node):
    """CT_StrData → list[str]"""
    if node is None: return []
    return _pts_to_list(parse_c_CT_StrVal,
                        _get_children(node, 'pt'),
                        _get_child(node, 'ptCount'))

# ==============================================================================
# CT_Lvl, CT_MultiLvlStrData
# ==============================================================================

def parse_c_CT_Lvl(node):
    """CT_Lvl → list[str]  (해당 레벨의 pt 목록)"""
    if node is None: return []
    pts = _get_children(node, 'pt')
    mapping = {}
    for pt in pts:
        idx, val = parse_c_CT_StrVal(pt)
        mapping[idx] = val
    if not mapping:
        return []
    total = max(mapping.keys()) + 1
    return [mapping.get(i, "") for i in range(total)]

def parse_c_CT_MultiLvlStrData(node):
    """CT_MultiLvlStrData → list[str]
    여러 lvl이 있으면 각 레벨을 ' > '로 조합해 계층 레이블을 구성한다.
    예) lvl[0]=['2024','2024'], lvl[1]=['Q1','Q2'] → ['2024 > Q1','2024 > Q2']
    ptCount는 카테고리 수를 알리며, 각 lvl의 pt 개수와 동일하다."""
    if node is None: return []
    ptcount = _get_child(node, 'ptCount')
    total = int(_get_val(ptcount) or 0) if ptcount is not None else None
    lvl_nodes = _get_children(node, 'lvl')
    if not lvl_nodes:
        return []
    # 각 레벨의 dense 리스트를 수집
    levels = [parse_c_CT_Lvl(lvl) for lvl in lvl_nodes]
    if total is None:
        total = max((len(lv) for lv in levels), default=0)
    result = []
    for i in range(total):
        parts = []
        for lv in levels:
            cell = lv[i] if i < len(lv) else ""
            if cell:
                parts.append(cell)
        result.append(" > ".join(parts))
    return result

# ==============================================================================
# 참조 / 리터럴 소스
# CT_NumRef, CT_StrRef, CT_MultiLvlStrRef
# CT_NumDataSource, CT_AxDataSource
# ==============================================================================

def parse_c_CT_NumRef(node):
    """CT_NumRef → list[str]  (캐시 우선, 없으면 수식 폴백)"""
    if node is None: return []
    cache = _get_child(node, 'numCache')
    if cache is not None:
        return parse_c_CT_NumData(cache)
    f = _get_child(node, 'f')
    formula = (f.text or "") if f is not None else ""
    return [f"[ref: {formula}]"] if formula else []

def parse_c_CT_StrRef(node):
    """CT_StrRef → list[str]  (캐시 우선, 없으면 수식 폴백)"""
    if node is None: return []
    cache = _get_child(node, 'strCache')
    if cache is not None:
        return parse_c_CT_StrData(cache)
    f = _get_child(node, 'f')
    formula = (f.text or "") if f is not None else ""
    return [f"[ref: {formula}]"] if formula else []

def parse_c_CT_MultiLvlStrRef(node):
    """CT_MultiLvlStrRef → list[str]  (캐시 우선, 없으면 수식 폴백)
    이전에는 첫 번째 lvl만 읽었으나 모든 lvl을 읽어 계층 레이블로 조합한다."""
    if node is None: return []
    cache = _get_child(node, 'multiLvlStrCache')
    if cache is not None:
        return parse_c_CT_MultiLvlStrData(cache)
    f = _get_child(node, 'f')
    formula = (f.text or "") if f is not None else ""
    return [f"[ref: {formula}]"] if formula else []

def parse_c_CT_NumDataSource(node):
    """CT_NumDataSource (numRef | numLit) → list[str]"""
    if node is None: return []
    numref = _get_child(node, 'numRef')
    if numref is not None:
        return parse_c_CT_NumRef(numref)
    numlit = _get_child(node, 'numLit')
    if numlit is not None:
        return parse_c_CT_NumData(numlit)
    return []

def parse_c_CT_AxDataSource(node):
    """CT_AxDataSource (multiLvlStrRef | numRef | numLit | strRef | strLit) → list[str]"""
    if node is None: return []
    multiref = _get_child(node, 'multiLvlStrRef')
    if multiref is not None:
        return parse_c_CT_MultiLvlStrRef(multiref)
    numref = _get_child(node, 'numRef')
    if numref is not None:
        return parse_c_CT_NumRef(numref)
    numlit = _get_child(node, 'numLit')
    if numlit is not None:
        return parse_c_CT_NumData(numlit)
    strref = _get_child(node, 'strRef')
    if strref is not None:
        return parse_c_CT_StrRef(strref)
    strlit = _get_child(node, 'strLit')
    if strlit is not None:
        return parse_c_CT_StrData(strlit)
    return []

# ==============================================================================
# 레이아웃
# CT_ManualLayout, CT_Layout
# ==============================================================================

def parse_c_CT_ManualLayout(node):
    """CT_ManualLayout → dict  (화면 배치 좌표/크기)
    layoutTarget, xMode, yMode, wMode, hMode, x, y, w, h 를 모두 읽는다."""
    if node is None: return {}
    return {
        'layoutTarget': _get_val(_get_child(node, 'layoutTarget')),
        'xMode':  _get_val(_get_child(node, 'xMode')),
        'yMode':  _get_val(_get_child(node, 'yMode')),
        'wMode':  _get_val(_get_child(node, 'wMode')),
        'hMode':  _get_val(_get_child(node, 'hMode')),
        'x': parse_c_CT_Double(_get_child(node, 'x')),
        'y': parse_c_CT_Double(_get_child(node, 'y')),
        'w': parse_c_CT_Double(_get_child(node, 'w')),
        'h': parse_c_CT_Double(_get_child(node, 'h')),
    }

def parse_c_CT_Layout(node):
    """CT_Layout → dict"""
    if node is None: return {}
    return {'manualLayout': parse_c_CT_ManualLayout(_get_child(node, 'manualLayout'))}

# ==============================================================================
# 차트 제목 관련
# CT_SerTx, CT_Tx, CT_Title
# ==============================================================================

def parse_c_CT_SerTx(node):
    """CT_SerTx (strRef | v) → str"""
    if node is None: return ""
    strref = _get_child(node, 'strRef')
    if strref is not None:
        vals = parse_c_CT_StrRef(strref)
        return vals[0] if vals else ""
    v = _get_child(node, 'v')
    if v is not None:
        return v.text or ""
    return ""

def parse_c_CT_Tx(node):
    """CT_Tx (strRef | rich) → str"""
    if node is None: return ""
    strref = _get_child(node, 'strRef')
    if strref is not None:
        vals = parse_c_CT_StrRef(strref)
        return vals[0] if vals else ""
    rich = _get_child(node, 'rich')
    if rich is not None:
        return parse_a_CT_TextBody(rich)
    return ""

def parse_c_CT_Title(node):
    """CT_Title → dict  {'text': str, 'layout': dict, 'overlay': bool|None}
    spPr(도형 서식)·txPr(텍스트 서식)는 a:CT_ShapeProperties / a:CT_TextBody 타입으로
    dml-main.xsd 영역이므로 이 파일 범위 밖이다. 파싱 없이 None으로 기록한다."""
    if node is None: return {'text': "", 'layout': {}, 'overlay': None, 'spPr': None, 'txPr': None}
    return {
        'text':    parse_c_CT_Tx(_get_child(node, 'tx')),
        'layout':  parse_c_CT_Layout(_get_child(node, 'layout')),
        'overlay': parse_c_CT_Boolean(_get_child(node, 'overlay')),
        # spPr: a:CT_ShapeProperties — dml-main.xsd 영역, 시각 서식 전용이므로 파싱 생략
        'spPr': None,
        # txPr: a:CT_TextBody — 동일 이유로 파싱 생략 (텍스트 자체는 tx에서 이미 추출)
        'txPr': None,
    }

# ==============================================================================
# 숫자 포맷
# CT_NumFmt
# ==============================================================================

def parse_c_CT_NumFmt(node):
    """CT_NumFmt → dict  {'formatCode': str, 'sourceLinked': bool|None}"""
    if node is None: return {}
    sc = node.attrib.get('sourceLinked')
    return {
        'formatCode':   node.attrib.get('formatCode', ''),
        'sourceLinked': (sc.lower() in ('1','true')) if sc is not None else None,
    }

# ==============================================================================
# 마커
# CT_MarkerStyle, CT_MarkerSize, CT_Marker
# ==============================================================================

def parse_c_CT_Marker(node):
    """CT_Marker → dict  {'symbol': str|None, 'size': int|None}
    spPr는 a:CT_ShapeProperties — dml-main.xsd 영역, 마커 시각 서식 전용이므로 생략."""
    if node is None: return {}
    return {
        'symbol': _get_val(_get_child(node, 'symbol')),
        'size':   parse_c_CT_UnsignedInt(_get_child(node, 'size')),
        # spPr: 마커 도형 서식 — Markdown 출력에 불필요
        'spPr': None,
    }

# ==============================================================================
# 그림 옵션
# CT_PictureOptions
# ==============================================================================

def parse_c_CT_PictureOptions(node):
    """CT_PictureOptions → dict
    막대/영역 차트에 이미지를 채울 때 사용하는 설정.
    Markdown 테이블은 이미지 렌더링을 하지 않으므로 파싱만 하고 렌더러에서 사용 안 함."""
    if node is None: return {}
    return {
        'applyToFront': parse_c_CT_Boolean(_get_child(node, 'applyToFront')),
        'applyToSides': parse_c_CT_Boolean(_get_child(node, 'applyToSides')),
        'applyToEnd':   parse_c_CT_Boolean(_get_child(node, 'applyToEnd')),
        'pictureFormat':    _get_val(_get_child(node, 'pictureFormat')),
        'pictureStackUnit': parse_c_CT_Double(_get_child(node, 'pictureStackUnit')),
    }

# ==============================================================================
# 개별 데이터 포인트
# CT_DPt
# ==============================================================================

def parse_c_CT_DPt(node):
    """CT_DPt → dict  특정 인덱스의 데이터 포인트 서식 재정의.
    idx(포인트 번호)는 기록하지만 나머지 서식 정보는 Markdown 셀 값에 영향 없으므로
    파서 구조 완결성을 위해 읽기만 하고 렌더러에서는 사용하지 않는다."""
    if node is None: return {}
    return {
        'idx':             parse_c_CT_UnsignedInt(_get_child(node, 'idx')),
        'invertIfNegative': parse_c_CT_Boolean(_get_child(node, 'invertIfNegative')),
        'marker':           parse_c_CT_Marker(_get_child(node, 'marker')),
        'bubble3D':         parse_c_CT_Boolean(_get_child(node, 'bubble3D')),
        'explosion':        parse_c_CT_UnsignedInt(_get_child(node, 'explosion')),
        # spPr: a:CT_ShapeProperties — dml-main.xsd 영역, 시각 서식 전용
        'spPr': None,
        'pictureOptions': parse_c_CT_PictureOptions(_get_child(node, 'pictureOptions')),
    }

# ==============================================================================
# 추세선
# CT_TrendlineLbl, CT_Trendline
# ==============================================================================

def parse_c_CT_TrendlineLbl(node):
    """CT_TrendlineLbl → dict  추세선 레이블 (위치, 텍스트, 숫자 포맷).
    layout·spPr·txPr는 시각 서식 전용이므로 생략."""
    if node is None: return {}
    return {
        'layout': parse_c_CT_Layout(_get_child(node, 'layout')),
        'tx':     parse_c_CT_Tx(_get_child(node, 'tx')),
        'numFmt': parse_c_CT_NumFmt(_get_child(node, 'numFmt')),
        # spPr, txPr: 시각 서식 — Markdown 출력에 불필요
        'spPr': None,
        'txPr': None,
    }

def parse_c_CT_Trendline(node):
    """CT_Trendline → dict  추세선 정의.
    파싱은 완전하게 하지만 Markdown 테이블은 추세선을 표현할 수 없으므로
    렌더러에서 이 dict를 사용하지 않는다."""
    if node is None: return {}
    name_el = _get_child(node, 'name')
    return {
        'name':          (name_el.text or "") if name_el is not None else "",
        # spPr: 추세선 선 서식 — 시각 전용
        'spPr':          None,
        'trendlineType': _get_val(_get_child(node, 'trendlineType')),
        'order':         _get_val(_get_child(node, 'order')),
        'period':        _get_val(_get_child(node, 'period')),
        'forward':       parse_c_CT_Double(_get_child(node, 'forward')),
        'backward':      parse_c_CT_Double(_get_child(node, 'backward')),
        'intercept':     parse_c_CT_Double(_get_child(node, 'intercept')),
        'dispRSqr':      parse_c_CT_Boolean(_get_child(node, 'dispRSqr')),
        'dispEq':        parse_c_CT_Boolean(_get_child(node, 'dispEq')),
        'trendlineLbl':  parse_c_CT_TrendlineLbl(_get_child(node, 'trendlineLbl')),
    }

# ==============================================================================
# 오류 막대
# CT_ErrBars
# ==============================================================================

def parse_c_CT_ErrBars(node):
    """CT_ErrBars → dict  오류 막대 정의.
    plus·minus는 CT_NumDataSource로 수치 데이터를 포함하지만,
    오류 범위는 Markdown 테이블 셀에 표현하지 않으므로 렌더러에서 사용 안 함."""
    if node is None: return {}
    return {
        'errDir':     _get_val(_get_child(node, 'errDir')),
        'errBarType': _get_val(_get_child(node, 'errBarType')),
        'errValType': _get_val(_get_child(node, 'errValType')),
        'noEndCap':   parse_c_CT_Boolean(_get_child(node, 'noEndCap')),
        'plus':       parse_c_CT_NumDataSource(_get_child(node, 'plus')),
        'minus':      parse_c_CT_NumDataSource(_get_child(node, 'minus')),
        'val':        parse_c_CT_Double(_get_child(node, 'val')),
        # spPr: 오류 막대 선 서식 — 시각 전용
        'spPr': None,
    }

# ==============================================================================
# 데이터 레이블
# EG_DLblShared, Group_DLbl, CT_DLbl, Group_DLbls, CT_DLbls
# ==============================================================================

def parse_c_EG_DLblShared(node):
    """EG_DLblShared → dict  데이터 레이블 공유 속성.
    showVal·showCatName 등 표시 여부 플래그를 포함하지만
    Markdown 테이블은 별도 레이블 열을 만들지 않으므로 렌더러에서 사용 안 함.
    spPr·txPr는 시각 서식 전용이므로 생략."""
    if node is None: return {}
    return {
        'numFmt':          parse_c_CT_NumFmt(_get_child(node, 'numFmt')),
        # spPr, txPr: 레이블 도형/텍스트 서식 — 시각 전용
        'spPr': None,
        'txPr': None,
        'dLblPos':         _get_val(_get_child(node, 'dLblPos')),
        'showLegendKey':   parse_c_CT_Boolean(_get_child(node, 'showLegendKey')),
        'showVal':         parse_c_CT_Boolean(_get_child(node, 'showVal')),
        'showCatName':     parse_c_CT_Boolean(_get_child(node, 'showCatName')),
        'showSerName':     parse_c_CT_Boolean(_get_child(node, 'showSerName')),
        'showPercent':     parse_c_CT_Boolean(_get_child(node, 'showPercent')),
        'showBubbleSize':  parse_c_CT_Boolean(_get_child(node, 'showBubbleSize')),
        'separator':       (_get_child(node, 'separator').text or "")
                           if _get_child(node, 'separator') is not None else None,
    }

def parse_c_Group_DLbl(node):
    """Group_DLbl → dict  개별 데이터 레이블 내용."""
    if node is None: return {}
    return {
        'layout': parse_c_CT_Layout(_get_child(node, 'layout')),
        'tx':     parse_c_CT_Tx(_get_child(node, 'tx')),
        **parse_c_EG_DLblShared(node),
    }

def parse_c_CT_DLbl(node):
    """CT_DLbl → dict  특정 인덱스의 데이터 레이블.
    delete=true면 해당 레이블 숨김. 파싱하되 렌더러에서 미사용."""
    if node is None: return {}
    idx = parse_c_CT_UnsignedInt(_get_child(node, 'idx'))
    delete_el = _get_child(node, 'delete')
    if delete_el is not None:
        return {'idx': idx, 'delete': parse_c_CT_Boolean(delete_el)}
    return {'idx': idx, 'delete': False, **parse_c_Group_DLbl(node)}

def parse_c_Group_DLbls(node):
    """Group_DLbls → dict  차트/시리즈 전체 레이블 설정."""
    if node is None: return {}
    return {
        **parse_c_EG_DLblShared(node),
        'showLeaderLines': parse_c_CT_Boolean(_get_child(node, 'showLeaderLines')),
        # leaderLines: CT_ChartLines — spPr만 있는 시각 서식 전용
        'leaderLines': None,
    }

def parse_c_CT_DLbls(node):
    """CT_DLbls → dict  시리즈/차트 레벨의 데이터 레이블 집합.
    개별 포인트 레이블 텍스트(dLbl[].tx)는 파싱하지만,
    Markdown 테이블은 레이블 텍스트를 별도 열로 출력하지 않으므로 렌더러 미사용."""
    if node is None: return {}
    dlbls = [parse_c_CT_DLbl(d) for d in _get_children(node, 'dLbl')]
    delete_el = _get_child(node, 'delete')
    if delete_el is not None:
        return {'dLbl': dlbls, 'delete': parse_c_CT_Boolean(delete_el)}
    return {'dLbl': dlbls, 'delete': False, **parse_c_Group_DLbls(node)}

# ==============================================================================
# 업/다운 바, 차트 선
# CT_UpDownBar, CT_UpDownBars, CT_ChartLines
# ==============================================================================

def parse_c_CT_UpDownBar(node):
    """CT_UpDownBar → dict  (spPr만 포함)
    spPr는 시각 서식 전용이므로 생략. 구조 완결성을 위해 함수만 존재."""
    if node is None: return {}
    return {'spPr': None}  # spPr: a:CT_ShapeProperties — 시각 전용

def parse_c_CT_UpDownBars(node):
    """CT_UpDownBars → dict  주식 차트의 상승/하락 바.
    Markdown 테이블은 바 색상/서식을 표현하지 않으므로 렌더러 미사용."""
    if node is None: return {}
    return {
        'gapWidth': _get_val(_get_child(node, 'gapWidth')),
        'upBars':   parse_c_CT_UpDownBar(_get_child(node, 'upBars')),
        'downBars': parse_c_CT_UpDownBar(_get_child(node, 'downBars')),
    }

def parse_c_CT_ChartLines(node):
    """CT_ChartLines → dict  (dropLines, hiLowLines, serLines 등에 사용)
    spPr만 포함하는 순수 시각 서식 컨테이너이므로 렌더러 미사용."""
    if node is None: return {}
    return {'spPr': None}  # spPr: a:CT_ShapeProperties — 시각 전용

# ==============================================================================
# EG_SerShared 그룹
# ==============================================================================

def parse_c_EG_SerShared(node):
    """EG_SerShared → dict  모든 시리즈 타입의 공통 필드.
    spPr(도형 서식)는 a:CT_ShapeProperties — dml-main.xsd 영역, 시각 서식 전용이므로
    파싱하지 않고 None으로 기록한다."""
    if node is None: return {'idx': 0, 'order': 0, 'name': '', 'spPr': None}
    return {
        'idx':   int(_get_val(_get_child(node, 'idx')) or 0),
        'order': int(_get_val(_get_child(node, 'order')) or 0),
        'name':  parse_c_CT_SerTx(_get_child(node, 'tx')),
        # spPr: 시리즈 막대/선 색상 등 — Markdown 출력에 불필요
        'spPr': None,
    }

# ==============================================================================
# 시리즈 파서 (8종)
# 각각 SeriesData dict 반환:
#   name, idx, order, categories, values, x_values, y_values, bubble_size,
#   marker, dpt, dlbls, trendlines, err_bars, smooth, invertIfNegative,
#   pictureOptions, shape, explosion, bubble3D, spPr
# ==============================================================================

def _make_series(shared, **kwargs):
    """SeriesData 기본 뼈대 생성."""
    return {
        'name':  shared['name'],
        'idx':   shared['idx'],
        'order': shared['order'],
        # spPr: 시리즈 도형 서식 — 시각 전용
        'spPr':  None,
        'categories':  None,
        'values':      [],
        'x_values':    None,
        'y_values':    None,
        'bubble_size': None,
        'marker':          None,
        'dpt':             [],
        'dlbls':           None,
        'trendlines':      [],
        'err_bars':        [],
        'smooth':          None,
        'invertIfNegative': None,
        'pictureOptions':  None,
        'shape':           None,
        'explosion':       None,
        'bubble3D':        None,
        **kwargs,
    }

def parse_c_CT_BarSer(node):
    """CT_BarSer → SeriesData
    XSD: EG_SerShared, invertIfNegative, pictureOptions, dPt[], dLbls,
         trendline[], errBars, cat, val, shape"""
    if node is None: return _make_series({'idx':0,'order':0,'name':'','spPr':None})
    shared = parse_c_EG_SerShared(node)
    cat = parse_c_CT_AxDataSource(_get_child(node, 'cat'))
    val = parse_c_CT_NumDataSource(_get_child(node, 'val'))
    return _make_series(shared,
        invertIfNegative = parse_c_CT_Boolean(_get_child(node, 'invertIfNegative')),
        pictureOptions   = parse_c_CT_PictureOptions(_get_child(node, 'pictureOptions')),
        dpt              = [parse_c_CT_DPt(d) for d in _get_children(node, 'dPt')],
        dlbls            = parse_c_CT_DLbls(_get_child(node, 'dLbls')),
        trendlines       = [parse_c_CT_Trendline(t) for t in _get_children(node, 'trendline')],
        err_bars         = [parse_c_CT_ErrBars(e) for e in _get_children(node, 'errBars')],
        categories       = cat or None,
        values           = val,
        shape            = _get_val(_get_child(node, 'shape')),
    )

def parse_c_CT_LineSer(node):
    """CT_LineSer → SeriesData
    XSD: EG_SerShared, marker, dPt[], dLbls, trendline[], errBars, cat, val, smooth"""
    if node is None: return _make_series({'idx':0,'order':0,'name':'','spPr':None})
    shared = parse_c_EG_SerShared(node)
    cat = parse_c_CT_AxDataSource(_get_child(node, 'cat'))
    val = parse_c_CT_NumDataSource(_get_child(node, 'val'))
    return _make_series(shared,
        marker     = parse_c_CT_Marker(_get_child(node, 'marker')),
        dpt        = [parse_c_CT_DPt(d) for d in _get_children(node, 'dPt')],
        dlbls      = parse_c_CT_DLbls(_get_child(node, 'dLbls')),
        trendlines = [parse_c_CT_Trendline(t) for t in _get_children(node, 'trendline')],
        err_bars   = [parse_c_CT_ErrBars(e) for e in _get_children(node, 'errBars')],
        categories = cat or None,
        values     = val,
        smooth     = parse_c_CT_Boolean(_get_child(node, 'smooth')),
    )

def parse_c_CT_AreaSer(node):
    """CT_AreaSer → SeriesData
    XSD: EG_SerShared, pictureOptions, dPt[], dLbls, trendline[], errBars(최대2), cat, val"""
    if node is None: return _make_series({'idx':0,'order':0,'name':'','spPr':None})
    shared = parse_c_EG_SerShared(node)
    cat = parse_c_CT_AxDataSource(_get_child(node, 'cat'))
    val = parse_c_CT_NumDataSource(_get_child(node, 'val'))
    return _make_series(shared,
        pictureOptions = parse_c_CT_PictureOptions(_get_child(node, 'pictureOptions')),
        dpt        = [parse_c_CT_DPt(d) for d in _get_children(node, 'dPt')],
        dlbls      = parse_c_CT_DLbls(_get_child(node, 'dLbls')),
        trendlines = [parse_c_CT_Trendline(t) for t in _get_children(node, 'trendline')],
        err_bars   = [parse_c_CT_ErrBars(e) for e in _get_children(node, 'errBars')],
        categories = cat or None,
        values     = val,
    )

def parse_c_CT_PieSer(node):
    """CT_PieSer → SeriesData
    XSD: EG_SerShared, explosion, dPt[], dLbls, cat, val"""
    if node is None: return _make_series({'idx':0,'order':0,'name':'','spPr':None})
    shared = parse_c_EG_SerShared(node)
    cat = parse_c_CT_AxDataSource(_get_child(node, 'cat'))
    val = parse_c_CT_NumDataSource(_get_child(node, 'val'))
    return _make_series(shared,
        explosion  = parse_c_CT_UnsignedInt(_get_child(node, 'explosion')),
        dpt        = [parse_c_CT_DPt(d) for d in _get_children(node, 'dPt')],
        dlbls      = parse_c_CT_DLbls(_get_child(node, 'dLbls')),
        categories = cat or None,
        values     = val,
    )

def parse_c_CT_RadarSer(node):
    """CT_RadarSer → SeriesData
    XSD: EG_SerShared, marker, dPt[], dLbls, cat, val"""
    if node is None: return _make_series({'idx':0,'order':0,'name':'','spPr':None})
    shared = parse_c_EG_SerShared(node)
    cat = parse_c_CT_AxDataSource(_get_child(node, 'cat'))
    val = parse_c_CT_NumDataSource(_get_child(node, 'val'))
    return _make_series(shared,
        marker     = parse_c_CT_Marker(_get_child(node, 'marker')),
        dpt        = [parse_c_CT_DPt(d) for d in _get_children(node, 'dPt')],
        dlbls      = parse_c_CT_DLbls(_get_child(node, 'dLbls')),
        categories = cat or None,
        values     = val,
    )

def parse_c_CT_SurfaceSer(node):
    """CT_SurfaceSer → SeriesData
    XSD: EG_SerShared, cat, val  (다른 시리즈 타입보다 단순)"""
    if node is None: return _make_series({'idx':0,'order':0,'name':'','spPr':None})
    shared = parse_c_EG_SerShared(node)
    cat = parse_c_CT_AxDataSource(_get_child(node, 'cat'))
    val = parse_c_CT_NumDataSource(_get_child(node, 'val'))
    return _make_series(shared,
        categories = cat or None,
        values     = val,
    )

def parse_c_CT_ScatterSer(node):
    """CT_ScatterSer → SeriesData  (cat/val 없이 xVal/yVal 사용)
    XSD: EG_SerShared, marker, dPt[], dLbls, trendline[], errBars(최대2), xVal, yVal, smooth"""
    if node is None: return _make_series({'idx':0,'order':0,'name':'','spPr':None})
    shared = parse_c_EG_SerShared(node)
    return _make_series(shared,
        marker     = parse_c_CT_Marker(_get_child(node, 'marker')),
        dpt        = [parse_c_CT_DPt(d) for d in _get_children(node, 'dPt')],
        dlbls      = parse_c_CT_DLbls(_get_child(node, 'dLbls')),
        trendlines = [parse_c_CT_Trendline(t) for t in _get_children(node, 'trendline')],
        err_bars   = [parse_c_CT_ErrBars(e) for e in _get_children(node, 'errBars')],
        x_values   = parse_c_CT_AxDataSource(_get_child(node, 'xVal')) or None,
        y_values   = parse_c_CT_NumDataSource(_get_child(node, 'yVal')),
        smooth     = parse_c_CT_Boolean(_get_child(node, 'smooth')),
    )

def parse_c_CT_BubbleSer(node):
    """CT_BubbleSer → SeriesData
    XSD: EG_SerShared, invertIfNegative, dPt[], dLbls, trendline[], errBars(최대2),
         xVal, yVal, bubbleSize, bubble3D"""
    if node is None: return _make_series({'idx':0,'order':0,'name':'','spPr':None})
    shared = parse_c_EG_SerShared(node)
    return _make_series(shared,
        invertIfNegative = parse_c_CT_Boolean(_get_child(node, 'invertIfNegative')),
        dpt        = [parse_c_CT_DPt(d) for d in _get_children(node, 'dPt')],
        dlbls      = parse_c_CT_DLbls(_get_child(node, 'dLbls')),
        trendlines = [parse_c_CT_Trendline(t) for t in _get_children(node, 'trendline')],
        err_bars   = [parse_c_CT_ErrBars(e) for e in _get_children(node, 'errBars')],
        x_values   = parse_c_CT_AxDataSource(_get_child(node, 'xVal')) or None,
        y_values   = parse_c_CT_NumDataSource(_get_child(node, 'yVal')),
        bubble_size= parse_c_CT_NumDataSource(_get_child(node, 'bubbleSize')) or None,
        bubble3D   = parse_c_CT_Boolean(_get_child(node, 'bubble3D')),
    )

# ==============================================================================
# 공유 그룹 헬퍼 — 시리즈 리스트 수집
# EG_LineChartShared, EG_BarChartShared, EG_AreaChartShared,
# EG_PieChartShared, EG_SurfaceChartShared
# ==============================================================================

def parse_c_EG_LineChartShared(node):
    """EG_LineChartShared → dict
    XSD: grouping(required), varyColors, ser[], dLbls, dropLines
    grouping·varyColors·dropLines는 시각 렌더링 속성이므로 Markdown 테이블에서 미사용."""
    if node is None: return {'series': [], 'grouping': None, 'varyColors': None, 'dlbls': None}
    return {
        'grouping':   _get_val(_get_child(node, 'grouping')),
        'varyColors': parse_c_CT_Boolean(_get_child(node, 'varyColors')),
        'series':     [parse_c_CT_LineSer(s) for s in _get_children(node, 'ser')],
        'dlbls':      parse_c_CT_DLbls(_get_child(node, 'dLbls')),
        # dropLines: CT_ChartLines — 시각 서식 전용
        'dropLines': None,
    }

def parse_c_EG_BarChartShared(node):
    """EG_BarChartShared → dict
    XSD: barDir(required), grouping, varyColors, ser[], dLbls
    barDir는 차트 타입 분류에 사용되며 각 CT_BarChart 파서에서 별도 처리."""
    if node is None: return {'series': [], 'grouping': None, 'varyColors': None, 'dlbls': None}
    return {
        'barDir':     _get_val(_get_child(node, 'barDir')),
        'grouping':   _get_val(_get_child(node, 'grouping')),
        'varyColors': parse_c_CT_Boolean(_get_child(node, 'varyColors')),
        'series':     [parse_c_CT_BarSer(s) for s in _get_children(node, 'ser')],
        'dlbls':      parse_c_CT_DLbls(_get_child(node, 'dLbls')),
    }

def parse_c_EG_AreaChartShared(node):
    """EG_AreaChartShared → dict
    XSD: grouping, varyColors, ser[], dLbls, dropLines"""
    if node is None: return {'series': [], 'grouping': None, 'varyColors': None, 'dlbls': None}
    return {
        'grouping':   _get_val(_get_child(node, 'grouping')),
        'varyColors': parse_c_CT_Boolean(_get_child(node, 'varyColors')),
        'series':     [parse_c_CT_AreaSer(s) for s in _get_children(node, 'ser')],
        'dlbls':      parse_c_CT_DLbls(_get_child(node, 'dLbls')),
        # dropLines: CT_ChartLines — 시각 서식 전용
        'dropLines': None,
    }

def parse_c_EG_PieChartShared(node):
    """EG_PieChartShared → dict
    XSD: varyColors, ser[], dLbls"""
    if node is None: return {'series': [], 'varyColors': None, 'dlbls': None}
    return {
        'varyColors': parse_c_CT_Boolean(_get_child(node, 'varyColors')),
        'series':     [parse_c_CT_PieSer(s) for s in _get_children(node, 'ser')],
        'dlbls':      parse_c_CT_DLbls(_get_child(node, 'dLbls')),
    }

def parse_c_EG_SurfaceChartShared(node):
    """EG_SurfaceChartShared → dict
    XSD: wireframe, ser[], bandFmts
    wireframe·bandFmts는 3D 시각 서식 전용이므로 Markdown 테이블에서 미사용."""
    if node is None: return {'series': []}
    return {
        'wireframe': parse_c_CT_Boolean(_get_child(node, 'wireframe')),
        'series':    [parse_c_CT_SurfaceSer(s) for s in _get_children(node, 'ser')],
        # bandFmts: CT_BandFmts — 색상 밴드 서식, 시각 전용
        'bandFmts': None,
    }

# ==============================================================================
# 차트 타입 파서 (16종)
# 각각 {'chart_type': str, 'series': list, ...chart-level props...} 반환
# ==============================================================================

def _chart(chart_type, shared_data, **extra):
    return {'chart_type': chart_type,
            'series': shared_data.get('series', []),
            **extra}

def parse_c_CT_LineChart(node):
    """CT_LineChart
    XSD: EG_LineChartShared, hiLowLines, upDownBars, marker(bool), smooth(bool), axId[2]
    hiLowLines·upDownBars·marker·smooth·axId는 시각/축 연결 정보이므로 렌더러 미사용."""
    if node is None: return _chart("Line Chart", {})
    shared = parse_c_EG_LineChartShared(node)
    return _chart("Line Chart", shared,
        hiLowLines = parse_c_CT_ChartLines(_get_child(node, 'hiLowLines')),
        upDownBars = parse_c_CT_UpDownBars(_get_child(node, 'upDownBars')),
        marker     = parse_c_CT_Boolean(_get_child(node, 'marker')),
        smooth     = parse_c_CT_Boolean(_get_child(node, 'smooth')),
        axId       = [parse_c_CT_UnsignedInt(a) for a in _get_children(node, 'axId')],
    )

def parse_c_CT_Line3DChart(node):
    """CT_Line3DChart
    XSD: EG_LineChartShared, gapDepth, axId[3]"""
    if node is None: return _chart("3-D Line Chart", {})
    shared = parse_c_EG_LineChartShared(node)
    return _chart("3-D Line Chart", shared,
        gapDepth = _get_val(_get_child(node, 'gapDepth')),
        axId     = [parse_c_CT_UnsignedInt(a) for a in _get_children(node, 'axId')],
    )

def parse_c_CT_StockChart(node):
    """CT_StockChart
    XSD: ser[3..4], dLbls, dropLines, hiLowLines, upDownBars, axId[2]"""
    if node is None: return _chart("Stock Chart", {})
    series = [parse_c_CT_LineSer(s) for s in _get_children(node, 'ser')]
    return {
        'chart_type': "Stock Chart",
        'series':     series,
        'dlbls':      parse_c_CT_DLbls(_get_child(node, 'dLbls')),
        'dropLines':  None,  # CT_ChartLines — 시각 전용
        'hiLowLines': parse_c_CT_ChartLines(_get_child(node, 'hiLowLines')),
        'upDownBars': parse_c_CT_UpDownBars(_get_child(node, 'upDownBars')),
        'axId':       [parse_c_CT_UnsignedInt(a) for a in _get_children(node, 'axId')],
    }

def parse_c_CT_BarChart(node):
    """CT_BarChart
    XSD: EG_BarChartShared, gapWidth, overlap, serLines[], axId[2]
    barDir(bar→가로막대, col→세로막대)로 차트 타입명 결정."""
    if node is None: return _chart("Bar Chart", {})
    shared = parse_c_EG_BarChartShared(node)
    direction = shared.get('barDir') or "col"
    chart_type = "Bar Chart" if direction == "bar" else "Column Chart"
    return _chart(chart_type, shared,
        gapWidth  = _get_val(_get_child(node, 'gapWidth')),
        overlap   = _get_val(_get_child(node, 'overlap')),
        # serLines: CT_ChartLines[] — 시각 전용
        serLines  = None,
        axId      = [parse_c_CT_UnsignedInt(a) for a in _get_children(node, 'axId')],
    )

def parse_c_CT_Bar3DChart(node):
    """CT_Bar3DChart
    XSD: EG_BarChartShared, gapWidth, gapDepth, shape, axId[2..3]"""
    if node is None: return _chart("3-D Bar Chart", {})
    shared = parse_c_EG_BarChartShared(node)
    direction = shared.get('barDir') or "col"
    chart_type = "3-D Bar Chart" if direction == "bar" else "3-D Column Chart"
    return _chart(chart_type, shared,
        gapWidth  = _get_val(_get_child(node, 'gapWidth')),
        gapDepth  = _get_val(_get_child(node, 'gapDepth')),
        shape     = _get_val(_get_child(node, 'shape')),
        axId      = [parse_c_CT_UnsignedInt(a) for a in _get_children(node, 'axId')],
    )

def parse_c_CT_AreaChart(node):
    """CT_AreaChart
    XSD: EG_AreaChartShared, axId[2]"""
    if node is None: return _chart("Area Chart", {})
    shared = parse_c_EG_AreaChartShared(node)
    return _chart("Area Chart", shared,
        axId = [parse_c_CT_UnsignedInt(a) for a in _get_children(node, 'axId')],
    )

def parse_c_CT_Area3DChart(node):
    """CT_Area3DChart
    XSD: EG_AreaChartShared, gapDepth, axId[2..3]"""
    if node is None: return _chart("3-D Area Chart", {})
    shared = parse_c_EG_AreaChartShared(node)
    return _chart("3-D Area Chart", shared,
        gapDepth = _get_val(_get_child(node, 'gapDepth')),
        axId     = [parse_c_CT_UnsignedInt(a) for a in _get_children(node, 'axId')],
    )

def parse_c_CT_ScatterChart(node):
    """CT_ScatterChart
    XSD: scatterStyle(required), varyColors, ser[], dLbls, axId[2]
    scatterStyle은 마커/선 표시 방식으로 Markdown 테이블에서 미사용."""
    if node is None: return _chart("Scatter Chart", {})
    series = [parse_c_CT_ScatterSer(s) for s in _get_children(node, 'ser')]
    return {
        'chart_type':   "Scatter Chart",
        'series':       series,
        'scatterStyle': _get_val(_get_child(node, 'scatterStyle')),
        'varyColors':   parse_c_CT_Boolean(_get_child(node, 'varyColors')),
        'dlbls':        parse_c_CT_DLbls(_get_child(node, 'dLbls')),
        'axId':         [parse_c_CT_UnsignedInt(a) for a in _get_children(node, 'axId')],
    }

def parse_c_CT_RadarChart(node):
    """CT_RadarChart
    XSD: radarStyle(required), varyColors, ser[], dLbls, axId[2]"""
    if node is None: return _chart("Radar Chart", {})
    series = [parse_c_CT_RadarSer(s) for s in _get_children(node, 'ser')]
    return {
        'chart_type':  "Radar Chart",
        'series':      series,
        'radarStyle':  _get_val(_get_child(node, 'radarStyle')),
        'varyColors':  parse_c_CT_Boolean(_get_child(node, 'varyColors')),
        'dlbls':       parse_c_CT_DLbls(_get_child(node, 'dLbls')),
        'axId':        [parse_c_CT_UnsignedInt(a) for a in _get_children(node, 'axId')],
    }

def parse_c_CT_PieChart(node):
    """CT_PieChart
    XSD: EG_PieChartShared, firstSliceAng
    firstSliceAng는 파이 첫 조각 시작 각도 — 시각 전용이므로 렌더러 미사용."""
    if node is None: return _chart("Pie Chart", {})
    shared = parse_c_EG_PieChartShared(node)
    return _chart("Pie Chart", shared,
        firstSliceAng = parse_c_CT_UnsignedInt(_get_child(node, 'firstSliceAng')),
    )

def parse_c_CT_Pie3DChart(node):
    """CT_Pie3DChart
    XSD: EG_PieChartShared"""
    if node is None: return _chart("3-D Pie Chart", {})
    return _chart("3-D Pie Chart", parse_c_EG_PieChartShared(node))

def parse_c_CT_DoughnutChart(node):
    """CT_DoughnutChart
    XSD: EG_PieChartShared, firstSliceAng, holeSize
    holeSize는 도넛 구멍 크기 — 시각 전용."""
    if node is None: return _chart("Doughnut Chart", {})
    shared = parse_c_EG_PieChartShared(node)
    return _chart("Doughnut Chart", shared,
        firstSliceAng = parse_c_CT_UnsignedInt(_get_child(node, 'firstSliceAng')),
        holeSize      = _get_val(_get_child(node, 'holeSize')),
    )

def parse_c_CT_OfPieChart(node):
    """CT_OfPieChart
    XSD: ofPieType(required), EG_PieChartShared, gapWidth, splitType, splitPos,
         custSplit, secondPieSize, serLines[]
    ofPieType(pie→원형분리, bar→막대분리)로 차트 타입명 결정."""
    if node is None: return _chart("Pie-of-Pie Chart", {})
    ptype = _get_val(_get_child(node, 'ofPieType')) or "pie"
    chart_type = "Bar-of-Pie Chart" if ptype == "bar" else "Pie-of-Pie Chart"
    shared = parse_c_EG_PieChartShared(node)
    return _chart(chart_type, shared,
        ofPieType     = ptype,
        gapWidth      = _get_val(_get_child(node, 'gapWidth')),
        splitType     = _get_val(_get_child(node, 'splitType')),
        splitPos      = parse_c_CT_Double(_get_child(node, 'splitPos')),
        # custSplit: CT_CustSplit — 분리 포인트 인덱스 목록, 데이터 레이아웃 전용
        custSplit     = None,
        secondPieSize = _get_val(_get_child(node, 'secondPieSize')),
        # serLines: CT_ChartLines[] — 시각 전용
        serLines      = None,
    )

def parse_c_CT_BubbleChart(node):
    """CT_BubbleChart
    XSD: varyColors, ser[], dLbls, bubble3D, bubbleScale, showNegBubbles,
         sizeRepresents, axId[2]
    bubble3D·bubbleScale·showNegBubbles·sizeRepresents는 시각/크기 표현 설정
    — Markdown 테이블에서 미사용."""
    if node is None: return _chart("Bubble Chart", {})
    series = [parse_c_CT_BubbleSer(s) for s in _get_children(node, 'ser')]
    return {
        'chart_type':       "Bubble Chart",
        'series':           series,
        'varyColors':       parse_c_CT_Boolean(_get_child(node, 'varyColors')),
        'dlbls':            parse_c_CT_DLbls(_get_child(node, 'dLbls')),
        'bubble3D':         parse_c_CT_Boolean(_get_child(node, 'bubble3D')),
        'bubbleScale':      _get_val(_get_child(node, 'bubbleScale')),
        'showNegBubbles':   parse_c_CT_Boolean(_get_child(node, 'showNegBubbles')),
        'sizeRepresents':   _get_val(_get_child(node, 'sizeRepresents')),
        'axId':             [parse_c_CT_UnsignedInt(a) for a in _get_children(node, 'axId')],
    }

def parse_c_CT_SurfaceChart(node):
    """CT_SurfaceChart
    XSD: EG_SurfaceChartShared, axId[2..3]"""
    if node is None: return _chart("Surface Chart", {})
    shared = parse_c_EG_SurfaceChartShared(node)
    return _chart("Surface Chart", shared,
        axId = [parse_c_CT_UnsignedInt(a) for a in _get_children(node, 'axId')],
    )

def parse_c_CT_Surface3DChart(node):
    """CT_Surface3DChart
    XSD: EG_SurfaceChartShared, axId[3]"""
    if node is None: return _chart("3-D Surface Chart", {})
    shared = parse_c_EG_SurfaceChartShared(node)
    return _chart("3-D Surface Chart", shared,
        axId = [parse_c_CT_UnsignedInt(a) for a in _get_children(node, 'axId')],
    )

# ==============================================================================
# 축 파서
# CT_Scaling, EG_AxShared, CT_CatAx, CT_DateAx, CT_SerAx, CT_ValAx
# CT_DispUnitsLbl, CT_DispUnits
# ==============================================================================

def parse_c_CT_Scaling(node):
    """CT_Scaling → dict  (logBase, orientation, max, min)
    축 눈금 범위/방향 — Markdown 테이블에서 미사용이지만 완전 파싱."""
    if node is None: return {}
    return {
        'logBase':     parse_c_CT_Double(_get_child(node, 'logBase')),
        'orientation': _get_val(_get_child(node, 'orientation')),
        'max':         parse_c_CT_Double(_get_child(node, 'max')),
        'min':         parse_c_CT_Double(_get_child(node, 'min')),
    }

def parse_c_EG_AxShared(node):
    """EG_AxShared → dict  모든 축 타입의 공통 필드.
    axId·scaling·axPos·crossAx 등을 파싱한다.
    title(CT_Title)은 축 레이블 텍스트를 포함하지만
    Markdown 테이블은 축 이름을 별도 출력하지 않으므로 렌더러 미사용.
    spPr·txPr는 시각 서식 — 생략."""
    if node is None: return {}
    return {
        'axId':            parse_c_CT_UnsignedInt(_get_child(node, 'axId')),
        'scaling':         parse_c_CT_Scaling(_get_child(node, 'scaling')),
        'delete':          parse_c_CT_Boolean(_get_child(node, 'delete')),
        'axPos':           _get_val(_get_child(node, 'axPos')),
        # majorGridlines, minorGridlines: CT_ChartLines — 시각 전용
        'majorGridlines':  None,
        'minorGridlines':  None,
        # title: CT_Title — 축 이름 텍스트 포함이지만 Markdown 테이블에서 미사용
        'title':           parse_c_CT_Title(_get_child(node, 'title')),
        'numFmt':          parse_c_CT_NumFmt(_get_child(node, 'numFmt')),
        'majorTickMark':   _get_val(_get_child(node, 'majorTickMark')),
        'minorTickMark':   _get_val(_get_child(node, 'minorTickMark')),
        'tickLblPos':      _get_val(_get_child(node, 'tickLblPos')),
        # spPr, txPr: 시각 서식 — 생략
        'spPr': None,
        'txPr': None,
        'crossAx': parse_c_CT_UnsignedInt(_get_child(node, 'crossAx')),
        'crosses':   _get_val(_get_child(node, 'crosses')),
        'crossesAt': parse_c_CT_Double(_get_child(node, 'crossesAt')),
    }

def parse_c_CT_DispUnitsLbl(node):
    """CT_DispUnitsLbl → dict  축 단위 레이블 (위치, 텍스트)."""
    if node is None: return {}
    return {
        'layout': parse_c_CT_Layout(_get_child(node, 'layout')),
        'tx':     parse_c_CT_Tx(_get_child(node, 'tx')),
        # spPr, txPr: 시각 서식 — 생략
        'spPr': None,
        'txPr': None,
    }

def parse_c_CT_DispUnits(node):
    """CT_DispUnits → dict  축 값 단위 표시 설정 (custUnit | builtInUnit).
    Markdown 테이블은 축 단위 표시를 별도 출력하지 않으므로 렌더러 미사용."""
    if node is None: return {}
    cust = parse_c_CT_Double(_get_child(node, 'custUnit'))
    builtin = _get_val(_get_child(node, 'builtInUnit'))
    return {
        'custUnit':    cust,
        'builtInUnit': builtin,
        'dispUnitsLbl': parse_c_CT_DispUnitsLbl(_get_child(node, 'dispUnitsLbl')),
    }

def parse_c_CT_CatAx(node):
    """CT_CatAx → dict  카테고리 축.
    XSD: EG_AxShared, auto, lblAlgn, lblOffset, tickLblSkip, tickMarkSkip, noMultiLvlLbl
    축 설정은 Markdown 테이블 데이터와 무관하므로 렌더러 미사용."""
    if node is None: return {}
    return {
        **parse_c_EG_AxShared(node),
        'auto':          parse_c_CT_Boolean(_get_child(node, 'auto')),
        'lblAlgn':       _get_val(_get_child(node, 'lblAlgn')),
        'lblOffset':     _get_val(_get_child(node, 'lblOffset')),
        'tickLblSkip':   parse_c_CT_UnsignedInt(_get_child(node, 'tickLblSkip')),
        'tickMarkSkip':  parse_c_CT_UnsignedInt(_get_child(node, 'tickMarkSkip')),
        'noMultiLvlLbl': parse_c_CT_Boolean(_get_child(node, 'noMultiLvlLbl')),
    }

def parse_c_CT_DateAx(node):
    """CT_DateAx → dict  날짜 축.
    XSD: EG_AxShared, auto, lblOffset, baseTimeUnit, majorUnit, majorTimeUnit,
         minorUnit, minorTimeUnit"""
    if node is None: return {}
    return {
        **parse_c_EG_AxShared(node),
        'auto':          parse_c_CT_Boolean(_get_child(node, 'auto')),
        'lblOffset':     _get_val(_get_child(node, 'lblOffset')),
        'baseTimeUnit':  _get_val(_get_child(node, 'baseTimeUnit')),
        'majorUnit':     parse_c_CT_Double(_get_child(node, 'majorUnit')),
        'majorTimeUnit': _get_val(_get_child(node, 'majorTimeUnit')),
        'minorUnit':     parse_c_CT_Double(_get_child(node, 'minorUnit')),
        'minorTimeUnit': _get_val(_get_child(node, 'minorTimeUnit')),
    }

def parse_c_CT_SerAx(node):
    """CT_SerAx → dict  계열 축.
    XSD: EG_AxShared, tickLblSkip, tickMarkSkip"""
    if node is None: return {}
    return {
        **parse_c_EG_AxShared(node),
        'tickLblSkip':  parse_c_CT_UnsignedInt(_get_child(node, 'tickLblSkip')),
        'tickMarkSkip': parse_c_CT_UnsignedInt(_get_child(node, 'tickMarkSkip')),
    }

def parse_c_CT_ValAx(node):
    """CT_ValAx → dict  값 축.
    XSD: EG_AxShared, crossBetween, majorUnit, minorUnit, dispUnits"""
    if node is None: return {}
    return {
        **parse_c_EG_AxShared(node),
        'crossBetween': _get_val(_get_child(node, 'crossBetween')),
        'majorUnit':    parse_c_CT_Double(_get_child(node, 'majorUnit')),
        'minorUnit':    parse_c_CT_Double(_get_child(node, 'minorUnit')),
        'dispUnits':    parse_c_CT_DispUnits(_get_child(node, 'dispUnits')),
    }

# ==============================================================================
# 데이터 테이블
# CT_DTable
# ==============================================================================

def parse_c_CT_DTable(node):
    """CT_DTable → dict  차트 하단에 붙는 데이터 테이블 표시 설정.
    Markdown 출력에서는 차트 데이터를 이미 테이블로 렌더링하므로
    별도의 DTable 내용을 추가 출력할 필요가 없어 렌더러 미사용.
    spPr·txPr는 시각 서식 — 생략."""
    if node is None: return {}
    return {
        'showHorzBorder': parse_c_CT_Boolean(_get_child(node, 'showHorzBorder')),
        'showVertBorder': parse_c_CT_Boolean(_get_child(node, 'showVertBorder')),
        'showOutline':    parse_c_CT_Boolean(_get_child(node, 'showOutline')),
        'showKeys':       parse_c_CT_Boolean(_get_child(node, 'showKeys')),
        # spPr, txPr: 시각 서식 — 생략
        'spPr': None,
        'txPr': None,
    }

# ==============================================================================
# PlotArea 디스패처
# ==============================================================================

CHART_TYPE_DISPATCH = {
    'lineChart':      parse_c_CT_LineChart,
    'line3DChart':    parse_c_CT_Line3DChart,
    'stockChart':     parse_c_CT_StockChart,
    'barChart':       parse_c_CT_BarChart,
    'bar3DChart':     parse_c_CT_Bar3DChart,
    'areaChart':      parse_c_CT_AreaChart,
    'area3DChart':    parse_c_CT_Area3DChart,
    'scatterChart':   parse_c_CT_ScatterChart,
    'radarChart':     parse_c_CT_RadarChart,
    'pieChart':       parse_c_CT_PieChart,
    'pie3DChart':     parse_c_CT_Pie3DChart,
    'doughnutChart':  parse_c_CT_DoughnutChart,
    'ofPieChart':     parse_c_CT_OfPieChart,
    'bubbleChart':    parse_c_CT_BubbleChart,
    'surfaceChart':   parse_c_CT_SurfaceChart,
    'surface3DChart': parse_c_CT_Surface3DChart,
}

AX_DISPATCH = {
    'catAx':  parse_c_CT_CatAx,
    'dateAx': parse_c_CT_DateAx,
    'serAx':  parse_c_CT_SerAx,
    'valAx':  parse_c_CT_ValAx,
}

def parse_c_CT_PlotArea(node):
    """CT_PlotArea → dict
    XSD: layout, <choice:16 chart types>(1+), <choice:4 axis types>(0+), dTable, spPr
    axes는 파싱하지만 Markdown 테이블에서 축 설정을 출력하지 않으므로 렌더러 미사용.
    dTable도 파싱하지만 위의 CT_DTable 주석 참조."""
    if node is None: return {'charts': [], 'axes': [], 'dTable': None}
    charts = []
    axes = []
    for child in node:
        tag = _get_tag(child)
        fn_chart = CHART_TYPE_DISPATCH.get(tag)
        if fn_chart is not None:
            charts.append(fn_chart(child))
            continue
        fn_ax = AX_DISPATCH.get(tag)
        if fn_ax is not None:
            axes.append(fn_ax(child))
    return {
        'layout': parse_c_CT_Layout(_get_child(node, 'layout')),
        'charts': charts,
        # axes: 축 설정 파싱 완료. Markdown 테이블은 축 이름·범위를 별도 출력하지 않음
        'axes':   axes,
        # dTable: 파싱 완료. Markdown에서 이미 테이블로 렌더링하므로 중복 출력 불필요
        'dTable': parse_c_CT_DTable(_get_child(node, 'dTable')),
        # spPr: CT_PlotArea 배경 서식 — 시각 전용
        'spPr':   None,
    }

# ==============================================================================
# 범례
# EG_LegendEntryData, CT_LegendEntry, CT_Legend
# ==============================================================================

def parse_c_EG_LegendEntryData(node):
    """EG_LegendEntryData → dict  (txPr만 포함)
    txPr는 텍스트 서식 — 시각 전용이므로 생략."""
    if node is None: return {}
    return {'txPr': None}  # txPr: a:CT_TextBody — 시각 서식 전용

def parse_c_CT_LegendEntry(node):
    """CT_LegendEntry → dict  개별 범례 항목.
    delete=true면 해당 항목 숨김."""
    if node is None: return {}
    idx = parse_c_CT_UnsignedInt(_get_child(node, 'idx'))
    delete_el = _get_child(node, 'delete')
    if delete_el is not None:
        return {'idx': idx, 'delete': parse_c_CT_Boolean(delete_el)}
    return {'idx': idx, 'delete': False, **parse_c_EG_LegendEntryData(node)}

def parse_c_CT_Legend(node):
    """CT_Legend → dict  차트 범례 전체 설정.
    legendPos(b/tr/l/r/t)·legendEntry[]·overlay를 파싱하지만
    Markdown 테이블은 범례를 별도 섹션으로 출력하지 않으므로 렌더러 미사용.
    spPr·txPr는 시각 서식 — 생략."""
    if node is None: return {}
    return {
        'legendPos':    _get_val(_get_child(node, 'legendPos')),
        'legendEntry':  [parse_c_CT_LegendEntry(e) for e in _get_children(node, 'legendEntry')],
        'layout':       parse_c_CT_Layout(_get_child(node, 'layout')),
        'overlay':      parse_c_CT_Boolean(_get_child(node, 'overlay')),
        # spPr, txPr: 시각 서식 — 생략
        'spPr': None,
        'txPr': None,
    }

# ==============================================================================
# 3D 뷰, 바닥/측벽/후벽
# CT_View3D, CT_Surface (3D 차트 벽면용)
# ==============================================================================

def parse_c_CT_View3D(node):
    """CT_View3D → dict  3D 차트 시점 설정 (회전각, 깊이, 원근).
    Markdown 테이블은 2D 평면 출력이므로 렌더러 미사용."""
    if node is None: return {}
    return {
        'rotX':         parse_c_CT_Double(_get_child(node, 'rotX')),
        'hPercent':     _get_val(_get_child(node, 'hPercent')),
        'rotY':         parse_c_CT_Double(_get_child(node, 'rotY')),
        'depthPercent': _get_val(_get_child(node, 'depthPercent')),
        'rAngAx':       parse_c_CT_Boolean(_get_child(node, 'rAngAx')),
        'perspective':  parse_c_CT_Double(_get_child(node, 'perspective')),
    }

def parse_c_CT_Surface(node):
    """CT_Surface → dict  3D 차트의 바닥(floor)/측벽(sideWall)/후벽(backWall).
    XSD: thickness, spPr, pictureOptions
    spPr는 벽면 색상 — 시각 전용이므로 생략. Markdown 테이블에서 미사용."""
    if node is None: return {}
    return {
        'thickness':      _get_val(_get_child(node, 'thickness')),
        # spPr: 벽면 도형 서식 — 시각 전용
        'spPr':           None,
        'pictureOptions': parse_c_CT_PictureOptions(_get_child(node, 'pictureOptions')),
    }

# ==============================================================================
# 피벗 서식
# CT_PivotFmt, CT_PivotFmts
# ==============================================================================

def parse_c_CT_PivotFmt(node):
    """CT_PivotFmt → dict  피벗 차트의 개별 포인트 서식 재정의.
    spPr·txPr는 시각 서식 — 생략. Markdown 테이블에서 미사용."""
    if node is None: return {}
    return {
        'idx':    parse_c_CT_UnsignedInt(_get_child(node, 'idx')),
        # spPr, txPr: 시각 서식 — 생략
        'spPr':   None,
        'txPr':   None,
        'marker': parse_c_CT_Marker(_get_child(node, 'marker')),
        'dLbl':   parse_c_CT_DLbl(_get_child(node, 'dLbl')),
    }

def parse_c_CT_PivotFmts(node):
    """CT_PivotFmts → list[dict]  피벗 서식 목록.
    Markdown 테이블에서 미사용."""
    if node is None: return []
    return [parse_c_CT_PivotFmt(p) for p in _get_children(node, 'pivotFmt')]

# ==============================================================================
# 최상위: CT_Chart, CT_PivotSource, CT_Protection
# CT_ExternalData, CT_PrintSettings, CT_ChartSpace
# ==============================================================================

def parse_c_CT_PivotSource(node):
    """CT_PivotSource → dict  피벗 테이블 연결 정보.
    차트 데이터가 피벗 테이블 기반임을 알리지만 데이터 자체는 시리즈 캐시에서 추출하므로
    Markdown 렌더러에서 미사용."""
    if node is None: return {}
    name_el = _get_child(node, 'name')
    return {
        'name':  (name_el.text or "") if name_el is not None else "",
        'fmtId': parse_c_CT_UnsignedInt(_get_child(node, 'fmtId')),
    }

def parse_c_CT_Protection(node):
    """CT_Protection → dict  차트 편집 잠금 설정.
    chartObject·data·formatting·selection·userInterface 플래그.
    Markdown 변환은 읽기 전용이므로 보호 설정은 렌더러 미사용."""
    if node is None: return {}
    return {
        'chartObject':   parse_c_CT_Boolean(_get_child(node, 'chartObject')),
        'data':          parse_c_CT_Boolean(_get_child(node, 'data')),
        'formatting':    parse_c_CT_Boolean(_get_child(node, 'formatting')),
        'selection':     parse_c_CT_Boolean(_get_child(node, 'selection')),
        'userInterface': parse_c_CT_Boolean(_get_child(node, 'userInterface')),
    }

def parse_c_CT_ExternalData(node):
    """CT_ExternalData → dict  외부 데이터 소스 관계 ID와 자동 갱신 여부.
    데이터는 이미 캐시(numCache/strCache)에서 추출했으므로 렌더러 미사용."""
    if node is None: return {}
    return {
        'rid':        _get_val(node, 'id'),  # r:id 속성
        'autoUpdate': parse_c_CT_Boolean(_get_child(node, 'autoUpdate')),
    }

def parse_c_CT_HeaderFooter(node):
    """CT_HeaderFooter → dict  차트 인쇄 머리글/바닥글.
    Markdown 출력은 인쇄 레이아웃을 지원하지 않으므로 렌더러 미사용."""
    if node is None: return {}
    def txt(tag):
        el = _get_child(node, tag)
        return (el.text or "") if el is not None else ""
    return {
        'oddHeader': txt('oddHeader'), 'oddFooter': txt('oddFooter'),
        'evenHeader': txt('evenHeader'), 'evenFooter': txt('evenFooter'),
        'firstHeader': txt('firstHeader'), 'firstFooter': txt('firstFooter'),
        'alignWithMargins':  node.attrib.get('alignWithMargins'),
        'differentOddEven':  node.attrib.get('differentOddEven'),
        'differentFirst':    node.attrib.get('differentFirst'),
    }

def parse_c_CT_PageMargins(node):
    """CT_PageMargins → dict  인쇄 여백 (l/r/t/b/header/footer).
    Markdown 출력과 무관하므로 렌더러 미사용."""
    if node is None: return {}
    return {k: node.attrib.get(k) for k in ('l','r','t','b','header','footer')}

def parse_c_CT_PageSetup(node):
    """CT_PageSetup → dict  인쇄 설정 (용지 크기, 방향 등).
    Markdown 출력과 무관하므로 렌더러 미사용."""
    if node is None: return {}
    keys = ('paperSize','paperHeight','paperWidth','firstPageNumber',
            'orientation','blackAndWhite','draft','useFirstPageNumber',
            'horizontalDpi','verticalDpi','copies')
    return {k: node.attrib.get(k) for k in keys}

def parse_c_CT_PrintSettings(node):
    """CT_PrintSettings → dict  인쇄 설정 묶음 (headerFooter, pageMargins, pageSetup).
    Markdown 출력과 무관하므로 렌더러 미사용."""
    if node is None: return {}
    return {
        'headerFooter': parse_c_CT_HeaderFooter(_get_child(node, 'headerFooter')),
        'pageMargins':  parse_c_CT_PageMargins(_get_child(node, 'pageMargins')),
        'pageSetup':    parse_c_CT_PageSetup(_get_child(node, 'pageSetup')),
    }

def parse_c_CT_Chart(node):
    """CT_Chart → dict
    XSD: title, autoTitleDeleted, pivotFmts, view3D, floor, sideWall, backWall,
         plotArea(required), legend, plotVisOnly, dispBlanksAs, showDLblsOverMax
    title과 plotArea의 charts만 렌더러에서 사용한다.
    나머지는 파싱하되 Markdown 테이블 출력에 불필요하므로 미사용."""
    if node is None: return {'title': None, 'charts': [], 'axes': []}
    plot = parse_c_CT_PlotArea(_get_child(node, 'plotArea'))
    title_dict = parse_c_CT_Title(_get_child(node, 'title'))
    return {
        # title['text']만 렌더러에서 사용
        'title':  title_dict['text'] or None,
        # autoTitleDeleted: 자동 제목 삭제 여부 — Markdown에서 미사용
        'autoTitleDeleted': parse_c_CT_Boolean(_get_child(node, 'autoTitleDeleted')),
        # pivotFmts: 피벗 서식 재정의 — Markdown에서 미사용
        'pivotFmts': parse_c_CT_PivotFmts(_get_child(node, 'pivotFmts')),
        # view3D: 3D 시점 — Markdown 2D 테이블과 무관
        'view3D':    parse_c_CT_View3D(_get_child(node, 'view3D')),
        # floor/sideWall/backWall: 3D 벽면 서식 — Markdown과 무관
        'floor':     parse_c_CT_Surface(_get_child(node, 'floor')),
        'sideWall':  parse_c_CT_Surface(_get_child(node, 'sideWall')),
        'backWall':  parse_c_CT_Surface(_get_child(node, 'backWall')),
        'charts':    plot['charts'],
        'axes':      plot['axes'],
        'dTable':    plot['dTable'],
        # legend: 파싱 완료, Markdown 테이블은 시리즈명이 열 헤더에 이미 포함되므로 별도 출력 불필요
        'legend':    parse_c_CT_Legend(_get_child(node, 'legend')),
        # plotVisOnly: 숨겨진 셀 무시 여부 — 데이터 필터 기준, Markdown에서 미사용
        'plotVisOnly':       parse_c_CT_Boolean(_get_child(node, 'plotVisOnly')),
        # dispBlanksAs: 빈 셀을 gap/zero/span 중 어떻게 처리할지 — 렌더러는 "" 고정
        'dispBlanksAs':      _get_val(_get_child(node, 'dispBlanksAs')),
        # showDLblsOverMax: 최대값 초과 레이블 표시 여부 — 시각 전용
        'showDLblsOverMax':  parse_c_CT_Boolean(_get_child(node, 'showDLblsOverMax')),
    }

def parse_c_CT_ChartSpace(root, ctx=None):
    """CT_ChartSpace → dict  루트 진입점.
    XSD: date1904, lang, roundedCorners, style, clrMapOvr, pivotSource,
         protection, chart(required), spPr, txPr, externalData,
         printSettings, userShapes
    chart 자식의 title·charts·axes만 렌더러에서 사용.
    나머지는 파싱하되 Markdown 테이블 출력에 불필요하므로 미사용.

    ctx: ZipContext | None — 제공되면 userShapes rid를 .rels 경유로 해석해
         cdr:CT_Drawing 파트를 파싱한다."""
    if root is None: return {'title': None, 'charts': [], 'axes': []}
    if _get_tag(root) == 'chartSpace':
        cs = root
    else:
        cs = root  # 루트가 chartSpace가 아닌 경우 그대로 처리

    # date1904를 chart 파싱 전에 읽어 모듈 수준 플래그 설정
    # parse_c_CT_NumData 의 날짜 직렬 변환에서 이 값을 참조한다
    global _date1904
    _date1904_val = parse_c_CT_Boolean(_get_child(cs, 'date1904'))
    _date1904 = bool(_date1904_val)

    chart_node = _get_child(cs, 'chart')
    chart = parse_c_CT_Chart(chart_node)
    return {
        # date1904: 날짜 직렬 방식 — 모듈 플래그(_date1904)로 parse_c_CT_NumData에 전파됨.
        'date1904':      _date1904_val,
        # lang: 언어 태그 — Markdown 출력 언어는 입력 데이터 그대로 사용하므로 미사용
        'lang':          _get_val(_get_child(cs, 'lang')),
        # roundedCorners: 차트 모서리 둥글기 — 시각 전용
        'roundedCorners': parse_c_CT_Boolean(_get_child(cs, 'roundedCorners')),
        # style: 1~48 번호 테마 스타일 — 색상/서식 테마, 시각 전용
        'style':         _get_val(_get_child(cs, 'style')),
        # clrMapOvr: a:CT_ColorMapping — dml-main.xsd 영역, 색상 매핑, 시각 전용
        'clrMapOvr':     None,
        # pivotSource: 피벗 연결 정보 — 데이터는 캐시에서 추출했으므로 미사용
        'pivotSource':   parse_c_CT_PivotSource(_get_child(cs, 'pivotSource')),
        # protection: 편집 잠금 — 읽기 전용 변환기에서 미사용
        'protection':    parse_c_CT_Protection(_get_child(cs, 'protection')),
        'title':         chart['title'],
        'charts':        chart['charts'],
        'axes':          chart['axes'],
        'dTable':        chart.get('dTable'),
        'legend':        chart.get('legend'),
        # spPr: 차트 공간 배경 서식 — 시각 전용
        'spPr':          None,
        # txPr: 기본 텍스트 서식 — 시각 전용
        'txPr':          None,
        # externalData: 외부 데이터 소스 참조 — 캐시 우선 원칙상 미사용
        'externalData':  parse_c_CT_ExternalData(_get_child(cs, 'externalData')),
        # printSettings: 인쇄 설정 — Markdown 출력과 무관
        'printSettings': parse_c_CT_PrintSettings(_get_child(cs, 'printSettings')),
        # userShapes: CT_RelId — r:id로 chartDrawing 파트를 참조한다.
        # ctx(ZipContext)가 있으면 .rels를 통해 cdr:CT_Drawing 파트를 실제로 파싱한다.
        # ctx가 없으면 rid 문자열만 보존하고 cdr 파싱은 생략한다.
        'userShapes':    _resolve_user_shapes(_get_child(cs, 'userShapes'), ctx),
    }


def _resolve_user_shapes(node, ctx):
    """userShapes CT_RelId 노드와 ZipContext를 받아 cdr:CT_Drawing dict 또는 rid str 반환.

    ctx가 있으면 rid → ZipContext.resolve() → cdr 파트 파싱.
    ctx가 없으면 rid 문자열만 반환한다 (단독 .xml 실행 시 fallback)."""
    rid = parse_c_CT_RelId(node)
    if rid is None:
        return None
    if ctx is None:
        return rid
    drawing_ctx = ctx.resolve(rid)
    if drawing_ctx is None:
        return rid
    try:
        drawing_root = drawing_ctx.parse_xml()
        return parse_cdr_CT_Drawing(drawing_root)
    except Exception:
        return rid

# ==============================================================================
# CT_RelId  (c: 네임스페이스 — dml-chart.xsd 정의)
# ==============================================================================

def parse_c_CT_RelId(node):
    """CT_RelId → str | None   r:id 관계 식별자만 담는 단순 타입.

    XSD 정의 (dml-chart.xsd):
        <xsd:complexType name="CT_RelId">
          <xsd:attribute ref="r:id" use="required"/>
        </xsd:complexType>

    r:id 속성값(예: "rId1")을 문자열로 반환한다.
    실제 대상 파트(userShapes drawing XML 등)를 읽으려면 .rels 파일을 통해
    ZipContext.resolve(rid)를 호출해야 한다.
    """
    if node is None:
        return None
    # r:id 속성은 네임스페이스 프리픽스 포함 또는 로컬명 'id' 형태로 나타날 수 있다
    for k, v in node.attrib.items():
        local = k.split('}')[-1] if '}' in k else k
        if local == 'id':
            return v
    return None


# ==============================================================================
# cdr:* 파서  (dml-chartDrawing.xsd)
# c:chartSpace 의 userShapes 요소가 참조하는 별도 파트(chart#Drawing.xml)의 루트 타입.
# ==============================================================================

def parse_cdr_CT_ShapeNonVisual(node):
    """cdr:CT_ShapeNonVisual → dict
    XSD: cNvPr(CT_NonVisualDrawingProps) + cNvSpPr(CT_NonVisualDrawingShapeProps)
    a: 네임스페이스 비시각 속성 — id/name/title 등 메타만 보존.
    Markdown 출력에서 직접 사용하지 않는다 (도형 이름이 필요할 때만 참조)."""
    if node is None:
        return {}
    cNvPr = _get_child(node, 'cNvPr')
    return {
        'id':   cNvPr.get('id') if cNvPr is not None else None,
        'name': cNvPr.get('name') if cNvPr is not None else None,
    }

def parse_cdr_CT_Shape(node):
    """cdr:CT_Shape → dict
    XSD: nvSpPr + spPr + style? + txBody?
    Markdown에서는 txBody 텍스트만 추출한다.
    spPr(도형 서식), style(테마 스타일)은 시각 전용이므로 미사용."""
    if node is None:
        return {}
    return {
        'nvSpPr': parse_cdr_CT_ShapeNonVisual(_get_child(node, 'nvSpPr')),
        # spPr: CT_ShapeProperties — 위치·크기·채우기·선 서식 등 시각 전용
        'spPr':   None,
        # style: CT_ShapeStyle — 테마 스타일 참조 시각 전용
        'style':  None,
        # txBody: 도형 내 텍스트 — Markdown에 텍스트가 있으면 포함 가능
        'txBody': parse_a_CT_TextBody(_get_child(node, 'txBody')),
        # 속성
        'macro':       node.get('macro'),
        'textlink':    node.get('textlink'),
        'fLocksText':  node.get('fLocksText', 'true'),
        'fPublished':  node.get('fPublished', 'false'),
    }

def parse_cdr_CT_ConnectorNonVisual(node):
    """cdr:CT_ConnectorNonVisual → dict
    XSD: cNvPr + cNvCxnSpPr
    연결선 비시각 속성 — 시각 전용, Markdown에서 미사용."""
    if node is None:
        return {}
    cNvPr = _get_child(node, 'cNvPr')
    return {
        'id':   cNvPr.get('id') if cNvPr is not None else None,
        'name': cNvPr.get('name') if cNvPr is not None else None,
    }

def parse_cdr_CT_Connector(node):
    """cdr:CT_Connector → dict
    XSD: nvCxnSpPr + spPr + style?
    연결선(선 도형) — 시작·끝 연결 정보 포함. Markdown 출력에서 미사용.
    이유: 연결선은 도형 간 관계를 나타내며 텍스트 데이터가 없다."""
    if node is None:
        return {}
    return {
        'nvCxnSpPr': parse_cdr_CT_ConnectorNonVisual(_get_child(node, 'nvCxnSpPr')),
        'macro':      node.get('macro'),
        'fPublished': node.get('fPublished', 'false'),
    }

def parse_cdr_CT_PictureNonVisual(node):
    """cdr:CT_PictureNonVisual → dict
    XSD: cNvPr + cNvPicPr
    그림 비시각 속성 — 시각 전용, Markdown에서 미사용."""
    if node is None:
        return {}
    cNvPr = _get_child(node, 'cNvPr')
    return {
        'id':   cNvPr.get('id') if cNvPr is not None else None,
        'name': cNvPr.get('name') if cNvPr is not None else None,
    }

def parse_cdr_CT_Picture(node):
    """cdr:CT_Picture → dict
    XSD: nvPicPr + blipFill + spPr + style?
    이미지 삽입 도형. blipFill의 r:embed rid는 ZipContext로 이미지 파일을 가리킨다.
    Markdown 출력에서는 이미지를 ![](path) 형태로 표현할 여지가 있으나
    현재 렌더러는 차트 데이터 테이블만 출력하므로 미사용."""
    if node is None:
        return {}
    # blipFill > a:blip r:embed 에서 이미지 rid 추출
    blip_fill = _get_child(node, 'blipFill')
    blip = _get_child(blip_fill, 'blip') if blip_fill is not None else None
    img_rid = None
    if blip is not None:
        for k, v in blip.attrib.items():
            if k.split('}')[-1] == 'embed':
                img_rid = v
                break
    return {
        'nvPicPr':  parse_cdr_CT_PictureNonVisual(_get_child(node, 'nvPicPr')),
        'img_rid':  img_rid,
        'macro':    node.get('macro', ''),
        'fPublished': node.get('fPublished', 'false'),
    }

def parse_cdr_CT_GraphicFrameNonVisual(node):
    """cdr:CT_GraphicFrameNonVisual → dict
    XSD: cNvPr + cNvGraphicFramePr
    그래픽 프레임 비시각 속성 — 시각 전용, Markdown에서 미사용."""
    if node is None:
        return {}
    cNvPr = _get_child(node, 'cNvPr')
    return {
        'id':   cNvPr.get('id') if cNvPr is not None else None,
        'name': cNvPr.get('name') if cNvPr is not None else None,
    }

def parse_cdr_CT_GraphicFrame(node):
    """cdr:CT_GraphicFrame → dict
    XSD: nvGraphicFramePr + xfrm + a:graphic
    차트 드로잉 안에 다시 a:graphic(예: 내장 차트)을 담을 수 있다.
    Markdown 출력에서는 내부 그래픽 데이터를 재귀 처리하지 않는다.
    이유: userShapes는 차트 위에 덧그린 장식 도형이며, 그 안의 중첩 차트는
    별도 파싱 파이프라인이 필요해 현재 범위를 벗어난다."""
    if node is None:
        return {}
    return {
        'nvGraphicFramePr': parse_cdr_CT_GraphicFrameNonVisual(
            _get_child(node, 'nvGraphicFramePr')),
        'macro':      node.get('macro'),
        'fPublished': node.get('fPublished', 'false'),
    }

def parse_cdr_CT_GroupShapeNonVisual(node):
    """cdr:CT_GroupShapeNonVisual → dict
    XSD: cNvPr + cNvGrpSpPr
    그룹 도형 비시각 속성 — 시각 전용, Markdown에서 미사용."""
    if node is None:
        return {}
    cNvPr = _get_child(node, 'cNvPr')
    return {
        'id':   cNvPr.get('id') if cNvPr is not None else None,
        'name': cNvPr.get('name') if cNvPr is not None else None,
    }

def parse_cdr_EG_ObjectChoices(node):
    """cdr:EG_ObjectChoices — sp | grpSp | graphicFrame | cxnSp | pic 중 하나.
    XSD group으로 정의되며 CT_RelSizeAnchor / CT_AbsSizeAnchor 내부에서 사용.
    해당하는 자식 태그를 찾아 각 파서로 디스패치한다."""
    if node is None:
        return None
    _dispatch = {
        'sp':           parse_cdr_CT_Shape,
        'grpSp':        parse_cdr_CT_GroupShape,
        'graphicFrame': parse_cdr_CT_GraphicFrame,
        'cxnSp':        parse_cdr_CT_Connector,
        'pic':          parse_cdr_CT_Picture,
    }
    for child in node:
        tag = _get_tag(child)
        if tag in _dispatch:
            return _dispatch[tag](child)
    return None

def parse_cdr_CT_GroupShape(node):
    """cdr:CT_GroupShape → dict
    XSD: nvGrpSpPr + grpSpPr + (sp | grpSp | graphicFrame | cxnSp | pic)*
    그룹 도형 — 재귀 구조. 텍스트가 있는 sp(도형) 자식만 수집한다.
    grpSpPr(그룹 서식)은 시각 전용이므로 미사용."""
    if node is None:
        return {}
    shapes = []
    for child in node:
        tag = _get_tag(child)
        if tag == 'sp':
            shapes.append(parse_cdr_CT_Shape(child))
        elif tag == 'grpSp':
            sub = parse_cdr_CT_GroupShape(child)
            shapes.extend(sub.get('shapes', []))
        elif tag == 'graphicFrame':
            shapes.append(parse_cdr_CT_GraphicFrame(child))
        elif tag == 'cxnSp':
            shapes.append(parse_cdr_CT_Connector(child))
        elif tag == 'pic':
            shapes.append(parse_cdr_CT_Picture(child))
    return {
        'nvGrpSpPr': parse_cdr_CT_GroupShapeNonVisual(_get_child(node, 'nvGrpSpPr')),
        'shapes': shapes,
    }

def parse_cdr_ST_MarkerCoordinate(val):
    """cdr:ST_MarkerCoordinate → float   0.0–1.0 범위 상대 좌표.
    XSD: xsd:double { minInclusive='0.0' maxInclusive='1.0' }
    cdr:CT_Marker의 x/y 요소 값으로 사용된다.
    시각 레이아웃 좌표 — Markdown에서 미사용."""
    try:
        return float(val)
    except (TypeError, ValueError):
        return None


def parse_cdr_CT_Marker(node):
    """cdr:CT_Marker → dict  상대 좌표 마커 (c:CT_Marker와 별개).
    XSD: x(ST_MarkerCoordinate 0.0–1.0) + y(ST_MarkerCoordinate 0.0–1.0)
    차트 영역 내 상대 위치를 나타낸다. 시각 레이아웃 전용이므로 Markdown에서 미사용."""
    if node is None:
        return {}
    x_node = _get_child(node, 'x')
    y_node = _get_child(node, 'y')
    return {
        'x': float(x_node.text) if x_node is not None and x_node.text else None,
        'y': float(y_node.text) if y_node is not None and y_node.text else None,
    }

def parse_cdr_CT_RelSizeAnchor(node):
    """cdr:CT_RelSizeAnchor → dict
    XSD: from(CT_Marker) + to(CT_Marker) + EG_ObjectChoices
    상대 크기 앵커 — from/to 좌표가 모두 0.0–1.0 상대값. 시각 레이아웃 전용.
    Markdown 출력에서 도형 텍스트는 수집하되 위치 정보는 미사용."""
    if node is None:
        return {}
    return {
        'from':   parse_cdr_CT_Marker(_get_child(node, 'from')),
        'to':     parse_cdr_CT_Marker(_get_child(node, 'to')),
        'object': parse_cdr_EG_ObjectChoices(node),
    }

def parse_cdr_CT_AbsSizeAnchor(node):
    """cdr:CT_AbsSizeAnchor → dict
    XSD: from(CT_Marker) + ext(a:CT_PositiveSize2D) + EG_ObjectChoices
    절대 크기 앵커 — from은 상대 좌표, ext는 EMU 단위 절대 크기. 시각 레이아웃 전용.
    Markdown 출력에서 도형 텍스트는 수집하되 크기/위치 정보는 미사용."""
    if node is None:
        return {}
    ext = _get_child(node, 'ext')
    return {
        'from': parse_cdr_CT_Marker(_get_child(node, 'from')),
        'ext':  {'cx': ext.get('cx') if ext is not None else None,
                 'cy': ext.get('cy') if ext is not None else None},
        'object': parse_cdr_EG_ObjectChoices(node),
    }

def parse_cdr_EG_Anchor(node):
    """cdr:EG_Anchor — relSizeAnchor | absSizeAnchor 중 하나.
    CT_Drawing의 반복 자식으로 등장한다."""
    if node is None:
        return None
    tag = _get_tag(node)
    if tag == 'relSizeAnchor':
        return parse_cdr_CT_RelSizeAnchor(node)
    if tag == 'absSizeAnchor':
        return parse_cdr_CT_AbsSizeAnchor(node)
    return None

def parse_cdr_CT_Drawing(node):
    """cdr:CT_Drawing → dict   chartDrawing 파트의 루트 타입.
    XSD: (relSizeAnchor | absSizeAnchor)*
    차트 위에 덧그린 도형들의 컨테이너이다.
    각 앵커 안의 sp(도형) txBody 텍스트만 Markdown에 포함한다.
    이유: 앵커 위치/크기·연결선·그래픽 프레임 등은 Markdown 텍스트로 표현 불가."""
    if node is None:
        return {'anchors': []}
    anchors = []
    for child in node:
        tag = _get_tag(child)
        if tag in ('relSizeAnchor', 'absSizeAnchor'):
            anchors.append(parse_cdr_EG_Anchor(child))
    return {'anchors': anchors}


# ==============================================================================
# Markdown 렌더러
# ==============================================================================

def _escape_pipe(s):
    return str(s).replace('|', r'\|')


def _fmt_num(v):
    """Format a numeric string for display: trim IEEE-754 floating-point noise.

    '0.15837999999999999' → '0.158380'  (6 significant figures, trailing-zero-stripped)
    '36'                  → '36'
    non-numeric strings   → returned as-is
    """
    if not v:
        return v
    try:
        f = float(v)
        formatted = f"{f:.6g}"
        return formatted
    except (ValueError, TypeError):
        return str(v)


def _render_series_table(series_list, chart_type):
    """시리즈 데이터를 Markdown 테이블로 렌더링.

    Scatter → 시리즈별 X|Y 테이블.
    Bubble  → 시리즈별 X|Y|Bubble Size 테이블.
    그 외    → 카테고리 행 × 시리즈 열 통합 테이블.

    각 series dict에 파싱된 marker·dpt·dlbls·trendlines·err_bars 등은
    이 함수 내에서 사용하지 않는다.
    이유: Markdown 테이블 셀은 숫자/문자 값만 표현하며 색상·마커 모양·
    추세선·오류 범위 같은 시각적 메타정보를 담을 방법이 없기 때문이다."""
    if not series_list:
        return "_No series data._"

    # Sort by 'order' field so series appear in chart display order,
    # not XML file order (which may differ, e.g. order=0 series stored last).
    series_list = sorted(series_list, key=lambda s: s.get('order', 0))

    is_scatter = (chart_type == "Scatter Chart")
    is_bubble  = (chart_type == "Bubble Chart")

    if is_scatter:
        # Build a single unified table: each series contributes two columns (X, Y).
        # Rows are aligned by index; shorter series are padded with empty cells.
        series_data = []
        for s in series_list:
            xs = [_fmt_num(v) for v in (s.get('x_values') or [])]
            ys = [_fmt_num(v) for v in (s.get('y_values') or [])]
            series_data.append((s['name'] or "Series", xs, ys))

        max_rows = max(
            (max(len(xs), len(ys)) for _, xs, ys in series_data),
            default=0,
        )
        if max_rows == 0:
            return "_No series data._"

        limit = min(max_rows, _MAX_SCATTER_ROWS)

        # 열 헤더 결정
        # Scatter 시리즈의 이름(tx)은 ECMA-376 EG_SerShared.tx — 범례 레이블이다.
        # X/Y 열 자체에 대한 별도 이름은 스펙에 없으므로 tx를 식별자로 활용한다.
        #
        # 단일 시리즈: X / Y
        #   시리즈가 하나뿐이므로 시리즈명으로 구분할 필요가 없다.
        #   "X (name)" 형식은 name이 우연히 "Y 값"인 경우 "X (Y 값)"처럼
        #   Y축 레이블로 오인될 수 있어 단순 X / Y가 더 명확하다.
        #
        # 다중 시리즈: X (name) / Y (name)
        #   어떤 X-Y 쌍이 어느 시리즈에 속하는지 구분하기 위해 시리즈명이 필요하다.
        headers = []
        if len(series_data) == 1:
            headers = ["X", "Y"]
        else:
            for name, _, _ in series_data:
                n = _escape_pipe(name)
                headers += [f"X ({n})", f"Y ({n})"]

        rows = []
        for i in range(limit):
            row = []
            for _, xs, ys in series_data:
                x = _escape_pipe(xs[i] if i < len(xs) else "")
                y = _escape_pipe(ys[i] if i < len(ys) else "")
                row += [x, y]
            # Skip rows where every cell is empty
            if not any(c for c in row):
                continue
            rows.append(row)

        if not rows:
            return "_No series data._"

        tbl = _format_table(headers, rows)
        if max_rows > limit:
            tbl += f"\n_(... {max_rows - limit} more rows)_"
        return tbl

    if is_bubble:
        # Single unified table: X / Y / Size columns per series.
        series_data = []
        for s in series_list:
            xs = [_fmt_num(v) for v in (s.get('x_values') or [])]
            ys = [_fmt_num(v) for v in (s.get('y_values') or [])]
            bs = [_fmt_num(v) for v in (s.get('bubble_size') or [])]
            series_data.append((s['name'] or "Series", xs, ys, bs))

        max_rows = max(
            (max(len(xs), len(ys), len(bs)) for _, xs, ys, bs in series_data),
            default=0,
        )
        if max_rows == 0:
            return "_No series data._"

        limit = min(max_rows, _MAX_SCATTER_ROWS)

        # Scatter와 동일한 이유로 단일/다중 시리즈를 분기한다.
        # 단일 시리즈: X / Y / Size
        # 다중 시리즈: X (name) / Y (name) / Size (name)
        headers = []
        if len(series_data) == 1:
            headers = ["X", "Y", "Size"]
        else:
            for name, _, _, _ in series_data:
                n = _escape_pipe(name)
                headers += [f"X ({n})", f"Y ({n})", f"Size ({n})"]

        rows = []
        for i in range(limit):
            row = []
            for _, xs, ys, bs in series_data:
                x = _escape_pipe(xs[i] if i < len(xs) else "")
                y = _escape_pipe(ys[i] if i < len(ys) else "")
                b = _escape_pipe(bs[i] if i < len(bs) else "")
                row += [x, y, b]
            if not any(c for c in row):
                continue
            rows.append(row)

        if not rows:
            return "_No series data._"

        tbl = _format_table(headers, rows)
        if max_rows > limit:
            tbl += f"\n_(... {max_rows - limit} more rows)_"
        return tbl

    # 일반: 카테고리 × 시리즈 그리드
    categories = None
    for s in series_list:
        if s.get('categories'):
            categories = s['categories']
            break

    max_vals = max((len(s['values']) for s in series_list), default=0)
    n_rows = max(len(categories) if categories else 0, max_vals)

    if n_rows == 0:
        return "_No data points._"

    headers = [""] + [s['name'] or f"Series {i+1}"
                               for i, s in enumerate(series_list)]
    rows = []
    for i in range(n_rows):
        cat_label = (categories[i] if categories and i < len(categories) else "")
        row = [_escape_pipe(cat_label)]
        for s in series_list:
            vals = s['values']
            row.append(_escape_pipe(_fmt_num(vals[i] if i < len(vals) else "")))
        rows.append(row)

    return _format_table(headers, rows)

def _format_table(headers, rows):
    """열 너비 패딩이 적용된 Markdown 파이프 테이블 생성."""
    col_widths = [len(str(h)) for h in headers]
    for row in rows:
        for j, cell in enumerate(row):
            if j < len(col_widths):
                col_widths[j] = max(col_widths[j], len(str(cell)))

    def fmt_row(cells):
        padded = []
        for j, cell in enumerate(cells):
            w = col_widths[j] if j < len(col_widths) else len(str(cell))
            padded.append(str(cell).ljust(w))
        return "| " + " | ".join(padded) + " |"

    sep = "| " + " | ".join("-" * w for w in col_widths) + " |"
    lines = [fmt_row(headers), sep] + [fmt_row(r) for r in rows]
    return "\n".join(lines)

_MAX_SCATTER_ROWS = 50  # limit for scatter/bubble data rows


def _collect_drawing_texts(obj) -> list[str]:
    """cdr 도형 객체(또는 userShapes dict)에서 비어있지 않은 텍스트를 재귀 수집."""
    if obj is None or isinstance(obj, str):
        return []
    texts = []
    # CT_Drawing top-level: has 'anchors' list
    if 'anchors' in obj:
        for anchor in obj.get('anchors', []):
            texts.extend(_collect_drawing_texts(anchor.get('object')))
        return texts
    # CT_Shape: has 'txBody' string
    tx = obj.get('txBody', '')
    if isinstance(tx, str) and tx.strip():
        texts.append(tx.strip())
    # CT_GroupShape: has 'shapes' list
    for shape in obj.get('shapes', []):
        texts.extend(_collect_drawing_texts(shape))
    return texts


def _render_chart_to_markdown(title, charts, axes=None, user_shapes=None):
    """차트 전체(제목 + 축 레이블 + 서브 차트 목록 + chartDrawing 텍스트)를 Markdown으로 렌더링."""
    lines = []
    if title:
        lines.append(f"**{title}**")
        lines.append("")

    # Render axis labels if present (skip deleted axes)
    if axes:
        ax_lines = []
        for ax in axes:
            if ax.get('delete') is True:
                continue
            ax_pos = ax.get('axPos', '')
            ax_title_dict = ax.get('title') or {}
            ax_text = ax_title_dict.get('text', '') if isinstance(ax_title_dict, dict) else ''
            ax_text = (ax_text or '').strip()
            if not ax_text:
                continue
            if ax_pos in ('b', 't'):
                ax_lines.append(f"X-axis: {ax_text}")
            elif ax_pos in ('l', 'r'):
                ax_lines.append(f"Y-axis: {ax_text}")
        if ax_lines:
            lines.extend(ax_lines)
            lines.append("")

    _NON_GRID = {"Scatter Chart", "Bubble Chart"}
    can_merge = len(charts) > 1 and not any(
        c.get('chart_type') in _NON_GRID for c in charts
    )
    if can_merge:
        # 콤보 차트: 모든 서브 차트의 시리즈를 하나의 테이블로 병합
        unique_types = list(dict.fromkeys(c.get('chart_type', 'Chart') for c in charts))
        all_series = [s for c in charts for s in c.get('series', [])]
        lines.append(f"**Chart type:** {', '.join(unique_types)}")
        lines.append("")
        lines.append(_render_series_table(all_series, unique_types[0]))
        lines.append("")
    else:
        for chart in charts:
            chart_type = chart.get('chart_type', 'Chart')
            series = chart.get('series', [])
            lines.append(f"**Chart type:** {chart_type}")
            lines.append("")
            lines.append(_render_series_table(series, chart_type))
            lines.append("")

    # chartDrawing texts (labels/callouts drawn on top of the chart)
    if user_shapes and not isinstance(user_shapes, str):
        drawing_texts = _collect_drawing_texts(user_shapes)
        for t in drawing_texts:
            lines.append(t)
        if drawing_texts:
            lines.append("")

    return "\n".join(lines).rstrip() + "\n"


def _render_chart_to_markdown_excel(title, chart_types, axes, excel_table):
    """Excel 테이블 기반 렌더러 (chart_source='excel' 경로).

    제목·차트 타입·축 레이블은 XML 파싱 결과에서, 데이터 테이블은 Excel에서 가져온다.
    여러 sub-chart(콤보 차트)가 있으면 타입명을 쉼표로 연결한다."""
    lines = []
    if title:
        lines.append(f"**{title}**")
        lines.append("")

    # Render axis labels (same logic as _render_chart_to_markdown)
    if axes:
        ax_lines = []
        for ax in axes:
            if ax.get('delete') is True:
                continue
            ax_pos = ax.get('axPos', '')
            ax_title_dict = ax.get('title') or {}
            ax_text = ax_title_dict.get('text', '') if isinstance(ax_title_dict, dict) else ''
            ax_text = (ax_text or '').strip()
            if not ax_text:
                continue
            if ax_pos in ('b', 't'):
                ax_lines.append(f"X-axis: {ax_text}")
            elif ax_pos in ('l', 'r'):
                ax_lines.append(f"Y-axis: {ax_text}")
        if ax_lines:
            lines.extend(ax_lines)
            lines.append("")

    if chart_types:
        lines.append(f"**Chart type:** {', '.join(dict.fromkeys(chart_types))}")
        lines.append("")

    lines.append(excel_table)
    return "\n".join(lines).rstrip() + "\n"

# ==============================================================================
# I/O 헬퍼
# ==============================================================================

_PML_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
# 차트 관계 타입 식별 문자열 (전통 DrawingML + chartEx)
_CHART_REL_TYPE_FRAGMENTS = (
    "relationships/chart",    # .../officeDocument/2006/relationships/chart
    "relationships/chartEx",  # .../office/2014/relationships/chartEx
)


def _load_pptx_charts_in_slide_order(
    zf: zipfile.ZipFile,
) -> list[tuple[ET.Element, "ZipContext"]]:
    """PPTX의 presentation.xml → sldIdLst 순으로 차트 파트를 반환한다.

    파일명 정렬 대신 슬라이드 순서를 따르므로 chartEx 파일이 전통 차트와
    섞여 있어도 슬라이드에 배치된 순서가 그대로 유지된다.
    """
    names = set(zf.namelist())
    prs_root = ET.fromstring(zf.read("ppt/presentation.xml"))
    prs_rels = ET.fromstring(zf.read("ppt/_rels/presentation.xml.rels"))
    rid_to_target = {r.get("Id"): r.get("Target") for r in prs_rels}

    results: list[tuple[ET.Element, ZipContext]] = []
    for sld_id_el in prs_root.findall(
        f".//{{{_PML_NS}}}sldIdLst/{{{_PML_NS}}}sldId"
    ):
        rid = sld_id_el.get(f"{{{_REL_NS}}}id")
        target = rid_to_target.get(rid or "", "")
        if not target:
            continue

        # 슬라이드 경로 → .rels 경로
        sld_path = "ppt/" + target
        sld_rels_path = (
            sld_path
            .replace("slides/slide", "slides/_rels/slide")
            .replace(".xml", ".xml.rels")
        )
        if sld_rels_path not in names:
            continue

        sld_rels = ET.fromstring(zf.read(sld_rels_path))
        for rel in sld_rels:
            rel_type = rel.get("Type", "")
            if not any(t in rel_type for t in _CHART_REL_TYPE_FRAGMENTS):
                continue
            chart_target = rel.get("Target", "")
            if not chart_target:
                continue
            # Target은 "../charts/chartN.xml" 형태 (slides/ 기준 상대 경로)
            chart_path = "ppt/charts/" + chart_target.split("/")[-1]
            if chart_path not in names:
                continue
            try:
                chart_root = ET.fromstring(zf.read(chart_path))
            except ET.ParseError as e:
                print(f"Warning: could not parse {chart_path}: {e}", file=sys.stderr)
                continue
            results.append((chart_root, ZipContext(zf, chart_path)))

    return results


def load_chart_parts(path: str) -> list[tuple[ET.Element, Optional["ZipContext"]]]:
    """차트 파트를 (ET.Element, ZipContext | None) 쌍의 목록으로 반환.

    .pptx: presentation.xml의 sldIdLst 순으로 탐색해 슬라이드 순서를 보장한다.
    .xlsx / .docx (zip): 슬라이드 개념이 없으므로 파일명 natural sort로 탐색한다.
    .xml: ET.Element만 반환, ZipContext는 None.

    ZipContext가 사용 가능하면 각 chart 파트의 ZipContext도 함께 반환해
    userShapes 같은 관계 참조를 resolve할 수 있게 한다.
    """
    if path.lower().endswith('.xml'):
        try:
            root = ET.parse(path).getroot()
            return [(root, None)]
        except ET.ParseError as e:
            print(f"Error parsing {path}: {e}", file=sys.stderr)
            return []

    def _natural_key(s):
        import re
        return [int(c) if c.isdigit() else c for c in re.split(r'(\d+)', s)]

    results = []
    try:
        zf = zipfile.ZipFile(path, 'r')

        if path.lower().endswith('.pptx'):
            # PPTX: 슬라이드 순서 보장
            results = _load_pptx_charts_in_slide_order(zf)
        else:
            # xlsx / docx: 파일명 natural sort
            names = zf.namelist()
            chart_paths = sorted(
                (n for n in names
                 if n.endswith('.xml') and '/charts/chart' in n),
                key=_natural_key,
            )
            if not chart_paths:
                chart_paths = sorted(
                    (n for n in names
                     if 'chart' in n.lower() and n.endswith('.xml') and 'charts' in n),
                    key=_natural_key,
                )
            for cp in chart_paths:
                try:
                    with zf.open(cp) as f:
                        root = ET.parse(f).getroot()
                    results.append((root, ZipContext(zf, cp)))
                except ET.ParseError as e:
                    print(f"Warning: could not parse {cp}: {e}", file=sys.stderr)

        # ZipContext를 사용 중이면 zf를 열어 두어야 하므로 닫지 않는다.
        # 결과가 없으면 닫는다.
        if not results:
            zf.close()
    except zipfile.BadZipFile:
        try:
            root = ET.parse(path).getroot()
            results.append((root, None))
        except ET.ParseError as e:
            print(f"Error parsing {path}: {e}", file=sys.stderr)
    return results


# ==============================================================================
# 공개 API
# ==============================================================================

def convert_chart(
    root: ET.Element,
    ctx: Optional["ZipContext"] = None,
    chart_source: str = "xml",
) -> str:
    """차트 XML 루트 Element를 Markdown 테이블 문자열로 변환한다.

    Args:
        root:         c:chartSpace 루트 ET.Element.
        ctx:          ZipContext 인스턴스 (userShapes 등 관계 참조 해석에 사용).
                      None이면 관계 참조를 건너뛴다.
        chart_source: "xml" (기본값) 또는 "excel".
                      "excel"이면 ctx를 통해 내장 .xlsx 워크북을 우선 조회하고,
                      없거나 실패하면 XML 파싱 결과로 자동 폴백한다.

    Returns:
        제목(있을 경우)과 시리즈 데이터를 담은 Markdown 테이블 문자열.
    """
    # XML 파싱은 항상 수행 — 메타(제목·타입·축)는 XML에서 가져온다.
    result = parse_c_CT_ChartSpace(root, ctx)

    if chart_source == "excel" and ctx is not None:
        for _rid, child_ctx in ctx.resolve_all().items():
            if child_ctx.path.lower().endswith((".xlsx", ".xlsm")):
                try:
                    from .excel2md import convert_excel_to_table
                    excel_bytes = child_ctx.read_bytes()
                    excel_table = convert_excel_to_table(excel_bytes)
                    if excel_table:
                        # 제목·차트타입·축은 XML에서, 데이터는 Excel에서
                        chart_types = [
                            c.get('chart_type', 'Chart')
                            for c in result.get('charts', [])
                        ]
                        return _render_chart_to_markdown_excel(
                            result['title'],
                            chart_types,
                            result.get('axes'),
                            excel_table,
                        )
                except Exception:
                    pass  # Excel 실패 시 XML 폴백

    return _render_chart_to_markdown(
        result['title'],
        result['charts'],
        result.get('axes'),
        result.get('userShapes'),
    )


# ==============================================================================
# main
# ==============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Convert ECMA-376 chart XML to Markdown tables.")
    parser.add_argument('input', help="Input file (.xml, .pptx, .xlsx, .docx)")
    parser.add_argument('-o', '--output', help="Output file (default: stdout)")
    parser.add_argument(
        '--chart-source', choices=('xml', 'excel'), default='xml',
        help="Data source: 'xml' (default) parses DrawingML cache; "
             "'excel' uses the embedded Excel workbook (falls back to xml if not found).",
    )
    args = parser.parse_args()

    contexts = load_chart_parts(args.input)
    if not contexts:
        print("No chart data found.", file=sys.stderr)
        sys.exit(1)

    parts = []
    seen_zf = set()
    for root, ctx in contexts:
        md = convert_chart(root, ctx, chart_source=args.chart_source)
        parts.append(md)
        if ctx is not None and id(ctx.zf) not in seen_zf:
            seen_zf.add(id(ctx.zf))

    # 공유 ZipFile 닫기
    for root, ctx in contexts:
        if ctx is not None:
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

if __name__ == '__main__':
    main()

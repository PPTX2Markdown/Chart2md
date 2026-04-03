"""
ooxml_context.py — OOXML zip 컨테이너 내 파일 컨텍스트 관리

Office Open XML 파일(.xlsx, .docx, .pptx)은 ZIP 아카이브이며 내부 파트들은
.rels 파일을 통해 서로 참조한다. ZipContext는 열린 ZipFile과 현재 파트 경로를
함께 캡슐화하고, r:id → 대상 ZipContext 해석을 lazy하게 수행한다.

사용 예:
    ctx = ZipContext.from_path("workbook.xlsx")
    chart_ctx = ctx.resolve("rId1")        # .rels 경유 해석
    root = chart_ctx.parse_xml()           # ET.Element 반환
    raw  = chart_ctx.read_bytes()          # OLE 바이너리 등

여러 변환 모듈(chart2md, omml2md, ole2md, smartArt2md)이 공유한다.

Public API
----------
ZipContext
    열린 ZIP 아카이브와 특정 파트를 함께 나타내는 컨텍스트 객체.
iter_parts_matching(zf, pattern) -> list[ZipContext]
    ZIP 파일에서 glob 패턴에 맞는 파트를 ZipContext 리스트로 반환한다.
"""

from __future__ import annotations

__all__ = ["ZipContext", "iter_parts_matching"]

import posixpath
import urllib.parse
import xml.etree.ElementTree as ET
import zipfile
from typing import Optional


# ---------------------------------------------------------------------------
# 내부 헬퍼
# ---------------------------------------------------------------------------

# .rels 네임스페이스 (Transitional / Strict 모두 동일)
_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _rels_path(part_path: str) -> str:
    """파트 경로로부터 해당하는 .rels 파일 경로를 반환한다.

    ECMA-376 §13.2.4: 파트 /a/b/c.xml 의 관계 파일은 /a/b/_rels/c.xml.rels
    """
    directory = posixpath.dirname(part_path)
    filename = posixpath.basename(part_path)
    return posixpath.join(directory, "_rels", filename + ".rels")


def _parse_rels_file(data: bytes) -> dict[str, tuple[str, bool]]:
    """bytes로 읽은 .rels XML을 파싱해 {Id: (Target, is_external)} 매핑을 반환한다.

    Target이 절대 경로(/)로 시작하면 그대로, 상대 경로이면 나중에 기준 경로와
    결합해야 하므로 여기서는 raw Target 문자열을 보존한다.
    """
    root = ET.fromstring(data)
    result: dict[str, tuple[str, bool]] = {}
    for rel in root:
        tag = rel.tag
        # 네임스페이스 제거
        if "}" in tag:
            tag = tag.split("}", 1)[1]
        if tag == "Relationship":
            rid = rel.get("Id", "")
            target = rel.get("Target", "")
            is_external = rel.get("TargetMode") == "External"
            if rid:
                result[rid] = (target, is_external)
    return result


def _resolve_target(base_path: str, target: str) -> str:
    """기준 파트 경로와 Target 문자열로 정규화된 절대 파트 경로를 반환한다.

    OOXML 규칙:
      - Target이 '/'로 시작하면 ZIP 루트 기준 절대 경로 (앞 '/' 제거)
      - 그 외는 base_path의 디렉터리 기준 상대 경로
    """
    # ECMA-376 §13.3: Target 속성값은 URI-encoded일 수 있음 (%20 등)
    # ZIP 엔트리 이름은 raw 경로이므로 decode 후 비교
    target = urllib.parse.unquote(target)
    if target.startswith("/"):
        # 절대 경로: 선두 슬래시만 제거
        return target.lstrip("/")
    base_dir = posixpath.dirname(base_path)
    joined = posixpath.join(base_dir, target)
    # posixpath.normpath로 ../ 처리
    normalized = posixpath.normpath(joined)
    # normpath가 '.'을 반환하는 경우 방어
    return normalized.lstrip("/")


# ---------------------------------------------------------------------------
# ZipContext
# ---------------------------------------------------------------------------

class ZipContext:
    """열린 ZIP 아카이브와 그 안의 특정 파트를 함께 나타낸다.

    Attributes:
        zf:   열린 zipfile.ZipFile 인스턴스 (공유됨, 닫지 않는다)
        path: ZIP 내부 파트 경로 (예: "ppt/charts/chart1.xml")
    """

    def __init__(self, zf: zipfile.ZipFile, path: str) -> None:
        self.zf = zf
        self.path = path
        self._rels_cache: Optional[dict[str, tuple[str, bool]]] = None

    # ------------------------------------------------------------------
    # 팩토리 메서드
    # ------------------------------------------------------------------

    @classmethod
    def from_path(cls, zip_path: str, entry: str = "") -> "ZipContext":
        """파일 시스템 경로에서 ZipContext를 생성한다.

        entry를 지정하지 않으면 ZIP 루트(_rels/.rels)를 기준 파트로 삼는다.
        """
        zf = zipfile.ZipFile(zip_path, "r")
        return cls(zf, entry)

    # ------------------------------------------------------------------
    # .rels 해석
    # ------------------------------------------------------------------

    def _get_rels(self) -> dict[str, tuple[str, bool]]:
        """현재 파트의 .rels 매핑을 lazy하게 로드해 반환한다.

        .rels 파일이 존재하지 않으면 빈 dict를 반환한다.
        """
        if self._rels_cache is not None:
            return self._rels_cache

        rp = _rels_path(self.path)
        try:
            data = self.zf.read(rp)
            self._rels_cache = _parse_rels_file(data)
        except KeyError:
            # .rels 파일 없음
            self._rels_cache = {}
        return self._rels_cache

    def resolve(self, rid: str) -> Optional["ZipContext"]:
        """r:id 값으로 참조 대상 파트의 ZipContext를 반환한다.

        관계가 없거나, 대상 파트가 외부 링크(External)이거나, ZIP 안에 없으면 None을 반환한다.
        """
        rels = self._get_rels()
        if rid not in rels:
            return None
        target, is_external = rels[rid]
        if is_external:
            return None # Cannot resolve external Targets into ZipContext
            
        resolved = _resolve_target(self.path, target)
        # ZIP 안에 실제로 존재하는지 확인
        try:
            self.zf.getinfo(resolved)
        except KeyError:
            return None
        return ZipContext(self.zf, resolved)

    def resolve_all(self) -> dict[str, "ZipContext"]:
        """현재 파트의 모든 관계를 {rid: ZipContext} dict로 반환한다.

        ZIP 안에 없는 외부 참조(http:// 등)는 제외한다.
        """
        result: dict[str, ZipContext] = {}
        for rid, _ in self._get_rels().items():
            ctx = self.resolve(rid)
            if ctx is not None:
                result[rid] = ctx
        return result

    # ------------------------------------------------------------------
    # 파트 내용 접근
    # ------------------------------------------------------------------

    def read_bytes(self) -> bytes:
        """파트를 bytes로 읽어 반환한다 (OLE 바이너리 등에 사용)."""
        return self.zf.read(self.path)

    def parse_xml(self) -> ET.Element:
        """파트를 XML로 파싱해 루트 Element를 반환한다."""
        data = self.zf.read(self.path)
        return ET.fromstring(data)

    # ------------------------------------------------------------------
    # 편의 메서드
    # ------------------------------------------------------------------

    def exists(self) -> bool:
        """파트가 ZIP 안에 실제로 존재하는지 확인한다."""
        try:
            self.zf.getinfo(self.path)
            return True
        except KeyError:
            return False

    def sibling(self, relative_path: str) -> "ZipContext":
        """현재 파트와 같은 디렉터리 기준의 상대 경로로 새 ZipContext를 반환한다."""
        new_path = _resolve_target(self.path, relative_path)
        return ZipContext(self.zf, new_path)

    def child(self, sub_path: str) -> "ZipContext":
        """현재 파트를 디렉터리로 보고 하위 경로의 ZipContext를 반환한다."""
        base = self.path.rstrip("/") + "/"
        return ZipContext(self.zf, base + sub_path.lstrip("/"))

    def close(self) -> None:
        """소유한 ZipFile을 닫는다.

        ZipContext를 공유할 때는 최상위 컨텍스트에서만 호출한다.
        """
        self.zf.close()

    def __enter__(self) -> "ZipContext":
        return self

    def __exit__(self, *_: object) -> None:
        self.close()

    def __repr__(self) -> str:
        return f"ZipContext({self.path!r})"


# ---------------------------------------------------------------------------
# 패키지 루트 편의 함수
# ---------------------------------------------------------------------------

def iter_parts_matching(zf: zipfile.ZipFile, pattern: str) -> list[ZipContext]:
    """ZIP 안에서 glob 패턴에 맞는 파트를 ZipContext 리스트로 반환한다.

    pattern 예: "ppt/charts/chart*.xml"
    단순 prefix+suffix 매칭만 지원한다 (fnmatch 사용).
    """
    import fnmatch
    return [
        ZipContext(zf, name)
        for name in zf.namelist()
        if fnmatch.fnmatch(name, pattern)
    ]

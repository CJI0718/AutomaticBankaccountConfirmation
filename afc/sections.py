"""cck 섹션·컬럼 스펙 로더 + PDF 텍스트의 섹션 분할/상태 판정.

configs/sections.yaml 은 cck 산출물에서 동결한 정적 스펙(런타임 cck 비의존).
각 조회서 PDF는 본문에 'N. ...' 형태의 최상위 섹션(은행 1~10)을 가지며,
cck는 이를 2-1/2-2 처럼 세분한다. 여기서는 PDF 최상위 섹션으로 분할하고
각 섹션의 (상태=해당없음/내용있음, 원문 라인)을 추출한다.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path

import yaml

SPEC_PATH = Path(__file__).resolve().parent.parent / "configs" / "sections.yaml"

ID_COLUMNS = ["조서번호", "조회대상 회사", "사업자번호", "금융기관명", "조회기준일"]

# 거래 없음/미회수 신호.
NO_DATA_PHRASES = ("해당 없음", "해당없음", "해당 거래 없음", "당사 거래회사 아님",
                   "거래하지 않", "담보보증내역없음", "담보보증 내역없음")

# 데이터가 아닌 보일러플레이트(섹션 본문에서 제거).
_BOILERPLATE_RE = re.compile(
    r"(확인자 소속 및 성명|열 람 용|본 조회서는 원본|동일함을 보증|"
    r"발급기관|^\d+/\d+$|^[가-힣]+\s?은행$|^[가-힣]+은행$)"
)
_SECTION_HDR_RE = re.compile(r"^\s*([0-9]+)\.\s+\S")


@dataclass(frozen=True)
class SectionSpec:
    id: str            # cck 섹션 id (예: '2-1')
    top: str           # PDF 최상위 번호 (예: '2')
    header: str
    columns: tuple[str, ...]   # 조서번호 포함 표시 컬럼 (helper/trailing 제거)
    data_columns: tuple[str, ...]  # ID 5개 이후의 데이터 컬럼


@dataclass(frozen=True)
class SectionContent:
    top: str
    has_data: bool
    raw_lines: tuple[str, ...]
    summary_value: str | None = None   # §2-1 총한도액처럼 단일값


def _trim_columns(cols: list[str]) -> list[str]:
    """첫 빈칸 이전까지만 — cck의 _통화/_금액 helper 컬럼 제거."""
    out: list[str] = []
    for c in cols:
        if c == "":
            break
        out.append(c)
    return out


@lru_cache(maxsize=2)
def load_section_spec(path: str | None = None) -> dict[str, list[SectionSpec]]:
    raw = yaml.safe_load(Path(path or SPEC_PATH).read_text(encoding="utf-8"))
    result: dict[str, list[SectionSpec]] = {}
    for category, body in (raw.get("categories") or {}).items():
        specs: list[SectionSpec] = []
        for s in body.get("sections", []):
            cols = _trim_columns(list(s["columns"]))
            top = str(s["id"]).split("-")[0]
            specs.append(SectionSpec(
                id=str(s["id"]), top=top, header=s["header"],
                columns=tuple(cols), data_columns=tuple(cols[len(ID_COLUMNS):]),
            ))
        result[category] = specs
    return result


def split_pdf_sections(full_text: str) -> dict[str, SectionContent]:
    """PDF 본문을 최상위 섹션(1., 2., ...)으로 분할."""
    lines = full_text.splitlines()
    # 섹션 시작 인덱스 수집
    marks: list[tuple[int, str]] = []
    for i, ln in enumerate(lines):
        m = _SECTION_HDR_RE.match(ln.strip())
        if m:
            marks.append((i, m.group(1)))
    sections: dict[str, SectionContent] = {}
    for idx, (start, top) in enumerate(marks):
        end = marks[idx + 1][0] if idx + 1 < len(marks) else len(lines)
        body = [ln.strip() for ln in lines[start + 1:end] if ln.strip()]
        body = [ln for ln in body if not _BOILERPLATE_RE.search(ln)]
        joined = "\n".join(body)
        has_data = not any(p in joined for p in NO_DATA_PHRASES) and bool(body)
        summary = None
        mt = re.search(r"총\s*한도액\s*[:：]\s*(.+)", joined)
        if mt:
            summary = mt.group(1).strip()
        # 최초 등장만 보존 (동일 top 중복 시)
        if top not in sections:
            sections[top] = SectionContent(
                top=top, has_data=has_data, raw_lines=tuple(body), summary_value=summary
            )
    return sections

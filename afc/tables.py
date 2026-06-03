"""표 인식 기반 추출 (pymupdf find_tables).

평면 get_text()는 컬럼 경계를 잃어 ① 페이지 넘김 시 머리말이 셀에 섞이고
② 인출제한 같은 칼럼을 다음 행과 구분하지 못한다. find_tables()는 행·열을
보존하므로 모든 섹션을 칸별로 깨끗이 뽑을 수 있다.

각 표를 섹션 스펙(configs/sections.yaml)의 데이터 컬럼 헤더와 대조해 섹션 id를
부여하고, 페이지에 걸쳐 분리된 표는 같은 섹션으로 이어 붙인다.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

import fitz

from afc.sections import load_section_spec


# 단일값(표 아님)·텍스트 섹션은 표 매칭에서 제외 — 다른 섹션 헤더와 토큰 충돌 방지.
_SKIP_SECTIONS = {"2-1"}
_NO_DATA_PHRASES = ("해당 없음", "해당없음", "해당 거래 없음", "거래 없음", "당사 거래회사 아님")


def _tokens(text: str | None) -> set[str]:
    """문자열을 단어 토큰 집합으로. '담보 보증 및 관련약정' → {담보, 보증, 관련약정}."""
    out: set[str] = set()
    for part in re.split(r"[_·\s]+", text or ""):
        p = part.strip()
        if len(p) >= 2:
            out.add(p)
    return out


def _clean_cell(c: str | None) -> str:
    if c is None:
        return ""
    s = c.replace("\n", " ").strip()
    # '열 람 용' 워터마크가 셀 앞에 번지는 경우 제거.
    s = re.sub(r"^[열람용]\s+", "", s)
    return s.strip()


@dataclass(frozen=True)
class SectionSignature:
    section_id: str
    tokens: frozenset[str]


def _signatures(category: str) -> list[SectionSignature]:
    sigs: list[SectionSignature] = []
    for spec in load_section_spec()[category]:
        if spec.id in _SKIP_SECTIONS:
            continue
        toks: set[str] = set()
        for col in spec.data_columns:
            toks |= _tokens(col)
        if toks:
            sigs.append(SectionSignature(spec.id, frozenset(toks)))
    return sigs


def _match_section(cells: list[str], sigs: list[SectionSignature], threshold: float = 0.5) -> str | None:
    """헤더 토큰과 섹션 시그니처의 일치도. 분모를 max(시그니처, 헤더)로 둬서
    단어 일부만 겹치는 짧은 시그니처(총한도액의 '한도액' 등)의 오매칭을 막는다."""
    htoks: set[str] = set()
    for c in cells:
        htoks |= _tokens(c)
    if not htoks:
        return None
    best, best_score = None, 0.0
    for sig in sigs:
        inter = len(sig.tokens & htoks)
        score = inter / max(len(sig.tokens), len(htoks))
        if score > best_score:
            best, best_score = sig.section_id, score
    return best if best_score >= threshold else None


@dataclass
class SectionTable:
    section_id: str
    columns: tuple[str, ...]
    rows: list[list[str]] = field(default_factory=list)


def extract_section_tables(pdf_bytes: bytes, category: str) -> dict[str, SectionTable]:
    """카테고리 PDF에서 섹션 id → SectionTable(헤더 제외 데이터 행) 추출."""
    if category not in load_section_spec():
        return {}
    sigs = _signatures(category)
    specs = {s.id: s for s in load_section_spec()[category]}
    # 섹션별 컬럼 라벨 토큰 — 2단 병합헤더의 둘째 줄 등 '머리말 행'을 걸러내는 데 쓴다.
    label_tokens = {
        sid: {tok for col in spec.data_columns for tok in _tokens(col)}
        for sid, spec in specs.items()
    }
    out: dict[str, SectionTable] = {}
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        for pi in range(doc.page_count):
            for t in doc[pi].find_tables().tables:
                if t.row_count < 1 or t.col_count < 3:
                    continue
                raw = t.extract()
                header = [_clean_cell(c) for c in raw[0]]
                sid = _match_section(header, sigs)
                if sid is None:
                    continue
                labels = label_tokens[sid]
                data: list[list[str]] = []
                for row in raw[1:]:
                    cells = [_clean_cell(c) for c in row]
                    filled = [c for c in cells if c]
                    if len(filled) < 2:  # 단일셀 = 빈 섹션의 소제목/잡음
                        continue
                    joined = " ".join(cells)
                    if any(p in joined for p in _NO_DATA_PHRASES):  # '해당 없음' 행
                        continue
                    # 모든 셀이 컬럼 라벨 토큰뿐이면 머리말 행(2단 병합헤더 둘째 줄 포함).
                    if all(_tokens(c) and _tokens(c) <= labels for c in filled):
                        continue
                    # 실데이터 신호(숫자=계좌/금액/날짜, 또는 긴 텍스트)가 없으면 잡음
                    # (페이지 푸터의 '신한'/'은행' 같은 기관명 파편) → 제외.
                    if not any(re.search(r"\d", c) or len(c) >= 8 for c in filled):
                        continue
                    data.append(cells)
                if not data:
                    continue
                st = out.get(sid)
                if st is None:
                    out[sid] = SectionTable(sid, specs[sid].data_columns, data)
                else:
                    st.rows.extend(data)
    finally:
        doc.close()
    return out

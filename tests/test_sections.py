"""섹션 스펙 로더 + PDF 섹션 분할 테스트."""
from __future__ import annotations

from afc.sections import ID_COLUMNS, load_section_spec, split_pdf_sections


def test_spec_has_all_categories_and_id_columns():
    spec = load_section_spec()
    assert set(spec) == {"BANK", "GUARANTEE", "INSURANCE", "INVESTMENT"}
    # BANK 14 sections incl. sub-sections
    assert [s.id for s in spec["BANK"]] == \
        ["1", "2-1", "2-2", "3", "4", "5", "6-1", "6-2", "7-1", "7-2", "7-3", "8", "9", "10"]
    # 모든 섹션은 조서번호 + 공통 ID 컬럼으로 시작
    for sections in spec.values():
        for s in sections:
            assert list(s.columns[:len(ID_COLUMNS)]) == ID_COLUMNS
            assert s.columns[0] == "조서번호"


def test_split_detects_empty_vs_data():
    text = (
        "1. 조회기준일 현재 조회대상회사의 당 은행에 대한 금융상품의 내용은 다음과 같습니다.\n"
        "보통예금\n123456789012\nKRW 0 (0)\n-\n20251220\n00000000\n"
        "2. 조회기준일 현재 조회대상회사에 대한 당 은행의 대출거래의 내용은 다음과 같습니다.\n"
        "총 한도액 : KRW6,100,000,000등\n해당 없음\n"
        "3. 조회기준일 현재 ... 지급보증\n해당 없음\n"
    )
    secs = split_pdf_sections(text)
    assert secs["1"].has_data is True
    assert secs["2"].summary_value == "KRW6,100,000,000등"
    assert secs["3"].has_data is False  # '해당 없음'


def test_top_level_numbering():
    text = "1. 가\n데이터1234567890\n2. 나\n해당 없음\n10. 다\n해당 없음\n"
    secs = split_pdf_sections(text)
    assert set(secs) == {"1", "2", "10"}

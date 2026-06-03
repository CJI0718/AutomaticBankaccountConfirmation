"""표(find_tables) 기반 도출 + 멀티파일 출력 헬퍼 테스트.

find_tables 자체는 실제 PDF가 필요하므로, 여기서는 표 셀 → 구조화 매핑
(deposits_from_table / loans_from_table)과 금액 칼럼 분리 로직을 합성 데이터로 검증한다.
"""
from __future__ import annotations

from afc.extract import (
    ConfirmationHeader,
    SourceRef,
    deposits_from_table,
    loans_from_table,
)
from afc.institutions import load_institution_config
from afc.output import _expand_columns, _fmt_cell, _is_money_col
from afc.tables import SectionTable, _match_section, _signatures

CONFIG = load_institution_config()


def _hdr():
    return ConfirmationHeader(
        company_name="삼화전기(주)", business_no="315-81-00390", institution_name="KEB하나은행",
        institution_category="BANK", confirmation_date="2025-12-31",
        source=SourceRef(source_file="x.pdf"),
    )


def test_deposits_from_table_captures_restriction_and_dates():
    table = SectionTable("1", (), rows=[
        ["퇴직연금신탁", "20491012879652", "(KRW)1,030,851,184(1,033,783,491)",
         "-", "20251231", "00000000", "확정급여형퇴직연금(DB)"],
        ["보통예금", "17589000402505", "(KRW)0 ()", "-", "20251220", "00000000", ""],
    ])
    deps = deposits_from_table(table, _hdr(), CONFIG)
    assert len(deps) == 2
    d = deps[0]
    assert d.balance.currency == "KRW" and d.balance.amount == 1030851184.0
    assert d.restrictions == "확정급여형퇴직연금(DB)"   # 인출제한 캡처
    assert d.interest_rate is None                      # '-' → None
    assert d.last_interest_payment_date == "20251231"
    assert d.maturity_date == "00000000"


def test_loans_from_table_handles_comma_split_money():
    table = SectionTable("2-2", (), rows=[
        # KEB하나: 약정한도액이 쉼표에서 줄바꿈 → find_tables가 공백으로 결합
        ["매입외환(DP/DA)", "(KRW)4,000,000, 000", "(KRW)0", "20180427", "20260427",
         "-", "00000000", "만기일시상환", ""],
    ])
    loans = loans_from_table(table, _hdr(), CONFIG)
    assert len(loans) == 1
    L = loans[0]
    assert L.limit.currency == "KRW" and L.limit.amount == 4000000000.0  # 공백 정리됨
    assert L.drawn.amount == 0.0
    assert L.interest_rate is None and L.repayment == "만기일시상환"


def test_money_column_detection_and_expansion():
    # 금액성 칼럼은 분리, 순위/율/종류 등은 분리 안 함
    for money in ["부보금액", "총 한도액", "연대보증 등의 한도", "선순위 설정금액",
                  "누적적립금", "해약환급금", "예수금", "신용설정 보증금", "평가액", "금    액"]:
        assert _is_money_col(money), money
    for plain in ["설정순위", "계좌번호", "금융상품의 종류", "출자좌수",
                  "지급보증수수료율", "연이자율", "지분율"]:
        assert not _is_money_col(plain), plain
    cols, plan = _expand_columns(("금융상품의 종류", "계좌번호", "금액"))
    assert cols == ["금융상품의 종류", "계좌번호", "금액_통화", "금액_금액"]
    assert plan[-1] == ("money", 2)


def test_fmt_cell_dates():
    assert _fmt_cell("20251220") == "2025-12-20"
    assert _fmt_cell("00000000") == ""
    assert _fmt_cell("보통예금") == "보통예금"


def test_section_signature_matching():
    sigs = _signatures("BANK")
    # 대출 표 헤더(2단 병합 둘째줄 포함)가 §2-2로 매칭되어야 한다.
    loan_header = ["대출 종류", "금액", "대출일", "최종만기일", "이자", "상환방법", "담보 보증 및 관련약정"]
    assert _match_section(loan_header, sigs) == "2-2"
    # 지급보증 헤더는 §2-1(총한도액)으로 오매칭되면 안 된다(제외 처리).
    g_header = ["내용", "한도액", "실행금액", "지급보증수수료율", "기간", "담보지급보증"]
    assert _match_section(g_header, sigs) != "2-1"

"""계정분류·대사 레이어 테스트."""
from __future__ import annotations

import pytest

from afc.reconciliation import (
    build_recon_rows,
    classify_account,
    classify_loan,
    load_account_mapping,
)

MAP = load_account_mapping()
FISCAL = "2025-12-31"  # 조회기준일 (ISO)


@pytest.mark.parametrize("product,maturity,expected", [
    # 요구불 → 현금및현금성자산
    ("보통예금", "00000000", "현금및현금성자산"),
    ("기업자유예금", "00000000", "현금및현금성자산"),
    ("외화당좌예금", "00010101", "현금및현금성자산"),
    ("신한 스크랩거래계좌(보통예금)", "00000000", "현금및현금성자산"),
    # 정기성 + 만기로 단기/장기 분기
    ("신한 S드림 정기예금(기업뱅킹전용)", "20260707", "단기금융상품"),   # ~7개월
    ("국민수퍼고정금리형-만기일시지급식", "20260320", "단기금융상품"),    # ~3개월
    ("정기예금", "20300101", "장기금융상품"),                          # >1년
    # 퇴직연금(DB) → 사외적립자산
    ("[퇴직연금]확정급여형(DB)", "20600506", "사외적립자산"),
    ("퇴직연금신탁", "00000000", "사외적립자산"),
    # 신탁/연금 → 검토필요 (사람 판단)
    ("마켓프리미엄신탁-법인용(6등급)", "20430705", "검토필요"),
])
def test_classify(product, maturity, expected):
    category, _basis = classify_account(product, maturity, FISCAL, MAP)
    assert category == expected


def test_time_product_without_maturity_is_review():
    cat, basis = classify_account("정기예금", "00000000", FISCAL, MAP)
    assert cat.startswith("검토필요")
    assert "만기" in basis


def test_boundary_12_months_is_short_term():
    # 2025-12-31 + 12개월 = 2026-12-31 → 단기 (≤12)
    cat, _ = classify_account("정기예금", "20261231", FISCAL, MAP)
    assert cat == "단기금융상품"


@pytest.mark.parametrize("maturity,expected", [
    ("20260701", "단기차입금"),      # ~7개월
    ("20300401", "장기차입금"),      # >1년
    ("00000000", "검토필요"),        # 만기 미상
])
def test_classify_loan(maturity, expected):
    cat, _basis = classify_loan(maturity, FISCAL, MAP)
    assert cat == expected


def test_recon_includes_loans_as_liability():
    """대출이 차입금(부채)으로 대사 행에 편입되는지."""
    from afc.extract import ConfirmationRecord, LoanRow
    from afc.schema import ConfirmationHeader, MoneyAmount, SourceRef

    h = ConfirmationHeader(
        company_name="가나㈜", business_no="123-45-67890", institution_name="신한은행",
        institution_category="BANK", confirmation_date="2025-12-31",
        source=SourceRef(source_file="x.pdf"),
    )
    loan = LoanRow(
        loan_type="무역금융", limit=MoneyAmount("KRW 61억", "KRW", 6_100_000_000.0),
        drawn=MoneyAmount("KRW 0", "KRW", 0.0), loan_date="20190401",
        maturity_date="20260702", interest_rate=None, last_interest_date="00000000",
        repayment="일시상환", collateral="103494999",
    )
    rec = ConfirmationRecord(header=h, status="complete", source_file="x.pdf", pages=1,
                             deposits=(), loans=(loan,))
    rows = build_recon_rows([rec], MAP)
    assert len(rows) == 1
    assert rows[0].side == "부채" and rows[0].category == "단기차입금"

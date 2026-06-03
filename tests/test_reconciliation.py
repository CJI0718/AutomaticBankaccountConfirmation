"""계정분류·대사 레이어 테스트."""
from __future__ import annotations

import pytest

from afc.reconciliation import classify_account, load_account_mapping

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

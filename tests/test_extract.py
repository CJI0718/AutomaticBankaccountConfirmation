"""Regression tests for the deterministic extractor.

Fixtures are synthetic — they mirror the real PDF text-layer structure (wrapped
amounts, multi-line product names, foreign-currency rows) without embedding any
client data.
"""
from __future__ import annotations

import pytest

from afc.extract import (
    STATUS_COMPLETE,
    STATUS_EMPTY_REVIEW,
    STATUS_NON_DEALING,
    ConfirmationHeader,
    SourceRef,
    extract_record,
    format_date8,
    normalize_date,
    parse_filename,
    parse_loans,
    parse_money,
)
from afc.institutions import load_institution_config
from afc.sections import split_pdf_sections

CONFIG = load_institution_config()


def _bank_header(name="신한은행"):
    return ConfirmationHeader(
        company_name="삼화전기(주)", business_no="315-81-00390", institution_name=name,
        institution_category="BANK", confirmation_date="2025-12-31",
        source=SourceRef(source_file="x.pdf"),
    )


# ── filename / date ─────────────────────────────────────────────────────────

def test_parse_filename():
    name = "전자_[]_가나㈜_[123-45-67890]_국민은행_[2025년12월31일].pdf"
    parsed = parse_filename(name)
    assert parsed["company_name"] == "가나㈜"
    assert parsed["business_no"] == "123-45-67890"
    assert parsed["institution_name"] == "국민은행"
    assert parsed["confirmation_date"] == "2025년12월31일"


def test_parse_filename_no_match():
    assert parse_filename("random.pdf")["business_no"] is None


@pytest.mark.parametrize("raw,expected", [
    ("2025년12월31일", "2025-12-31"),
    ("2025-12-31", "2025-12-31"),
    ("2025년 1월 3일", "2025-01-03"),
    (None, None),
])
def test_normalize_date(raw, expected):
    assert normalize_date(raw) == expected


# ── classification ──────────────────────────────────────────────────────────

@pytest.mark.parametrize("name,category", [
    ("국민은행", "BANK"),
    ("KEB하나은행", "BANK"),
    ("서울보증보험", "GUARANTEE"),       # 보증보험 beats 보험
    ("현대해상화재보험", "INSURANCE"),
    ("에이아이지 손해보험", "INSURANCE"),
    ("현대차증권", "INVESTMENT"),
    ("듣보잡금고", "UNKNOWN"),
])
def test_classify(name, category):
    assert CONFIG.classify(name) == category


# ── money ───────────────────────────────────────────────────────────────────

@pytest.mark.parametrize("raw,currency,amount", [
    ("KRW 0 (0)", "KRW", 0.0),
    ("114,239.00 (0.00)", None, 114239.0),
    ("(KRW)1,030,851,184(1,033,783,491)", "KRW", 1030851184.0),
    ("EUR             0.00 (EUR0.00)", "EUR", 0.0),
    ("USD 421,013.26 (0)", "USD", 421013.26),
    ("KRW 2,414,000,000 (0)", "KRW", 2414000000.0),
])
def test_parse_money(raw, currency, amount):
    m = parse_money(raw, CONFIG)
    assert m.currency == currency
    assert m.amount == amount


# ── deposit parsing ─────────────────────────────────────────────────────────

DEPOSIT_HEADER = (
    "삼화전기(주)\n사업자등록번호 : 123-45-67890\n조회기준일 : 2025 년\n월\n일 현재\n12\n31\n"
    "1. 조회기준일 현재 조회대상회사의 당 은행에 대한 금융상품의 내용은 다음과 같습니다.\n"
    "금융상품의 종류\n계좌번호\n금액\n연이자율\n최종이자지급일\n만기일\n인출제한 등\n"
)
DEPOSIT_BODY = (
    "보통예금\n100023086052\nKRW 0 (0)\n-\n20251220\n00000000\n"
    "신한 스크랩거래계좌(\n보통예금)\n100031799552\nKRW 0 (0)\n-\n20251220\n00000000\n"  # wrapped product
    "국민수퍼고정금리형-만\n기일시지급식\n7124151772111001\n1,000,000,000.00\n(0.00)\n2.5000\n00000000\n20260320\n"  # wrapped amount
    "외화보통예금\n21586807100024\nEUR             0.00 (EUR\n0.00)\n0.0000\n00000000\n00000000\n"  # wrapped fx amount
    "확인자 소속 및 성명 :\n홍길동\n"
)
FILENAME = "전자_[]_가나㈜_[123-45-67890]_신한은행_[2025년12월31일].pdf"


def _record(body_text):
    return extract_record(FILENAME, DEPOSIT_HEADER + body_text, 1, CONFIG)


def test_deposit_count_and_fields():
    rec = _record(DEPOSIT_BODY)
    assert rec.status == STATUS_COMPLETE
    assert len(rec.deposits) == 4

    by_acct = {d.account_no: d for d in rec.deposits}
    assert by_acct["100031799552"].product_type == "신한 스크랩거래계좌(보통예금)"

    fixed = by_acct["7124151772111001"]
    assert fixed.product_type == "국민수퍼고정금리형-만기일시지급식"
    assert fixed.balance.amount == 1_000_000_000.0
    assert fixed.interest_rate == "2.5000"
    assert fixed.maturity_date == "20260320"

    fx = by_acct["21586807100024"]
    assert fx.balance.currency == "EUR"
    assert fx.balance.amount == 0.0
    assert fx.interest_rate == "0.0000"  # only literal "-" becomes None


@pytest.mark.parametrize("raw,expected", [
    ("20251220", "2025-12-20"), ("00000000", ""), ("00010101", ""),
    ("", ""), (None, ""), ("일시상환", "일시상환"),
])
def test_format_date8(raw, expected):
    assert format_date8(raw) == expected


def test_deposit_optional_rate_does_not_shift_date():
    """기업은행처럼 연이자율 줄이 없는 예금: 날짜가 연이자율 칸으로 밀리면 안 된다."""
    text = (
        "1. 조회기준일 현재 조회대상회사의 당 은행에 대한 금융상품의 내용은 다음과 같습니다.\n"
        "금융상품의 종류\n계좌번호\n금액\n연이자율\n최종이자지급일\n만기일\n인출제한 등\n"
        "기업자유\n05902661004017\nKRW             0.00 ()\n20251221\n00000000\n"
        "확인자 소속 및 성명 :\n"
    )
    rec = extract_record("전자_[]_삼화전기(주)_[315-81-00390]_기업은행_[2025년12월31일].pdf",
                         text, 1, CONFIG)
    assert len(rec.deposits) == 1
    d = rec.deposits[0]
    assert d.interest_rate is None                 # 연이자율 칸은 비어야 함 (날짜 아님)
    assert d.last_interest_payment_date == "20251221"
    assert d.maturity_date == "00000000"


def test_parse_loans_columnizes_money_dates():
    """대출 §2-2: 통화/금액 분리, 쉼표 줄분리·통화 별행 복원, 연이율 선택."""
    text = (
        "2. 조회기준일 현재 조회대상회사에 대한 당 은행의 대출거래의 내용은 다음과 같습니다.\n"
        "총 한도액 : (KRW)4,000,000,000\n"
        "대출 종류\n금액\n대출일\n최종만기일\n이자\n상환방법\n담보 보증 및 관련약정\n"
        "약정한도액\n대출금액\n연이율\n최종이자지급일\n"
        # 쉼표 줄분리 + 연이율 '-' (KEB하나 형태)
        "매입외환(DP/DA)\n(KRW)4,000,000,\n000\n(KRW)0\n20180427\n20260427\n-\n00000000\n만기일시상환\n"
        # 통화 별행 + 담보 숫자 (신한 형태, 연이율 생략)
        "무역금융\nKRW\n6,100,000,000\nKRW 0\n20190401\n20260702\n00000000\n일시상환\n103494999\n"
    )
    content = split_pdf_sections(text)["2"]
    loans = parse_loans(content, _bank_header(), CONFIG)
    assert len(loans) == 2
    a, b = loans
    assert (a.loan_type, a.limit.currency, a.limit.amount) == ("매입외환(DP/DA)", "KRW", 4000000000.0)
    assert a.interest_rate is None and a.repayment == "만기일시상환"
    assert (b.loan_type, b.limit.amount, b.collateral) == ("무역금융", 6100000000.0, "103494999")
    assert b.drawn.currency == "KRW" and b.drawn.amount == 0.0


def test_interest_rate_dash_becomes_none():
    rec = _record(DEPOSIT_BODY)
    assert {d.account_no: d for d in rec.deposits}["100023086052"].interest_rate is None


# ── status resolution ────────────────────────────────────────────────────────

def test_status_non_dealing_explicit():
    rec = extract_record(
        "전자_[]_가나㈜_[123-45-67890]_현대차증권_[2025년12월31일].pdf",
        "사업자등록번호 : 123-45-67890\n금융상품의 종류\n당사 거래회사 아님\n확인자 소속 및 성명 :",
        1, CONFIG,
    )
    assert rec.status == STATUS_NON_DEALING


def test_status_empty_review_when_no_data():
    rec = extract_record(
        "전자_[]_가나㈜_[123-45-67890]_서울보증보험_[2025년12월31일].pdf",
        "사업자등록번호 : 123-45-67890\n보증(보험)의 종류\n증권번호\n해당 거래 없음\n확인자 소속 및 성명 :",
        1, CONFIG,
    )
    assert rec.status == STATUS_EMPTY_REVIEW


def test_bizno_mismatch_finding():
    rec = extract_record(
        "전자_[]_가나㈜_[123-45-67890]_국민은행_[2025년12월31일].pdf",
        "사업자등록번호 : 999-99-99999\nKRW 100 (0)\n12345678901",
        1, CONFIG,
    )
    assert any(f.code == "BIZNO_MISMATCH" for f in rec.findings)

"""감사 계정분류 + 대사(對査)용 정규화 레이어.

조회서 raw(예금 행)를 감사조서 4200 '조회서대사' 시트 축으로 재구성한다:
  금융기관 · 계정분류 · 구분(자산/부채) · 계좌번호 · 금융상품종류 · 통화 ·
  환산전 금액 · 이자율 · 만기 · 사용제한
+ 통화별 소계 / 계정분류별 roll-up을 엑셀 단계에서 얹는다.

계정분류는 configs/account_mapping.yaml 규칙으로 결정론적으로 부여하되,
판단이 필요한 항목(퇴직연금·신탁·만기미상)은 '검토필요'로 분리한다.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from functools import lru_cache
from pathlib import Path

import yaml

from afc.extract import ConfirmationRecord
from afc.schema import BankDepositRow

from afc.paths import CONFIG_DIR

MAPPING_PATH = CONFIG_DIR / "account_mapping.yaml"


@dataclass(frozen=True)
class AccountMapping:
    short_term_months: int
    demand_keywords: tuple[str, ...]
    time_keywords: tuple[str, ...]
    plan_asset_keywords: tuple[str, ...]
    review_keywords: tuple[str, ...]
    label_short: str
    label_long: str
    label_plan_asset: str
    label_review: str
    label_unclassified: str


@lru_cache(maxsize=4)
def load_account_mapping(path: str | None = None) -> AccountMapping:
    p = Path(path) if path else MAPPING_PATH
    raw = yaml.safe_load(p.read_text(encoding="utf-8"))
    cats = raw.get("categories") or {}
    demand = tuple((cats.get("현금및현금성자산") or {}).get("demand_keywords") or ())
    time_kw = tuple((cats.get("금융상품") or {}).get("time_keywords") or ())
    labels = raw.get("labels") or {}
    return AccountMapping(
        short_term_months=int(raw.get("short_term_months", 12)),
        demand_keywords=demand,
        time_keywords=time_kw,
        plan_asset_keywords=tuple(raw.get("plan_asset_keywords") or ()),
        review_keywords=tuple(raw.get("review_keywords") or ()),
        label_short=labels.get("short_term", "단기금융상품"),
        label_long=labels.get("long_term", "장기금융상품"),
        label_plan_asset=labels.get("plan_asset", "사외적립자산"),
        label_review=labels.get("review", "검토필요"),
        label_unclassified=labels.get("unclassified", "검토필요(미분류)"),
    )


def _parse_flex_date(value: str | None) -> date | None:
    """Parse 'YYYYMMDD' (만기) or 'YYYY-MM-DD' (조회기준일). None on sentinel/blank."""
    if not value:
        return None
    digits = value.replace("-", "").replace(".", "").strip()
    if not digits.isdigit() or len(digits) != 8 or digits in ("00000000", "00010101"):
        return None
    try:
        return date(int(digits[:4]), int(digits[4:6]), int(digits[6:]))
    except ValueError:
        return None


def _months_between(start: date, end: date) -> int:
    return (end.year - start.year) * 12 + (end.month - start.month)


def classify_account(
    product: str, maturity: str | None, fiscal_date: str | None, mapping: AccountMapping
) -> tuple[str, str]:
    """(계정분류 라벨, 분류근거) 반환. 결정론적."""
    text = product or ""
    if any(kw in text for kw in mapping.plan_asset_keywords):
        return mapping.label_plan_asset, "퇴직연금 — 사외적립자산(DB 가정, DB/DC 구분 검토)"
    if any(kw in text for kw in mapping.review_keywords):
        return mapping.label_review, "신탁/연금 등 — 회계사 판단 필요"
    if any(kw in text for kw in mapping.demand_keywords):
        return "현금및현금성자산", "요구불·수시입출 상품"
    if any(kw in text for kw in mapping.time_keywords):
        mat = _parse_flex_date(maturity)
        ref = _parse_flex_date(fiscal_date)
        if mat is None:
            return mapping.label_unclassified, "정기성이나 만기 미상"
        if ref is None:
            return mapping.label_review, "조회기준일 미상 — 만기 판정 불가"
        months = _months_between(ref, mat)
        if months <= mapping.short_term_months:
            return mapping.label_short, f"잔여만기 {months}개월 (≤{mapping.short_term_months})"
        return mapping.label_long, f"잔여만기 {months}개월 (>{mapping.short_term_months})"
    return mapping.label_unclassified, "분류 키워드 미일치"


@dataclass(frozen=True)
class ReconRow:
    category: str          # 감사 계정분류
    side: str              # 자산 / 부채
    institution: str
    company: str
    business_no: str
    account_no: str
    product: str
    currency: str | None
    amount: float | None   # 환산전 (외화 또는 원화)
    interest_rate: str | None
    maturity: str | None
    restrictions: str | None
    basis: str             # 분류 근거 (검토용)

    # 계정분류 정렬 우선순위 — leadsheet roll-up 순서.
    _ORDER = ("현금및현금성자산", "단기금융상품", "장기금융상품", "사외적립자산")

    @property
    def sort_key(self) -> tuple:
        try:
            cat_rank = self._ORDER.index(self.category)
        except ValueError:
            cat_rank = len(self._ORDER) + 1
        return (cat_rank, self.category, self.currency or "ZZZ", self.institution, self.account_no)


def _row_from_deposit(d: BankDepositRow, mapping: AccountMapping) -> ReconRow:
    category, basis = classify_account(
        d.product_type, d.maturity_date, d.header.confirmation_date, mapping
    )
    side = "부채" if category.startswith("단기차입") or category.startswith("장기차입") else "자산"
    return ReconRow(
        category=category,
        side=side,
        institution=d.header.institution_name,
        company=d.header.company_name,
        business_no=d.header.business_no,
        account_no=d.account_no,
        product=d.product_type,
        currency=d.balance.currency,
        amount=d.balance.amount,
        interest_rate=d.interest_rate,
        maturity=d.maturity_date,
        restrictions=d.restrictions,
        basis=basis,
    )


def build_recon_rows(records: list[ConfirmationRecord], mapping: AccountMapping) -> list[ReconRow]:
    rows = [_row_from_deposit(d, mapping) for rec in records for d in rec.deposits]
    rows.sort(key=lambda r: r.sort_key)
    return rows

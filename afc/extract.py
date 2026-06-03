"""Deterministic extraction of canonical records from confirmation PDFs.

No LLM calls. Every field is derived from the PDF text layer (pymupdf) and the
filename, using anchors that are stable across institutions:

* Header        — filename (validated against the in-document business number).
* Status        — explicit non-dealing phrase, else COMPLETE; all-empty forms are
                  flagged for human review (see institutions.is_explicit_non_dealing).
* BANK deposits — section 1 rows, anchored on the account-number line whose next
                  line carries the parenthesised amount. This anchor is immune to
                  product names that themselves contain parentheses, e.g.
                  "[퇴직연금]확정급여형(DB)" or "마켓프리미엄신탁-법인용(6등급)".

Other sections (loans, guarantees, derivatives, insurance, investment …) are
inventoried but not yet field-parsed; see `SECTION_COVERAGE` and the README TODO.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from afc.institutions import InstitutionConfig
from afc.sections import SectionContent, split_pdf_sections
from afc.schema import (
    BankDepositRow,
    ConfirmationHeader,
    InstitutionCategory,
    MoneyAmount,
    SourceRef,
    ValidationFinding,
)

# Filename: 전자_[...]_<company>_[<biz_no>]_<institution>_[<date>].pdf
FILENAME_RE = re.compile(
    r"^전자_\[.*?\]_(?P<company>.*?)_\[(?P<business_no>\d{3}-\d{2}-\d{5})\]_"
    r"(?P<institution>.*?)_\[(?P<date>.*?)\]\.pdf$"
)
BIZ_NO_RE = re.compile(r"(\d{3}-\d{2}-\d{5})")
DEPOSIT_HEADER_END = "인출제한 등"
SECTION1_SIGNATURE = "금융상품의 내용은 다음과 같습니다"
DATE8_RE = re.compile(r"^\d{8}$")

# Lines that are page furniture, not data. Dropped before row parsing.
_FURNITURE_SUBSTRINGS = (
    "담당공인회계사", "귀중", "회계법인", "조회대상 회사", "사업자등록번호",
    "조회기준일", "확인자 소속", "열 람 용", "SCRIPT ERROR", "일 현재",
    "다음과 같습니다", "별첨", "붙임",
)
_PAGE_NO_RE = re.compile(r"^\d+\s*/\s*\d+$")

# Status keys (stable identifiers; display strings live in institutions.yaml).
STATUS_COMPLETE = "complete"
STATUS_NON_DEALING = "non_dealing"
STATUS_EMPTY_REVIEW = "empty_review"


@dataclass(frozen=True)
class ConfirmationRecord:
    """One returned confirmation PDF, fully reduced to canonical form."""

    header: ConfirmationHeader
    status: str  # one of STATUS_* keys
    source_file: str
    pages: int
    deposits: tuple[BankDepositRow, ...] = ()
    loans: tuple[LoanRow, ...] = ()
    findings: tuple[ValidationFinding, ...] = ()
    sections: dict[str, SectionContent] = field(default_factory=dict)  # PDF 최상위 §N → 내용
    section_tables: dict = field(default_factory=dict)  # find_tables: 섹션 id → SectionTable

    @property
    def has_any_data(self) -> bool:
        return bool(self.deposits)


# ── filename / header ──────────────────────────────────────────────────────

def parse_filename(name: str) -> dict[str, str | None]:
    match = FILENAME_RE.match(name)
    if not match:
        return {k: None for k in ("company_name", "business_no", "institution_name", "confirmation_date")}
    return {
        "company_name": match.group("company"),
        "business_no": match.group("business_no"),
        "institution_name": match.group("institution"),
        "confirmation_date": match.group("date"),
    }


def normalize_date(raw: str | None) -> str | None:
    """'2025년12월31일' / '2025-12-31' → '2025-12-31'. Returns input on no match."""
    if not raw:
        return raw
    m = re.search(r"(\d{4})\D*(\d{1,2})\D*(\d{1,2})", raw)
    if not m:
        return raw
    y, mo, d = m.groups()
    return f"{int(y):04d}-{int(mo):02d}-{int(d):02d}"


_NULL_DATES = {"00000000", "00010101", "10101", ""}


def format_date8(raw: str | None) -> str:
    """'20251220' → '2025-12-20'; sentinel/blank (00000000 등) → ''. Non-date → raw."""
    if not raw:
        return ""
    s = raw.strip()
    if s in _NULL_DATES:
        return ""
    if s.isdigit() and len(s) == 8:
        return f"{s[:4]}-{s[4:6]}-{s[6:]}"
    return s


def _looks_like_rate(line: str) -> bool:
    """연이자율 셀 판별: '-' 또는 소수(2.47, 0.0000). 8자리 날짜는 제외."""
    s = line.strip()
    if s == "-":
        return True
    return "." in s and bool(re.fullmatch(r"-?\d+(?:\.\d+)?%?", s))


def _is_date_field(line: str) -> bool:
    return line.isdigit() and len(line) == 8


# ── money ──────────────────────────────────────────────────────────────────

def parse_money(raw: str | None, config: InstitutionConfig) -> MoneyAmount:
    """Parse a confirmation amount cell into (currency, amount).

    Handles the layouts seen across institutions, e.g.::

        KRW 0 (0)                       -> (KRW, 0.0)
        114,239.00 (0.00)               -> (None, 114239.0)
        (KRW)1,030,851,184(1,033,783,491) -> (KRW, 1030851184.0)
        EUR             0.00 (EUR0.00)  -> (EUR, 0.0)
        USD 421,013.26 (0)              -> (USD, 421013.26)

    The leading number is the confirmed balance; the parenthesised trailing
    number (당좌대출 등) is intentionally ignored, matching the cck 금액_금액 column.
    """
    if raw is None:
        return MoneyAmount(raw=None, currency=None, amount=None)
    text = raw.strip()
    currency = None
    for cur in config.known_currencies:
        if cur in text:
            currency = cur
            break
    # First numeric token *outside* the trailing parenthetical.
    head = text.split("(", 1)[0] if "(" in text and not text.startswith("(") else text
    if text.startswith("("):  # "(KRW)1,030,..." — strip the leading (CUR)
        head = re.sub(r"^\([A-Za-z]{0,3}\)", "", text)
        head = head.split("(", 1)[0]
    num_match = re.search(r"-?[\d,]+(?:\.\d+)?", head)
    amount = float(num_match.group().replace(",", "")) if num_match else None
    return MoneyAmount(raw=text, currency=currency, amount=amount)


# ── text helpers ───────────────────────────────────────────────────────────

def _is_furniture(line: str) -> bool:
    if _PAGE_NO_RE.match(line):
        return True
    return any(sub in line for sub in _FURNITURE_SUBSTRINGS)


def _clean_lines(text: str) -> list[str]:
    out: list[str] = []
    for raw in text.splitlines():
        line = raw.replace("　", " ").strip()
        if not line or _is_furniture(line):
            continue
        out.append(line)
    return out


def _is_account(line: str) -> bool:
    return line.isdigit() and 8 <= len(line) <= 20 and not DATE8_RE.match(line)


_DATA_TOKEN_RE = re.compile(r"\b(?:KRW|USD|EUR|JPY|WON|CNY|HKD|GBP|CHF)\b|\d{10,}")


def _has_data_tokens(text: str) -> bool:
    """Heuristic: does the form carry any transaction data (currency code or a
    long account/policy number)? Used to flag all-empty non-bank forms for human
    review instead of silently marking them 완료. The business number contains
    dashes so it never matches the 10+ digit run."""
    return bool(_DATA_TOKEN_RE.search(text))


# ── BANK section-1 deposits ────────────────────────────────────────────────

def _deposit_body(full_text: str) -> str | None:
    """Slice from the deposit column header to the end of section 1."""
    if SECTION1_SIGNATURE not in full_text or DEPOSIT_HEADER_END not in full_text:
        return None
    start = full_text.index(DEPOSIT_HEADER_END) + len(DEPOSIT_HEADER_END)
    rest = full_text[start:]
    # Section 1 ends at the confirmer block or the start of section 2. Use only
    # unambiguous phrases — a bare "2." would false-match interest rates (2.5000).
    enders = []
    for marker in ("확인자 소속", "당 은행의 대출거래", "총 한도액"):
        idx = rest.find(marker)
        if idx != -1:
            enders.append(idx)
    return rest[: min(enders)] if enders else rest


def parse_bank_deposits(
    full_text: str, header: ConfirmationHeader, config: InstitutionConfig
) -> list[BankDepositRow]:
    body = _deposit_body(full_text)
    if body is None:
        return []
    lines = _clean_lines(body)
    rows: list[BankDepositRow] = []
    prev_end = -1  # index of the last consumed line (maturity of previous row)
    i = 0
    n = len(lines)
    while i < n:
        amount, after = _read_amount(lines, i + 1, n) if _is_account(lines[i]) else (None, None)
        if amount is not None:
            account = lines[i]
            # Trailing meta fields are [연이자율?, 최종이자지급일, 만기일] where 연이자율 is
            # optional (e.g. 기업은행 omits it). Consume by *type*, not fixed offset, so a
            # missing rate never shifts a date into the rate column.
            k = after
            rate = None
            if k < n and _looks_like_rate(lines[k]):
                rate = lines[k]
                k += 1
            last_int = lines[k] if k < n and _is_date_field(lines[k]) else None
            if last_int is not None:
                k += 1
            maturity = lines[k] if k < n and _is_date_field(lines[k]) else None
            if maturity is not None:
                k += 1
            product = "".join(lines[prev_end + 1 : i]).strip()
            rows.append(
                BankDepositRow(
                    header=header,
                    product_type=product,
                    account_no=account,
                    balance=parse_money(amount, config),
                    interest_rate=None if rate in (None, "-") else rate,
                    last_interest_payment_date=last_int,
                    maturity_date=maturity,
                    restrictions=None,
                    source=SourceRef(source_file=header.source.source_file, section_id="deposits"),
                )
            )
            prev_end = k - 1
            i = k
        else:
            i += 1
    return rows


def _read_amount(lines: list[str], start: int, n: int) -> tuple[str | None, int]:
    """Join lines from `start` into one amount string until ')' closes it.

    The amount cell can wrap across up to ~3 physical lines, e.g.
    ``1,000,000,000.00`` + ``(0.00)`` or ``EUR 0.00 (EUR`` + ``0.00)``.
    Returns (joined_amount, index_after_amount) or (None, start) if the lines
    at `start` are not a valid parenthesised amount.
    """
    parts: list[str] = []
    j = start
    while j < n and j - start < 3:
        parts.append(lines[j])
        j += 1
        if ")" in lines[j - 1]:
            break
    joined = "".join(parts)
    if "(" in joined and ")" in joined and any(c.isdigit() for c in joined):
        return joined, j
    return None, start


# ── BANK section 2-2 loans ──────────────────────────────────────────────────

_LOAN_LABELS = {
    "대출 종류", "대출종류", "금액", "대출일", "최종만기일", "이자", "상환방법",
    "담보 보증 및 관련약정", "약정한도액", "대출금액", "연이율", "최종이자지급일",
}
_CONFIRMER_RE = re.compile(r"센터|지점|차장|과장|대리|부장|<\d|영업부|영업점|귀중")
_CUR_RE = re.compile(r"^\(?([A-Z]{3})\)?\s*")


@dataclass(frozen=True)
class LoanRow:
    loan_type: str
    limit: MoneyAmount          # 약정한도액
    drawn: MoneyAmount          # 대출금액
    loan_date: str | None
    maturity_date: str | None
    interest_rate: str | None
    last_interest_date: str | None
    repayment: str | None
    collateral: str | None


def _read_loan_money(tokens: list[str], i: int, config: InstitutionConfig) -> tuple[MoneyAmount, int]:
    """Read one money cell that may span lines: 'KRW'+'4,000,000,000',
    '(KRW)4,000,000,'+'000', 'USD 6,000,000', '0.00', '(KRW)0'."""
    n = len(tokens)
    s = tokens[i]
    i += 1
    currency = None
    m = _CUR_RE.match(s)
    if m and m.group(1) in config.known_currencies:
        currency = m.group(1)
        s = s[m.end():]
    num = s.strip()
    # 통화만 있던 줄이거나 숫자가 쉼표에서 잘린 경우 다음 줄을 이어붙인다.
    while (num == "" or num.endswith(",")) and i < n and not _is_date_field(tokens[i]):
        nxt = tokens[i].strip()
        mm = _CUR_RE.match(nxt)  # 다음 줄도 통화로 시작하면 별도 셀 → 중단
        if mm and mm.group(1) in config.known_currencies:
            break
        num += nxt
        i += 1
    return parse_money(f"{currency or ''} {num}".strip(), config), i


def _loan_lines(content, header: ConfirmationHeader) -> list[str]:
    """§2 본문에서 컬럼헤더/총한도액/기관명/확인자 등을 제거한 데이터 라인만."""
    out: list[str] = []
    for ln in content.raw_lines:
        s = ln.strip()
        if (not s or s in _LOAN_LABELS or s.startswith("총 한도액")
                or s == header.institution_name or s.endswith(("은행", "증권", "보험"))
                or _CONFIRMER_RE.search(s)):
            continue
        out.append(s)
    return out


def _is_currency_money(s: str, config: InstitutionConfig) -> bool:
    m = _CUR_RE.match(s)
    return bool(m and m.group(1) in config.known_currencies)


def _is_money_start(s: str, config: InstitutionConfig) -> bool:
    if _is_currency_money(s, config):
        return True
    return bool(re.match(r"^[\d,]+(\.\d+)?$", s))  # 통화 없는 숫자 금액


def parse_loans(content, header: ConfirmationHeader, config: InstitutionConfig) -> list[LoanRow]:
    """§2-2 대출거래 행을 통화·날짜 앵커로 재조립 (은행별 포맷 편차 흡수)."""
    if content is None or not content.has_data:
        return []
    tokens = _loan_lines(content, header)
    rows: list[LoanRow] = []
    i, n = 0, len(tokens)
    while i < n:
        # 1) 대출종류 — money/date 전까지의 텍스트(줄바꿈 결합)
        start = i
        while i < n and not _is_money_start(tokens[i], config) and not _is_date_field(tokens[i]):
            i += 1
        if i == start:  # 종류 없이 시작 → 비정상, 한 줄 건너뜀
            i += 1
            continue
        loan_type = "".join(tokens[start:i])
        if i >= n or _is_date_field(tokens[i]):
            break
        # 2) 약정한도액, 대출금액
        limit, i = _read_loan_money(tokens, i, config)
        if i >= n or not _is_money_start(tokens[i], config):
            break
        drawn, i = _read_loan_money(tokens, i, config)
        # 3) 대출일, 최종만기일
        loan_date = tokens[i] if i < n and _is_date_field(tokens[i]) else None
        i += 1 if loan_date else 0
        maturity = tokens[i] if i < n and _is_date_field(tokens[i]) else None
        i += 1 if maturity else 0
        # 4) 연이율(선택) → 최종이자지급일
        rate = None
        if i < n and _looks_like_rate(tokens[i]):
            rate = tokens[i]
            i += 1
        last_int = tokens[i] if i < n and _is_date_field(tokens[i]) else None
        i += 1 if last_int else 0
        # 5) 상환방법(…상환), 담보보증(1줄)
        repayment = None
        if i < n and "상환" in tokens[i] and not _is_money_start(tokens[i], config):
            repayment = tokens[i]
            i += 1
        # 담보보증: 1줄. 계좌/약정번호(순수 숫자)나 알려진 문구만 취한다. 임의 텍스트는
        # 다음 행의 '대출종류'일 수 있으므로 침범하지 않는다(담보 공란 처리).
        collateral = None
        if i < n:
            tok = tokens[i]
            if (tok.replace(",", "").isdigit()
                    or re.search(r"항목참조|참조|해당|없음|담보|보증|약정", tok)):
                collateral = tok
                i += 1
        rows.append(LoanRow(
            loan_type=loan_type, limit=limit, drawn=drawn,
            loan_date=loan_date, maturity_date=maturity,
            interest_rate=None if rate in (None, "-") else rate,
            last_interest_date=last_int, repayment=repayment, collateral=collateral,
        ))
    return rows


# ── table-based derivation (find_tables) ────────────────────────────────────

def _money_cell(s: str | None) -> str:
    """금액 셀의 내부 공백 제거 — '(KRW)4,000,000, 000' → '(KRW)4,000,000,000'."""
    return (s or "").replace(" ", "")


def _date_cell(s: str | None) -> str | None:
    t = (s or "").strip()
    return t if t.isdigit() and len(t) == 8 else (t or None)


def deposits_from_table(table, header: ConfirmationHeader, config: InstitutionConfig) -> list[BankDepositRow]:
    """§1 예금 표(행·열 보존) → BankDepositRow. 인출제한 칼럼까지 포함."""
    rows: list[BankDepositRow] = []
    for cells in table.rows:
        c = (list(cells) + [""] * 7)[:7]
        product, account, amount, rate, last_int, maturity, restr = c
        account = account.replace(" ", "")
        if not account.isdigit():
            continue
        rows.append(BankDepositRow(
            header=header, product_type=product.strip(), account_no=account,
            balance=parse_money(_money_cell(amount), config),
            interest_rate=None if rate.strip() in ("", "-") else rate.strip(),
            last_interest_payment_date=_date_cell(last_int),
            maturity_date=_date_cell(maturity),
            restrictions=restr.strip() or None,
            source=SourceRef(source_file=header.source.source_file, section_id="deposits"),
        ))
    return rows


def loans_from_table(table, header: ConfirmationHeader, config: InstitutionConfig) -> list[LoanRow]:
    """§2-2 대출 표 → LoanRow (9열: 종류/약정/대출/대출일/만기/연이율/이자일/상환/담보)."""
    rows: list[LoanRow] = []
    for cells in table.rows:
        c = (list(cells) + [""] * 9)[:9]
        kind, limit, drawn, l_date, mat, rate, int_date, repay, coll = (x.strip() for x in c)
        rows.append(LoanRow(
            loan_type=kind,
            limit=parse_money(_money_cell(limit), config),
            drawn=parse_money(_money_cell(drawn), config),
            loan_date=_date_cell(l_date), maturity_date=_date_cell(mat),
            interest_rate=None if rate in ("", "-") else rate,
            last_interest_date=_date_cell(int_date),
            repayment=repay or None, collateral=coll or None,
        ))
    return rows


# ── top-level ──────────────────────────────────────────────────────────────

def extract_record(
    source_file: str, full_text: str, pages: int, config: InstitutionConfig,
    section_tables: dict | None = None,
) -> ConfirmationRecord:
    parsed = parse_filename(source_file)
    category: InstitutionCategory = config.classify(parsed["institution_name"])
    section_tables = section_tables or {}
    findings: list[ValidationFinding] = []

    # Cross-check the filename business number against the document body.
    doc_biz = BIZ_NO_RE.search(full_text)
    if doc_biz and parsed["business_no"] and doc_biz.group(1) != parsed["business_no"]:
        findings.append(
            ValidationFinding(
                code="BIZNO_MISMATCH",
                severity="warning",
                message=f"파일명 사업자번호({parsed['business_no']})와 본문({doc_biz.group(1)})이 다릅니다.",
                source_file=source_file,
            )
        )
    if category == "UNKNOWN":
        findings.append(
            ValidationFinding(
                code="CATEGORY_UNKNOWN",
                severity="warning",
                message=f"기관 분류 실패: {parsed['institution_name']!r} — configs/institutions.yaml 보강 필요",
                source_file=source_file,
            )
        )

    header = ConfirmationHeader(
        company_name=parsed["company_name"] or "",
        business_no=parsed["business_no"] or "",
        institution_name=parsed["institution_name"] or "",
        institution_category=category,
        confirmation_date=normalize_date(parsed["confirmation_date"]) or "",
        source=SourceRef(source_file=source_file, extraction_method="text_layer"),
    )

    # Prefer table extraction (find_tables) — preserves columns incl. 인출제한 and is
    # immune to page-break header repetition; fall back to the line parser if absent.
    deposits: list[BankDepositRow] = []
    loans: list[LoanRow] = []
    if category == "BANK":
        if "1" in section_tables:
            deposits = deposits_from_table(section_tables["1"], header, config)
        else:
            deposits = parse_bank_deposits(full_text, header, config)
        if "2-2" in section_tables:
            loans = loans_from_table(section_tables["2-2"], header, config)

    # Status resolution (deterministic, human-judgment cases flagged not guessed):
    #   1. explicit non-dealing phrase in the document  -> NON_DEALING
    #   2. no transaction data at all (no deposits, no currency/account tokens)
    #      -> EMPTY_REVIEW (cck sometimes hand-classifies these as 거래하지 않는
    #         금융기관 — not deterministically derivable, so we flag for review)
    #   3. otherwise                                      -> COMPLETE
    data_present = bool(deposits) or _has_data_tokens(full_text)
    if config.is_explicit_non_dealing(full_text):
        status = STATUS_NON_DEALING
    elif not data_present:
        status = STATUS_EMPTY_REVIEW
        findings.append(
            ValidationFinding(
                code="EMPTY_FORM",
                severity="info",
                message=f"{category} 조회서에서 거래 데이터가 추출되지 않았습니다 — 비거래/미회수 여부 확인 필요.",
                source_file=source_file,
            )
        )
    else:
        status = STATUS_COMPLETE

    return ConfirmationRecord(
        header=header,
        status=status,
        source_file=source_file,
        pages=pages,
        deposits=tuple(deposits),
        loans=tuple(loans),
        findings=tuple(findings),
        sections=split_pdf_sections(full_text),
        section_tables=section_tables,
    )

"""Write the standard confirmation workbook (요약 + 예금 명세).

This is *our* deterministic standard layout, not a byte-for-byte clone of the
cck/스마트리뷰어 sheet. The cck workbook is reproduced for reconciliation via
`afc.evaluation.diff_xlsx`; here we emit a clean, machine-checkable canonical view.
"""
from __future__ import annotations

import re
from collections.abc import Iterable
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from afc.extract import ConfirmationRecord, format_date8, parse_loans
from afc.institutions import InstitutionConfig
from afc.reconciliation import ReconRow, build_recon_rows, load_account_mapping
from afc.sections import ID_COLUMNS, SectionSpec, load_section_spec

_HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
_HEADER_FONT = Font(color="FFFFFF", bold=True)
_SUBTOTAL_FILL = PatternFill("solid", fgColor="DDEBF7")
_TOTAL_FILL = PatternFill("solid", fgColor="FCE4D6")
_PBC_FILL = PatternFill("solid", fgColor="FFF2CC")  # 감사인 입력 칸 (회사제시/환율)
_RECON_COLS = [
    "계정분류", "구분", "금융기관", "계좌번호", "금융상품종류", "통화",
    "환산전 금액", "적용환율", "환산후 금액(원화)",
    "이자율", "만기", "사용제한",
    "[회사제시] 환산후(원화)", "차이(원화)", "분류근거",
]
_FX_COL, _PRE_COL, _POST_COL = 8, 7, 9          # 적용환율 / 환산전 / 환산후
_PBC_POST_COL, _DIFF_COL = 13, 14
_KRW = {"KRW", "WON", None}
_SUMMARY_COLS = ["조서 번호", "금융기관명", "사업자등록번호", "취합 현황"]
_DEPOSIT_COLS = [
    "조회대상 회사", "사업자번호", "금융기관명", "조회기준일",
    "금융상품의 종류", "계좌번호", "통화", "금액",
    "연이자율", "최종이자지급일", "만기일", "인출제한 등",
]


def _style_header(ws, row: int, ncols: int) -> None:
    for c in range(1, ncols + 1):
        cell = ws.cell(row, c)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")


def _display_width(value) -> int:
    """한글·전각은 라틴 대비 폭이 넓어 약 1.8배로 환산."""
    s = str(value)
    return sum(2 if ord(ch) > 0x2E7F else 1 for ch in s)


def _autosize(ws, max_width: int = 50, min_row: int = 1) -> None:
    """열 너비 자동. min_row 이전 행(예: 1행의 긴 섹션 머리말)은 폭 계산에서 제외해
    열자마자 표가 보기 좋게 한다(머리말은 빈 오른쪽 칸으로 자연히 넘쳐 보임)."""
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        widths = [
            _display_width(c.value) for c in col
            if c.value is not None and c.row >= min_row
            and not (isinstance(c.value, str) and c.value.startswith("="))  # 수식은 결과가 짧음
        ]
        width = max(widths, default=8)
        ws.column_dimensions[letter].width = min(max_width, max(9, width + 2))


def _summary_rows(records: Iterable[ConfirmationRecord], config: InstitutionConfig):
    rows = []
    for rec in records:
        rows.append(
            (
                rec.header.institution_name,
                rec.header.business_no,
                config.status_label(rec.status),
            )
        )
    # cck ordering: institution name (Korean), then business number.
    rows.sort(key=lambda r: (r[0], r[1]))
    return rows


def _money_fmt(cell) -> None:
    cell.number_format = "#,##0.00"


def build_recon_sheet(ws, rows: list[ReconRow], client: str) -> None:
    """감사조서 4200 '조회서대사' 축으로 정리한 시트.

    계정분류 → 통화 순으로 묶고, (계정분류·통화) 소계와 계정분류별 환산후(원화)
    roll-up, 전체 총계를 SUM 수식으로 넣는다. 적용환율·회사제시 칸은 노란색으로
    비워 두어 감사인이 채우면 환산후/차이가 자동 계산된다.
    """
    ws.cell(1, 1, f"프로젝트명: {client}").font = Font(bold=True)
    ws.cell(2, 1, "금융기관조회서 ↔ 회사제시(PBC) 대사").font = Font(bold=True, size=12)
    hdr = 4
    for j, name in enumerate(_RECON_COLS, start=1):
        ws.cell(hdr, j, name)
    _style_header(ws, hdr, len(_RECON_COLS))

    r = hdr + 1
    post_rows_per_category: dict[str, list[int]] = {}

    # 계정분류 → 통화 그룹 경계에서 소계를 삽입.
    from itertools import groupby
    for category, cat_iter in groupby(rows, key=lambda x: x.category):
        cat_rows = list(cat_iter)
        cat_data_rows: list[int] = []
        for currency, cur_iter in groupby(cat_rows, key=lambda x: x.currency):
            cur_rows = list(cur_iter)
            first = r
            for row in cur_rows:
                is_krw = row.currency in _KRW
                ws.cell(r, 1, row.category)
                ws.cell(r, 2, row.side)
                ws.cell(r, 3, row.institution)
                ws.cell(r, 4, row.account_no)
                ws.cell(r, 5, row.product)
                ws.cell(r, 6, row.currency)
                _money_fmt(ws.cell(r, _PRE_COL, row.amount))
                fx = ws.cell(r, _FX_COL, 1 if is_krw else None)
                if not is_krw:
                    fx.fill = _PBC_FILL  # 외화 환율은 감사인 입력
                post = ws.cell(r, _POST_COL, f"=G{r}*H{r}")
                _money_fmt(post)
                ws.cell(r, 10, row.interest_rate)
                ws.cell(r, 11, row.maturity)
                ws.cell(r, 12, row.restrictions)
                ws.cell(r, _PBC_POST_COL).fill = _PBC_FILL          # 회사제시 입력칸
                _money_fmt(ws.cell(r, _DIFF_COL, f"=I{r}-M{r}"))    # 차이 = 조회서 - PBC
                ws.cell(r, 15, row.basis)
                cat_data_rows.append(r)
                r += 1
            # (계정분류·통화) 소계 — 환산전 합계
            ws.cell(r, 5, f"  ▷ {category} / {currency or '원화'} 소계")
            _money_fmt(ws.cell(r, _PRE_COL, f"=SUM(G{first}:G{r-1})"))
            _money_fmt(ws.cell(r, _POST_COL, f"=SUM(I{first}:I{r-1})"))
            for c in range(1, len(_RECON_COLS) + 1):
                ws.cell(r, c).fill = _SUBTOTAL_FILL
            ws.cell(r, 5).font = Font(italic=True)
            r += 1
        post_rows_per_category[category] = cat_data_rows

    # 전체 총계 (환산후 원화) — 데이터 행의 환산후만 합산.
    all_post = [rr for rows_ in post_rows_per_category.values() for rr in rows_]
    r += 1
    ws.cell(r, 5, "■ 총계 (환산후 원화)").font = Font(bold=True)
    total_cell = ws.cell(r, _POST_COL, "=" + "+".join(f"I{x}" for x in all_post) if all_post else 0)
    _money_fmt(total_cell)
    for c in range(1, len(_RECON_COLS) + 1):
        ws.cell(r, c).fill = _TOTAL_FILL

    ws.freeze_panes = "A5"
    _autosize(ws)


def build_pbc_recon_sheet(ws, result, client: str) -> None:
    """조회서 ↔ 회사제시(PBC) 계좌 단위 대사 + 양방향 완전성 예외.

    감사조서 4200 절차: 회사명세→조회서, 조회서→회사명세 두 방향을 모두 확인한다.
    """
    cols = ["계좌번호", "금융기관", "계정분류", "통화", "조회서 금액", "PBC 금액", "차이", "상태"]
    ws.cell(1, 1, f"프로젝트명: {client}").font = Font(bold=True)
    ws.cell(2, 1, "조회서 ↔ 회사제시(PBC) 대사").font = Font(bold=True, size=12)
    s = result.summary()
    ws.cell(3, 1, f"일치 {s['matched']-s['diff']} · 차이 {s['diff']} · "
                  f"조회서에만(명세누락) {s['confirm_only']} · 명세에만(미회수) {s['pbc_only']}")

    hdr = 5
    for j, name in enumerate(cols, start=1):
        ws.cell(hdr, j, name)
    _style_header(ws, hdr, len(cols))
    r = hdr + 1

    ws.cell(r, 1, "■ 계좌 대사 결과").font = Font(bold=True)
    r += 1
    for m in result.matched:
        ws.cell(r, 1, m.account_key)
        ws.cell(r, 2, m.institution)
        ws.cell(r, 3, m.category)
        ws.cell(r, 4, ", ".join(m.currencies))
        _money_fmt(ws.cell(r, 5, m.confirm_amount))
        _money_fmt(ws.cell(r, 6, m.pbc_amount))
        _money_fmt(ws.cell(r, 7, m.diff))
        st = ws.cell(r, 8, m.status)
        if m.status == "차이":
            for c in range(1, len(cols) + 1):
                ws.cell(r, c).fill = _TOTAL_FILL
            st.font = Font(bold=True, color="C00000")
        r += 1

    if result.confirm_only:
        r += 1
        ws.cell(r, 1, "■ 조회서에만 존재 — 회사명세 누락 (조회서→명세 완전성)").font = Font(bold=True, color="C00000")
        r += 1
        for key, inst, cat, amt in result.confirm_only:
            ws.cell(r, 1, key); ws.cell(r, 2, inst); ws.cell(r, 3, cat)
            _money_fmt(ws.cell(r, 5, amt))
            ws.cell(r, 8, "명세누락")
            for c in range(1, len(cols) + 1):
                ws.cell(r, c).fill = _SUBTOTAL_FILL
            r += 1

    if result.pbc_only:
        r += 1
        ws.cell(r, 1, "■ 회사명세에만 존재 — 조회서 미회수 (명세→조회서 완전성)").font = Font(bold=True, color="C00000")
        r += 1
        for p in result.pbc_only:
            ws.cell(r, 1, p.account_key); ws.cell(r, 2, p.institution)
            ws.cell(r, 4, p.currency)
            _money_fmt(ws.cell(r, 6, p.amount))
            ws.cell(r, 8, "미회수")
            for c in range(1, len(cols) + 1):
                ws.cell(r, c).fill = _PBC_FILL
            r += 1

    ws.freeze_panes = "A6"
    _autosize(ws)


def _aggregate(rows: list[ReconRow], *keys):
    """(key tuple) -> [건수, 환산전합계]. 통화는 항상 키에 포함해 외화 혼합 합산을 방지."""
    from collections import defaultdict
    agg: dict[tuple, list] = defaultdict(lambda: [0, 0.0])
    for r in rows:
        k = tuple((r.currency or "KRW") if key == "currency" else getattr(r, key) for key in keys)
        slot = agg[k]
        slot[0] += 1
        slot[1] += r.amount or 0.0
    return agg


def _write_summary(ws, rows: list[ReconRow], primary: str, primary_label: str, client: str) -> None:
    """primary(계정분류/금융기관/통화) → 통화 2단 roll-up. 통화별로 분리 합산."""
    is_currency_primary = primary == "currency"
    cols = ([primary_label, "건수", "환산전 합계"] if is_currency_primary
            else [primary_label, "통화", "건수", "환산전 합계", "환산후(원화)*"])
    ws.cell(1, 1, f"프로젝트명: {client}").font = Font(bold=True)
    ws.cell(2, 1, f"{primary_label} 정리").font = Font(bold=True, size=12)
    if not is_currency_primary:
        ws.cell(3, 1, "* 환산후(원화)는 KRW만 표시 — 외화는 '대사' 시트에서 환율 입력 후 합산").font = Font(size=9, italic=True)
    hdr = 4
    for j, name in enumerate(cols, start=1):
        ws.cell(hdr, j, name)
    _style_header(ws, hdr, len(cols))
    r = hdr + 1

    if is_currency_primary:
        agg = _aggregate(rows, "currency")
        order = ["KRW", "USD", "EUR", "JPY", "CNY", "HKD"]
        for key in sorted(agg, key=lambda x: (order.index(x[0]) if x[0] in order else 99, x[0])):
            cur = key[0]
            cnt, total = agg[key]
            ws.cell(r, 1, cur); ws.cell(r, 2, cnt)
            _money_fmt(ws.cell(r, 3, round(total, 2)))
            r += 1
        _autosize(ws)
        return

    agg = _aggregate(rows, primary, "currency")
    primaries = sorted({k[0] for k in agg}, key=lambda x: (ReconRow._ORDER.index(x) if x in ReconRow._ORDER else 99, x))
    for p in primaries:
        first = r
        for cur in sorted({k[1] for k in agg if k[0] == p}):
            cnt, total = agg[(p, cur)]
            ws.cell(r, 1, p); ws.cell(r, 2, cur); ws.cell(r, 3, cnt)
            _money_fmt(ws.cell(r, 4, round(total, 2)))
            if cur in _KRW:
                _money_fmt(ws.cell(r, 5, round(total, 2)))
            r += 1
        if r - first > 1:  # 소계 (KRW 환산후만 합산)
            ws.cell(r, 1, f"  ▷ {p} 소계").font = Font(italic=True)
            ws.cell(r, 3, f"=SUM(C{first}:C{r-1})")
            _money_fmt(ws.cell(r, 5, f"=SUM(E{first}:E{r-1})"))
            for c in range(1, len(cols) + 1):
                ws.cell(r, c).fill = _SUBTOTAL_FILL
            r += 1
    _autosize(ws)


def _section_content_text(content, spec: SectionSpec) -> str:
    """데이터 섹션의 원문을 컬럼헤더 echo 제거 후 ' / '로 결합 (원문 보존).

    cck 헤더는 '금액_약정한도액'처럼 병합셀이라 PDF에는 '금액'·'약정한도액'으로
    쪼개져 나온다 → 컬럼명을 _ · 공백으로 분해해 echo 라인을 폭넓게 제거한다.
    """
    labels = set(ID_COLUMNS)
    for col in spec.data_columns:
        labels.add(col)
        for part in re.split(r"[_·\s]+", col):
            if part:
                labels.add(part)
    lines = [
        ln for ln in content.raw_lines
        if ln not in labels and not ln.startswith("총 한도액")  # §2-1 값이 §2 본문에 섞임
    ]
    text = " / ".join(lines)
    return text if len(text) <= 900 else text[:900] + " …(원문 일부 — 전체는 JSONL)"


def build_category_sheet(ws, category: str, records: list[ConfirmationRecord],
                         specs: list[SectionSpec], client: str, config: InstitutionConfig) -> None:
    """cck 카테고리 시트(BANK/INSURANCE/GUARANTEE/INVESTMENT)와 동일한 섹션 구조로 렌더.

    각 섹션: 머리말 행 → 컬럼헤더 행(조서번호 포함) → 기관별 행.
    §1(은행 예금)은 구조화된 행, 그 외 데이터 섹션은 해당없음/원문보존(컬럼화는 차기).
    """
    records = sorted(records, key=lambda r: (r.header.institution_name, r.header.business_no))
    ws.cell(1, 1, f"프로젝트명: {client}").font = Font(bold=True)
    r = 3
    for spec in specs:
        ws.cell(r, 1, f"{spec.id}. {spec.header}").font = Font(bold=True)
        r += 1
        is_bank_deposit = category == "BANK" and spec.id == "1"
        is_bank_loan = category == "BANK" and spec.id == "2-2"
        # 금액 칼럼은 통화/금액 두 칼럼으로 분리해 보여준다.
        columns = list(spec.columns)
        if is_bank_deposit:
            amt_i = columns.index("금액")
            columns[amt_i:amt_i + 1] = ["통화", "금액"]
        elif is_bank_loan:
            columns = list(ID_COLUMNS) + [
                "대출 종류", "약정한도액_통화", "약정한도액_금액",
                "대출금액_통화", "대출금액_금액", "대출일", "최종만기일",
                "연이율", "최종이자지급일", "상환방법", "담보 보증 및 관련약정",
            ]
        for j, col in enumerate(columns, start=1):
            ws.cell(r, j, col)
        _style_header(ws, r, len(columns))
        r += 1
        ndata = len(spec.data_columns)
        for rec in records:
            h = rec.header
            base = ["", h.company_name, h.business_no, h.institution_name, h.confirmation_date]
            content = rec.sections.get(spec.top)
            if is_bank_deposit and rec.deposits:
                for d in rec.deposits:
                    vals = base + [d.product_type, d.account_no,
                                   d.balance.currency, d.balance.amount,
                                   d.interest_rate or "-",
                                   format_date8(d.last_interest_payment_date),
                                   format_date8(d.maturity_date), d.restrictions or ""]
                    for j, v in enumerate(vals[:len(columns)], start=1):
                        c = ws.cell(r, j, v)
                        if j == len(ID_COLUMNS) + 4:  # 금액
                            c.number_format = "#,##0.00"
                    r += 1
                continue
            if is_bank_loan:
                loans = parse_loans(content, rec.header, config)
                if loans:
                    for L in loans:
                        vals = base + [
                            L.loan_type, L.limit.currency, L.limit.amount,
                            L.drawn.currency, L.drawn.amount,
                            format_date8(L.loan_date), format_date8(L.maturity_date),
                            L.interest_rate or "-", format_date8(L.last_interest_date),
                            L.repayment, L.collateral,
                        ]
                        for j, v in enumerate(vals, start=1):
                            c = ws.cell(r, j, v)
                            if j in (len(ID_COLUMNS) + 3, len(ID_COLUMNS) + 5):  # 금액 칼럼들
                                c.number_format = "#,##0.00"
                        r += 1
                else:
                    for j, v in enumerate(base + ["해당 없음"], start=1):
                        ws.cell(r, j, v)
                    r += 1
                continue
            # 상태/원문 행 (기관 1행)
            if content is None or not content.has_data:
                first = "해당 없음"
            elif spec.id == "2-1" and content.summary_value:
                first = content.summary_value
            elif spec.id == "2-2" and not content.summary_value:
                first = "해당 없음"
            else:
                first = _section_content_text(content, spec)
            for j, v in enumerate(base, start=1):
                ws.cell(r, j, v)
            if ndata:
                ws.cell(r, len(ID_COLUMNS) + 1, first)
            r += 1
        r += 1  # 섹션 간 빈 행
    ws.freeze_panes = "A3"
    _autosize(ws, max_width=46)


def build_category_sheets(wb: Workbook, records: list[ConfirmationRecord], client: str,
                          config: InstitutionConfig) -> None:
    spec = load_section_spec()
    by_cat: dict[str, list[ConfirmationRecord]] = {}
    for rec in records:
        by_cat.setdefault(rec.header.institution_category, []).append(rec)
    for category in ("BANK", "INSURANCE", "GUARANTEE", "INVESTMENT"):
        recs = by_cat.get(category)
        if recs and category in spec:
            build_category_sheet(wb.create_sheet(category), category, recs, spec[category], client, config)


def build_workbook(
    records: list[ConfirmationRecord],
    config: InstitutionConfig,
    client: str,
    pbc_result=None,
) -> Workbook:
    wb = Workbook()

    # ── 요약 ────────────────────────────────────────────────────────────────
    ws = wb.active
    ws.title = "요약"
    ws.cell(1, 1, f"프로젝트명: {client}").font = Font(bold=True)
    for j, name in enumerate(_SUMMARY_COLS, start=1):
        ws.cell(3, j, name)
    _style_header(ws, 3, len(_SUMMARY_COLS))
    r = 4
    for inst, biz, status in _summary_rows(records, config):
        ws.cell(r, 1, None)  # 조서 번호 — assigned later by the auditor
        ws.cell(r, 2, inst)
        ws.cell(r, 3, biz)
        ws.cell(r, 4, status)
        r += 1
    ws.freeze_panes = "A4"
    _autosize(ws)

    # ── 카테고리 시트 (cck 동일 형식: 섹션 머리말 + 컬럼헤더 + 조서번호) ───────
    build_category_sheets(wb, records, client, config)

    # ── 대사 (계정분류·통화별 소계 + 환산후 roll-up) ─────────────────────────
    recon_rows = build_recon_rows(records, load_account_mapping())
    if recon_rows:
        build_recon_sheet(wb.create_sheet("대사"), recon_rows, client)
        _write_summary(wb.create_sheet("계정과목별"), recon_rows, "category", "계정과목", client)
        _write_summary(wb.create_sheet("금융기관별"), recon_rows, "institution", "금융기관", client)
        _write_summary(wb.create_sheet("통화별"), recon_rows, "currency", "통화", client)
    if pbc_result is not None:
        build_pbc_recon_sheet(wb.create_sheet("PBC대사"), pbc_result, client)

    # ── 예금 명세 (BANK section 1) ───────────────────────────────────────────
    ws2 = wb.create_sheet("예금")
    for j, name in enumerate(_DEPOSIT_COLS, start=1):
        ws2.cell(1, j, name)
    _style_header(ws2, 1, len(_DEPOSIT_COLS))
    r = 2
    deposits = [d for rec in records for d in rec.deposits]
    deposits.sort(key=lambda d: (d.header.institution_name, d.header.business_no, d.account_no))
    for d in deposits:
        h = d.header
        ws2.cell(r, 1, h.company_name)
        ws2.cell(r, 2, h.business_no)
        ws2.cell(r, 3, h.institution_name)
        ws2.cell(r, 4, h.confirmation_date)
        ws2.cell(r, 5, d.product_type)
        ws2.cell(r, 6, d.account_no)
        ws2.cell(r, 7, d.balance.currency)
        amt = ws2.cell(r, 8, d.balance.amount)
        amt.number_format = "#,##0.00"
        ws2.cell(r, 9, d.interest_rate)
        ws2.cell(r, 10, format_date8(d.last_interest_payment_date))
        ws2.cell(r, 11, format_date8(d.maturity_date))
        ws2.cell(r, 12, d.restrictions)
        r += 1
    ws2.freeze_panes = "A2"
    _autosize(ws2)
    return wb


def write_workbook(
    records: list[ConfirmationRecord],
    config: InstitutionConfig,
    client: str,
    path: Path,
    pbc_result=None,
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    build_workbook(records, config, client, pbc_result=pbc_result).save(path)

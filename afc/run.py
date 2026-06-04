"""End-to-end pipeline: confirmation zip -> canonical JSONL + standard xlsx + 검증 리포트.

매 실행 LLM 호출 0회. 기관별 가변성은 configs/institutions.yaml 로 동결.

사용:
    python -m afc.run "<...금융기관조회서(전자)_[client]_[date].zip>"
    python -m afc.run "<...zip>" --benchmark "<...cck 산출물.xlsx>"
"""
from __future__ import annotations

import argparse
import json
import re
import zipfile
from collections import Counter
from dataclasses import asdict
from datetime import datetime
from pathlib import Path

import fitz

from afc.evaluation.diff_xlsx import diff_xlsx
from afc.output import write_outputs
from afc.extract import (
    STATUS_EMPTY_REVIEW,
    STATUS_NON_DEALING,
    ConfirmationRecord,
    extract_record,
    parse_filename,
)
from afc.tables import extract_section_tables
from afc.institutions import InstitutionConfig, load_institution_config
from afc.pbc import load_pbc, reconcile
from afc.reconciliation import build_recon_rows, load_account_mapping

CLIENT_RE = re.compile(r"_\[(?P<client>[^\]]+)\]_\[(?P<date>[^\]]+)\]")


def parse_zip_meta(zip_path: Path) -> tuple[str, str]:
    """'금융기관조회서(전자)_[삼화전기]_[2025-12-31].zip' -> ('삼화전기', '2025-12-31')."""
    m = CLIENT_RE.search(zip_path.stem)
    if m:
        return m.group("client"), m.group("date")
    return zip_path.stem, ""


def extract_records(zip_path: Path, config: InstitutionConfig) -> list[ConfirmationRecord]:
    records: list[ConfirmationRecord] = []
    with zipfile.ZipFile(zip_path) as archive:
        for entry in archive.infolist():
            if not entry.filename.lower().endswith(".pdf"):
                continue
            data = archive.read(entry)
            doc = fitz.open(stream=data, filetype="pdf")
            try:
                full_text = "\n".join(doc[i].get_text() for i in range(doc.page_count))
                pages = doc.page_count
            finally:
                doc.close()
            category = config.classify(parse_filename(entry.orig_filename)["institution_name"])
            section_tables = extract_section_tables(data, category)
            records.append(
                extract_record(entry.orig_filename, full_text, pages, config, section_tables)
            )
    return records


def write_jsonl(records: list[ConfirmationRecord], path: Path) -> None:
    with path.open("w", encoding="utf-8") as f:
        for rec in records:
            f.write(json.dumps(asdict(rec), ensure_ascii=False, default=str) + "\n")


def write_validation_report(
    records: list[ConfirmationRecord], path: Path, client: str, config: InstitutionConfig,
    pbc_result=None,
) -> None:
    status_counts = Counter(config.status_label(r.status) for r in records)
    cat_counts = Counter(r.header.institution_category for r in records)
    total_deposits = sum(len(r.deposits) for r in records)
    findings = [(r.source_file, f) for r in records for f in r.findings]

    lines = [
        f"# 금융기관조회서 검증 리포트 — {client}",
        "",
        f"- 생성: {datetime.now():%Y-%m-%d %H:%M:%S}",
        f"- 조회서(PDF) 수: **{len(records)}**",
        f"- 추출 예금(BANK §1) 행: **{total_deposits}**",
        "",
        "## 취합 현황",
        "",
        "| 상태 | 건수 |",
        "| --- | ---: |",
    ]
    lines += [f"| {label} | {n} |" for label, n in sorted(status_counts.items())]
    lines += ["", "## 기관 분류", "", "| 분류 | 건수 |", "| --- | ---: |"]
    lines += [f"| {cat} | {n} |" for cat, n in sorted(cat_counts.items())]

    review = [r for r in records if r.status == STATUS_EMPTY_REVIEW]
    if review:
        lines += [
            "",
            "## ⚠ 수동 확인 필요 (예금 데이터 미추출)",
            "",
            "은행 조회서이나 §1 금융상품이 비어 있습니다. 비거래/미회수 여부를 확인하세요.",
            "",
        ]
        lines += [f"- {r.header.institution_name} ({r.header.business_no})" for r in review]

    non_dealing = [r for r in records if r.status == STATUS_NON_DEALING]
    if non_dealing:
        lines += ["", "## 비거래 (조회서상 명시)", ""]
        lines += [f"- {r.header.institution_name} ({r.header.business_no})" for r in non_dealing]

    if pbc_result is not None:
        s = pbc_result.summary()
        lines += [
            "", "## 회사제시(PBC) 대사",
            "",
            f"- 계좌 일치: **{s['matched'] - s['diff']}** / 금액 차이: **{s['diff']}**",
            f"- 조회서에만 존재(회사명세 누락): **{s['confirm_only']}**",
            f"- 회사명세에만 존재(조회서 미회수): **{s['pbc_only']}**",
        ]
        if any(m.status == "차이" for m in pbc_result.matched):
            lines += ["", "### 금액 차이 계좌", "", "| 계좌 | 금융기관 | 조회서 | PBC | 차이 |", "| --- | --- | ---: | ---: | ---: |"]
            for m in pbc_result.matched:
                if m.status == "차이":
                    lines.append(f"| {m.account_key} | {m.institution} | {m.confirm_amount:,.0f} | {m.pbc_amount:,.0f} | {m.diff:,.0f} |")
        if pbc_result.confirm_only:
            lines += ["", "### 조회서에만 존재 (회사명세 누락 — 완전성)", ""]
            lines += [f"- {inst} / {key} / {amt:,.0f}" for key, inst, _cat, amt in pbc_result.confirm_only]
        if pbc_result.pbc_only:
            lines += ["", "### 회사명세에만 존재 (조회서 미회수 — 완전성)", ""]
            lines += [f"- {p.institution} / {p.account_key} / {(p.amount or 0):,.0f}" for p in pbc_result.pbc_only]

    if findings:
        lines += ["", "## 검증 메시지", "", "| 파일 | 코드 | 심각도 | 메시지 |", "| --- | --- | --- | --- |"]
        for src, f in findings:
            lines.append(f"| {src} | {f.code} | {f.severity} | {f.message} |")

    lines.append("")
    path.write_text("\n".join(lines), encoding="utf-8")


def run_pipeline(
    zip_path: Path,
    pbc: Path | None = None,
    out_base: Path = Path("output"),
    log=print,
) -> tuple[Path, list[Path]]:
    """end-to-end. 결과 폴더와 생성 파일 목록 반환. log: 진행상황 콜백(GUI 등)."""
    config = load_institution_config()
    client, _date = parse_zip_meta(zip_path)

    log(f"[1/4] 조회서 추출 - {zip_path.name}")
    records = extract_records(zip_path, config)
    deposits = sum(len(r.deposits) for r in records)
    log(f"      PDF {len(records)}건 / 예금 {deposits}행")

    pbc_result = None
    if pbc and Path(pbc).exists():
        recon_rows = build_recon_rows(records, load_account_mapping())
        pbc_result = reconcile(recon_rows, load_pbc(Path(pbc)))
        s = pbc_result.summary()
        log(f"[2/4] 회사명세 대사 - 일치 {s['matched']-s['diff']} / 차이 {s['diff']} / "
            f"조회서에만 {s['confirm_only']} / 명세에만 {s['pbc_only']}")
    else:
        log("[2/4] 회사명세 대사 - (명세 미지정, 건너뜀)")

    run_ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    client_base = Path(out_base) / client
    out_dir = client_base / run_ts
    out_dir.mkdir(parents=True, exist_ok=True)
    (client_base / "_latest.txt").write_text(run_ts, encoding="utf-8")

    log("[3/4] 원시데이터(JSONL) 저장")
    write_jsonl(records, out_dir / "confirmations.jsonl")

    log("[4/4] 엑셀(다중 파일) + 검증리포트 생성")
    written = write_outputs(records, config, client, out_dir, pbc_result=pbc_result)
    write_validation_report(records, out_dir / "00_검증리포트.md", client, config, pbc_result)
    for p in written:
        log(f"      - {p.name}")

    log(f"[완료] {out_dir}")
    return out_dir, written


def main(zip_path: Path, benchmark: Path | None = None, pbc: Path | None = None) -> Path:
    out_dir, _ = run_pipeline(zip_path, pbc=pbc)
    return out_dir


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="금융기관조회서 PDF → 표준 xlsx/JSONL/검증리포트")
    parser.add_argument("zip_path", type=Path)
    parser.add_argument("--benchmark", type=Path, default=None, help="cck 산출물 xlsx (회귀 비교용)")
    parser.add_argument("--pbc", type=Path, default=None, help="회사제시 예금명세 xlsx (대사용)")
    args = parser.parse_args()
    main(args.zip_path, args.benchmark, args.pbc)

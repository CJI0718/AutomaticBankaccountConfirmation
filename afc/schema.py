from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Literal


InstitutionCategory = Literal["BANK", "INSURANCE", "GUARANTEE", "INVESTMENT", "UNKNOWN"]
ExtractionMethod = Literal["text_layer", "ocr", "llm_vision", "manual"]
Severity = Literal["info", "warning", "error"]


@dataclass(frozen=True)
class SourceRef:
    source_file: str
    page: int | None = None
    bbox: tuple[float, float, float, float] | None = None
    section_id: str | None = None
    extraction_method: ExtractionMethod = "text_layer"


@dataclass(frozen=True)
class MoneyAmount:
    raw: str | None
    currency: str | None
    amount: float | None


@dataclass(frozen=True)
class ConfirmationHeader:
    company_name: str
    business_no: str
    institution_name: str
    institution_category: InstitutionCategory
    confirmation_date: str
    source: SourceRef


@dataclass(frozen=True)
class BankDepositRow:
    header: ConfirmationHeader
    product_type: str
    account_no: str
    balance: MoneyAmount
    interest_rate: str | None
    last_interest_payment_date: str | None
    maturity_date: str | None
    restrictions: str | None
    source: SourceRef


@dataclass(frozen=True)
class ValidationFinding:
    code: str
    severity: Severity
    message: str
    source_file: str | None = None
    sheet: str | None = None
    cell: str | None = None


@dataclass
class RunManifest:
    client_name: str
    confirmation_date: str
    run_id: str
    created_at: datetime = field(default_factory=datetime.now)
    input_zip: str | None = None
    benchmark_xlsx: str | None = None
    findings: list[ValidationFinding] = field(default_factory=list)

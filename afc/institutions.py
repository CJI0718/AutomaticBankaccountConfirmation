"""Static institution configuration loader.

`configs/institutions.yaml` absorbs institution-level variability so the run hot
path stays deterministic (LLM 호출 0회). This module turns that YAML into typed
lookups: category classification, section signatures, and status phrases.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path

import yaml

from afc.paths import CONFIG_DIR
from afc.schema import InstitutionCategory


CONFIG_PATH = CONFIG_DIR / "institutions.yaml"

# Keyword fallbacks when an institution name is not an explicit alias.
# Order matters: 보증보험 must beat 보험, 증권 before generic.
_CATEGORY_KEYWORDS: list[tuple[str, InstitutionCategory]] = [
    ("보증보험", "GUARANTEE"),
    ("보증", "GUARANTEE"),
    ("증권", "INVESTMENT"),
    ("자산운용", "INVESTMENT"),
    ("손해보험", "INSURANCE"),
    ("화재보험", "INSURANCE"),
    ("생명", "INSURANCE"),
    ("보험", "INSURANCE"),
    ("은행", "BANK"),
]


@dataclass(frozen=True)
class CategoryConfig:
    name: InstitutionCategory
    aliases: tuple[str, ...]
    section_signatures: dict[str, str]


@dataclass(frozen=True)
class InstitutionConfig:
    categories: dict[str, CategoryConfig]
    no_transaction_phrases: tuple[str, ...]
    default_currency: str
    known_currencies: tuple[str, ...]
    status_labels: dict[str, str] = field(default_factory=dict)
    # Phrases printed in *every* section of a returned-but-non-dealing form.
    empty_section_phrases: tuple[str, ...] = field(
        default=("해당 거래 없음", "해당없음", "해당 없음")
    )

    def status_label(self, key: str) -> str:
        return self.status_labels.get(key, key)

    def classify(self, institution_name: str | None) -> InstitutionCategory:
        """Map an institution display name to a category.

        Exact-alias match first (config-driven), then keyword heuristics so new
        institutions classify without a config edit. Returns UNKNOWN when nothing
        matches — the caller should surface that as a validation finding.
        """
        if not institution_name:
            return "UNKNOWN"
        name = institution_name.strip()
        for cat in self.categories.values():
            if any(alias in name for alias in cat.aliases):
                return cat.name
        for keyword, category in _CATEGORY_KEYWORDS:
            if keyword in name:
                return category
        return "UNKNOWN"

    def is_explicit_non_dealing(self, text: str) -> bool:
        """True when the PDF explicitly states the institution does not deal with
        the company (deterministic). Empty-but-returned forms are *not* covered
        here — those are flagged for human review by the caller."""
        return any(phrase in text for phrase in self.no_transaction_phrases)


def _parse(raw: dict) -> InstitutionConfig:
    categories: dict[str, CategoryConfig] = {}
    for name, body in (raw.get("categories") or {}).items():
        categories[name] = CategoryConfig(
            name=name,  # type: ignore[arg-type]
            aliases=tuple(body.get("aliases") or ()),
            section_signatures=dict(body.get("section_signatures") or {}),
        )
    money = raw.get("money") or {}
    return InstitutionConfig(
        categories=categories,
        no_transaction_phrases=tuple(raw.get("no_transaction_phrases") or ()),
        default_currency=money.get("default_currency", "KRW"),
        known_currencies=tuple(money.get("known_currencies") or ()),
        status_labels=dict(raw.get("status_labels") or {}),
    )


@lru_cache(maxsize=4)
def load_institution_config(path: str | None = None) -> InstitutionConfig:
    config_path = Path(path) if path else CONFIG_PATH
    with config_path.open(encoding="utf-8") as f:
        raw = yaml.safe_load(f)
    return _parse(raw)

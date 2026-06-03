from __future__ import annotations

import argparse
import csv
import json
import re
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path

import fitz


FILENAME_RE = re.compile(
    r"^전자_\[.*?\]_(?P<company>.*?)_\[(?P<business_no>\d{3}-\d{2}-\d{5})\]_(?P<institution>.*?)_\[(?P<date>.*?)\]\.pdf$"
)
SECTION_RE = re.compile(r"^(?:[1-9]\.|[1-9]-[1-9]\.)\s+")


@dataclass(frozen=True)
class PdfInventoryItem:
    file: str
    company_name: str | None
    business_no: str | None
    institution_name: str | None
    confirmation_date: str | None
    pages: int
    producer: str | None
    encrypted: bool
    first_page_chars: int
    first_lines: list[str]
    section_lines: list[str]


def parse_confirmation_filename(name: str) -> dict[str, str | None]:
    match = FILENAME_RE.match(name)
    if not match:
        return {
            "company_name": None,
            "business_no": None,
            "institution_name": None,
            "confirmation_date": None,
        }
    return {
        "company_name": match.group("company"),
        "business_no": match.group("business_no"),
        "institution_name": match.group("institution"),
        "confirmation_date": match.group("date"),
    }


def section_candidates(lines: list[str]) -> list[str]:
    return [line for line in lines if len(line) <= 140 and SECTION_RE.match(line)]


def inspect_pdf_bytes(file_name: str, data: bytes) -> PdfInventoryItem:
    parsed = parse_confirmation_filename(Path(file_name).name)
    doc = fitz.open(stream=data, filetype="pdf")
    try:
        first_text = doc[0].get_text() if doc.page_count else ""
        sample_lines: list[str] = []
        for page_index in range(min(doc.page_count, 2)):
            sample_lines.extend(
                line.strip()
                for line in doc[page_index].get_text().splitlines()
                if line.strip()
            )
        return PdfInventoryItem(
            file=Path(file_name).name,
            company_name=parsed["company_name"],
            business_no=parsed["business_no"],
            institution_name=parsed["institution_name"],
            confirmation_date=parsed["confirmation_date"],
            pages=doc.page_count,
            producer=doc.metadata.get("producer"),
            encrypted=doc.is_encrypted,
            first_page_chars=len(first_text),
            first_lines=sample_lines[:8],
            section_lines=section_candidates(sample_lines)[:12],
        )
    finally:
        doc.close()


def inspect_zip(zip_path: Path) -> list[PdfInventoryItem]:
    items: list[PdfInventoryItem] = []
    with zipfile.ZipFile(zip_path) as archive:
        for entry in archive.infolist():
            if entry.filename.lower().endswith(".pdf"):
                items.append(inspect_pdf_bytes(entry.filename, archive.read(entry)))
    return items


def write_csv(items: list[PdfInventoryItem], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=list(asdict(items[0]).keys()) if items else [])
        writer.writeheader()
        for item in items:
            row = asdict(item)
            row["first_lines"] = " | ".join(item.first_lines)
            row["section_lines"] = " | ".join(item.section_lines)
            writer.writerow(row)


def main() -> None:
    parser = argparse.ArgumentParser(description="Inspect financial confirmation PDFs inside a zip.")
    parser.add_argument("zip_path", type=Path)
    parser.add_argument("--csv", type=Path, default=None, help="Optional CSV output path.")
    args = parser.parse_args()

    items = inspect_zip(args.zip_path)
    if args.csv:
        write_csv(items, args.csv)
    for item in items:
        print(json.dumps(asdict(item), ensure_ascii=False))


if __name__ == "__main__":
    main()

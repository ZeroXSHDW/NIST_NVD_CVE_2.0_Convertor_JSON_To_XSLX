"""
NIST CVE Data Integrity Validator
----------------------------------
Verifies the generated XLSX workbook against the source JSON files and exits
non-zero whenever the workbook is missing, malformed, truncated, or different
from the source records.
"""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from json_to_xlsx import (
    COLUMNS as ANNUAL_COLUMNS,
    DATA_DIR,
    INDEX_COLUMNS,
    OUTPUT_FILE,
    extract_entry,
)

ANNUAL_FIELDS = tuple(key for key, _ in ANNUAL_COLUMNS)
ANNUAL_HEADERS = tuple(header for _, header in ANNUAL_COLUMNS)
INDEX_FIELDS = tuple(key for key, _ in INDEX_COLUMNS)
INDEX_HEADERS = tuple(header for _, header in INDEX_COLUMNS)
YEAR_FILE = re.compile(r"^nvdcve-2\.0-(\d{4})\.json$")


def clean_text(text: Any) -> str:
    """Normalize text values before comparing JSON and workbook cells."""
    if not isinstance(text, str):
        return str(text or "")
    return text.replace("\r\n", "\n").replace("\n\r", "\n").replace("\r", "\n")


def normalize(value: Any) -> str:
    """Normalize Excel numeric values and text for stable comparisons."""
    if value is None or value == "":
        return ""
    try:
        return str(float(value))
    except (TypeError, ValueError):
        return clean_text(value)


def _row_to_record(row: tuple[Any, ...], fields: tuple[str, ...]) -> dict[str, Any]:
    """Map a worksheet row to named fields, padding missing trailing cells."""
    values = list(row[: len(fields)])
    values.extend([""] * (len(fields) - len(values)))
    return dict(zip(fields, values))


def _load_source_records(json_path: Path) -> tuple[list[dict[str, Any]], list[tuple[Any, ...]]]:
    """Load one source feed and return records plus source-format errors."""
    try:
        with json_path.open("r", encoding="utf-8") as source:
            data = json.load(source)
    except (OSError, json.JSONDecodeError) as exc:
        return [], [(json_path.name, "source_error", str(exc))]

    if not isinstance(data, dict):
        return [], [(json_path.name, "source_error", "top-level JSON value is not an object")]

    vulnerabilities = data.get("vulnerabilities")
    if vulnerabilities is None:
        return [], [(json_path.name, "source_error", "missing vulnerabilities list")]
    if not isinstance(vulnerabilities, list):
        return [], [(json_path.name, "source_error", "vulnerabilities is not a list")]

    errors: list[tuple[Any, ...]] = []
    records: list[dict[str, Any]] = []
    for position, vulnerability in enumerate(vulnerabilities, start=1):
        if not isinstance(vulnerability, dict):
            errors.append((json_path.name, "source_record", position, "record is not an object"))
            cve_entry: dict[str, Any] = {}
        else:
            raw_cve = vulnerability.get("cve")
            if raw_cve is not None and not isinstance(raw_cve, dict):
                errors.append((json_path.name, "source_record", position, "cve is not an object"))
            cve_entry = raw_cve if isinstance(raw_cve, dict) else {}
        records.append(extract_entry(cve_entry))
    return records, errors


def _compare_records(
    expected: list[dict[str, Any]],
    actual: list[dict[str, Any]],
    fields: tuple[str, ...],
    context: str,
    mismatches: list[tuple[Any, ...]],
) -> None:
    """Compare record counts and every field, including missing rows."""
    if len(actual) != len(expected):
        mismatches.append((context, "row_count", len(actual), len(expected)))

    for row_number, (expected_record, actual_record) in enumerate(
        zip(expected, actual), start=2
    ):
        for field in fields:
            actual_value = normalize(actual_record.get(field))
            expected_value = normalize(expected_record.get(field))
            if actual_value != expected_value:
                mismatches.append((context, row_number, field, actual_value, expected_value))


def validate_workbook(
    data_dir: Path = DATA_DIR,
    xlsx_path: Path = OUTPUT_FILE,
) -> list[tuple[Any, ...]]:
    """Return all source/workbook mismatches; an empty list means valid."""
    data_dir = Path(data_dir)
    xlsx_path = Path(xlsx_path)
    mismatches: list[tuple[Any, ...]] = []

    json_files = (
        sorted(path for path in data_dir.glob("nvdcve-2.0-*.json") if YEAR_FILE.fullmatch(path.name))
        if data_dir.is_dir()
        else []
    )
    source_years = {YEAR_FILE.fullmatch(path.name).group(1) for path in json_files}

    if not xlsx_path.is_file():
        return [("workbook", "missing", str(xlsx_path))]

    try:
        workbook = load_workbook(xlsx_path, read_only=True, data_only=False)
    except (OSError, ValueError) as exc:
        return [("workbook", "unreadable", str(exc))]

    try:
        if "INDEX" not in workbook.sheetnames:
            mismatches.append(("workbook", "missing_sheet", "INDEX"))

        actual_years = set(workbook.sheetnames) - {"INDEX"}
        for year in sorted(source_years - actual_years):
            mismatches.append(("workbook", "missing_sheet", year))
        for sheet in sorted(actual_years - source_years):
            mismatches.append(("workbook", "unexpected_sheet", sheet))

        expected_index: list[dict[str, Any]] = []
        for json_path in json_files:
            match = YEAR_FILE.fullmatch(json_path.name)
            assert match is not None  # filtered above
            year = match.group(1)
            expected_records, source_errors = _load_source_records(json_path)
            mismatches.extend(source_errors)

            for record in expected_records:
                expected_index.append(
                    {
                        "cve_id": record["cve_id"],
                        "year": year,
                        "severity": record["severity"],
                        "baseScore": record["baseScore"],
                        "summary": record["description"][:200]
                        + ("..." if len(record["description"]) > 200 else ""),
                    }
                )

            if year not in workbook.sheetnames:
                continue
            sheet = workbook[year]
            headers = tuple(
                sheet.cell(row=1, column=index).value
                for index in range(1, len(ANNUAL_FIELDS) + 1)
            )
            if headers != ANNUAL_HEADERS:
                mismatches.append((year, "headers", headers, ANNUAL_HEADERS))
            actual_records = [
                _row_to_record(row, ANNUAL_FIELDS)
                for row in sheet.iter_rows(min_row=2, values_only=True)
            ]
            _compare_records(expected_records, actual_records, ANNUAL_FIELDS, year, mismatches)

        if "INDEX" in workbook.sheetnames:
            index_sheet = workbook["INDEX"]
            headers = tuple(
                index_sheet.cell(row=1, column=index).value
                for index in range(1, len(INDEX_FIELDS) + 1)
            )
            if headers != INDEX_HEADERS:
                mismatches.append(("INDEX", "headers", headers, INDEX_HEADERS))
            expected_index.sort(key=lambda entry: entry["cve_id"], reverse=True)
            actual_index = [
                _row_to_record(row, INDEX_FIELDS)
                for row in index_sheet.iter_rows(min_row=2, values_only=True)
            ]
            _compare_records(expected_index, actual_index, INDEX_FIELDS, "INDEX", mismatches)
    finally:
        workbook.close()

    return mismatches


def main(data_dir: Path = DATA_DIR, xlsx_path: Path = OUTPUT_FILE) -> int:
    """Validate the workbook and return a shell-friendly status code."""
    print(f"Loading workbook {xlsx_path} for integrity validation...")
    mismatches = validate_workbook(data_dir=data_dir, xlsx_path=xlsx_path)
    if mismatches:
        print("Workbook validation failed:")
        for mismatch in mismatches[:10]:
            print(mismatch)
        print(f"Total mismatches: {len(mismatches)}")
        return 1

    print("All workbook rows and index entries match source JSON perfectly.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

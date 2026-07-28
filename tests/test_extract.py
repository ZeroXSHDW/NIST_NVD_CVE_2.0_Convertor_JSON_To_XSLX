"""Unit tests for json_to_xlsx extract helpers."""

import json
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from json_to_xlsx import clean_text, extract_entry  # noqa: E402

FIXTURE = Path(__file__).parent / "fixtures" / "sample_cve.json"


@pytest.fixture(scope="module")
def sample_vulns():
    with open(FIXTURE, encoding="utf-8") as f:
        data = json.load(f)
    return data["vulnerabilities"]


def test_clean_text_normalizes_line_endings():
    assert clean_text("a\r\nb\rc") == "a\nb\nc"
    assert clean_text(None) == ""
    assert clean_text(42) == "42"


def test_extract_entry_primary_cvss_v31(sample_vulns):
    extracted = extract_entry(sample_vulns[0]["cve"])
    assert extracted["cve_id"] == "CVE-2024-0001"
    assert extracted["description"] == "A sample vulnerability for unit tests.\nSecond line."
    assert extracted["published"] == "2024-01-15T10:00:00.000"
    assert extracted["lastModified"] == "2024-01-16T12:00:00.000"
    assert extracted["baseScore"] == 9.8
    assert extracted["severity"] == "CRITICAL"
    assert extracted["vectorString"].startswith("CVSS:3.1/")


def test_extract_entry_falls_back_to_secondary_metric(sample_vulns):
    extracted = extract_entry(sample_vulns[1]["cve"])
    assert extracted["cve_id"] == "CVE-2023-9999"
    assert extracted["baseScore"] == 5.0
    assert extracted["severity"] == "MEDIUM"
    assert extracted["vectorString"] == "AV:N/AC:L/Au:N/C:P/I:P/A:P"


def test_extract_entry_empty_cve():
    extracted = extract_entry({})
    assert extracted["cve_id"] == ""
    assert extracted["description"] == ""
    assert extracted["baseScore"] == ""
    assert extracted["severity"] == ""
    assert extracted["vectorString"] == ""

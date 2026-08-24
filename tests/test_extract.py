"""Unit tests for json_to_xlsx extract helpers and edge-case feeds."""

import io
import json
import sys
import zipfile
from pathlib import Path

import pytest
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import download_nvd  # noqa: E402
from download_nvd import (  # noqa: E402
    collect_feed_links,
    download_response_atomically,
    extract_zip_safely,
)
from json_to_xlsx import (  # noqa: E402
    build_workbook,
    clean_text,
    extract_entry,
    load_vulnerabilities,
)

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


def test_extract_entry_missing_and_null_fields():
    extracted = extract_entry(
        {
            "id": None,
            "descriptions": None,
            "metrics": None,
            "published": None,
            "lastModified": None,
        }
    )
    assert extracted["cve_id"] == ""
    assert extracted["description"] == ""
    assert extracted["published"] == ""
    assert extracted["lastModified"] == ""
    assert extracted["baseScore"] == ""


def test_extract_entry_non_dict_and_malformed_metrics():
    assert extract_entry(None)["cve_id"] == ""
    extracted = extract_entry(
        {
            "id": "CVE-2020-1",
            "descriptions": [{"lang": "en", "value": "ok"}, "skip-me"],
            "metrics": {
                "cvssMetricV31": [None, {"type": "Primary", "cvssData": None}],
            },
        }
    )
    assert extracted["cve_id"] == "CVE-2020-1"
    assert extracted["description"] == "ok"
    assert extracted["baseScore"] == ""
    assert extracted["vectorString"] == ""


def test_collect_feed_links_empty_and_relative(tmp_path):
    assert collect_feed_links("") == []
    assert collect_feed_links("<html></html>") == []
    html = """
    <a href="/feeds/json/cve/2.0/nvdcve-2.0-2024.json.zip">2024</a>
    <a href="https://nvd.nist.gov/feeds/json/cve/2.0/nvdcve-2.0-modified.json.zip">mod</a>
    <a href="/feeds/json/cve/1.1/nvdcve-1.1-2024.json.zip">ignore</a>
    """
    links = collect_feed_links(html)
    assert len(links) == 2
    assert all(link.startswith("https://nvd.nist.gov/") for link in links)
    assert any(link.endswith("nvdcve-2.0-2024.json.zip") for link in links)


def test_collect_feed_links_rejects_untrusted_hosts_and_query_urls():
    html = """
    <a href="https://evil.example/nvdcve-2.0-2024.json.zip">evil</a>
    <a href="http://nvd.nist.gov/feeds/json/cve/2.0/nvdcve-2.0-2024.json.zip">http</a>
    <a href="//evil.example/nvdcve-2.0-2024.json.zip">protocol-relative</a>
    <a href="/feeds/json/cve/2.0/nvdcve-2.0-2024.json.zip?download=1">query</a>
    <a href="/feeds/json/cve/2.0/nvdcve-2.0-not-a-feed.json.zip">invalid-name</a>
    <a href="/feeds/json/cve/2.0/nvdcve-2.0-recent.json.zip">recent</a>
    """
    assert collect_feed_links(html) == [
        "https://nvd.nist.gov/feeds/json/cve/2.0/nvdcve-2.0-recent.json.zip"
    ]


def test_extract_zip_safely_empty_and_corrupt(tmp_path):
    dest = tmp_path / "out"
    dest.mkdir()
    missing = tmp_path / "missing.zip"
    assert extract_zip_safely(missing, dest) is False

    empty = tmp_path / "empty.zip"
    empty.write_bytes(b"")
    assert extract_zip_safely(empty, dest) is False

    corrupt = tmp_path / "corrupt.zip"
    corrupt.write_bytes(b"not-a-zip")
    assert extract_zip_safely(corrupt, dest) is False

    good = tmp_path / "good.zip"
    with zipfile.ZipFile(good, "w") as zf:
        zf.writestr("nvdcve-2.0-2024.json", '{"vulnerabilities":[]}')
    assert extract_zip_safely(good, dest) is True
    assert (dest / "nvdcve-2.0-2024.json").is_file()


def test_extract_zip_safely_rejects_path_traversal(tmp_path):
    dest = tmp_path / "out"
    dest.mkdir()
    archive = tmp_path / "path-traversal.zip"
    outside = tmp_path / "outside.json"
    with zipfile.ZipFile(archive, "w") as zf:
        zf.writestr("../outside.json", "should not be written")

    assert extract_zip_safely(archive, dest) is False
    assert not outside.exists()


def test_download_response_atomically_preserves_previous_file_on_failure(tmp_path):
    destination = tmp_path / "feed.zip"
    destination.write_bytes(b"previous archive")

    class FailingResponse(io.BytesIO):
        def read(self, size=-1):
            if self.tell() > 0:
                raise OSError("simulated interrupted download")
            return super().read(size)

    with pytest.raises(OSError, match="interrupted"):
        download_response_atomically(FailingResponse(b"partial archive"), destination)

    assert destination.read_bytes() == b"previous archive"
    assert list(tmp_path.glob(".feed.zip.*.part")) == []


def test_download_response_atomically_replaces_after_success(tmp_path):
    destination = tmp_path / "feed.zip"
    destination.write_bytes(b"previous archive")

    download_response_atomically(io.BytesIO(b"complete archive"), destination)

    assert destination.read_bytes() == b"complete archive"
    assert list(tmp_path.glob(".feed.zip.*.part")) == []


def test_downloader_removes_corrupt_final_archive(monkeypatch, tmp_path):
    target_dir = tmp_path / "nvd_data"
    feed_name = "nvdcve-2.0-2024.json.zip"
    feed_url = f"https://nvd.nist.gov/feeds/json/cve/2.0/{feed_name}"
    feed_page = f'<a href="{feed_url}">2024</a>'.encode()
    responses = iter([io.BytesIO(feed_page), io.BytesIO(b"not-a-zip")])

    monkeypatch.setattr(download_nvd, "TARGET_DIR", target_dir)
    monkeypatch.setattr(
        download_nvd,
        "urlopen",
        lambda *args, **kwargs: next(responses),
    )

    assert download_nvd.download_and_extract_feeds() is False
    assert not (target_dir / feed_name).exists()
    assert list(target_dir.glob("*.part")) == []


def test_load_vulnerabilities_empty_and_invalid(tmp_path):
    empty_feed = tmp_path / "nvdcve-2.0-2024.json"
    empty_feed.write_text('{"vulnerabilities": []}', encoding="utf-8")
    assert load_vulnerabilities(empty_feed) == []

    missing_key = tmp_path / "nvdcve-2.0-2023.json"
    missing_key.write_text("{}", encoding="utf-8")
    assert load_vulnerabilities(missing_key) == []

    bad_type = tmp_path / "nvdcve-2.0-2022.json"
    bad_type.write_text('{"vulnerabilities": {}}', encoding="utf-8")
    assert load_vulnerabilities(bad_type) == []

    corrupt = tmp_path / "nvdcve-2.0-2021.json"
    corrupt.write_text("{not-json", encoding="utf-8")
    assert load_vulnerabilities(corrupt) == []


def test_build_workbook_empty_dir_and_empty_feed(tmp_path):
    empty_dir = tmp_path / "empty_data"
    empty_dir.mkdir()
    out = tmp_path / "empty.xlsx"
    build_workbook(data_dir=empty_dir, output_file=out)
    assert out.is_file()
    wb = load_workbook(out)
    try:
        assert wb.sheetnames == ["INDEX"]
        assert wb["INDEX"].cell(1, 1).value == "CVE ID"
    finally:
        wb.close()

    data_dir = tmp_path / "feeds"
    data_dir.mkdir()
    (data_dir / "nvdcve-2.0-2024.json").write_text(
        json.dumps(
            {
                "vulnerabilities": [
                    {},
                    {"cve": None},
                    {"cve": {"id": "CVE-2024-42", "descriptions": [], "metrics": {}}},
                ]
            }
        ),
        encoding="utf-8",
    )
    out2 = tmp_path / "partial.xlsx"
    build_workbook(data_dir=data_dir, output_file=out2)
    wb2 = load_workbook(out2)
    try:
        assert "INDEX" in wb2.sheetnames
        assert "2024" in wb2.sheetnames
        # three rows attempted (header + 3 vulns including blanks)
        assert wb2["2024"].max_row == 4
        assert wb2["2024"].cell(4, 1).value == "CVE-2024-42"
    finally:
        wb2.close()


def test_build_workbook_keeps_formula_like_imports_as_literal_text(tmp_path):
    data_dir = tmp_path / "feeds"
    data_dir.mkdir()
    (data_dir / "nvdcve-2.0-2024.json").write_text(
        json.dumps(
            {
                "vulnerabilities": [
                    {
                        "cve": {
                            "id": "@CVE-2024-42",
                            "descriptions": [
                                {
                                    "lang": "en",
                                    "value": '=HYPERLINK("https://evil.example")',
                                }
                            ],
                            "published": "+unsafe",
                            "lastModified": "-unsafe",
                            "metrics": {
                                "cvssMetricV31": [
                                    {
                                        "type": "Primary",
                                        "cvssData": {
                                            "baseScore": "=1+1",
                                            "baseSeverity": "@HIGH",
                                            "vectorString": "+VECTOR",
                                        },
                                    }
                                ]
                            },
                        }
                    }
                ]
            }
        ),
        encoding="utf-8",
    )
    output = tmp_path / "formula-safe.xlsx"
    build_workbook(data_dir=data_dir, output_file=output)

    wb = load_workbook(output, data_only=False)
    try:
        row = wb["2024"][2]
        assert [cell.data_type for cell in row] == ["s"] * 7
        assert row[1].value == '=HYPERLINK("https://evil.example")'
        assert wb["INDEX"].cell(2, 5).data_type == "s"
        assert wb["INDEX"].cell(2, 5).value == '=HYPERLINK("https://evil.example")'
    finally:
        wb.close()

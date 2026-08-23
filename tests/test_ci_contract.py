from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def test_ci_uses_the_reviewed_hash_locked_dependency_graph():
    workflow = (ROOT / ".github" / "workflows" / "ci.yml").read_text()
    lockfile = (ROOT / "requirements-ci.txt").read_text()
    dev_requirements = (ROOT / "requirements-dev.txt").read_text()

    assert "python -m pip install --require-hashes -r requirements-ci.txt" in workflow
    assert "pip install --upgrade pip" not in workflow
    assert "-r requirements.txt" not in workflow
    assert "--hash=sha256:" in lockfile
    assert "-r requirements-ci.txt" in dev_requirements
    for package in ("beautifulsoup4", "openpyxl", "pip-audit", "pytest"):
        assert f"{package}==" in lockfile


def test_readme_documents_the_operational_contract():
    readme = (ROOT / "README.md").read_text()

    for heading in (
        "## Architecture",
        "## Quick Start",
        "## Runtime configuration and output paths",
        "## Verification",
        "## Security",
        "## Deployment",
        "## Troubleshooting",
    ):
        assert heading in readme

    assert "credential-free batch utility" in readme
    assert "python /path/to/NIST_NVD_CVE_2.0_Convertor_JSON_To_XSLX/run_pipeline.py" in readme
    assert "NIST_CVE_Compiled.xlsx" in readme

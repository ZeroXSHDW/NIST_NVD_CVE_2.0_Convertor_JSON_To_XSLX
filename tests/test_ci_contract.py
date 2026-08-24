from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def test_ci_uses_the_reviewed_hash_locked_dependency_graph():
    workflow = (ROOT / ".github" / "workflows" / "ci.yml").read_text()
    lockfile = (ROOT / "requirements-ci.txt").read_text()
    dev_requirements = (ROOT / "requirements-dev.txt").read_text()

    assert "python -m pip install --require-hashes -r requirements-ci.txt" in workflow
    assert "pip install --upgrade pip" not in workflow
    assert "-r requirements.txt" not in workflow
    assert "git diff --check" in workflow
    assert "ubuntu-latest" not in workflow
    assert "runs-on: ubuntu-24.04" in workflow
    assert "--hash=sha256:" in lockfile
    assert "-r requirements-ci.txt" in dev_requirements
    for package in ("beautifulsoup4", "openpyxl", "pip-audit", "pytest"):
        assert f"{package}==" in lockfile

    checkout = workflow.index("actions/checkout@")
    hygiene = workflow.index("name: Check patch hygiene")
    setup_python = workflow.index("actions/setup-python@")
    assert checkout < hygiene < setup_python
    assert workflow.count("git diff --check") == 1


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
    assert "fail-closed" in readme
    assert "annual-sheet and `INDEX` row counts" in readme
    assert "git diff --check" in readme


def test_readme_states_the_current_review_posture_without_a_production_overclaim():
    readme = (ROOT / "README.md").read_text()

    assert "## Current release status" in readme
    assert "codex/workflow-permissions" in readme
    assert "PR #3" in readme
    assert "25 fixture-backed tests" in readme
    assert "not a production data-publication approval" in readme
    assert "five medium Dependabot alerts" in readme

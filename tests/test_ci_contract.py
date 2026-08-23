from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def test_ci_uses_the_reviewed_hash_locked_dependency_graph():
    workflow = (ROOT / ".github" / "workflows" / "ci.yml").read_text()
    lockfile = (ROOT / "requirements-ci.txt").read_text()

    assert "python -m pip install --require-hashes -r requirements-ci.txt" in workflow
    assert "pip install --upgrade pip" not in workflow
    assert "-r requirements.txt" not in workflow
    assert "--hash=sha256:" in lockfile
    for package in ("beautifulsoup4", "openpyxl", "pip-audit", "pytest"):
        assert f"{package}==" in lockfile

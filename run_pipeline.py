"""Run the NVD download, conversion, and validation stages in order."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
STAGES: tuple[tuple[str, str], ...] = (
    ("download_nvd.py", "Downloading and Extracting NIST Data Feeds"),
    ("json_to_xlsx.py", "Converting JSON Feeds to Master XLSX"),
    ("validate_xlsx.py", "Validating XLSX Data Integrity"),
)


def run_command(
    script_path: Path,
    description: str,
    *,
    project_root: Path = PROJECT_ROOT,
) -> bool:
    """Run one pipeline stage with the repository as its working directory."""
    print(f"\n{'=' * 60}")
    print(f"STEP: {description}")
    print(f"{'=' * 60}")

    try:
        subprocess.run(
            [sys.executable, str(script_path)],
            cwd=project_root,
            check=True,
        )
    except subprocess.CalledProcessError as error:
        print(f"\n❌ FAILED: {description} (Exit code: {error.returncode})")
        return False
    except OSError as error:
        print(f"\n❌ ERROR: Could not start {description}: {error}")
        return False
    return True


def main() -> int:
    """Run every pipeline stage and return a process exit status."""
    for script_name, description in STAGES:
        script_path = PROJECT_ROOT / script_name
        if not script_path.is_file():
            print(f"❌ ERROR: Could not find {script_name} in {PROJECT_ROOT}")
            return 1

        if not run_command(script_path, description):
            print(f"\n🛑 Pipeline halted due to error in {script_name}.")
            return 1

    print(f"\n{'=' * 60}")
    print("✅ FULL PIPELINE COMPLETED SUCCESSFULLY!")
    print(f"{'=' * 60}")
    print("Final Output: NIST_CVE_Compiled.xlsx is ready for use.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

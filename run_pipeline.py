"""
NIST CVE Data Pipeline Orchestrator
------------------------------------
A master controller script that sequentially runs the download,
conversion, and validation stages of the NIST CVE Data Repository.
"""

import sys
import subprocess
from pathlib import Path

def run_command(command, description):
    """Utility to run a command and report status."""
    print(f"\n{'='*60}")
    print(f"STEP: {description}")
    print(f"{'='*60}")
    
    try:
        # We use sys.executable to ensure we use the same python interpreter
        subprocess.run([sys.executable, command], check=True)
        return True
    except subprocess.CalledProcessError as e:
        print(f"\n❌ FAILED: {description} (Exit code: {e.returncode})")
        return False
    except Exception as e:
        print(f"\n❌ ERROR: Unexpected error during {description}: {e}")
        return False

def main():
    """
    Main orchestrator for the NIST CVE Data Pipeline.
    1. Download/Extract
    2. Convert to XLSX
    3. Validate
    """
    project_root = Path(__file__).parent
    
    steps = [
        ("download_nvd.py", "Downloading and Extracting NIST Data Feeds"),
        ("json_to_xlsx.py", "Converting JSON Feeds to Master XLSX"),
        ("validate_xlsx.py", "Validating XLSX Data Integrity")
    ]
    
    for script, description in steps:
        script_path = project_root / script
        if not script_path.exists():
            print(f"❌ ERROR: Could not find {script} in {project_root}")
            sys.exit(1)
            
        success = run_command(script, description)
        if not success:
            print(f"\n🛑 Pipeline halted due to error in {script}.")
            sys.exit(1)

    print(f"\n{'='*60}")
    print("✅ FULL PIPELINE COMPLETED SUCCESSFULLY!")
    print(f"{'='*60}")
    print("Final Output: NIST_CVE_Compiled.xlsx is ready for use.")

if __name__ == "__main__":
    main()

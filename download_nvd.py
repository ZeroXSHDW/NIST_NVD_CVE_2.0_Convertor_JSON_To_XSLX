"""
NVD CVE 2.0 Data Feed Downloader
--------------------------------
Scrapes the NIST NVD data feeds page for CVE 2.0 JSON ZIP links,
downloads them, and extracts the JSON files for processing.

Features:
- Skip historical yearly files if already present.
- Always fetch 'modified' and 'recent' feeds for latest updates.
- Robust error handling with a descriptive User-Agent.
"""

from bs4 import BeautifulSoup
import zipfile
from pathlib import Path
from urllib.request import Request, urlopen

# --- Configuration ---
BASE_URL = "https://nvd.nist.gov"
FEEDS_URL = "https://nvd.nist.gov/vuln/data-feeds"
TARGET_DIR = Path(__file__).parent / "nvd_data"
USER_AGENT = (
    "NIST-NVD-CVE-2.0-Converter/1.0 "
    "(+https://github.com/ZeroXSHDW/NIST_NVD_CVE_2.0_Convertor_JSON_To_XSLX; "
    "research/data-pipeline)"
)


def collect_feed_links(html: str, base_url: str = BASE_URL) -> list[str]:
    """Parse feed page HTML and return unique CVE 2.0 JSON ZIP URLs."""
    if not html or not html.strip():
        return []

    soup = BeautifulSoup(html, "html.parser")
    links: list[str] = []
    for link in soup.find_all("a", href=True):
        href = link["href"]
        if "nvdcve-2.0-" in href and href.endswith(".json.zip"):
            if not href.startswith("http"):
                href = f"{base_url.rstrip('/')}/{href.lstrip('/')}"
            links.append(href)
    return sorted(set(links))


def extract_zip_safely(zip_path: Path, dest_dir: Path) -> bool:
    """
    Extract a ZIP if it contains at least one member.
    Returns False for missing, empty, or corrupt archives.
    """
    if not zip_path.is_file() or zip_path.stat().st_size == 0:
        return False
    try:
        with zipfile.ZipFile(zip_path, "r") as z:
            names = z.namelist()
            if not names:
                return False
            destination = dest_dir.resolve()
            for member in z.infolist():
                member_path = (dest_dir / member.filename).resolve()
                try:
                    member_path.relative_to(destination)
                except ValueError:
                    return False
            z.extractall(dest_dir)
        return True
    except (zipfile.BadZipFile, OSError):
        return False


def download_and_extract_feeds():
    """
    Downloads and extracts all CVE 2.0 JSON ZIP feeds from the NVD website.
    Always fetches the latest 'modified' and 'recent' files.
    Skips historical yearly files if they already exist locally to save bandwidth.
    """
    if not TARGET_DIR.exists():
        print(f"Creating directory: {TARGET_DIR}")
        TARGET_DIR.mkdir(parents=True, exist_ok=True)

    print(f"Fetching feeds list from {FEEDS_URL}...")
    headers = {
        "User-Agent": USER_AGENT,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    }
    try:
        request = Request(FEEDS_URL, headers=headers)
        with urlopen(request, timeout=30) as response:
            feeds_html = response.read().decode("utf-8", errors="replace")
    except Exception as e:
        print(f"Error fetching the feeds page: {e}")
        return False

    links = collect_feed_links(feeds_html)
    if not links:
        print("No CVE 2.0 JSON ZIP feeds found on the page.")
        return False

    print(f"Found {len(links)} feed ZIPs. Starting download/extraction...")

    success_count = 0
    for url in links:
        zip_name = url.split("/")[-1]
        local_zip_path = TARGET_DIR / zip_name
        json_name = zip_name.replace(".zip", "")
        local_json_path = TARGET_DIR / json_name

        # Logic from user: Always re-download "modified" and "recent".
        # Skip yearly files if both ZIP and JSON exist.
        is_dynamic = "modified" in zip_name or "recent" in zip_name
        if not is_dynamic and local_zip_path.exists() and local_json_path.exists():
            print(f"✓ Skipping {zip_name} (historical file already exists)")
            success_count += 1
            continue

        print(f"↓ Downloading {zip_name}...", end=" ", flush=True)
        try:
            request = Request(url, headers=headers)
            with urlopen(request, timeout=60) as response, open(
                local_zip_path, "wb"
            ) as f:
                while chunk := response.read(8192):
                    f.write(chunk)

            if not extract_zip_safely(local_zip_path, TARGET_DIR):
                print("Failed! Error: empty or corrupt ZIP")
                continue

            print("Done.")
            success_count += 1
        except Exception as e:
            print(f"Failed! Error: {e}")

    print(f"\nSummary: {success_count}/{len(links)} feeds processed successfully.")
    print(f"Data directory: {TARGET_DIR.absolute()}")
    return success_count == len(links)


if __name__ == "__main__":
    download_and_extract_feeds()

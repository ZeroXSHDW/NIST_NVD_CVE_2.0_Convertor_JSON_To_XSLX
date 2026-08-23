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

import os
import re
import tempfile
import zipfile
from pathlib import Path
from urllib.parse import urljoin, urlparse
from urllib.request import Request, urlopen

from bs4 import BeautifulSoup

# --- Configuration ---
BASE_URL = "https://nvd.nist.gov"
FEEDS_URL = "https://nvd.nist.gov/vuln/data-feeds"
TARGET_DIR = Path(__file__).parent / "nvd_data"
USER_AGENT = (
    "NIST-NVD-CVE-2.0-Converter/1.0 "
    "(+https://github.com/ZeroXSHDW/NIST_NVD_CVE_2.0_Convertor_JSON_To_XSLX; "
    "research/data-pipeline)"
)
FEED_FILENAME = re.compile(r"^nvdcve-2\.0-(?:\d{4}|modified|recent)\.json\.zip$")


def collect_feed_links(html: str, base_url: str = BASE_URL) -> list[str]:
    """Parse the NVD feed page and return only approved HTTPS feed URLs."""
    if not html or not html.strip():
        return []

    base = urlparse(base_url)
    allowed_host = base.netloc.lower()
    if base.scheme != "https" or not allowed_host:
        return []

    soup = BeautifulSoup(html, "html.parser")
    links: list[str] = []
    for link in soup.find_all("a", href=True):
        href = link["href"]
        if not isinstance(href, str):
            continue
        candidate = urljoin(f"{base_url.rstrip('/')}/", href)
        parsed = urlparse(candidate)
        filename = Path(parsed.path).name
        if (
            parsed.scheme == "https"
            and parsed.netloc.lower() == allowed_host
            and not parsed.params
            and not parsed.query
            and not parsed.fragment
            and FEED_FILENAME.fullmatch(filename)
        ):
            links.append(candidate)
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


def download_response_atomically(response, destination: Path) -> None:
    """Write a downloaded response without exposing a partial final file."""

    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{destination.name}.",
        suffix=".part",
        dir=destination.parent,
    )
    try:
        with os.fdopen(descriptor, "wb") as output:
            while chunk := response.read(8192):
                output.write(chunk)
            output.flush()
            os.fsync(output.fileno())
        os.replace(temporary_name, destination)
    except Exception:
        try:
            os.unlink(temporary_name)
        except FileNotFoundError:
            pass
        raise


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
            with urlopen(request, timeout=60) as response:
                download_response_atomically(response, local_zip_path)

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

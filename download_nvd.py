"""
NVD CVE 2.0 Data Feed Downloader
--------------------------------
Scrapes the NIST NVD data feeds page for CVE 2.0 JSON ZIP links,
downloads them, and extracts the JSON files for processing.

Features:
- Skip historical yearly files if already present.
- Always fetch 'modified' and 'recent' feeds for latest updates.
- Robust error handling and User-Agent spoofing for WAF bypass.
"""

import requests
from bs4 import BeautifulSoup
import zipfile
from pathlib import Path

# --- Configuration ---
BASE_URL = "https://nvd.nist.gov"
FEEDS_URL = "https://nvd.nist.gov/vuln/data-feeds"
TARGET_DIR = Path(__file__).parent / "nvd_data"

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
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    try:
        response = requests.get(FEEDS_URL, headers=headers, timeout=30)
        response.raise_for_status()
    except Exception as e:
        print(f"Error fetching the feeds page: {e}")
        return False
    
    soup = BeautifulSoup(response.text, 'html.parser')
    
    # We are looking for links that match nvdcve-2.0-.*\.json\.zip
    links = []
    for link in soup.find_all('a', href=True):
        href = link['href']
        if "nvdcve-2.0-" in href and href.endswith(".json.zip"):
            if not href.startswith("http"):
                href = f"{BASE_URL.rstrip('/')}/{href.lstrip('/')}"
            links.append(href)
    
    links = sorted(list(set(links)))
    
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
            r = requests.get(url, headers=headers, stream=True, timeout=60)
            r.raise_for_status()
            
            # Save ZIP to disk (useful for the "exists" check next time)
            with open(local_zip_path, "wb") as f:
                for chunk in r.iter_content(chunk_size=8192):
                    f.write(chunk)
            
            # Extract
            with zipfile.ZipFile(local_zip_path, "r") as z:
                z.extractall(TARGET_DIR)
                
            print("Done.")
            success_count += 1
        except Exception as e:
            print(f"Failed! Error: {e}")

    print(f"\nSummary: {success_count}/{len(links)} feeds processed successfully.")
    print(f"Data directory: {TARGET_DIR.absolute()}")
    return success_count == len(links)

if __name__ == "__main__":
    download_and_extract_feeds()

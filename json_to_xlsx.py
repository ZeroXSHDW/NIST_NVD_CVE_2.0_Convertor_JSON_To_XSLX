"""
NVD CVE JSON to XLSX Converter
------------------------------
Processes a collection of NVD CVE 2.0 JSON files and compiles them into
a single, highly-structured Excel workbook with annual sheets and a master index.
"""

import json
import re
from pathlib import Path
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# --- Configuration ---
DATA_DIR = Path(__file__).parent / "nvd_data"
OUTPUT_FILE = Path(__file__).parent / 'NIST_CVE_Compiled.xlsx'

# Define columns we want in the annual sheets
COLUMNS = [
    ('cve_id', 'CVE ID'),
    ('description', 'Description (EN)'),
    ('published', 'Published Date'),
    ('lastModified', 'Last Modified Date'),
    ('baseScore', 'Base Score'),
    ('severity', 'Severity'),
    ('vectorString', 'CVSS Vector'),
]

# Columns for the Master INDEX sheet
INDEX_COLUMNS = [
    ('cve_id', 'CVE ID'),
    ('year', 'Year/Sheet'),
    ('severity', 'Severity'),
    ('baseScore', 'Base Score'),
    ('summary', 'Summary (Short Description)'),
]

FORMULA_PREFIXES = ('=', '+', '-', '@')


def write_cell(cell, value):
    """Write imported text as a literal cell, never an executable formula."""
    cell.value = value
    if isinstance(value, str) and value.startswith(FORMULA_PREFIXES):
        cell.data_type = 's'


def clean_text(text):
    if not isinstance(text, str):
        return str(text or '')
    return text.replace('\r\n', '\n').replace('\n\r', '\n').replace('\r', '\n')


def extract_entry(cve_entry):
    if not isinstance(cve_entry, dict):
        cve_entry = {}

    cve_id = str(cve_entry.get('id') or '')
    description = ''
    descriptions = cve_entry.get('descriptions') or []
    if isinstance(descriptions, list):
        for desc in descriptions:
            if not isinstance(desc, dict):
                continue
            if desc.get('lang') == 'en':
                description = clean_text(desc.get('value', ''))
                break
    published = str(cve_entry.get('published') or '')
    last_modified = str(cve_entry.get('lastModified') or '')

    base_score = ''
    severity = ''
    vector = ''
    metrics = cve_entry.get('metrics') or {}
    if not isinstance(metrics, dict):
        metrics = {}

    metric_keys = ['cvssMetricV40', 'cvssMetricV31', 'cvssMetricV30', 'cvssMetricV2']
    found_metric = None
    for key in metric_keys:
        for m in metrics.get(key) or []:
            if not isinstance(m, dict):
                continue
            if m.get('type') == 'Primary':
                found_metric = (key, m)
                break
        if found_metric:
            break

    if not found_metric:
        for key in metric_keys:
            for m in metrics.get(key) or []:
                if not isinstance(m, dict):
                    continue
                found_metric = (key, m)
                break
            if found_metric:
                break

    if found_metric:
        key, m = found_metric
        cvss_data = m.get('cvssData') or {}
        if not isinstance(cvss_data, dict):
            cvss_data = {}
        ds = cvss_data.get('baseScore', '')
        base_score = ds if ds is not None else ''
        severity = cvss_data.get('baseSeverity', '') or m.get('baseSeverity', '') or m.get('severity', '') or ''
        vector = cvss_data.get('vectorString', '') or ''

    return {
        'cve_id': cve_id,
        'description': description,
        'published': published,
        'lastModified': last_modified,
        'baseScore': base_score,
        'severity': str(severity),
        'vectorString': str(vector),
    }


def load_vulnerabilities(json_path: Path) -> list:
    """Load a feed JSON and return its vulnerabilities list (empty on edge cases)."""
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except (OSError, json.JSONDecodeError) as exc:
        print(f"Warning: skipping unreadable feed {json_path.name}: {exc}")
        return []

    if not isinstance(data, dict):
        print(f"Warning: skipping non-object feed {json_path.name}")
        return []

    vulns = data.get('vulnerabilities')
    if vulns is None:
        return []
    if not isinstance(vulns, list):
        print(f"Warning: skipping feed with non-list vulnerabilities: {json_path.name}")
        return []
    return vulns


def build_workbook(data_dir: Path = DATA_DIR, output_file: Path = OUTPUT_FILE) -> Path:
    """
    Build the compiled XLSX from yearly NVD JSON feeds under data_dir.
    Empty feeds and missing fields are tolerated; INDEX is always created.
    """
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    master_index = []
    data_dir = Path(data_dir)
    if not data_dir.is_dir():
        print(f"Warning: data directory missing: {data_dir}")
        json_files = []
    else:
        json_files = sorted(data_dir.glob('nvdcve-2.0-*.json'))

    if not json_files:
        print("No yearly NVD JSON feeds found; writing INDEX-only workbook.")

    for json_path in json_files:
        match = re.search(r'nvdcve-2\.0-(\d{4})\.json', json_path.name)
        if not match:
            continue
        year = match.group(1)
        print(f"Processing {year}...")
        ws = wb.create_sheet(title=year)
        for col_idx, (_, header) in enumerate(COLUMNS, start=1):
            write_cell(ws.cell(row=1, column=col_idx), header)

        vulns = load_vulnerabilities(json_path)
        row_num = 2
        for vuln in vulns:
            if not isinstance(vuln, dict):
                continue
            cve_entry = vuln.get('cve')
            extracted = extract_entry(cve_entry if isinstance(cve_entry, dict) else {})
            for col_idx, (key, _) in enumerate(COLUMNS, start=1):
                write_cell(ws.cell(row=row_num, column=col_idx), extracted.get(key, ''))

            summary = extracted['description'][:200] + (
                '...' if len(extracted['description']) > 200 else ''
            )
            master_index.append({
                'cve_id': extracted['cve_id'],
                'year': year,
                'severity': extracted['severity'],
                'baseScore': extracted['baseScore'],
                'summary': summary,
            })
            row_num += 1

        for col_idx, (key, header) in enumerate(COLUMNS, start=1):
            ws.column_dimensions[get_column_letter(col_idx)].width = max(len(header), 15)

    print("Creating Master INDEX sheet...")
    idx_ws = wb.create_sheet(title='INDEX', index=0)
    for col_idx, (_, header) in enumerate(INDEX_COLUMNS, start=1):
        write_cell(idx_ws.cell(row=1, column=col_idx), header)

    master_index.sort(key=lambda x: x['cve_id'], reverse=True)

    for row_num, entry in enumerate(master_index, start=2):
        for col_idx, (key, _) in enumerate(INDEX_COLUMNS, start=1):
            write_cell(idx_ws.cell(row=row_num, column=col_idx), entry.get(key, ''))

    for col_idx, (key, header) in enumerate(INDEX_COLUMNS, start=1):
        idx_ws.column_dimensions[get_column_letter(col_idx)].width = max(len(header), 15)

    output_file = Path(output_file)
    wb.save(output_file)
    print(f'Workbook saved with INDEX to {output_file}')
    return output_file


def main():
    build_workbook()


if __name__ == '__main__':
    main()

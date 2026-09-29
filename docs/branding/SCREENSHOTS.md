# Branding screenshots — NIST NVD CVE 2.0 Convertor (JSON → XLSX)

Assets here come from a **real converter run**, not a mocked UI.

| File | What it shows | How produced |
| --- | --- | --- |
| `sample-NIST_CVE_Compiled.xlsx` | Workbook with `INDEX` + `2024` sheets | `json_to_xlsx.build_workbook()` on `tests/fixtures/sample_cve.json` staged as `nvdcve-2.0-2024.json` |
| `hero-sample-xlsx.jpg` | Preview of the INDEX sheet rows | Rendered from the real openpyxl workbook cells |

Brand site: https://ZeroDevLLC.com  ·  Store: https://zerodevllc.store

Development twin: https://github.com/ZeroXSHDW/NIST_NVD_CVE_2.0_Convertor_JSON_To_XSLX-dev

# z/OS Product Compatibility

Streamlit app: pick a z/OS version (3.2 / 3.1 / 2.5) and see which vendor products are compatible,
which need **PTFs / usermods** (shown in **yellow**, with the IDs), EOS dates and latest versions.
Download the result as Excel. Written answers (Tab 3) use a local Granite/Llama model through Ollama
(optional, no API key, nothing leaves your machine).

## PTF scraper (tab 5)

Tab 5 reads PTF tables from every vendor link (up to 10 pages per source, login optional) and exports one Excel sheet per vendor. Pages behind a vendor login show "Access denied" unless you enter credentials. See `ptf_scraper.py`.

## Quick start

Windows: double-click `run.bat`   |   Mac/Linux: `./run.sh`

or manually:
    python -m venv venv
    .\venv\Scripts\activate
    pip install -r requirements.txt
    streamlit run app.py -> http://localhost:8501

## What you get

| Tab | Purpose |
|---|---|
| 1 z/OS Version | Filter the whole matrix. PTF cell is yellow when PTFs/usermods are required. Download workbook. |
| 2 Check my products | Pick products + your VRM (or upload a CSV/XLSX with Vendor, Product, VRM). Verdict + PTF worklist. |
| 3 Ask | Plain-English questions (works without AI; shows matching rows). |
| 4 Sources | Official vendor links. |

Under each table, **PTF / usermod worklist** lists every product that needs fixes, the extracted IDs
(e.g. `LU16693`), a copy-ready ID list, a `.txt` download and a starter SMP/E RECEIVE/APPLY CHECK template
(a template only: review zone names and options before running). The Excel export has the same yellow
cells plus a **PTF Worklist** sheet.

## Updating the data (the easy way: through Excel)

1. Download a workbook from the app, or use your own copy.
2. Edit the status / PTF-usermod / EOS / latest-version cells in Excel (keep the header row).
3. Import it back:

        python tools/import_xlsx.py my_corrected.xlsx --dry-run     (preview)
        python tools/import_xlsx.py my_corrected.xlsx               (apply to data/*.csv)

4. Restart the app. Rows are matched on Vendor + Product + VRM; empty cells never overwrite data.

Any cell that starts with `Yes` in "If Yes, PTF Required?" turns yellow. Put the IDs after it, e.g.
`Yes - LU16693, LU18696 (FIXCAT ...)`. IDs like `LU05157/58/59` are expanded automatically.

## Optional: written answers with Granite or Llama

    ollama pull granite3.3:8b          (or llama3:latest)
    set OLLAMA_MODEL=llama3:latest       (optional; default granite3.3:8b)
    set OLLAMA_HOST=http://localhost:11434

If Ollama is not running, Tab 3 simply shows the matching rows.

## Tests

    python tests/test_project.py

## Data notes

* z/OS 3.2 PTF/usermod numbers come from `tests/reference_zOS_3_2_filtered.xlsx` (240 rows, 29 need PTFs).
* z/OS 3.1 PTF cells mostly say "see PTF list on Broadcom grid" (no numbers yet); 2.5 is entirely Unverified.
  Import a 3.1 or 2.5 workbook with the tool above to fill them in.
* Always confirm on the live vendor page before production changes.

# Kalani Model Automation

This repo automates the "blue cell" input workflow for the A.CRE Hotel Development model.

## Prerequisites

- Python 3.10+
- `pip install -r requirements.txt` (only needs `openpyxl` & `PyYAML`)

## Files

- `rebuild_kalani.py` – CLI script that writes inputs into the template.
- `kalani_config.yaml` – Human-readable assumptions (grouped by tab/section).
- `A.CRE-Hotel-Development-Model-beta-v1.57.xlsx` – Original template (keep in repo root).

## Usage (Command-line)

1. **Edit `kalani_config.yaml`:**
   - `summary`: Summary tab inputs (property info, exit cap, debt settings, etc.).
   - `budget`: Acquisition/soft/hard/FF&E/financing buckets.
   - `cell_inputs`: Explicit blue-cell overrides for tabs like `OpCashFlow` and `MonthlyCF`.
   - `row_ranges`: Row-wide fills (e.g., `MonthlyCF!P206:GN206`).

2. **Run the script:**
   ```bash
   python rebuild_kalani.py --template A.CRE-Hotel-Development-Model-beta-v1.57.xlsx
   ```
   Optional flags:
   - `--config path/to/yaml`
   - `--template path/to/template.xlsx`
   - `--output path/to/output.xlsx`

3. **Open the output workbook (`updated_kalani_model_v2.xlsx`)** in Excel to recalc/inspect.

## Notes

- The script only writes to cells listed in the YAML; all formulas remain intact.
- To add new inputs, drop them under the appropriate section (e.g., `cell_inputs.OpCashFlow`) using Excel coordinates.
- For whole-row data (like Monthly cash-flow rates), use the `row_ranges` block to avoid hundreds of manual entries.

## Web UI (non-technical workflow)

1. Install dependencies (`pip install -r requirements.txt`).
2. Run the web app:
   ```bash
   # optional: USE_HTTPS=1 python app.py  (enables an auto-generated TLS cert)
   python app.py
   ```
3. Visit http://127.0.0.1:5000/ in a browser.
4. Pick the template type from the dropdown and click **Load Template Inputs** to swap in the matching YAML blueprint (e.g., A.CRE vs. UNDERWRITING v11).
5. Upload the corresponding `.xlsx`, review/edit the visible input fields, and click **Generate Workbook**.
6. After the success banner appears, click **Download Latest Workbook** (top-right button) to fetch the XLSX.

The UI mirrors every blue cell, so the client never has to edit YAML by hand (though the raw file is still available under `kalani_config.yaml` if needed).

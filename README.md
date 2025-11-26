
## Licence Plate Time Aggregation

Small, cross-platform Python utility that scans Excel exports with licence plate events (`01 ENTRY` / `02 EXIT`), pairs the events per plate, and aggregates how long each vehicle stayed at the parking lot. Results are exported to a new Excel workbook for reporting.

### Requirements

- Python 3.9+ (works on macOS and Windows)
- Python packages listed in `requirements.txt`
  ```bash
  python3 -m venv .venv          # optional but recommended
  source .venv/bin/activate      # Windows: .venv\Scripts\activate
  pip install -r requirements.txt
  ```

> `pandas` depends on `openpyxl` (for `.xlsx`) and `xlrd` (for legacy `.xls`). They are included in `requirements.txt`; no other external tools are needed.

### Configure

Update the fields inside `config.json`:
    - `source_folder`: folder containing your camera exports (e.g. `/Users/mgolc/Documents/Kameros - masinu numeriai`).
    - `output_file`: where the aggregated Excel file will be stored (relative paths are resolved from the repo root).
    - `timestamp_format`: optional format string if the timestamp column is plain text.
    - `columns`: rename if your Excel files use different headers for plate/event/timestamp.
    - `entry_marker` / `exit_marker`: text used in the event column (compared in uppercase).
    - `recursive`: set `true` to include Excel files in sub-folders.

### Run

**macOS / Linux**
```bash
cd /Users/mgolc/repos/LicensePlateReporting
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python aggregate.py --config config.json
```

**Windows (PowerShell)**
```powershell
cd C:\Users\mgolc\repos\LicensePlateReporting
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python aggregate.py --config config.json
```

- Use `--source-folder`, `--output-file`, or `--timestamp-format` to override values without editing the JSON.
- Add `--recursive` if you want to scan sub-folders even when `recursive` is `false` in the config.

### Output

The script scans every `.xls` / `.xlsx` file in the source folder, standardises the columns, and creates three sheets in the destination workbook:

- `intervals`: every matched ENTRY→EXIT pair with entry time, exit time, and duration (minutes).
- `monthly_totals`: visits, total minutes, and hours per plate per calendar month.
- `issues` (only when needed): unmatched entries/exits or inconsistent timestamps that you might want to fix manually.

If no spreadsheets are found or the expected columns/timestamps are missing, the script prints a warning and exits with a non-zero status so you can adjust the configuration.


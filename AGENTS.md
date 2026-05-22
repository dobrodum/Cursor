# AGENTS.md

## Cursor Cloud specific instructions

This repository contains standalone Python CLI scripts for sales/business data analysis. There is no web server, database, or service infrastructure.

### Repository structure

The `main` branch contains only `README.md`. All code lives on feature branches:

- `cursor/parameter-selection-bc51` — `parameter_search.py` and `sales_capture_quarter_sensitivity.py`
- `cursor/fix-xlwings-script-c576` — `rolling_average_excel_model.py`

### Scripts overview

| Script | Dependencies | Linux support |
|---|---|---|
| `parameter_search.py` | Python stdlib only | Yes |
| `sales_capture_quarter_sensitivity.py` | stdlib + optional `openpyxl` | Yes |
| `rolling_average_excel_model.py` | `xlwings` + desktop Excel | **No** (macOS/Windows only) |

### Running scripts

Each script is a standalone CLI tool invoked with `python3 <script> --help`. To test on the feature branches:

```bash
git checkout origin/cursor/parameter-selection-bc51 -- parameter_search.py sales_capture_quarter_sensitivity.py
python3 parameter_search.py --data <csv> --target-col <col> --candidate-cols <cols>
python3 sales_capture_quarter_sensitivity.py --data <csv> --quarter-col <col> --db-sales-col <col> --total-sales-col <col>
```

### Important caveats

- `rolling_average_excel_model.py` will raise `RuntimeError` on Linux — it requires a macOS/Windows machine with Microsoft Excel installed. On Cloud Agent VMs, you can only syntax-check it (`python3 -c "import py_compile; py_compile.compile('rolling_average_excel_model.py')"`) and run `--help`.
- There is no `requirements.txt`, `pyproject.toml`, or formal dependency management. The update script installs `openpyxl` via pip.
- There are no automated tests, linting configuration, or build steps in this repository.

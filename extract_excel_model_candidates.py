#!/usr/bin/env python3
"""Extract empirical and regression candidates from Excel model files.

This script:
1) Opens each source workbook once.
2) Processes both "Empirical Model" and "Regression Model" sheets while open.
3) Closes the workbook without saving any source changes.
4) Writes one consolidated output workbook with:
   - empirical_candidates
   - regression_candidates
"""

from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# =========================
# Configure these folders
# =========================
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")


EMPIRICAL_COLUMNS = [
    "model",
    "ticker",
    "model_period",
    "model_date",
    "method",
    "parameter_name",
    "parameter_value",
    "num_quarters_used",
    "last_quarter_used",
    "forecast_value",
    "actual_value",
    "forecast_max",
    "forecast_min",
    "range_width",
    "avg_penetration_pct",
    "quarterly_sales",
    "reported_sales",
    "growth_rate_pct",
    "sales_captured_in_db_pct",
    "source_file",
]

REGRESSION_COLUMNS = [
    "model",
    "ticker",
    "model_period",
    "model_date",
    "method",
    "parameter_name",
    "parameter_value",
    "num_quarters_used",
    "forecast_value",
    "actual_value",
    "forecast_max",
    "forecast_min",
    "range_width",
    "intercept",
    "slope",
    "source_file",
]

MONTH_TO_NUM = {
    "jan": 1,
    "feb": 2,
    "mar": 3,
    "apr": 4,
    "may": 5,
    "jun": 6,
    "jul": 7,
    "aug": 8,
    "sep": 9,
    "oct": 10,
    "nov": 11,
    "dec": 12,
}

PERIOD_DAY = {"early": 5, "mid": 15, "late": 25}

N_QUARTERS = 10


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def is_number(value: Any) -> bool:
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return True
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if text.endswith("%"):
            text = text[:-1]
        if not text:
            return False
        try:
            float(text)
            return True
        except ValueError:
            return False
    return False


def to_float(value: Any) -> Optional[float]:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        pct = text.endswith("%")
        if pct:
            text = text[:-1]
        if not text:
            return None
        try:
            num = float(text)
            return num / 100.0 if pct else num
        except ValueError:
            return None
    return None


def safe_div(numerator: Any, denominator: Any) -> Optional[float]:
    n = to_float(numerator)
    d = to_float(denominator)
    if n is None or d in (None, 0):
        return None
    return n / d


def safe_sub(a: Any, b: Any) -> Optional[float]:
    av = to_float(a)
    bv = to_float(b)
    if av is None or bv is None:
        return None
    return av - bv


def values_close(a: Any, b: Any, tol: float = 1e-10) -> bool:
    af = to_float(a)
    bf = to_float(b)
    if af is None and bf is None:
        return True
    if af is None or bf is None:
        return False
    return abs(af - bf) <= tol


def set_formula2(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        # Fallback for Excel versions without Formula2 support.
        cell.formula = formula


def ensure_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def get_sheet_case_insensitive(book: xw.Book, name: str) -> Optional[xw.Sheet]:
    target = name.strip().lower()
    for sheet in book.sheets:
        if str(sheet.name).strip().lower() == target:
            return sheet
    return None


def find_anchor_cell(sheet: xw.Sheet, text: str = "max") -> Optional[Tuple[int, int]]:
    # Fast path via Excel Find.
    try:
        found = sheet.api.Cells.Find(What=text, LookAt=1, MatchCase=False)
        if found is not None:
            return int(found.Row), int(found.Column)
    except Exception:
        pass

    # Fallback scan of used range.
    try:
        used = sheet.used_range
        values = ensure_2d(used.value)
        start_row = int(used.row)
        start_col = int(used.column)
        target = normalize_label(text)
        for r_idx, row_vals in enumerate(values):
            for c_idx, v in enumerate(row_vals):
                if normalize_label(v) == target:
                    return start_row + r_idx, start_col + c_idx
    except Exception:
        pass

    return None


def header_map_for_row(
    sheet: xw.Sheet, header_row: int, start_col: int, end_col: int
) -> Dict[str, int]:
    if end_col < start_col:
        return {}
    row_values = sheet.range((header_row, start_col), (header_row, end_col)).value
    if not isinstance(row_values, list):
        row_values = [row_values]
    mapping: Dict[str, int] = {}
    for idx, raw in enumerate(row_values):
        norm = normalize_label(raw)
        if norm:
            mapping[norm] = start_col + idx
    return mapping


def find_col_by_keywords(
    normalized_to_col: Dict[str, int],
    keywords: Sequence[str],
    fallback: Optional[int] = None,
) -> Optional[int]:
    for key in keywords:
        key_norm = normalize_label(key)
        for label, col in normalized_to_col.items():
            if key_norm in label:
                return col
    return fallback


def parse_model_fields(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    parts = [p.strip() for p in stem.split(" - ")]

    ticker = ""
    period_raw = ""
    if len(parts) >= 3:
        ticker = parts[1].strip()
        period_raw = parts[2].split("_")[0].strip()
    elif len(parts) >= 2:
        ticker = parts[1].strip()

    # Fallback ticker if the expected pattern is not present.
    if not ticker:
        ticker = re.sub(r"\W+", "", stem).upper()[:12]

    # Expected pattern: EarlyJan2026 / MidJan2026 / LateJan2026
    model_period = ""
    model_date = ""
    m = re.match(r"^(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(\d{4})$", period_raw, re.IGNORECASE)
    if m:
        phase = m.group(1).title()
        month_text = m.group(2)
        year = int(m.group(3))
        month_abbrev = month_text[:3].lower()
        month_num = MONTH_TO_NUM.get(month_abbrev)
        if month_num is not None:
            model_period = f"{phase}{month_text[:3].title()}_{year}"
            model_day = PERIOD_DAY[phase.lower()]
            model_date = date(year, month_num, model_day).isoformat()

    # Fallback period/date if parse fails.
    if not model_period:
        model_period = period_raw or "unknown_period"
    if not model_date:
        model_date = ""

    model = f"{ticker}_{model_period}"

    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def safe_close_book(book: xw.Book) -> None:
    try:
        book.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        book.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        book.close()
    except Exception:
        pass


def normalize_range_values(values: Any) -> List[Any]:
    if values is None:
        return []
    if isinstance(values, list):
        return values
    return [values]


def process_empirical_sheet(
    book: xw.Book, file_fields: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    sheet = get_sheet_case_insensitive(book, "Empirical Model")
    if sheet is None:
        print(f"  - Empirical Model sheet missing in {source_file}; skipped empirical.")
        return []

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"  - Could not find 'max' anchor in Empirical Model ({source_file}); skipped empirical.")
        return []

    anchor_row, anchor_col = anchor
    header_start = max(1, anchor_col - 40)
    header_end = anchor_col + 40
    headers = header_map_for_row(sheet, anchor_row, header_start, header_end)

    max_col = anchor_col
    min_col = find_col_by_keywords(headers, ["min"], fallback=anchor_col + 1)
    num_quarters_col = find_col_by_keywords(
        headers, ["num_quarters_used", "num quarters", "quarters used", "qtrs used", "n quarters"]
    )
    last_quarter_col = find_col_by_keywords(headers, ["last_quarter_used", "last quarter"])
    forecast_col = find_col_by_keywords(headers, ["estimated total sold", "forecast"])
    reported_sales_col = find_col_by_keywords(headers, ["reported sales", "actual sales", "actual value"])
    quarterly_sales_col = find_col_by_keywords(headers, ["quarterly sales", "quarter sales"])
    growth_rate_col = find_col_by_keywords(headers, ["growth rate"])
    captured_pct_col = find_col_by_keywords(headers, ["sales captured in db", "captured in db", "captured pct"])

    data_start = anchor_row + 1
    candidate_rows = [data_start + i for i in range(N_QUARTERS)]

    # Average penetration formulas are written to temporary scratch cells in this sheet.
    scratch_col = max(anchor_col + 60, 200)
    scratch_start_row = 2
    avg_formula_rows: List[int] = []
    for i, row_idx in enumerate(candidate_rows):
        formula_cell = sheet.range((scratch_start_row + i, scratch_col))
        if captured_pct_col is not None:
            formula = (
                f'=IFERROR(AVERAGE(R{data_start}C{captured_pct_col}:'
                f'R{row_idx}C{captured_pct_col}),"")'
            )
            set_formula2(formula_cell, formula)
            avg_formula_rows.append(i)
        elif quarterly_sales_col is not None and reported_sales_col is not None:
            formula = (
                f'=IFERROR(AVERAGE('
                f'R{data_start}C{quarterly_sales_col}:R{row_idx}C{quarterly_sales_col}/'
                f'R{data_start}C{reported_sales_col}:R{row_idx}C{reported_sales_col}'
                f'),"")'
            )
            set_formula2(formula_cell, formula)
            avg_formula_rows.append(i)

    if avg_formula_rows:
        book.app.calculate()

    avg_values = normalize_range_values(
        sheet.range(
            (scratch_start_row, scratch_col),
            (scratch_start_row + N_QUARTERS - 1, scratch_col),
        ).value
    )

    rows: List[Dict[str, Any]] = []
    for i, row_idx in enumerate(candidate_rows, start=1):
        num_quarters_used = (
            sheet.range((row_idx, num_quarters_col)).value
            if num_quarters_col is not None
            else i
        )
        last_quarter_used = (
            sheet.range((row_idx, last_quarter_col)).value
            if last_quarter_col is not None
            else None
        )
        forecast_value = (
            sheet.range((row_idx, forecast_col)).value
            if forecast_col is not None
            else None
        )
        forecast_max = sheet.range((row_idx, max_col)).value
        forecast_min = (
            sheet.range((row_idx, min_col)).value if min_col is not None else None
        )
        reported_sales = (
            sheet.range((row_idx, reported_sales_col)).value
            if reported_sales_col is not None
            else None
        )
        quarterly_sales = (
            sheet.range((row_idx, quarterly_sales_col)).value
            if quarterly_sales_col is not None
            else None
        )
        growth_rate_pct = (
            sheet.range((row_idx, growth_rate_col)).value
            if growth_rate_col is not None
            else None
        )
        sales_captured_in_db_pct = (
            sheet.range((row_idx, captured_pct_col)).value
            if captured_pct_col is not None
            else safe_div(quarterly_sales, reported_sales)
        )
        avg_penetration_pct = avg_values[i - 1] if i - 1 < len(avg_values) else None

        # Fallback forecast estimate if no direct column is available.
        if forecast_value is None:
            q = to_float(quarterly_sales)
            avg_p = to_float(avg_penetration_pct)
            if q is not None and avg_p not in (None, 0):
                forecast_value = q / avg_p

        actual_value = reported_sales
        range_width = safe_sub(forecast_max, forecast_min)

        row_has_payload = any(
            value not in (None, "")
            for value in (
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                quarterly_sales,
                reported_sales,
                avg_penetration_pct,
            )
        )
        if not row_has_payload:
            continue

        row = {
            "model": file_fields["model"],
            "ticker": file_fields["ticker"],
            "model_period": file_fields["model_period"],
            "model_date": file_fields["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration_pct,
            "num_quarters_used": num_quarters_used,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "avg_penetration_pct": avg_penetration_pct,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate_pct,
            "sales_captured_in_db_pct": sales_captured_in_db_pct,
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def process_regression_sheet(
    book: xw.Book, file_fields: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    sheet = get_sheet_case_insensitive(book, "Regression Model")
    if sheet is None:
        print(f"  - Regression Model sheet missing in {source_file}; skipped regression.")
        return []

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"  - Could not find 'max' anchor in Regression Model ({source_file}); skipped regression.")
        return []

    anchor_row, anchor_col = anchor
    header_start = max(1, anchor_col - 40)
    header_end = anchor_col + 40
    headers = header_map_for_row(sheet, anchor_row, header_start, header_end)

    max_col = anchor_col
    min_col = find_col_by_keywords(headers, ["min"], fallback=anchor_col + 1)
    forecast_col = find_col_by_keywords(
        headers,
        [
            "tot fcst w o sa",
            "tot fcst wo sa",
            "tot fcst without sa",
            "total forecast without sa",
            "tot fcst",
        ],
    )
    actual_col = find_col_by_keywords(headers, ["actual value", "actual sales", "reported sales"])
    num_quarters_col = find_col_by_keywords(
        headers, ["num_quarters_used", "num quarters", "quarters used", "qtrs used", "n quarters"]
    )

    # Explicit requirement for regression anchor offsets.
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if x_col < 1 or y_col < 1:
        print(f"  - Invalid x/y offsets in Regression Model ({source_file}); skipped regression.")
        return []

    used_last_row = int(sheet.used_range.last_cell.row)
    source_last_row = min(anchor_row - 1, used_last_row)
    if source_last_row < 1:
        return []

    x_values = normalize_range_values(sheet.range((1, x_col), (source_last_row, x_col)).value)
    y_values = normalize_range_values(sheet.range((1, y_col), (source_last_row, y_col)).value)

    valid_rows: List[int] = []
    for idx, (xv, yv) in enumerate(zip(x_values, y_values), start=1):
        if is_number(xv) and is_number(yv):
            valid_rows.append(idx)

    if not valid_rows:
        print(f"  - No valid regression source rows found in {source_file}.")
        return []

    max_n = min(N_QUARTERS, len(valid_rows))
    scratch_col_x = max(anchor_col + 60, 200)
    scratch_col_y = scratch_col_x + 1
    scratch_col_formula = scratch_col_x + 3
    scratch_base_row = max(anchor_row + 20, 40)

    intercept_cells: Dict[int, xw.Range] = {}
    slope_cells: Dict[int, xw.Range] = {}
    latest_x_by_n: Dict[int, Any] = {}

    for n in range(1, max_n + 1):
        rows_used = valid_rows[-n:]
        block_row = scratch_base_row + (n - 1) * 20
        x_block = [[x_values[r - 1]] for r in rows_used]
        y_block = [[y_values[r - 1]] for r in rows_used]
        latest_x_by_n[n] = x_block[-1][0] if x_block else None

        sheet.range((block_row, scratch_col_x), (block_row + n - 1, scratch_col_x)).value = x_block
        sheet.range((block_row, scratch_col_y), (block_row + n - 1, scratch_col_y)).value = y_block

        intercept_cell = sheet.range((block_row, scratch_col_formula))
        slope_cell = sheet.range((block_row + 1, scratch_col_formula))
        set_formula2(
            intercept_cell,
            (
                f'=IFERROR(INTERCEPT('
                f'R{block_row}C{scratch_col_y}:R{block_row + n - 1}C{scratch_col_y},'
                f'R{block_row}C{scratch_col_x}:R{block_row + n - 1}C{scratch_col_x}'
                f'),"")'
            ),
        )
        set_formula2(
            slope_cell,
            (
                f'=IFERROR(SLOPE('
                f'R{block_row}C{scratch_col_y}:R{block_row + n - 1}C{scratch_col_y},'
                f'R{block_row}C{scratch_col_x}:R{block_row + n - 1}C{scratch_col_x}'
                f'),"")'
            ),
        )
        intercept_cells[n] = intercept_cell
        slope_cells[n] = slope_cell

    if intercept_cells:
        book.app.calculate()

    rows: List[Dict[str, Any]] = []
    for n in range(1, max_n + 1):
        row_idx = anchor_row + n
        num_quarters_used = (
            sheet.range((row_idx, num_quarters_col)).value
            if num_quarters_col is not None
            else n
        )
        intercept = intercept_cells[n].value
        slope = slope_cells[n].value

        forecast_value = (
            sheet.range((row_idx, forecast_col)).value
            if forecast_col is not None
            else None
        )
        if forecast_value is None and is_number(intercept) and is_number(slope):
            latest_x = to_float(latest_x_by_n.get(n))
            if latest_x is not None:
                forecast_value = to_float(intercept) + to_float(slope) * latest_x

        actual_value = (
            sheet.range((row_idx, actual_col)).value if actual_col is not None else None
        )
        forecast_max = sheet.range((row_idx, max_col)).value
        forecast_min = (
            sheet.range((row_idx, min_col)).value if min_col is not None else None
        )
        range_width = safe_sub(forecast_max, forecast_min)

        row = {
            "model": file_fields["model"],
            "ticker": file_fields["ticker"],
            "model_period": file_fields["model_period"],
            "model_date": file_fields["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        # Prevent duplicate final row by comparing calculated values to previous row.
        if rows:
            prev = rows[-1]
            duplicate = (
                values_close(prev["intercept"], row["intercept"])
                and values_close(prev["slope"], row["slope"])
                and values_close(prev["forecast_value"], row["forecast_value"])
                and values_close(prev["forecast_max"], row["forecast_max"])
                and values_close(prev["forecast_min"], row["forecast_min"])
            )
            if duplicate:
                continue

        payload = (
            row["forecast_value"],
            row["forecast_max"],
            row["forecast_min"],
            row["intercept"],
            row["slope"],
        )
        if all(v in (None, "") for v in payload):
            continue

        rows.append(row)

    return rows


def next_output_path(output_folder: Path, input_folder_name: str) -> Path:
    base = output_folder / f"{input_folder_name}_PARAM.xlsx"
    if not base.exists():
        return base

    i = 1
    while True:
        candidate = output_folder / f"{input_folder_name}_PARAM.{i}.xlsx"
        if not candidate.exists():
            return candidate
        i += 1


def write_sheet(ws, columns: List[str], rows: List[Dict[str, Any]]) -> None:
    ws.append(columns)
    for row in rows:
        ws.append([row.get(col) for col in columns])

    # Format header row.
    header_font = Font(bold=True)
    for cell in ws[1]:
        cell.font = header_font

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # Set reasonable widths based on content.
    for col_idx, header in enumerate(columns, start=1):
        max_len = len(header)
        for cell in ws[get_column_letter(col_idx)]:
            if cell.value is None:
                continue
            max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 40)


def write_output_workbook(
    output_path: Path, empirical_rows: List[Dict[str, Any]], regression_rows: List[Dict[str, Any]]
) -> None:
    wb = Workbook()
    default_ws = wb.active
    wb.remove(default_ws)

    empirical_ws = wb.create_sheet("empirical_candidates")
    regression_ws = wb.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def iter_input_files(folder: Path) -> List[Path]:
    entries = sorted(folder.iterdir(), key=lambda p: p.name.lower())
    selected: List[Path] = []
    for path in entries:
        if not path.is_file():
            continue
        if path.name.startswith("~"):
            print(f"Skipped {path.name}: temp file prefix '~'.")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"Skipped {path.name}: not an .xlsx file.")
            continue
        selected.append(path)
    return selected


def main() -> None:
    in_dir = Path(input_dir)
    out_dir = Path(output_dir)

    if not in_dir.exists():
        raise FileNotFoundError(f"input_dir does not exist: {in_dir}")

    out_dir.mkdir(parents=True, exist_ok=True)
    output_path = next_output_path(out_dir, in_dir.name)
    files = iter_input_files(in_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app: Optional[xw.App] = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        for file_path in files:
            print(f"Processing {file_path.name} ...")
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                fields = parse_model_fields(file_path.name)

                empirical_rows.extend(process_empirical_sheet(wb, fields, file_path.name))
                regression_rows.extend(process_regression_sheet(wb, fields, file_path.name))

                files_processed += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error ({exc}).")
            finally:
                if wb is not None:
                    safe_close_book(wb)
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output written: {output_path}")
    print(f"Files processed: {files_processed}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

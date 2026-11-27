from __future__ import annotations

import math
import re
from datetime import datetime
from pathlib import Path
from statistics import pstdev
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# ---------- user-configurable paths ----------
input_dir = Path("/path/to/input_folder")
output_dir = Path("/path/to/output_folder")


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

N_QUARTERS = 10

# Anchor-based column offsets from the "max" cell in each model sheet.
# These preserve an offset-driven strategy and can be adjusted if the workbook layout changes.
EMPIRICAL_OFFSETS = {
    "quarter_col": -12,
    "quarterly_sales_col": -10,
    "reported_sales_col": -9,
    "penetration_col": -8,
    "growth_rate_col": -7,
    "sales_captured_col": -6,
}

DAY_MAP = {"Early": 5, "Mid": 15, "Late": 25}


def is_number(value: Any) -> bool:
    if not isinstance(value, (int, float)) or isinstance(value, bool):
        return False
    try:
        return not math.isnan(value)
    except TypeError:
        return True


def as_float(value: Any) -> float | None:
    if value is None:
        return None
    if is_number(value):
        return float(value)
    try:
        parsed = float(str(value).strip())
        return parsed if not math.isnan(parsed) else None
    except (TypeError, ValueError):
        return None


def normalize_pct(value: float | None) -> float | None:
    if value is None:
        return None
    if abs(value) > 1.5:
        return value / 100.0
    return value


def to_list(values: Any) -> list[Any]:
    if isinstance(values, list):
        return values
    if isinstance(values, tuple):
        return list(values)
    return [values]


def column_values(sheet: xw.Sheet, col: int, start_row: int, end_row: int) -> list[Any]:
    if end_row < start_row:
        return []
    values = sheet.range((start_row, col), (end_row, col)).value
    return to_list(values)


def safe_get(sheet: xw.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    try:
        return sheet.range((row, col)).value
    except Exception:
        return None


def set_formula2(target_range: xw.Range, formula: str) -> None:
    # formula2 is preferred for modern Excel dynamic/formula behavior.
    try:
        target_range.formula2 = formula
    except Exception:
        target_range.formula = formula


def find_anchor(sheet: xw.Sheet, anchor_text: str = "max") -> tuple[int, int]:
    used = sheet.used_range
    start_row = used.row
    start_col = used.column
    row_count = used.rows.count
    col_count = used.columns.count
    values = used.value

    if row_count == 1 and col_count == 1:
        cell_value = str(values).strip().lower() if values is not None else ""
        if cell_value == anchor_text:
            return start_row, start_col
        raise ValueError(f"anchor '{anchor_text}' not found")

    if row_count == 1:
        matrix = [to_list(values)]
    elif col_count == 1:
        matrix = [[v] for v in to_list(values)]
    else:
        matrix = [to_list(row) for row in to_list(values)]

    target = anchor_text.strip().lower()
    for r_idx, row in enumerate(matrix):
        for c_idx, cell_value in enumerate(row):
            if isinstance(cell_value, str) and cell_value.strip().lower() == target:
                return start_row + r_idx, start_col + c_idx

    raise ValueError(f"anchor '{anchor_text}' not found")


def parse_model_labels(file_path: Path) -> dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = ""
    if len(parts) >= 2 and parts[1]:
        ticker = parts[1].upper()
    else:
        ticker_match = re.search(r"\b[A-Z]{2,8}\b", stem)
        ticker = ticker_match.group(0) if ticker_match else "UNKNOWN"

    period_match = re.search(r"(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})", stem, flags=re.IGNORECASE)
    if period_match:
        period_bucket = period_match.group(1).title()
        month_token = period_match.group(2).title()
        year = int(period_match.group(3))
        month_short = month_token[:3].title()
        try:
            month_num = datetime.strptime(month_short, "%b").month
        except ValueError:
            month_num = 1
        day = DAY_MAP.get(period_bucket, 15)
        model_period = f"{period_bucket}{month_short}_{year}"
        model_date = f"{year:04d}-{month_num:02d}-{day:02d}"
    else:
        model_period = "unknown_period"
        model_date = ""

    model = f"{ticker}_{model_period}" if ticker and model_period else stem
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def latest_output_path(input_folder: Path, output_folder: Path) -> Path:
    output_folder.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_folder.name}_PARAM"
    candidate = output_folder / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = output_folder / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def safe_close_workbook(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        # Fall through to COM/API close fallback.
        pass

    try:
        workbook.api.Close(SaveChanges=False)
    except Exception:
        workbook.close()


def numeric_row_window(sheet: xw.Sheet, col: int, anchor_row: int) -> tuple[int, int]:
    values = column_values(sheet, col, 1, max(1, anchor_row - 1))
    first = None
    last = None
    for idx, raw in enumerate(values, start=1):
        if as_float(raw) is not None:
            if first is None:
                first = idx
            last = idx
    if first is None or last is None:
        raise ValueError(f"no numeric values found in column {col}")
    return first, last


def extract_empirical_rows(workbook: xw.Book, labels: dict[str, str], source_file: str) -> list[dict[str, Any]]:
    sheet = workbook.sheets["Empirical Model"]
    anchor_row, anchor_col = find_anchor(sheet, "max")

    penetration_col = anchor_col + EMPIRICAL_OFFSETS["penetration_col"]
    quarterly_sales_col = anchor_col + EMPIRICAL_OFFSETS["quarterly_sales_col"]
    reported_sales_col = anchor_col + EMPIRICAL_OFFSETS["reported_sales_col"]
    growth_rate_col = anchor_col + EMPIRICAL_OFFSETS["growth_rate_col"]
    sales_captured_col = anchor_col + EMPIRICAL_OFFSETS["sales_captured_col"]
    quarter_col = anchor_col + EMPIRICAL_OFFSETS["quarter_col"]

    first_row, last_row = numeric_row_window(sheet, penetration_col, anchor_row)
    available = last_row - first_row + 1
    n_to_run = min(N_QUARTERS, available)
    if n_to_run <= 0:
        return []

    pen_values = [normalize_pct(as_float(v)) for v in column_values(sheet, penetration_col, first_row, last_row)]
    quarterly_values = [as_float(v) for v in column_values(sheet, quarterly_sales_col, first_row, last_row)]
    reported_values = [as_float(v) for v in column_values(sheet, reported_sales_col, first_row, last_row)]
    growth_values = [normalize_pct(as_float(v)) for v in column_values(sheet, growth_rate_col, first_row, last_row)]
    captured_values = [normalize_pct(as_float(v)) for v in column_values(sheet, sales_captured_col, first_row, last_row)]
    quarter_values = column_values(sheet, quarter_col, first_row, last_row)

    helper_start_row = anchor_row + 2
    helper_col = anchor_col + 2

    for n in range(1, n_to_run + 1):
        start_row = last_row - n + 1
        helper_cell = sheet.range((helper_start_row + n - 1, helper_col))
        avg_formula = f"=AVERAGE(R{start_row}C{penetration_col}:R{last_row}C{penetration_col})"
        set_formula2(helper_cell, avg_formula)

    workbook.app.calculate()
    avg_values = to_list(sheet.range((helper_start_row, helper_col), (helper_start_row + n_to_run - 1, helper_col)).value)

    rows: list[dict[str, Any]] = []
    for idx, avg_raw in enumerate(avg_values, start=1):
        n = idx
        start_idx = len(pen_values) - n

        avg_pen = normalize_pct(as_float(avg_raw))
        window_pen = [v for v in pen_values[start_idx:] if v is not None and v > 0]
        window_quarterly = [v for v in quarterly_values[start_idx:] if v is not None]
        window_reported = [v for v in reported_values[start_idx:] if v is not None]

        reported_sales = window_reported[-1] if window_reported else None
        quarterly_sales = window_quarterly[-1] if window_quarterly else None

        forecast_value = None
        forecast_max = None
        forecast_min = None

        if reported_sales is not None and avg_pen is not None and avg_pen > 0:
            forecast_value = reported_sales / avg_pen
            max_pen = max(window_pen) if window_pen else None
            min_pen = min(window_pen) if window_pen else None
            if min_pen and min_pen > 0:
                forecast_max = reported_sales / min_pen
            if max_pen and max_pen > 0:
                forecast_min = reported_sales / max_pen

        row_idx = len(pen_values) - 1
        row = {
            **labels,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_pen,
            "num_quarters_used": n,
            "last_quarter_used": quarter_values[row_idx] if quarter_values else "",
            "forecast_value": forecast_value,
            "actual_value": reported_sales,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": (forecast_max - forecast_min) if (forecast_max is not None and forecast_min is not None) else None,
            "avg_penetration_pct": avg_pen,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_values[row_idx] if growth_values else None,
            "sales_captured_in_db_pct": captured_values[row_idx] if captured_values else None,
            "source_file": source_file,
        }
        rows.append(row)

    # Remove temporary formulas to keep opened workbook clean before close (still closed without save).
    sheet.range((helper_start_row, helper_col), (helper_start_row + n_to_run - 1, helper_col)).clear_contents()
    return rows


def rows_approximately_equal(a: dict[str, Any], b: dict[str, Any], keys: list[str], tol: float = 1e-9) -> bool:
    for key in keys:
        av = a.get(key)
        bv = b.get(key)
        af = as_float(av)
        bf = as_float(bv)
        if af is not None and bf is not None:
            if abs(af - bf) > tol:
                return False
        else:
            if av != bv:
                return False
    return True


def extract_regression_rows(workbook: xw.Book, labels: dict[str, str], source_file: str) -> list[dict[str, Any]]:
    sheet = workbook.sheets["Regression Model"]
    anchor_row, anchor_col = find_anchor(sheet, "max")

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    x_values = [as_float(v) for v in column_values(sheet, x_col, 1, max(1, anchor_row - 1))]
    y_values = [as_float(v) for v in column_values(sheet, y_col, 1, max(1, anchor_row - 1))]

    valid_rows = [idx for idx, (xv, yv) in enumerate(zip(x_values, y_values), start=1) if xv is not None and yv is not None]
    if not valid_rows:
        return []

    first_row = valid_rows[0]
    last_row = valid_rows[-1]
    available = len(valid_rows)
    n_to_run = min(N_QUARTERS, available)
    if n_to_run <= 0:
        return []

    helper_start_row = anchor_row + 2
    intercept_col = anchor_col + 2
    slope_col = anchor_col + 3

    for n in range(1, n_to_run + 1):
        start_row = max(first_row, last_row - n + 1)
        y_range = f"R{start_row}C{y_col}:R{last_row}C{y_col}"
        x_range = f"R{start_row}C{x_col}:R{last_row}C{x_col}"
        intercept_formula = f"=INTERCEPT({y_range},{x_range})"
        slope_formula = f"=SLOPE({y_range},{x_range})"

        set_formula2(sheet.range((helper_start_row + n - 1, intercept_col)), intercept_formula)
        set_formula2(sheet.range((helper_start_row + n - 1, slope_col)), slope_formula)

    workbook.app.calculate()

    intercept_values = to_list(
        sheet.range((helper_start_row, intercept_col), (helper_start_row + n_to_run - 1, intercept_col)).value
    )
    slope_values = to_list(sheet.range((helper_start_row, slope_col), (helper_start_row + n_to_run - 1, slope_col)).value)

    x_future = as_float(safe_get(sheet, last_row + 1, x_col))
    if x_future is None:
        x_future = as_float(safe_get(sheet, last_row, x_col))

    y_future_actual = as_float(safe_get(sheet, last_row + 1, y_col))

    rows: list[dict[str, Any]] = []
    for idx, (intercept_raw, slope_raw) in enumerate(zip(intercept_values, slope_values), start=1):
        n = idx
        start_row = max(first_row, last_row - n + 1)
        intercept_val = as_float(intercept_raw)
        slope_val = as_float(slope_raw)

        forecast_value = None
        if intercept_val is not None and slope_val is not None and x_future is not None:
            forecast_value = intercept_val + (slope_val * x_future)

        residuals: list[float] = []
        if intercept_val is not None and slope_val is not None:
            for row_idx in range(start_row, last_row + 1):
                xv = x_values[row_idx - 1]
                yv = y_values[row_idx - 1]
                if xv is None or yv is None:
                    continue
                residuals.append(yv - (intercept_val + slope_val * xv))

        sigma = pstdev(residuals) if len(residuals) > 1 else 0.0
        forecast_max = (forecast_value + sigma) if forecast_value is not None else None
        forecast_min = (forecast_value - sigma) if forecast_value is not None else None

        row = {
            **labels,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n,
            "num_quarters_used": n,
            "forecast_value": forecast_value,
            "actual_value": y_future_actual if y_future_actual is not None else "",
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": (forecast_max - forecast_min) if (forecast_max is not None and forecast_min is not None) else None,
            "intercept": intercept_val,
            "slope": slope_val,
            "source_file": source_file,
        }

        if rows and rows_approximately_equal(
            rows[-1],
            row,
            keys=["forecast_value", "intercept", "slope", "forecast_max", "forecast_min"],
        ):
            continue
        rows.append(row)

    # Remove temporary formulas to keep opened workbook clean before close (still closed without save).
    sheet.range((helper_start_row, intercept_col), (helper_start_row + n_to_run - 1, slope_col)).clear_contents()
    return rows


def write_sheet(workbook: Workbook, sheet_name: str, columns: list[str], rows: list[dict[str, Any]]) -> None:
    worksheet = workbook.create_sheet(title=sheet_name)
    worksheet.append(columns)

    for row in rows:
        worksheet.append([row.get(col, "") for col in columns])

    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    worksheet.freeze_panes = "A2"
    last_col = get_column_letter(len(columns))
    last_row = max(1, worksheet.max_row)
    worksheet.auto_filter.ref = f"A1:{last_col}{last_row}"

    for col_idx, col_name in enumerate(columns, start=1):
        values = [col_name]
        values.extend(str(row.get(col_name, "")) for row in rows if row.get(col_name, "") is not None)
        width = max(len(v) for v in values) + 2
        worksheet.column_dimensions[get_column_letter(col_idx)].width = min(max(width, 12), 48)


def collect_input_files(input_folder: Path) -> list[Path]:
    files: list[Path] = []
    for path in sorted(input_folder.iterdir()):
        if not path.is_file():
            continue
        if path.name.startswith("~"):
            print(f"skipped {path.name}: temp file")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"skipped {path.name}: not an .xlsx file")
            continue
        files.append(path)
    return files


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"input_dir does not exist: {input_dir}")

    files = collect_input_files(input_dir)
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.calculation = "manual"

    try:
        for file_path in files:
            print(f"processing {file_path.name}")
            workbook = None
            try:
                # Required safe open mode for source workbook handling.
                workbook = app.books.open(str(file_path), update_links=False)
                labels = parse_model_labels(file_path)
                file_had_output = False

                try:
                    extracted = extract_empirical_rows(workbook, labels, file_path.name)
                    empirical_rows.extend(extracted)
                    file_had_output = file_had_output or bool(extracted)
                except KeyError:
                    print(f"skipped {file_path.name}: missing sheet 'Empirical Model'")
                except Exception as exc:
                    print(f"skipped {file_path.name}: empirical extraction failed ({exc})")

                try:
                    extracted = extract_regression_rows(workbook, labels, file_path.name)
                    regression_rows.extend(extracted)
                    file_had_output = file_had_output or bool(extracted)
                except KeyError:
                    print(f"skipped {file_path.name}: missing sheet 'Regression Model'")
                except Exception as exc:
                    print(f"skipped {file_path.name}: regression extraction failed ({exc})")

                if file_had_output:
                    processed_files += 1
                else:
                    print(f"skipped {file_path.name}: no extractable rows")
            except Exception as exc:
                print(f"skipped {file_path.name}: {exc}")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        app.quit()

    output_path = latest_output_path(input_dir, output_dir)
    output_book = Workbook()
    output_book.remove(output_book.active)
    write_sheet(output_book, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(output_book, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    output_book.save(output_path)

    print(f"output_path: {output_path}")
    print(f"files_processed: {processed_files}")
    print(f"empirical_rows: {len(empirical_rows)}")
    print(f"regression_rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

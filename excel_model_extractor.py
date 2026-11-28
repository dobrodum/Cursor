from __future__ import annotations

import calendar
import re
from datetime import date
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -------------------------------
# Configure inputs/outputs here.
# -------------------------------
input_dir = r"/path/to/input"
output_dir = r"/path/to/output"


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

DAY_MAP = {"early": 5, "mid": 15, "late": 25}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip().lower()


def to_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def safe_number(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def safe_int(value: Any) -> int | None:
    number = safe_number(value)
    if number is None:
        return None
    return int(number)


def safe_col(col: int | None, fallback: int = 1) -> int:
    if col is None or col < 1:
        return fallback
    return col


def get_used_matrix(sheet: xw.Sheet) -> tuple[int, int, int, int, list[list[Any]]]:
    used = sheet.used_range
    first_row = used.row
    first_col = used.column
    n_rows = used.rows.count
    n_cols = used.columns.count
    values = to_2d(used.value)
    return first_row, first_col, n_rows, n_cols, values


def matrix_value(
    values: list[list[Any]],
    first_row: int,
    first_col: int,
    row: int,
    col: int,
) -> Any:
    row_idx = row - first_row
    col_idx = col - first_col
    if row_idx < 0 or col_idx < 0:
        return None
    if row_idx >= len(values):
        return None
    if col_idx >= len(values[row_idx]):
        return None
    return values[row_idx][col_idx]


def find_max_anchor(
    values: list[list[Any]],
    first_row: int,
    first_col: int,
) -> tuple[int, int] | None:
    for r_idx, row_values in enumerate(values):
        for c_idx, value in enumerate(row_values):
            if normalize_text(value) == "max":
                return first_row + r_idx, first_col + c_idx
    return None


def find_col_by_keywords(
    header_values: list[Any],
    first_col: int,
    keyword_groups: list[tuple[str, ...]],
    fallback_col: int | None = None,
) -> int | None:
    for offset, raw_value in enumerate(header_values):
        text = normalize_text(raw_value)
        if not text:
            continue
        for keywords in keyword_groups:
            if all(keyword in text for keyword in keywords):
                return first_col + offset
    return fallback_col


def read_cell(sheet: xw.Sheet, row: int, col: int) -> Any:
    try:
        return sheet.range((row, col)).value
    except Exception:
        return None


def set_formula2(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.api.Formula2 = formula_r1c1
        return
    except Exception:
        pass
    try:
        cell.formula2 = formula_r1c1
        return
    except Exception:
        pass
    cell.formula = formula_r1c1


def safe_close_source_workbook(wb: xw.Book) -> None:
    # Required: close source workbooks without saving.
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.api.Close(False)
    except Exception:
        wb.close()


def parse_file_labels(file_name: str) -> dict[str, str]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split("-")]

    ticker = parts[1] if len(parts) > 1 and parts[1] else "UNKNOWN"
    raw_period = parts[2] if len(parts) > 2 else ""
    raw_period = raw_period.split("_")[0].strip()

    period_match = re.search(
        r"(?i)\b(early|mid|late)\s*([a-z]{3,9})\s*([0-9]{4})\b",
        raw_period,
    )

    if period_match:
        bucket = period_match.group(1).lower()
        month_token = period_match.group(2).lower()
        year = int(period_match.group(3))
        month_abbr = month_token[:3].title()
        month_lookup = {abbr.lower(): idx for idx, abbr in enumerate(calendar.month_abbr) if abbr}
        month_num = month_lookup.get(month_abbr.lower())
        if month_num is None:
            model_period = re.sub(r"\s+", "", raw_period) if raw_period else "Unknown"
            model_date = ""
        else:
            day = DAY_MAP[bucket]
            model_period = f"{bucket.title()}{month_abbr}_{year}"
            model_date = date(year, month_num, day).isoformat()
    else:
        model_period = re.sub(r"\s+", "", raw_period) if raw_period else "Unknown"
        model_date = ""

    model = f"{ticker}_{model_period}"
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def extract_empirical_candidates(
    wb: xw.Book,
    source_file: str,
    meta: dict[str, str],
) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        return []

    first_row, first_col, _, n_cols, values = get_used_matrix(sheet)
    anchor = find_max_anchor(values, first_row, first_col)
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    header_values = [
        matrix_value(values, first_row, first_col, anchor_row, col)
        for col in range(first_col, first_col + n_cols)
    ]

    max_col = safe_col(find_col_by_keywords(header_values, first_col, [("max",)], fallback_col=anchor_col))
    min_col = safe_col(find_col_by_keywords(header_values, first_col, [("min",)], fallback_col=anchor_col + 1))
    forecast_col = safe_col(find_col_by_keywords(
        header_values,
        first_col,
        [
            ("estimated", "total", "sold"),
            ("forecast",),
            ("tot", "fcst"),
        ],
        fallback_col=anchor_col - 1,
    ))
    actual_col = safe_col(find_col_by_keywords(
        header_values,
        first_col,
        [("reported", "sales"), ("actual",)],
        fallback_col=anchor_col + 2,
    ))
    quarterly_sales_col = safe_col(find_col_by_keywords(
        header_values,
        first_col,
        [("quarterly", "sales"), ("quarter", "sales")],
        fallback_col=anchor_col - 4,
    ))
    growth_col = safe_col(find_col_by_keywords(
        header_values,
        first_col,
        [("growth", "rate")],
        fallback_col=anchor_col - 3,
    ))
    captured_col = safe_col(find_col_by_keywords(
        header_values,
        first_col,
        [("captured", "db"), ("sales", "captured")],
        fallback_col=anchor_col - 2,
    ))
    last_quarter_col = safe_col(find_col_by_keywords(
        header_values,
        first_col,
        [("last", "quarter"), ("quarter", "used")],
        fallback_col=anchor_col - 5,
    ))
    penetration_col = safe_col(find_col_by_keywords(
        header_values,
        first_col,
        [("penetration",), ("pen", "%")],
        fallback_col=anchor_col - 6,
    ))

    temp_avg_col = first_col + n_cols + 2
    start_data_row = anchor_row + 1
    n_quarters = 10

    # Write all temporary R1C1 formulas first, then calculate once.
    for i in range(1, n_quarters + 1):
        row = start_data_row + (i - 1)
        avg_start_row = max(1, row - i + 1)
        formula = (
            f"=AVERAGE(R{avg_start_row}C{penetration_col}:R{row}C{penetration_col})"
        )
        set_formula2(sheet.range((row, temp_avg_col)), formula)

    wb.app.calculate()

    rows: list[dict[str, Any]] = []
    for i in range(1, n_quarters + 1):
        row = start_data_row + (i - 1)

        avg_penetration = read_cell(sheet, row, temp_avg_col)
        forecast_value = read_cell(sheet, row, forecast_col)
        actual_value = read_cell(sheet, row, actual_col)
        forecast_max = read_cell(sheet, row, max_col)
        forecast_min = read_cell(sheet, row, min_col)
        last_quarter_used = read_cell(sheet, row, last_quarter_col)
        quarterly_sales = read_cell(sheet, row, quarterly_sales_col)
        growth_rate = read_cell(sheet, row, growth_col)
        captured_pct = read_cell(sheet, row, captured_col)

        max_num = safe_number(forecast_max)
        min_num = safe_number(forecast_min)
        range_width = (max_num - min_num) if max_num is not None and min_num is not None else None

        # Skip empty rows when workbook table is shorter than n_quarters.
        if all(
            value in (None, "")
            for value in (
                avg_penetration,
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
            )
        ):
            continue

        rows.append(
            {
                "model": meta["model"],
                "ticker": meta["ticker"],
                "model_period": meta["model_period"],
                "model_date": meta["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": i,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": actual_value,
                "growth_rate_pct": growth_rate,
                "sales_captured_in_db_pct": captured_pct,
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_candidates(
    wb: xw.Book,
    source_file: str,
    meta: dict[str, str],
) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        return []

    first_row, first_col, _, n_cols, values = get_used_matrix(sheet)
    anchor = find_max_anchor(values, first_row, first_col)
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    y_col = safe_col(anchor_col - 7)
    x_col = safe_col(anchor_col - 11)

    header_values = [
        matrix_value(values, first_row, first_col, anchor_row, col)
        for col in range(first_col, first_col + n_cols)
    ]

    max_col = safe_col(find_col_by_keywords(header_values, first_col, [("max",)], fallback_col=anchor_col))
    min_col = safe_col(find_col_by_keywords(header_values, first_col, [("min",)], fallback_col=anchor_col + 1))
    forecast_col = safe_col(find_col_by_keywords(
        header_values,
        first_col,
        [("tot", "fcst", "w/o", "sa"), ("forecast",), ("tot", "fcst")],
        fallback_col=anchor_col - 1,
    ))
    actual_col = find_col_by_keywords(
        header_values,
        first_col,
        [("actual",), ("reported", "sales")],
        fallback_col=None,
    )

    n_quarters = 10
    start_data_row = anchor_row + 1
    temp_intercept_col = first_col + n_cols + 4
    temp_slope_col = temp_intercept_col + 1

    # Write all temporary R1C1 formulas first, then calculate once.
    for i in range(1, n_quarters + 1):
        row = start_data_row + (i - 1)
        start_row = max(1, row - i + 1)
        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{row}C{y_col},R{start_row}C{x_col}:R{row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{row}C{y_col},R{start_row}C{x_col}:R{row}C{x_col})"
        )
        set_formula2(sheet.range((row, temp_intercept_col)), intercept_formula)
        set_formula2(sheet.range((row, temp_slope_col)), slope_formula)

    wb.app.calculate()

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None

    for i in range(1, n_quarters + 1):
        row = start_data_row + (i - 1)
        intercept = read_cell(sheet, row, temp_intercept_col)
        slope = read_cell(sheet, row, temp_slope_col)
        forecast_value = read_cell(sheet, row, forecast_col)
        forecast_max = read_cell(sheet, row, max_col)
        forecast_min = read_cell(sheet, row, min_col)

        if forecast_value in (None, ""):
            next_x = safe_number(read_cell(sheet, row + 1, x_col))
            intercept_num = safe_number(intercept)
            slope_num = safe_number(slope)
            if (
                next_x is not None
                and intercept_num is not None
                and slope_num is not None
            ):
                forecast_value = intercept_num + slope_num * next_x

        actual_value = read_cell(sheet, row, actual_col) if actual_col is not None else None

        max_num = safe_number(forecast_max)
        min_num = safe_number(forecast_min)
        range_width = (max_num - min_num) if max_num is not None and min_num is not None else None

        if all(
            value in (None, "")
            for value in (intercept, slope, forecast_value, forecast_max, forecast_min)
        ):
            continue

        signature = (
            safe_number(intercept),
            safe_number(slope),
            safe_number(forecast_value),
            safe_number(forecast_max),
            safe_number(forecast_min),
        )
        if previous_signature is not None and signature == previous_signature:
            # Avoid duplicate final row when Excel model yields repeated results.
            continue

        rows.append(
            {
                "model": meta["model"],
                "ticker": meta["ticker"],
                "model_period": meta["model_period"],
                "model_date": meta["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": i,
                "num_quarters_used": i,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )
        previous_signature = signature

    return rows


def output_path_for_run(input_folder: Path, target_output_dir: Path) -> Path:
    base_name = f"{input_folder.name}_PARAM"
    candidate = target_output_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    index = 1
    while True:
        candidate = target_output_dir / f"{base_name}.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def write_sheet(ws, columns: list[str], rows: list[dict[str, Any]]) -> None:
    ws.append(columns)
    for row in rows:
        ws.append([row.get(col) for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 42)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    empirical_ws = wb.create_sheet("empirical_candidates")
    regression_ws = wb.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def main() -> None:
    source_dir = Path(input_dir).expanduser().resolve()
    target_dir = Path(output_dir).expanduser().resolve()
    target_dir.mkdir(parents=True, exist_ok=True)

    if not source_dir.exists() or not source_dir.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {source_dir}")

    output_path = output_path_for_run(source_dir, target_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.enable_events = False

    try:
        for file_path in sorted(source_dir.iterdir()):
            if not file_path.is_file():
                continue

            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temp file")
                continue

            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not .xlsx")
                continue

            wb: xw.Book | None = None
            try:
                meta = parse_file_labels(file_path.name)
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_candidates(wb, file_path.name, meta))
                regression_rows.extend(extract_regression_candidates(wb, file_path.name, meta))
                processed_files += 1
                print(f"Processed {file_path.name}")
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    safe_close_source_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

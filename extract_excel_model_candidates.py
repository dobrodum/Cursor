from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure these paths before running.
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")

N_QUARTERS = 10

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

MONTH_LOOKUP = {
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

PERIOD_DAY = {"Early": 5, "Mid": 15, "Late": 25}
PERIOD_PATTERN = re.compile(r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(\d{4})", re.IGNORECASE)


@dataclass
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


def to_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "")
        if not cleaned:
            return None
        pct = cleaned.endswith("%")
        if pct:
            cleaned = cleaned[:-1]
        try:
            parsed = float(cleaned)
            return parsed / 100.0 if pct else parsed
        except ValueError:
            return None
    return None


def normalize_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def normalize_1d(values: Any) -> list[Any]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [values]
    if values and isinstance(values[0], list):
        return [row[0] if row else None for row in values]
    return values


def parse_file_labels(file_name: str) -> FileLabels:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split("-")]
    ticker = "UNKNOWN"
    if len(parts) >= 2 and parts[1]:
        ticker = parts[1].upper()

    period_match = PERIOD_PATTERN.search(stem)
    if not period_match:
        model_period = "unknown_period"
        return FileLabels(
            model=f"{ticker}_{model_period}",
            ticker=ticker,
            model_period=model_period,
            model_date="",
        )

    period_word = period_match.group(1).title()
    month_token = period_match.group(2)[:3].lower()
    year = int(period_match.group(3))
    month_number = MONTH_LOOKUP.get(month_token)
    if month_number is None:
        model_period = f"{period_word}{period_match.group(2)}_{year}"
        model_date = ""
    else:
        month_label = month_token.title()
        model_period = f"{period_word}{month_label}_{year}"
        model_date = date(year, month_number, PERIOD_DAY[period_word]).isoformat()

    return FileLabels(
        model=f"{ticker}_{model_period}",
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def build_output_path(source_dir: Path, destination_dir: Path) -> Path:
    base_name = f"{source_dir.name}_PARAM"
    candidate = destination_dir / f"{base_name}.xlsx"
    counter = 1
    while candidate.exists():
        candidate = destination_dir / f"{base_name}.{counter}.xlsx"
        counter += 1
    return candidate


def find_label_positions(sheet: xw.Sheet, labels: set[str]) -> dict[str, tuple[int, int]]:
    used = sheet.used_range
    values = normalize_2d(used.value)
    positions: dict[str, tuple[int, int]] = {}
    for row_idx, row_values in enumerate(values):
        for col_idx, cell_value in enumerate(row_values):
            if not isinstance(cell_value, str):
                continue
            normalized = cell_value.strip().lower()
            if normalized in labels and normalized not in positions:
                positions[normalized] = (used.row + row_idx, used.column + col_idx)
    return positions


def get_numeric_near_label(sheet: xw.Sheet, row: int, col: int) -> float | None:
    for row_delta, col_delta in (
        (0, 1),
        (0, -1),
        (1, 0),
        (-1, 0),
        (1, 1),
        (-1, 1),
        (1, -1),
        (-1, -1),
    ):
        candidate = to_float(sheet.cells(row + row_delta, col + col_delta).value)
        if candidate is not None:
            return candidate
    return None


def find_contiguous_numeric_rows(
    sheet: xw.Sheet,
    value_col: int,
    anchor_row: int,
    max_scan_rows: int = 250,
) -> list[int]:
    end_scan = anchor_row - 1
    if end_scan < 1:
        return []
    start_scan = max(1, end_scan - max_scan_rows + 1)
    col_values = normalize_1d(sheet.range((start_scan, value_col), (end_scan, value_col)).value)

    last_numeric_index: int | None = None
    for idx in range(len(col_values) - 1, -1, -1):
        if to_float(col_values[idx]) is not None:
            last_numeric_index = idx
            break

    if last_numeric_index is None:
        return []

    first_numeric_index = last_numeric_index
    while first_numeric_index >= 0 and to_float(col_values[first_numeric_index]) is not None:
        first_numeric_index -= 1
    first_numeric_index += 1

    start_row = start_scan + first_numeric_index
    end_row = start_scan + last_numeric_index
    return list(range(start_row, end_row + 1))


def safe_subtract(left: float | None, right: float | None) -> float | None:
    if left is None or right is None:
        return None
    return left - right


def close_workbook_no_save(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.close(False)
        return
    except Exception:
        pass

    workbook.api.Close(SaveChanges=False)


def process_empirical_sheet(
    workbook: xw.Book,
    labels: FileLabels,
    source_file: str,
) -> list[dict[str, Any]]:
    if "Empirical Model" not in [sheet.name for sheet in workbook.sheets]:
        return []

    sheet = workbook.sheets["Empirical Model"]
    label_positions = find_label_positions(sheet, {"max", "min"})
    if "max" not in label_positions:
        return []

    anchor_row, anchor_col = label_positions["max"]
    max_value = get_numeric_near_label(sheet, anchor_row, anchor_col)
    min_value = (
        get_numeric_near_label(sheet, *label_positions["min"])
        if "min" in label_positions
        else None
    )

    quarterly_sales_col = anchor_col - 11
    reported_sales_col = anchor_col - 7
    quarter_label_col = max(1, quarterly_sales_col - 1)

    rows = find_contiguous_numeric_rows(sheet, reported_sales_col, anchor_row)
    if not rows:
        return []

    row_start = rows[0]
    row_end = rows[-1]
    quarter_labels = normalize_1d(
        sheet.range((row_start, quarter_label_col), (row_end, quarter_label_col)).value
    )
    quarterly_sales_values = [
        to_float(v)
        for v in normalize_1d(
            sheet.range((row_start, quarterly_sales_col), (row_end, quarterly_sales_col)).value
        )
    ]
    reported_sales_values = [
        to_float(v)
        for v in normalize_1d(
            sheet.range((row_start, reported_sales_col), (row_end, reported_sales_col)).value
        )
    ]

    helper_cell = sheet.cells(anchor_row + 2, anchor_col + 2)
    empirical_rows: list[dict[str, Any]] = []
    max_quarters = min(N_QUARTERS, len(rows))

    for num_quarters in range(1, max_quarters + 1):
        start_idx = len(rows) - num_quarters
        start_row = rows[start_idx]

        helper_cell.formula2 = (
            f'=IFERROR(AVERAGE(R{start_row}C{reported_sales_col}:R{row_end}C{reported_sales_col}'
            f'/R{start_row}C{quarterly_sales_col}:R{row_end}C{quarterly_sales_col}),"")'
        )
        workbook.app.calculate()
        avg_penetration_pct = to_float(helper_cell.value)

        subset_quarterly = quarterly_sales_values[start_idx:]
        subset_reported = reported_sales_values[start_idx:]
        valid_pairs = [
            (quarterly, reported)
            for quarterly, reported in zip(subset_quarterly, subset_reported)
            if quarterly not in (None, 0) and reported is not None
        ]
        if not valid_pairs:
            continue

        actual_value = subset_reported[-1]
        quarterly_sales = subset_quarterly[-1]
        if actual_value is None:
            continue

        forecast_value = None
        if avg_penetration_pct not in (None, 0):
            forecast_value = actual_value / avg_penetration_pct

        penetration_values = [reported / quarterly for quarterly, reported in valid_pairs]
        computed_forecast_max = None
        computed_forecast_min = None
        if penetration_values:
            pen_max = max(penetration_values)
            pen_min = min(penetration_values)
            if pen_min != 0:
                computed_forecast_max = actual_value / pen_min
            if pen_max != 0:
                computed_forecast_min = actual_value / pen_max

        forecast_max = max_value if max_value is not None else computed_forecast_max
        forecast_min = min_value if min_value is not None else computed_forecast_min
        range_width = safe_subtract(forecast_max, forecast_min)

        previous_actual = subset_reported[-2] if len(subset_reported) > 1 else None
        growth_rate_pct = None
        if previous_actual not in (None, 0) and actual_value is not None:
            growth_rate_pct = (actual_value - previous_actual) / previous_actual

        sales_captured_in_db_pct = None
        if forecast_value not in (None, 0):
            sales_captured_in_db_pct = actual_value / forecast_value

        last_quarter_used = quarter_labels[start_idx]
        if last_quarter_used in (None, ""):
            last_quarter_used = f"row_{start_row}"

        empirical_rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": actual_value,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    return empirical_rows


def process_regression_sheet(
    workbook: xw.Book,
    labels: FileLabels,
    source_file: str,
) -> list[dict[str, Any]]:
    if "Regression Model" not in [sheet.name for sheet in workbook.sheets]:
        return []

    sheet = workbook.sheets["Regression Model"]
    label_positions = find_label_positions(sheet, {"max", "min"})
    if "max" not in label_positions:
        return []

    anchor_row, anchor_col = label_positions["max"]
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    max_value = get_numeric_near_label(sheet, anchor_row, anchor_col)
    min_value = (
        get_numeric_near_label(sheet, *label_positions["min"])
        if "min" in label_positions
        else None
    )

    rows = find_contiguous_numeric_rows(sheet, y_col, anchor_row)
    if len(rows) < 2:
        return []

    row_start = rows[0]
    row_end = rows[-1]
    x_values = [
        to_float(v)
        for v in normalize_1d(sheet.range((row_start, x_col), (row_end, x_col)).value)
    ]
    last_x_value = x_values[-1]
    next_x_cell = to_float(sheet.cells(row_end + 1, x_col).value)
    next_x_value = next_x_cell if next_x_cell is not None else (last_x_value + 1 if last_x_value is not None else None)

    intercept_cell = sheet.cells(anchor_row + 2, anchor_col + 3)
    slope_cell = sheet.cells(anchor_row + 2, anchor_col + 4)

    regression_rows: list[dict[str, Any]] = []
    last_signature: tuple[Any, ...] | None = None
    max_quarters = min(N_QUARTERS, len(rows))

    for num_quarters in range(2, max_quarters + 1):
        start_row = rows[-num_quarters]

        intercept_cell.formula2 = (
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{row_end}C{y_col},'
            f'R{start_row}C{x_col}:R{row_end}C{x_col}),"")'
        )
        slope_cell.formula2 = (
            f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{row_end}C{y_col},'
            f'R{start_row}C{x_col}:R{row_end}C{x_col}),"")'
        )
        workbook.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)
        if intercept is None or slope is None:
            continue

        forecast_total_without_sa = None
        if next_x_value is not None:
            forecast_total_without_sa = intercept + slope * next_x_value

        forecast_max = max_value
        forecast_min = min_value
        range_width = safe_subtract(forecast_max, forecast_min)

        signature = (
            round(intercept, 10),
            round(slope, 10),
            round(forecast_total_without_sa, 10) if forecast_total_without_sa is not None else None,
            round(forecast_max, 10) if forecast_max is not None else None,
            round(forecast_min, 10) if forecast_min is not None else None,
        )
        if signature == last_signature:
            continue

        regression_rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters,
                "num_quarters_used": num_quarters,
                "forecast_value": forecast_total_without_sa,
                "actual_value": "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )
        last_signature = signature

    return regression_rows


def write_sheet(worksheet, columns: list[str], rows: list[dict[str, Any]]) -> None:
    worksheet.append(columns)
    column_widths = [len(name) for name in columns]

    for row in rows:
        output_values = [row.get(column_name, "") for column_name in columns]
        worksheet.append(output_values)
        for idx, value in enumerate(output_values):
            if value is None:
                continue
            column_widths[idx] = max(column_widths[idx], len(str(value)))

    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for idx, width in enumerate(column_widths, start=1):
        worksheet.column_dimensions[get_column_letter(idx)].width = min(max(width + 2, 12), 42)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    workbook = Workbook()
    empirical_sheet = workbook.active
    empirical_sheet.title = "empirical_candidates"
    regression_sheet = workbook.create_sheet("regression_candidates")

    write_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)
    workbook.save(output_path)


def process_workbooks(source_dir: Path, destination_dir: Path) -> None:
    destination_dir.mkdir(parents=True, exist_ok=True)
    output_path = build_output_path(source_dir, destination_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(source_dir.iterdir()):
            if not file_path.is_file():
                continue

            if file_path.name.startswith("~"):
                print(f"skipped: {file_path.name} (temporary file)")
                continue

            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped: {file_path.name} (not an .xlsx file)")
                continue

            labels = parse_file_labels(file_path.name)
            workbook = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(process_empirical_sheet(workbook, labels, file_path.name))
                regression_rows.extend(process_regression_sheet(workbook, labels, file_path.name))
                processed_files += 1
                print(f"processed: {file_path.name}")
            except Exception as exc:
                print(f"skipped: {file_path.name} ({exc})")
            finally:
                if workbook is not None:
                    close_workbook_no_save(workbook)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"output: {output_path}")
    print(f"files_processed: {processed_files}")
    print(f"empirical_rows: {len(empirical_rows)}")
    print(f"regression_rows: {len(regression_rows)}")


def main() -> None:
    source_dir = Path(input_dir).expanduser().resolve()
    destination_dir = Path(output_dir).expanduser().resolve()

    if not source_dir.exists() or not source_dir.is_dir():
        raise FileNotFoundError(f"input_dir is not a folder: {source_dir}")

    process_workbooks(source_dir, destination_dir)


if __name__ == "__main__":
    main()

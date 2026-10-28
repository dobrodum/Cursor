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


# User-configurable paths.
input_dir = Path("input")
output_dir = Path("output")


EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"
N_QUARTERS = 10


# Column offsets are relative to the located "max" anchor cell.
FORECAST_VALUE_COL_OFFSET = -1
FORECAST_MAX_COL_OFFSET = 0
FORECAST_MIN_COL_OFFSET = 1
Y_COL_OFFSET = -7
X_COL_OFFSET = -11
LAST_QUARTER_COL_OFFSET = -12
GROWTH_RATE_COL_OFFSET = -5
SALES_CAPTURED_COL_OFFSET = -3
EMP_TEMP_FORMULA_COL_OFFSET = 4
REG_INTERCEPT_TEMP_COL_OFFSET = 4
REG_SLOPE_TEMP_COL_OFFSET = 5


EMPIRICAL_HEADERS = [
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

REGRESSION_HEADERS = [
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


@dataclass(frozen=True)
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


def parse_file_labels(file_name: str) -> FileLabels:
    stem = Path(file_name).stem
    pattern = re.compile(
        r"-\s*([A-Za-z0-9]+)\s*-\s*(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(\d{4})",
        re.IGNORECASE,
    )
    match = pattern.search(stem)

    if not match:
        fallback_ticker = stem.split("-")[0].strip().replace(" ", "_")
        return FileLabels(
            model=f"{fallback_ticker}_Unknown",
            ticker=fallback_ticker,
            model_period="Unknown",
            model_date="",
        )

    ticker = match.group(1).upper()
    period_token = match.group(2).title()
    month_token = match.group(3).title()
    year = int(match.group(4))

    month_map = {
        "Jan": 1,
        "January": 1,
        "Feb": 2,
        "February": 2,
        "Mar": 3,
        "March": 3,
        "Apr": 4,
        "April": 4,
        "May": 5,
        "Jun": 6,
        "June": 6,
        "Jul": 7,
        "July": 7,
        "Aug": 8,
        "August": 8,
        "Sep": 9,
        "Sept": 9,
        "September": 9,
        "Oct": 10,
        "October": 10,
        "Nov": 11,
        "November": 11,
        "Dec": 12,
        "December": 12,
    }

    month_num = month_map.get(month_token[:3].title(), month_map.get(month_token, 1))
    day_map = {"Early": 5, "Mid": 15, "Late": 25}
    model_day = day_map[period_token]
    model_date = date(year, month_num, model_day).isoformat()
    month_short = date(year, month_num, 1).strftime("%b")
    model_period = f"{period_token}{month_short}_{year}"
    model = f"{ticker}_{model_period}"

    return FileLabels(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def ensure_output_path(in_dir: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    base = f"{in_dir.name}_PARAM"
    candidate = out_dir / f"{base}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = out_dir / f"{base}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def iter_input_files(in_dir: Path) -> tuple[list[Path], list[tuple[str, str]]]:
    files: list[Path] = []
    skipped: list[tuple[str, str]] = []

    if not in_dir.exists():
        return files, [(str(in_dir), "input directory does not exist")]

    for path in sorted(in_dir.iterdir()):
        if not path.is_file():
            continue

        name = path.name
        lower_name = name.lower()

        if name.startswith("~"):
            skipped.append((name, "temporary file"))
            continue
        if not lower_name.endswith(".xlsx"):
            skipped.append((name, "not an .xlsx file"))
            continue

        files.append(path)

    return files, skipped


def safe_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def safe_int(value: Any) -> int | None:
    as_float = safe_float(value)
    if as_float is None:
        return None
    return int(as_float)


def safe_divide(numerator: Any, denominator: Any) -> float | None:
    num = safe_float(numerator)
    den = safe_float(denominator)
    if num is None or den is None or den == 0:
        return None
    return num / den


def get_cell_value(sheet: xw.main.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    return sheet.range((row, col)).value


def get_cell_float(sheet: xw.main.Sheet, row: int, col: int) -> float | None:
    return safe_float(get_cell_value(sheet, row, col))


def find_anchor(sheet: xw.main.Sheet, anchor_text: str = "max") -> tuple[int, int] | None:
    used_range = sheet.used_range
    values = used_range.value
    if values is None:
        return None

    if not isinstance(values, list):
        values = [[values]]
    elif values and not isinstance(values[0], list):
        values = [values]

    start_row = used_range.row
    start_col = used_range.column

    target = anchor_text.strip().lower()
    for row_idx, row_values in enumerate(values):
        for col_idx, value in enumerate(row_values):
            if isinstance(value, str) and value.strip().lower() == target:
                return start_row + row_idx, start_col + col_idx
    return None


def set_formula2_r1c1(target_range: xw.main.Range, formulas: list[list[str]]) -> None:
    try:
        target_range.formula2 = formulas
    except Exception:
        target_range.api.Formula2R1C1 = formulas


def close_workbook_safely(wb: xw.main.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.close(False)
        return
    except Exception:
        pass

    wb.api.Close(SaveChanges=False)


def process_empirical_sheet(
    wb: xw.main.Book,
    labels: FileLabels,
    source_file: str,
) -> list[dict[str, Any]]:
    if EMPIRICAL_SHEET_NAME not in [sht.name for sht in wb.sheets]:
        return []

    sheet = wb.sheets[EMPIRICAL_SHEET_NAME]
    anchor = find_anchor(sheet, "max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    x_col = anchor_col + X_COL_OFFSET
    y_col = anchor_col + Y_COL_OFFSET
    temp_col = anchor_col + EMP_TEMP_FORMULA_COL_OFFSET

    candidate_start_row = anchor_row + 1
    candidate_end_row = anchor_row + N_QUARTERS
    data_end_row = anchor_row - 1

    formulas: list[list[str]] = []
    for n_quarters in range(1, N_QUARTERS + 1):
        data_start_row = max(1, data_end_row - n_quarters + 1)
        formula = (
            f'=IFERROR(AVERAGE(R{data_start_row}C{x_col}:R{data_end_row}C{x_col}/'
            f'R{data_start_row}C{y_col}:R{data_end_row}C{y_col}),"")'
        )
        formulas.append([formula])

    temp_range = sheet.range(
        (candidate_start_row, temp_col),
        (candidate_end_row, temp_col),
    )
    set_formula2_r1c1(temp_range, formulas)
    wb.app.calculate()

    rows: list[dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        row_num = candidate_start_row + idx
        num_quarters_used = idx + 1

        avg_penetration = safe_float(get_cell_value(sheet, row_num, temp_col))
        quarterly_sales = get_cell_float(sheet, row_num, x_col)
        reported_sales = get_cell_float(sheet, row_num, y_col)
        growth_rate_pct = get_cell_float(sheet, row_num, anchor_col + GROWTH_RATE_COL_OFFSET)
        last_quarter_used = get_cell_value(sheet, row_num, anchor_col + LAST_QUARTER_COL_OFFSET)
        sales_captured_in_db_pct = get_cell_float(
            sheet, row_num, anchor_col + SALES_CAPTURED_COL_OFFSET
        )

        if sales_captured_in_db_pct is None:
            sales_captured_in_db_pct = safe_divide(quarterly_sales, reported_sales)

        forecast_value = get_cell_float(sheet, row_num, anchor_col + FORECAST_VALUE_COL_OFFSET)
        if forecast_value is None and avg_penetration not in (None, 0) and quarterly_sales is not None:
            forecast_value = quarterly_sales / avg_penetration

        forecast_max = get_cell_float(sheet, row_num, anchor_col + FORECAST_MAX_COL_OFFSET)
        forecast_min = get_cell_float(sheet, row_num, anchor_col + FORECAST_MIN_COL_OFFSET)
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        row = {
            "model": labels.model,
            "ticker": labels.ticker,
            "model_period": labels.model_period,
            "model_date": labels.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration,
            "num_quarters_used": num_quarters_used,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_value,
            "actual_value": reported_sales,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "avg_penetration_pct": avg_penetration,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate_pct,
            "sales_captured_in_db_pct": sales_captured_in_db_pct,
            "source_file": source_file,
        }

        has_signal = any(
            row[key] is not None
            for key in (
                "parameter_value",
                "forecast_value",
                "forecast_max",
                "forecast_min",
                "quarterly_sales",
                "reported_sales",
            )
        )
        if has_signal:
            rows.append(row)

    temp_range.value = None
    return rows


def normalized_signature(values: tuple[Any, ...]) -> tuple[Any, ...]:
    normalized: list[Any] = []
    for value in values:
        as_float = safe_float(value)
        if as_float is None:
            normalized.append(value)
        else:
            normalized.append(round(as_float, 10))
    return tuple(normalized)


def process_regression_sheet(
    wb: xw.main.Book,
    labels: FileLabels,
    source_file: str,
) -> list[dict[str, Any]]:
    if REGRESSION_SHEET_NAME not in [sht.name for sht in wb.sheets]:
        return []

    sheet = wb.sheets[REGRESSION_SHEET_NAME]
    anchor = find_anchor(sheet, "max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    candidate_start_row = anchor_row + 1
    candidate_end_row = anchor_row + N_QUARTERS
    data_end_row = anchor_row - 1

    intercept_col = anchor_col + REG_INTERCEPT_TEMP_COL_OFFSET
    slope_col = anchor_col + REG_SLOPE_TEMP_COL_OFFSET

    intercept_formulas: list[list[str]] = []
    slope_formulas: list[list[str]] = []
    for n_quarters in range(1, N_QUARTERS + 1):
        data_start_row = max(1, data_end_row - n_quarters + 1)
        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{data_start_row}C{y_col}:R{data_end_row}C{y_col},'
            f'R{data_start_row}C{x_col}:R{data_end_row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{data_start_row}C{y_col}:R{data_end_row}C{y_col},'
            f'R{data_start_row}C{x_col}:R{data_end_row}C{x_col}),"")'
        )
        intercept_formulas.append([intercept_formula])
        slope_formulas.append([slope_formula])

    intercept_range = sheet.range(
        (candidate_start_row, intercept_col),
        (candidate_end_row, intercept_col),
    )
    slope_range = sheet.range(
        (candidate_start_row, slope_col),
        (candidate_end_row, slope_col),
    )
    set_formula2_r1c1(intercept_range, intercept_formulas)
    set_formula2_r1c1(slope_range, slope_formulas)
    wb.app.calculate()

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None

    for idx in range(N_QUARTERS):
        row_num = candidate_start_row + idx
        default_num_quarters = idx + 1
        maybe_num_quarters = safe_int(get_cell_value(sheet, row_num, anchor_col - 2))
        num_quarters_used = maybe_num_quarters if maybe_num_quarters is not None else default_num_quarters

        intercept = get_cell_float(sheet, row_num, intercept_col)
        slope = get_cell_float(sheet, row_num, slope_col)
        forecast_value = get_cell_float(sheet, row_num, anchor_col + FORECAST_VALUE_COL_OFFSET)
        forecast_max = get_cell_float(sheet, row_num, anchor_col + FORECAST_MAX_COL_OFFSET)
        forecast_min = get_cell_float(sheet, row_num, anchor_col + FORECAST_MIN_COL_OFFSET)
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        actual_value = get_cell_float(sheet, row_num, y_col)

        row = {
            "model": labels.model,
            "ticker": labels.ticker,
            "model_period": labels.model_period,
            "model_date": labels.model_date,
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

        signature = normalized_signature(
            (
                row["num_quarters_used"],
                row["intercept"],
                row["slope"],
                row["forecast_value"],
                row["forecast_max"],
                row["forecast_min"],
            )
        )

        is_final_row = idx == N_QUARTERS - 1
        if is_final_row and previous_signature is not None and signature == previous_signature:
            continue

        has_signal = any(
            row[key] is not None
            for key in ("intercept", "slope", "forecast_value", "forecast_max", "forecast_min")
        )
        if has_signal:
            rows.append(row)
            previous_signature = signature

    intercept_range.value = None
    slope_range.value = None
    return rows


def write_sheet(ws: Any, headers: list[str], records: list[dict[str, Any]]) -> None:
    ws.append(headers)
    for record in records:
        ws.append([record.get(column, "") for column in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for idx, header in enumerate(headers, start=1):
        col_letter = get_column_letter(idx)
        max_len = len(header)
        for row in ws.iter_rows(min_row=2, min_col=idx, max_col=idx, values_only=True):
            value = row[0]
            value_len = len(str(value)) if value is not None else 0
            if value_len > max_len:
                max_len = value_len
        ws.column_dimensions[col_letter].width = min(max(12, max_len + 2), 48)


def main() -> None:
    in_dir = input_dir.expanduser().resolve()
    out_dir = output_dir.expanduser().resolve()
    output_path = ensure_output_path(in_dir, out_dir)

    files, skipped = iter_input_files(in_dir)
    for file_name, reason in skipped:
        print(f"Skipped: {file_name} ({reason})")

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    old_calculation = app.calculation
    app.calculation = "manual"

    try:
        for file_path in files:
            wb: xw.main.Book | None = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                labels = parse_file_labels(file_path.name)

                empirical_rows.extend(process_empirical_sheet(wb, labels, file_path.name))
                regression_rows.extend(process_regression_sheet(wb, labels, file_path.name))

                files_processed += 1
                print(f"Processed: {file_path.name}")
            except Exception as exc:
                print(f"Skipped: {file_path.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    close_workbook_safely(wb)
    finally:
        app.calculation = old_calculation
        app.quit()

    output_wb = Workbook()
    default_sheet = output_wb.active
    output_wb.remove(default_sheet)

    empirical_ws = output_wb.create_sheet("empirical_candidates")
    regression_ws = output_wb.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_HEADERS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_HEADERS, regression_rows)

    output_wb.save(output_path)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

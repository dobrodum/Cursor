from __future__ import annotations

import calendar
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -----------------------------
# User-configurable directories
# -----------------------------
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


@dataclass
class FileLabels:
    ticker: str
    model_period: str
    model_date: str
    model: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[_\-\s]+", " ", text)
    return text


def is_number(value: Any) -> bool:
    if value is None or value == "":
        return False
    try:
        float(value)
        return True
    except (TypeError, ValueError):
        return False


def to_float(value: Any) -> float | None:
    if not is_number(value):
        return None
    return float(value)


def clean_value(value: Any) -> Any:
    if value is None:
        return ""
    return value


def coerce_to_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if isinstance(values, (list, tuple)):
        if not values:
            return []
        first = values[0]
        if isinstance(first, (list, tuple)):
            return [list(row) for row in values]
        return [list(values)]
    return [[values]]


def coerce_to_1d(values: Any, expected_len: int) -> list[Any]:
    if expected_len <= 0:
        return []
    if isinstance(values, (list, tuple)):
        if values and isinstance(values[0], (list, tuple)):
            flat = [row[0] if row else None for row in values]
        else:
            flat = list(values)
    else:
        flat = [values]
    if len(flat) < expected_len:
        flat.extend([None] * (expected_len - len(flat)))
    return flat[:expected_len]


def find_sheet_case_insensitive(wb: xw.Book, target_name: str) -> xw.Sheet | None:
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == target_name.strip().lower():
            return sheet
    return None


def read_used_matrix(sheet: xw.Sheet) -> tuple[list[list[Any]], int, int, int, int]:
    used = sheet.used_range
    matrix = coerce_to_2d(used.value)
    start_row = used.row
    start_col = used.column
    end_row = used.last_cell.row
    end_col = used.last_cell.column
    return matrix, start_row, start_col, end_row, end_col


def find_max_anchor(
    matrix: list[list[Any]], start_row: int, start_col: int
) -> tuple[int, int] | None:
    candidates: list[tuple[int, int]] = []
    for r_idx, row in enumerate(matrix):
        for c_idx, value in enumerate(row):
            if normalize_text(value) == "max":
                abs_row = start_row + r_idx
                abs_col = start_col + c_idx
                candidates.append((abs_row, abs_col))

    if not candidates:
        return None

    for abs_row, abs_col in candidates:
        rel_row = abs_row - start_row
        rel_col = abs_col - start_col
        row_values = matrix[rel_row]
        for delta in (-2, -1, 1, 2):
            c = rel_col + delta
            if 0 <= c < len(row_values):
                if normalize_text(row_values[c]) == "min":
                    return abs_row, abs_col
    return candidates[0]


def infer_min_column(
    matrix: list[list[Any]], start_row: int, start_col: int, anchor_row: int, anchor_col: int
) -> int:
    rel_row = anchor_row - start_row
    rel_col = anchor_col - start_col
    if 0 <= rel_row < len(matrix):
        row_values = matrix[rel_row]
        for delta in range(-4, 5):
            c = rel_col + delta
            if 0 <= c < len(row_values) and normalize_text(row_values[c]) == "min":
                return start_col + c
    return anchor_col + 1


def find_header_column(
    matrix: list[list[Any]],
    start_row: int,
    start_col: int,
    anchor_row: int,
    search_terms: list[str],
) -> int | None:
    if not matrix:
        return None

    target_rows = [anchor_row - 1, anchor_row, anchor_row + 1]

    def matches(value: Any) -> bool:
        label = normalize_text(value)
        return any(term in label for term in search_terms)

    for abs_row in target_rows:
        rel_row = abs_row - start_row
        if not (0 <= rel_row < len(matrix)):
            continue
        row_values = matrix[rel_row]
        for c_idx, value in enumerate(row_values):
            if matches(value):
                return start_col + c_idx

    # Fallback: one full scan if the local header rows do not match.
    for r_idx, row_values in enumerate(matrix):
        for c_idx, value in enumerate(row_values):
            if matches(value):
                return start_col + c_idx
    return None


def set_formula2(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        cell.formula = formula_r1c1


def safe_close_workbook(wb: xw.Book | None) -> None:
    if wb is None:
        return
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    # Safe fallback if close(save=False) is unsupported by backend/platform.
    try:
        wb.api.Close(False)
        return
    except Exception:
        pass

    try:
        wb.close()
    except Exception:
        pass


def compute_range_width(max_value: Any, min_value: Any) -> float | str:
    max_num = to_float(max_value)
    min_num = to_float(min_value)
    if max_num is None or min_num is None:
        return ""
    return max_num - min_num


def parse_file_labels(file_name: str) -> FileLabels:
    stem = Path(file_name).stem

    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1].upper() if len(parts) >= 2 else ""

    period_match = re.search(r"\b(Early|Mid|Late)\s*([A-Za-z]+)\s*(\d{4})\b", stem, re.IGNORECASE)

    model_period = ""
    model_date = ""
    if period_match:
        period_word = period_match.group(1).title()
        month_token = period_match.group(2)
        year = int(period_match.group(3))
        month_idx = parse_month_index(month_token)
        if month_idx is not None:
            month_abbr = calendar.month_abbr[month_idx]
            model_period = f"{period_word}{month_abbr}_{year}"
            day = {"Early": 5, "Mid": 15, "Late": 25}[period_word]
            model_date = date(year, month_idx, day).isoformat()

    if not ticker:
        ticker_match = re.search(r"\b[A-Z]{2,6}\b", stem)
        ticker = ticker_match.group(0) if ticker_match else ""

    model = f"{ticker}_{model_period}".strip("_")
    return FileLabels(ticker=ticker, model_period=model_period, model_date=model_date, model=model)


def parse_month_index(token: str) -> int | None:
    token_norm = token.strip().lower()
    month_map: dict[str, int] = {}
    for idx in range(1, 13):
        month_map[calendar.month_abbr[idx].lower()] = idx
        month_map[calendar.month_name[idx].lower()] = idx

    if token_norm in month_map:
        return month_map[token_norm]
    if token_norm[:3] in month_map:
        return month_map[token_norm[:3]]
    return None


def build_output_path(input_folder: Path, output_folder: Path) -> Path:
    output_folder.mkdir(parents=True, exist_ok=True)
    base_stem = f"{input_folder.name}_PARAM"
    candidate = output_folder / f"{base_stem}.xlsx"
    if not candidate.exists():
        return candidate

    idx = 1
    while True:
        candidate = output_folder / f"{base_stem}.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def extract_empirical_rows(
    wb: xw.Book, labels: FileLabels, source_file: str
) -> list[dict[str, Any]]:
    sheet = find_sheet_case_insensitive(wb, "Empirical Model")
    if sheet is None:
        print(f"Skipped empirical extraction in {source_file}: sheet 'Empirical Model' not found")
        return []

    matrix, start_row, start_col, end_row, end_col = read_used_matrix(sheet)
    anchor = find_max_anchor(matrix, start_row, start_col)
    if anchor is None:
        print(f"Skipped empirical extraction in {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    min_col = infer_min_column(matrix, start_row, start_col, anchor_row, anchor_col)

    col_num_quarters = find_header_column(
        matrix,
        start_row,
        start_col,
        anchor_row,
        ["num quarters", "quarters used", "n quarters", "num_q"],
    )
    col_last_quarter = find_header_column(
        matrix,
        start_row,
        start_col,
        anchor_row,
        ["last quarter"],
    )
    col_forecast = find_header_column(
        matrix,
        start_row,
        start_col,
        anchor_row,
        ["estimated total sold", "total sold", "tot fcst", "forecast"],
    )
    col_reported = find_header_column(
        matrix,
        start_row,
        start_col,
        anchor_row,
        ["reported sales", "actual sales", "actual"],
    )
    col_quarterly_sales = find_header_column(
        matrix,
        start_row,
        start_col,
        anchor_row,
        ["quarterly sales", "qtr sales"],
    )
    col_growth_rate = find_header_column(
        matrix,
        start_row,
        start_col,
        anchor_row,
        ["growth rate"],
    )
    col_captured = find_header_column(
        matrix,
        start_row,
        start_col,
        anchor_row,
        ["captured in db", "captured"],
    )
    col_avg_pen = find_header_column(
        matrix,
        start_row,
        start_col,
        anchor_row,
        ["avg penetration", "average penetration", "avg pen"],
    )

    # Temporary formula column to keep source model untouched after close(save=False).
    temp_avg_col = max(end_col + 2, anchor_col + 2)
    formula_writes = 0
    for i in range(1, N_QUARTERS + 1):
        row = anchor_row + i
        if row > end_row + N_QUARTERS:
            break
        avg_start_col = max(1, anchor_col - i)
        avg_end_col = max(1, anchor_col - 1)
        formula = f"=AVERAGE(R{row}C{avg_start_col}:R{row}C{avg_end_col})"
        set_formula2(sheet.range((row, temp_avg_col)), formula)
        formula_writes += 1

    if formula_writes:
        wb.app.calculate()

    rows: list[dict[str, Any]] = []
    for i in range(1, N_QUARTERS + 1):
        row = anchor_row + i
        if row > end_row + N_QUARTERS:
            break

        num_quarters_used = (
            sheet.range((row, col_num_quarters)).value if col_num_quarters else i
        )
        if not is_number(num_quarters_used):
            num_quarters_used = i

        last_quarter_used = (
            sheet.range((row, col_last_quarter)).value if col_last_quarter else ""
        )
        forecast_value = sheet.range((row, col_forecast)).value if col_forecast else ""
        actual_value = sheet.range((row, col_reported)).value if col_reported else ""
        forecast_max = sheet.range((row, anchor_col)).value
        forecast_min = sheet.range((row, min_col)).value
        avg_penetration_pct = (
            sheet.range((row, col_avg_pen)).value
            if col_avg_pen
            else sheet.range((row, temp_avg_col)).value
        )
        quarterly_sales = (
            sheet.range((row, col_quarterly_sales)).value if col_quarterly_sales else ""
        )
        reported_sales = actual_value
        growth_rate_pct = sheet.range((row, col_growth_rate)).value if col_growth_rate else ""
        sales_captured_in_db_pct = sheet.range((row, col_captured)).value if col_captured else ""

        signal_values = [
            forecast_value,
            actual_value,
            forecast_max,
            forecast_min,
            avg_penetration_pct,
            quarterly_sales,
            growth_rate_pct,
            sales_captured_in_db_pct,
        ]
        if not any(value not in ("", None) for value in signal_values):
            continue

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": clean_value(avg_penetration_pct),
                "num_quarters_used": clean_value(num_quarters_used),
                "last_quarter_used": clean_value(last_quarter_used),
                "forecast_value": clean_value(forecast_value),
                "actual_value": clean_value(actual_value),
                "forecast_max": clean_value(forecast_max),
                "forecast_min": clean_value(forecast_min),
                "range_width": clean_value(compute_range_width(forecast_max, forecast_min)),
                "avg_penetration_pct": clean_value(avg_penetration_pct),
                "quarterly_sales": clean_value(quarterly_sales),
                "reported_sales": clean_value(reported_sales),
                "growth_rate_pct": clean_value(growth_rate_pct),
                "sales_captured_in_db_pct": clean_value(sales_captured_in_db_pct),
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_rows(
    wb: xw.Book, labels: FileLabels, source_file: str
) -> list[dict[str, Any]]:
    sheet = find_sheet_case_insensitive(wb, "Regression Model")
    if sheet is None:
        print(f"Skipped regression extraction in {source_file}: sheet 'Regression Model' not found")
        return []

    matrix, start_row, start_col, end_row, end_col = read_used_matrix(sheet)
    anchor = find_max_anchor(matrix, start_row, start_col)
    if anchor is None:
        print(f"Skipped regression extraction in {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    min_col = infer_min_column(matrix, start_row, start_col, anchor_row, anchor_col)

    x_col = anchor_col - 11
    y_col = anchor_col - 7
    if x_col < 1 or y_col < 1:
        print(f"Skipped regression extraction in {source_file}: anchor offsets produce invalid x/y columns")
        return []

    history_end_row = max(anchor_row - 1, 1)
    x_vals = coerce_to_1d(sheet.range((1, x_col), (history_end_row, x_col)).value, history_end_row)
    y_vals = coerce_to_1d(sheet.range((1, y_col), (history_end_row, y_col)).value, history_end_row)
    history_rows = [
        idx + 1
        for idx, (x_value, y_value) in enumerate(zip(x_vals, y_vals))
        if is_number(x_value) and is_number(y_value)
    ]

    if len(history_rows) < 2:
        print(f"Skipped regression extraction in {source_file}: insufficient numeric history for regression")
        return []

    col_num_quarters = find_header_column(
        matrix,
        start_row,
        start_col,
        anchor_row,
        ["num quarters", "quarters used", "n quarters", "num_q"],
    )
    col_forecast = find_header_column(
        matrix,
        start_row,
        start_col,
        anchor_row,
        ["tot fcst w/o sa", "tot fcst without sa", "forecast w/o sa", "tot fcst"],
    )
    col_actual = find_header_column(
        matrix,
        start_row,
        start_col,
        anchor_row,
        ["actual", "reported sales"],
    )

    n_rows = min(N_QUARTERS, len(history_rows))
    temp_row_start = max(end_row + 2, anchor_row + 20)
    temp_col_start = max(end_col + 2, anchor_col + 4)

    regression_slices: list[tuple[int, int, int, int, int]] = []
    for idx in range(n_rows):
        n_quarters = idx + 1
        start_hist_row = history_rows[-n_quarters]
        end_hist_row = history_rows[-1]
        data_row = anchor_row + n_quarters
        temp_row = temp_row_start + idx

        intercept_formula = (
            f"=INTERCEPT(R{start_hist_row}C{y_col}:R{end_hist_row}C{y_col},"
            f"R{start_hist_row}C{x_col}:R{end_hist_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_hist_row}C{y_col}:R{end_hist_row}C{y_col},"
            f"R{start_hist_row}C{x_col}:R{end_hist_row}C{x_col})"
        )

        next_x_row = end_hist_row + 1
        next_x_value = sheet.range((next_x_row, x_col)).value
        if not is_number(next_x_value):
            next_x_row = end_hist_row

        forecast_formula = f"=RC[-2]+RC[-1]*R{next_x_row}C{x_col}"

        set_formula2(sheet.range((temp_row, temp_col_start)), intercept_formula)
        set_formula2(sheet.range((temp_row, temp_col_start + 1)), slope_formula)
        set_formula2(sheet.range((temp_row, temp_col_start + 2)), forecast_formula)
        regression_slices.append((n_quarters, data_row, temp_row, start_hist_row, end_hist_row))

    if regression_slices:
        wb.app.calculate()

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None

    for n_quarters, data_row, temp_row, _start_hist_row, _end_hist_row in regression_slices:
        num_quarters_used = (
            sheet.range((data_row, col_num_quarters)).value if col_num_quarters else n_quarters
        )
        if not is_number(num_quarters_used):
            num_quarters_used = n_quarters

        intercept = sheet.range((temp_row, temp_col_start)).value
        slope = sheet.range((temp_row, temp_col_start + 1)).value
        forecast_calc = sheet.range((temp_row, temp_col_start + 2)).value

        forecast_value_sheet = sheet.range((data_row, col_forecast)).value if col_forecast else None
        forecast_value = (
            forecast_value_sheet
            if forecast_value_sheet not in (None, "")
            else forecast_calc
        )

        actual_value = sheet.range((data_row, col_actual)).value if col_actual else ""
        forecast_max = sheet.range((data_row, anchor_col)).value
        forecast_min = sheet.range((data_row, min_col)).value

        signature = (
            to_float(num_quarters_used),
            to_float(intercept),
            to_float(slope),
            to_float(forecast_value),
            to_float(forecast_max),
            to_float(forecast_min),
        )
        if previous_signature == signature:
            continue
        previous_signature = signature

        signal_values = [intercept, slope, forecast_value, forecast_max, forecast_min]
        if not any(value not in ("", None) for value in signal_values):
            continue

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": clean_value(num_quarters_used),
                "num_quarters_used": clean_value(num_quarters_used),
                "forecast_value": clean_value(forecast_value),
                "actual_value": clean_value(actual_value),
                "forecast_max": clean_value(forecast_max),
                "forecast_min": clean_value(forecast_min),
                "range_width": clean_value(compute_range_width(forecast_max, forecast_min)),
                "intercept": clean_value(intercept),
                "slope": clean_value(slope),
                "source_file": source_file,
            }
        )

    return rows


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    write_sheet(wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    wb.save(output_path)


def write_sheet(
    wb: Workbook, sheet_name: str, columns: list[str], rows: list[dict[str, Any]]
) -> None:
    ws = wb.create_sheet(title=sheet_name)
    ws.append(columns)
    for row in rows:
        ws.append([row.get(col, "") for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(ws.max_row, 1)}"

    for col_idx, header in enumerate(columns, start=1):
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = max(12, min(max_len + 2, 42))


def main() -> None:
    if not input_dir.exists() or not input_dir.is_dir():
        raise FileNotFoundError(f"Input folder not found: {input_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = build_output_path(input_dir, output_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    original_calculation = app.calculation

    try:
        app.calculation = "manual"
        for file_path in sorted(input_dir.iterdir(), key=lambda p: p.name.lower()):
            if not file_path.is_file():
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file")
                continue

            print(f"Processing {file_path.name}")
            wb: xw.Book | None = None
            try:
                labels = parse_file_labels(file_path.name)
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_rows(wb, labels, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, labels, file_path.name))
                processed_files += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error ({exc})")
            finally:
                safe_close_workbook(wb)
    finally:
        try:
            app.calculation = original_calculation
        except Exception:
            pass
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

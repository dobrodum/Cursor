from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import re
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Update these two paths to fit your environment.
input_dir = Path("/workspace/input")
output_dir = Path("/workspace/output")

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
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip().lower()


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def to_int(value: Any) -> Optional[int]:
    as_float = to_float(value)
    if as_float is None:
        return None
    return int(round(as_float))


def get_cell_value(sheet: xw.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    return sheet.cells(row, col).value


def set_formula2(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        cell.formula = formula_r1c1


def close_workbook_safely(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.api.Close(False)
        return
    except Exception:
        pass

    try:
        wb.close()
    except Exception:
        pass


def next_output_path(base_input_dir: Path, target_output_dir: Path) -> Path:
    target_output_dir.mkdir(parents=True, exist_ok=True)
    base_stem = f"{base_input_dir.name}_PARAM"

    first_path = target_output_dir / f"{base_stem}.xlsx"
    if not first_path.exists():
        return first_path

    idx = 1
    while True:
        candidate = target_output_dir / f"{base_stem}.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def parse_file_labels(file_name: str) -> FileLabels:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = parts[1] if len(parts) > 1 else ""
    ticker = re.sub(r"[^A-Za-z0-9_]", "", ticker).upper()

    period_token = parts[2] if len(parts) > 2 else ""
    period_match = re.search(r"(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})", period_token, re.IGNORECASE)

    model_period = ""
    model_date = ""
    if period_match:
        day_bucket, month_text, year_text = period_match.groups()
        day_bucket_title = day_bucket.title()
        month_short = month_text[:3].title()
        try:
            month_num = datetime.strptime(month_short, "%b").month
        except ValueError:
            month_num = 1

        day_lookup = {
            "Early": 5,
            "Mid": 15,
            "Late": 25,
        }
        day_num = day_lookup.get(day_bucket_title, 15)
        model_period = f"{day_bucket_title}{month_short}_{year_text}"
        model_date = f"{year_text}-{month_num:02d}-{day_num:02d}"

    model = f"{ticker}_{model_period}" if ticker and model_period else ticker or stem
    return FileLabels(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def iter_xlsx_files(folder: Path) -> Iterable[Path]:
    for file_path in sorted(folder.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"skipped: {file_path.name} (temp file)")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"skipped: {file_path.name} (not .xlsx)")
            continue
        yield file_path


def find_sheet_case_insensitive(wb: xw.Book, target_name: str) -> Optional[xw.Sheet]:
    normalized_target = normalize_text(target_name)
    for sheet in wb.sheets:
        if normalize_text(sheet.name) == normalized_target:
            return sheet
    return None


def matrix_from_used_range(sheet: xw.Sheet) -> Tuple[List[List[Any]], int, int]:
    used = sheet.used_range
    start_row = used.row
    start_col = used.column
    raw_values = used.value

    if raw_values is None:
        return [], start_row, start_col

    if not isinstance(raw_values, list):
        return [[raw_values]], start_row, start_col

    if raw_values and not isinstance(raw_values[0], list):
        return [raw_values], start_row, start_col

    return raw_values, start_row, start_col


def find_max_anchor(sheet: xw.Sheet) -> Optional[Tuple[int, int]]:
    matrix, start_row, start_col = matrix_from_used_range(sheet)
    if not matrix:
        return None

    best_anchor: Optional[Tuple[int, int, int]] = None
    total_rows = len(matrix)

    for r_idx in range(total_rows):
        row_vals = matrix[r_idx]
        for c_idx, value in enumerate(row_vals):
            if normalize_text(value) != "max":
                continue

            score = 1
            for delta in (-1, 1):
                n_col = c_idx + delta
                if 0 <= n_col < len(row_vals) and normalize_text(row_vals[n_col]) == "min":
                    score += 2

            if r_idx + 1 < total_rows and c_idx < len(matrix[r_idx + 1]):
                below = matrix[r_idx + 1][c_idx]
                if isinstance(below, (int, float)):
                    score += 1

            absolute_row = start_row + r_idx
            absolute_col = start_col + c_idx
            candidate = (score, absolute_row, absolute_col)

            if best_anchor is None or candidate > best_anchor:
                best_anchor = candidate

    if best_anchor is None:
        return None
    return best_anchor[1], best_anchor[2]


def read_header_texts(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    left_span: int = 24,
    right_span: int = 12,
) -> Dict[int, str]:
    min_col = max(1, anchor_col - left_span)
    max_col = max(1, anchor_col + right_span)
    headers: Dict[int, str] = {}
    for col in range(min_col, max_col + 1):
        headers[col] = normalize_text(get_cell_value(sheet, anchor_row, col))
    return headers


def find_offset(
    headers: Dict[int, str],
    anchor_col: int,
    keyword_sets: Sequence[Sequence[str]],
    default: int,
) -> int:
    for keywords in keyword_sets:
        for col, header in headers.items():
            if all(keyword in header for keyword in keywords):
                return col - anchor_col
    return default


def has_any_data(values: Sequence[Any]) -> bool:
    for value in values:
        if value is None:
            continue
        if isinstance(value, str) and value.strip() == "":
            continue
        return True
    return False


def maybe_round_for_signature(value: Optional[float]) -> Optional[float]:
    if value is None:
        return None
    return round(float(value), 8)


def extract_empirical_rows(
    wb: xw.Book,
    labels: FileLabels,
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = find_sheet_case_insensitive(wb, "Empirical Model")
    if sheet is None:
        print(f"skipped empirical extraction: {source_file} (missing sheet 'Empirical Model')")
        return []

    anchor = find_max_anchor(sheet)
    if anchor is None:
        print(f"skipped empirical extraction: {source_file} (missing 'max' anchor)")
        return []

    anchor_row, anchor_col = anchor
    headers = read_header_texts(sheet, anchor_row, anchor_col)

    offsets = {
        "num_quarters_used": find_offset(
            headers,
            anchor_col,
            keyword_sets=[("num", "quarter"), ("quarters", "used")],
            default=-14,
        ),
        "last_quarter_used": find_offset(
            headers,
            anchor_col,
            keyword_sets=[("last", "quarter"), ("quarter", "used")],
            default=-13,
        ),
        "forecast_value": find_offset(
            headers,
            anchor_col,
            keyword_sets=[("estimated", "total", "sold"), ("forecast", "total"), ("tot", "fcst")],
            default=-3,
        ),
        "actual_value": find_offset(
            headers,
            anchor_col,
            keyword_sets=[("reported", "sales"), ("actual", "sales"), ("actual",)],
            default=-2,
        ),
        "forecast_min": find_offset(
            headers,
            anchor_col,
            keyword_sets=[("min",)],
            default=1,
        ),
        "avg_penetration_pct": find_offset(
            headers,
            anchor_col,
            keyword_sets=[("avg", "penetration"), ("penetration", "pct"), ("penetration",)],
            default=-9,
        ),
        "quarterly_sales": find_offset(
            headers,
            anchor_col,
            keyword_sets=[("quarterly", "sales"), ("qtr", "sales"), ("quarter", "sales")],
            default=-12,
        ),
        "reported_sales": find_offset(
            headers,
            anchor_col,
            keyword_sets=[("reported", "sales"), ("actual", "sales")],
            default=-2,
        ),
        "growth_rate_pct": find_offset(
            headers,
            anchor_col,
            keyword_sets=[("growth", "rate"), ("growth",)],
            default=-6,
        ),
        "sales_captured_in_db_pct": find_offset(
            headers,
            anchor_col,
            keyword_sets=[("captured", "db"), ("sales", "captured")],
            default=-5,
        ),
    }

    max_col = anchor_col
    penetration_col = anchor_col + offsets["avg_penetration_pct"]
    first_data_row = anchor_row + 1
    helper_col = max(anchor_col + 20, sheet.used_range.last_cell.column + 2)

    rows: List[Dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        row_num = first_data_row + idx

        num_quarters_used = to_int(get_cell_value(sheet, row_num, anchor_col + offsets["num_quarters_used"]))
        if num_quarters_used is None:
            num_quarters_used = idx + 1

        last_quarter_used = get_cell_value(sheet, row_num, anchor_col + offsets["last_quarter_used"])
        forecast_value = to_float(get_cell_value(sheet, row_num, anchor_col + offsets["forecast_value"]))
        actual_value = to_float(get_cell_value(sheet, row_num, anchor_col + offsets["actual_value"]))
        forecast_max = to_float(get_cell_value(sheet, row_num, max_col))
        forecast_min = to_float(get_cell_value(sheet, row_num, anchor_col + offsets["forecast_min"]))
        quarterly_sales = to_float(get_cell_value(sheet, row_num, anchor_col + offsets["quarterly_sales"]))
        reported_sales = to_float(get_cell_value(sheet, row_num, anchor_col + offsets["reported_sales"]))
        growth_rate_pct = to_float(get_cell_value(sheet, row_num, anchor_col + offsets["growth_rate_pct"]))
        sales_captured_in_db_pct = to_float(
            get_cell_value(sheet, row_num, anchor_col + offsets["sales_captured_in_db_pct"])
        )

        avg_penetration_pct = to_float(get_cell_value(sheet, row_num, penetration_col))
        if num_quarters_used > 0:
            average_start = max(first_data_row, row_num - num_quarters_used + 1)
            avg_formula = f"=AVERAGE(R{average_start}C{penetration_col}:R{row_num}C{penetration_col})"
            avg_formula_cell = sheet.cells(row_num, helper_col)
            set_formula2(avg_formula_cell, avg_formula)
            wb.app.calculate()
            formula_avg = to_float(avg_formula_cell.value)
            avg_formula_cell.value = None
            if formula_avg is not None:
                avg_penetration_pct = formula_avg

        if not has_any_data(
            [
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
                quarterly_sales,
                reported_sales,
            ]
        ):
            continue

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
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
        )

    return rows


def extract_regression_rows(
    wb: xw.Book,
    labels: FileLabels,
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = find_sheet_case_insensitive(wb, "Regression Model")
    if sheet is None:
        print(f"skipped regression extraction: {source_file} (missing sheet 'Regression Model')")
        return []

    anchor = find_max_anchor(sheet)
    if anchor is None:
        print(f"skipped regression extraction: {source_file} (missing 'max' anchor)")
        return []

    anchor_row, anchor_col = anchor
    headers = read_header_texts(sheet, anchor_row, anchor_col)

    offsets = {
        "num_quarters_used": find_offset(
            headers,
            anchor_col,
            keyword_sets=[("num", "quarter"), ("quarters", "used")],
            default=-13,
        ),
        "forecast_value": find_offset(
            headers,
            anchor_col,
            keyword_sets=[("tot", "fcst", "w/o"), ("forecast", "without", "sa"), ("fcst", "without", "sa")],
            default=-3,
        ),
        "actual_value": find_offset(
            headers,
            anchor_col,
            keyword_sets=[("actual",), ("reported", "sales")],
            default=-2,
        ),
        "forecast_min": find_offset(
            headers,
            anchor_col,
            keyword_sets=[("min",)],
            default=1,
        ),
        "intercept": find_offset(
            headers,
            anchor_col,
            keyword_sets=[("intercept",)],
            default=2,
        ),
        "slope": find_offset(
            headers,
            anchor_col,
            keyword_sets=[("slope",)],
            default=3,
        ),
    }

    max_col = anchor_col
    x_col = anchor_col - 11
    y_col = anchor_col - 7
    first_data_row = anchor_row + 1
    helper_col = max(anchor_col + 20, sheet.used_range.last_cell.column + 2)

    rows: List[Dict[str, Any]] = []
    last_signature: Optional[Tuple[Any, ...]] = None

    for idx in range(N_QUARTERS):
        row_num = first_data_row + idx

        num_quarters_used = to_int(get_cell_value(sheet, row_num, anchor_col + offsets["num_quarters_used"]))
        if num_quarters_used is None:
            num_quarters_used = idx + 1

        if num_quarters_used <= 0:
            continue

        sample_start_row = max(first_data_row, row_num - num_quarters_used + 1)
        intercept_formula = (
            f"=INTERCEPT(R{sample_start_row}C{y_col}:R{row_num}C{y_col},"
            f"R{sample_start_row}C{x_col}:R{row_num}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{sample_start_row}C{y_col}:R{row_num}C{y_col},"
            f"R{sample_start_row}C{x_col}:R{row_num}C{x_col})"
        )

        intercept_cell = sheet.cells(row_num, helper_col)
        slope_cell = sheet.cells(row_num, helper_col + 1)
        set_formula2(intercept_cell, intercept_formula)
        set_formula2(slope_cell, slope_formula)
        wb.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)
        intercept_cell.value = None
        slope_cell.value = None

        if intercept is None:
            intercept = to_float(get_cell_value(sheet, row_num, anchor_col + offsets["intercept"]))
        if slope is None:
            slope = to_float(get_cell_value(sheet, row_num, anchor_col + offsets["slope"]))

        forecast_value = to_float(get_cell_value(sheet, row_num, anchor_col + offsets["forecast_value"]))
        actual_value = to_float(get_cell_value(sheet, row_num, anchor_col + offsets["actual_value"]))
        forecast_max = to_float(get_cell_value(sheet, row_num, max_col))
        forecast_min = to_float(get_cell_value(sheet, row_num, anchor_col + offsets["forecast_min"]))

        if not has_any_data([forecast_value, forecast_max, forecast_min, intercept, slope]):
            continue

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        signature = (
            num_quarters_used,
            maybe_round_for_signature(forecast_value),
            maybe_round_for_signature(forecast_max),
            maybe_round_for_signature(forecast_min),
            maybe_round_for_signature(intercept),
            maybe_round_for_signature(slope),
        )

        # Some workbooks repeat the final row in the regression section.
        if last_signature == signature:
            continue
        last_signature = signature

        rows.append(
            {
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
        )

    return rows


def autosize_columns(ws) -> None:
    for col_idx in range(1, ws.max_column + 1):
        column_letter = get_column_letter(col_idx)
        max_len = 0
        for row_idx in range(1, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            value_length = len(str(value))
            if value_length > max_len:
                max_len = value_length
        ws.column_dimensions[column_letter].width = min(max(12, max_len + 2), 48)


def write_sheet(ws, columns: List[str], rows: List[Dict[str, Any]]) -> None:
    ws.append(columns)
    for row in rows:
        ws.append([row.get(column) for column in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    autosize_columns(ws)


def build_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    out_wb = Workbook()
    out_wb.remove(out_wb.active)

    ws_empirical = out_wb.create_sheet("empirical_candidates")
    ws_regression = out_wb.create_sheet("regression_candidates")

    write_sheet(ws_empirical, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(ws_regression, REGRESSION_COLUMNS, regression_rows)

    out_wb.save(output_path)


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"input directory does not exist: {input_dir}")

    output_path = next_output_path(input_dir, output_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in iter_xlsx_files(input_dir):
            wb: Optional[xw.Book] = None
            try:
                print(f"processing: {file_path.name}")
                wb = app.books.open(str(file_path), update_links=False)
                labels = parse_file_labels(file_path.name)

                empirical = extract_empirical_rows(wb, labels, file_path.name)
                regression = extract_regression_rows(wb, labels, file_path.name)

                empirical_rows.extend(empirical)
                regression_rows.extend(regression)
                processed_files += 1
                print(
                    f"processed: {file_path.name} "
                    f"(empirical_rows={len(empirical)}, regression_rows={len(regression)})"
                )
            except Exception as exc:
                print(f"skipped: {file_path.name} (error: {exc})")
            finally:
                if wb is not None:
                    close_workbook_safely(wb)
    finally:
        app.quit()

    build_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"output path: {output_path}")
    print(f"number of files processed: {processed_files}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

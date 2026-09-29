from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from statistics import pstdev
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ---------------------------
# User-configurable locations
# ---------------------------
input_dir = "/path/to/input"
output_dir = "/path/to/output"

EMPIRICAL_SHEET = "Empirical Model"
REGRESSION_SHEET = "Regression Model"
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

MONTH_MAP = {
    "jan": 1,
    "january": 1,
    "feb": 2,
    "february": 2,
    "mar": 3,
    "march": 3,
    "apr": 4,
    "april": 4,
    "may": 5,
    "jun": 6,
    "june": 6,
    "jul": 7,
    "july": 7,
    "aug": 8,
    "august": 8,
    "sep": 9,
    "sept": 9,
    "september": 9,
    "oct": 10,
    "october": 10,
    "nov": 11,
    "november": 11,
    "dec": 12,
    "december": 12,
}

PERIOD_DAY = {"early": 5, "mid": 15, "late": 25}
PERIOD_RE = re.compile(r"(Early|Mid|Late)([A-Za-z]+)(\d{4})", re.IGNORECASE)


@dataclass
class ModelMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str


def as_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values
    return [values]


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
            return None
        try:
            return float(text)
        except ValueError:
            return None
    return None


def safe_subtract(a: Any, b: Any) -> Optional[float]:
    aa = to_float(a)
    bb = to_float(b)
    if aa is None or bb is None:
        return None
    return aa - bb


def norm_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def get_matrix_value(
    matrix: Sequence[Sequence[Any]],
    top_row: int,
    left_col: int,
    row: int,
    col: int,
) -> Any:
    r = row - top_row
    c = col - left_col
    if r < 0 or c < 0 or r >= len(matrix):
        return None
    if c >= len(matrix[r]):
        return None
    return matrix[r][c]


def find_anchor_max(
    matrix: Sequence[Sequence[Any]],
    top_row: int,
    left_col: int,
) -> Optional[Tuple[int, int]]:
    for r_idx, row in enumerate(matrix):
        for c_idx, value in enumerate(row):
            if norm_text(value) == "max":
                return top_row + r_idx, left_col + c_idx
    return None


def find_col_in_row_by_patterns(
    matrix: Sequence[Sequence[Any]],
    top_row: int,
    left_col: int,
    row_abs: int,
    patterns: Sequence[str],
    prefer_near_col: Optional[int] = None,
) -> Optional[int]:
    r = row_abs - top_row
    if r < 0 or r >= len(matrix):
        return None

    matches: List[Tuple[int, int]] = []
    row = matrix[r]
    lowered_patterns = [p.lower() for p in patterns]
    for c_idx, value in enumerate(row):
        text = norm_text(value)
        if not text:
            continue
        if any(p in text for p in lowered_patterns):
            col_abs = left_col + c_idx
            distance = (
                abs(col_abs - prefer_near_col) if prefer_near_col is not None else 0
            )
            matches.append((distance, col_abs))

    if not matches:
        return None
    matches.sort(key=lambda x: x[0])
    return matches[0][1]


def get_unique_output_path(input_path: Path, output_path: Path) -> Path:
    base_name = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    idx = 1
    while True:
        candidate = output_path / f"{base_name}.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def parse_model_meta(file_name: str) -> ModelMeta:
    stem = Path(file_name).stem
    parts = [p.strip() for p in stem.split(" - ") if p.strip()]

    ticker = ""
    if len(parts) >= 2:
        ticker = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper()
    if not ticker:
        match = re.search(r"\b[A-Z]{2,8}\b", stem)
        ticker = match.group(0) if match else ""

    period_match = PERIOD_RE.search(stem)
    if not period_match and len(parts) >= 3:
        period_match = PERIOD_RE.search(parts[2])

    period_name = ""
    model_date = ""
    if period_match:
        period_token = period_match.group(1).title()
        month_token = period_match.group(2).strip()
        year = int(period_match.group(3))
        month_key = month_token.lower()
        if month_key not in MONTH_MAP and len(month_key) >= 3:
            month_key = month_key[:3]
        month_num = MONTH_MAP.get(month_key)

        if month_num is not None:
            month_abbrev = date(year, month_num, 1).strftime("%b")
            period_name = f"{period_token}{month_abbrev}_{year}"
            day_num = PERIOD_DAY[period_token.lower()]
            model_date = date(year, month_num, day_num).isoformat()

    if not period_name:
        period_name = "unknown_period"
    if not model_date:
        model_date = ""

    model = f"{ticker}_{period_name}" if ticker else period_name
    return ModelMeta(
        model=model,
        ticker=ticker,
        model_period=period_name,
        model_date=model_date,
    )


def safe_close_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        try:
            wb.close(False)
            return
        except Exception:
            pass
    except Exception:
        pass

    try:
        wb.api.Close(SaveChanges=False)
        return
    except Exception:
        pass


def sheet_exists(wb: xw.Book, sheet_name: str) -> bool:
    return any(s.name == sheet_name for s in wb.sheets)


def extract_empirical_rows(
    wb: xw.Book,
    file_path: Path,
    meta: ModelMeta,
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []

    if not sheet_exists(wb, EMPIRICAL_SHEET):
        return rows

    sheet = wb.sheets[EMPIRICAL_SHEET]
    used = sheet.used_range
    matrix = as_2d(used.value)
    if not matrix:
        return rows

    top_row = used.row
    left_col = used.column
    bottom_row = top_row + len(matrix) - 1
    right_col = left_col + max(len(r) for r in matrix) - 1

    anchor = find_anchor_max(matrix, top_row, left_col)
    if anchor is None:
        return rows

    anchor_row, anchor_col = anchor

    min_col = find_col_in_row_by_patterns(
        matrix,
        top_row,
        left_col,
        anchor_row,
        patterns=[" min ", "min", "minimum"],
        prefer_near_col=anchor_col + 1,
    )
    if min_col is None:
        min_col = anchor_col + 1

    forecast_value_col = find_col_in_row_by_patterns(
        matrix,
        top_row,
        left_col,
        anchor_row,
        patterns=["estimated total sold", "est total sold", "forecast", "tot fcst"],
        prefer_near_col=anchor_col - 1,
    )
    if forecast_value_col is None:
        forecast_value_col = anchor_col - 1

    actual_value_col = find_col_in_row_by_patterns(
        matrix,
        top_row,
        left_col,
        anchor_row,
        patterns=["reported sales", "actual", "actual sales"],
        prefer_near_col=anchor_col - 2,
    )
    if actual_value_col is None:
        actual_value_col = anchor_col - 2

    quarterly_sales_col = find_col_in_row_by_patterns(
        matrix,
        top_row,
        left_col,
        anchor_row,
        patterns=["quarterly sales", "quarter sales", "quarter sales used"],
        prefer_near_col=anchor_col - 3,
    )
    if quarterly_sales_col is None:
        quarterly_sales_col = anchor_col - 3

    growth_col = find_col_in_row_by_patterns(
        matrix,
        top_row,
        left_col,
        anchor_row,
        patterns=["growth", "growth rate"],
        prefer_near_col=anchor_col - 4,
    )
    if growth_col is None:
        growth_col = anchor_col - 4

    captured_col = find_col_in_row_by_patterns(
        matrix,
        top_row,
        left_col,
        anchor_row,
        patterns=["captured in db", "sales captured", "captured"],
        prefer_near_col=anchor_col - 5,
    )
    if captured_col is None:
        captured_col = anchor_col - 5

    last_quarter_col = find_col_in_row_by_patterns(
        matrix,
        top_row,
        left_col,
        anchor_row,
        patterns=["last quarter", "last qtr", "last quarter used"],
        prefer_near_col=anchor_col - 6,
    )
    if last_quarter_col is None:
        last_quarter_col = anchor_col - 6

    penetration_col = find_col_in_row_by_patterns(
        matrix,
        top_row,
        left_col,
        anchor_row,
        patterns=["penetration", "pen %", "pen pct"],
        prefer_near_col=anchor_col - 7,
    )
    if penetration_col is None:
        penetration_col = anchor_col - 7

    data_start_row = anchor_row + 1
    scratch_col = right_col + 2

    planned_rows: List[Tuple[int, int]] = []
    for idx in range(N_QUARTERS):
        n_quarters = idx + 1
        row_abs = data_start_row + idx
        if row_abs > bottom_row:
            break

        # Skip empty candidates quickly.
        anchor_value = get_matrix_value(matrix, top_row, left_col, row_abs, anchor_col)
        min_value = get_matrix_value(matrix, top_row, left_col, row_abs, min_col)
        fcst_value = get_matrix_value(
            matrix, top_row, left_col, row_abs, forecast_value_col
        )
        if anchor_value is None and min_value is None and fcst_value is None:
            continue

        avg_start_row = max(data_start_row, row_abs - (n_quarters - 1))
        formula = (
            f"=AVERAGE(R{avg_start_row}C{penetration_col}:"
            f"R{row_abs}C{penetration_col})"
        )
        sheet.range((row_abs, scratch_col)).formula2 = formula
        planned_rows.append((n_quarters, row_abs))

    if not planned_rows:
        return rows

    wb.app.calculate()

    for n_quarters, row_abs in planned_rows:
        avg_pen = sheet.range((row_abs, scratch_col)).value
        forecast_max = sheet.range((row_abs, anchor_col)).value
        forecast_min = sheet.range((row_abs, min_col)).value
        forecast_value = sheet.range((row_abs, forecast_value_col)).value
        actual_value = sheet.range((row_abs, actual_value_col)).value
        quarterly_sales = sheet.range((row_abs, quarterly_sales_col)).value
        growth_rate = sheet.range((row_abs, growth_col)).value
        sales_captured = sheet.range((row_abs, captured_col)).value
        last_quarter_used = sheet.range((row_abs, last_quarter_col)).value

        if forecast_value is None:
            avg_num = to_float(avg_pen)
            q_sales_num = to_float(quarterly_sales)
            if avg_num is not None and q_sales_num is not None:
                forecast_value = avg_num * q_sales_num

        row: Dict[str, Any] = {
            "model": meta.model,
            "ticker": meta.ticker,
            "model_period": meta.model_period,
            "model_date": meta.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_pen,
            "num_quarters_used": n_quarters,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": safe_subtract(forecast_max, forecast_min),
            "avg_penetration_pct": avg_pen,
            "quarterly_sales": quarterly_sales,
            "reported_sales": actual_value,
            "growth_rate_pct": growth_rate,
            "sales_captured_in_db_pct": sales_captured,
            "source_file": file_path.name,
        }
        rows.append(row)

    return rows


def rounded_key(values: Iterable[Any], places: int = 8) -> Tuple[Any, ...]:
    key: List[Any] = []
    for value in values:
        num = to_float(value)
        if num is None:
            key.append(value)
        else:
            key.append(round(num, places))
    return tuple(key)


def regression_residual_std(
    matrix: Sequence[Sequence[Any]],
    top_row: int,
    left_col: int,
    rows: Sequence[int],
    x_col: int,
    y_col: int,
    intercept: Any,
    slope: Any,
) -> Optional[float]:
    intercept_num = to_float(intercept)
    slope_num = to_float(slope)
    if intercept_num is None or slope_num is None:
        return None

    residuals: List[float] = []
    for row_abs in rows:
        x_val = to_float(get_matrix_value(matrix, top_row, left_col, row_abs, x_col))
        y_val = to_float(get_matrix_value(matrix, top_row, left_col, row_abs, y_col))
        if x_val is None or y_val is None:
            continue
        residuals.append(y_val - (intercept_num + slope_num * x_val))
    if len(residuals) < 2:
        return None
    return float(pstdev(residuals))


def extract_regression_rows(
    wb: xw.Book,
    file_path: Path,
    meta: ModelMeta,
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []

    if not sheet_exists(wb, REGRESSION_SHEET):
        return rows

    sheet = wb.sheets[REGRESSION_SHEET]
    used = sheet.used_range
    matrix = as_2d(used.value)
    if not matrix:
        return rows

    top_row = used.row
    left_col = used.column
    bottom_row = top_row + len(matrix) - 1
    right_col = left_col + max(len(r) for r in matrix) - 1

    anchor = find_anchor_max(matrix, top_row, left_col)
    if anchor is None:
        return rows

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    min_col = find_col_in_row_by_patterns(
        matrix,
        top_row,
        left_col,
        anchor_row,
        patterns=[" min ", "min", "minimum"],
        prefer_near_col=anchor_col + 1,
    )
    if min_col is None:
        min_col = anchor_col + 1

    forecast_col = find_col_in_row_by_patterns(
        matrix,
        top_row,
        left_col,
        anchor_row,
        patterns=[
            "tot fcst w/o sa",
            "tot fcst without sa",
            "forecast w/o sa",
            "forecast without sa",
            "tot fcst",
        ],
        prefer_near_col=anchor_col - 1,
    )
    if forecast_col is None:
        forecast_col = anchor_col - 1

    actual_col = find_col_in_row_by_patterns(
        matrix,
        top_row,
        left_col,
        anchor_row,
        patterns=["actual", "reported sales"],
        prefer_near_col=anchor_col - 2,
    )

    data_rows: List[int] = []
    for row_abs in range(anchor_row + 1, bottom_row + 1):
        x_val = to_float(get_matrix_value(matrix, top_row, left_col, row_abs, x_col))
        y_val = to_float(get_matrix_value(matrix, top_row, left_col, row_abs, y_col))
        if x_val is not None and y_val is not None:
            data_rows.append(row_abs)

    if len(data_rows) < 2:
        for row_abs in range(top_row, bottom_row + 1):
            if row_abs == anchor_row:
                continue
            x_val = to_float(get_matrix_value(matrix, top_row, left_col, row_abs, x_col))
            y_val = to_float(get_matrix_value(matrix, top_row, left_col, row_abs, y_col))
            if x_val is not None and y_val is not None:
                data_rows.append(row_abs)

    if len(data_rows) < 2:
        return rows

    max_n = min(len(data_rows), N_QUARTERS)
    scratch_col = right_col + 2
    scratch_start_row = bottom_row + 2

    planned: List[Tuple[int, Sequence[int], int]] = []
    out_idx = 0
    for n_quarters in range(2, max_n + 1):
        sampled_rows = data_rows[-n_quarters:]
        out_row = scratch_start_row + out_idx
        out_idx += 1

        first_row = sampled_rows[0]
        last_row = sampled_rows[-1]

        intercept_formula = (
            f"=INTERCEPT(R{first_row}C{y_col}:R{last_row}C{y_col},"
            f"R{first_row}C{x_col}:R{last_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{first_row}C{y_col}:R{last_row}C{y_col},"
            f"R{first_row}C{x_col}:R{last_row}C{x_col})"
        )
        sheet.range((out_row, scratch_col)).formula2 = intercept_formula
        sheet.range((out_row, scratch_col + 1)).formula2 = slope_formula

        planned.append((n_quarters, sampled_rows, out_row))

    if not planned:
        return rows

    wb.app.calculate()

    last_key: Optional[Tuple[Any, ...]] = None
    for n_quarters, sampled_rows, out_row in planned:
        intercept = sheet.range((out_row, scratch_col)).value
        slope = sheet.range((out_row, scratch_col + 1)).value

        last_data_row = sampled_rows[-1]
        x_last = to_float(
            get_matrix_value(matrix, top_row, left_col, last_data_row, x_col)
        )
        forecast_value = None
        i_num = to_float(intercept)
        s_num = to_float(slope)
        if i_num is not None and s_num is not None and x_last is not None:
            forecast_value = i_num + (s_num * x_last)

        mapped_row = anchor_row + (n_quarters - 1)
        mapped_forecast = get_matrix_value(
            matrix, top_row, left_col, mapped_row, forecast_col
        )
        if mapped_forecast is not None:
            forecast_value = mapped_forecast

        forecast_max = get_matrix_value(matrix, top_row, left_col, mapped_row, anchor_col)
        forecast_min = get_matrix_value(matrix, top_row, left_col, mapped_row, min_col)

        if forecast_max is None or forecast_min is None:
            resid_std = regression_residual_std(
                matrix=matrix,
                top_row=top_row,
                left_col=left_col,
                rows=sampled_rows,
                x_col=x_col,
                y_col=y_col,
                intercept=intercept,
                slope=slope,
            )
            f_num = to_float(forecast_value)
            if resid_std is not None and f_num is not None:
                if forecast_max is None:
                    forecast_max = f_num + resid_std
                if forecast_min is None:
                    forecast_min = f_num - resid_std

        actual_value = None
        if actual_col is not None:
            actual_value = get_matrix_value(
                matrix, top_row, left_col, sampled_rows[-1], actual_col
            )

        current_key = rounded_key([intercept, slope, forecast_value, forecast_max, forecast_min])
        if current_key == last_key:
            continue
        last_key = current_key

        row: Dict[str, Any] = {
            "model": meta.model,
            "ticker": meta.ticker,
            "model_period": meta.model_period,
            "model_date": meta.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n_quarters,
            "num_quarters_used": n_quarters,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": safe_subtract(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": file_path.name,
        }
        rows.append(row)

    return rows


def write_rows_to_sheet(
    ws: Any,
    headers: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    ws.append(list(headers))
    for col_idx in range(1, len(headers) + 1):
        ws.cell(row=1, column=col_idx).font = Font(bold=True)

    for row in rows:
        ws.append([row.get(h) for h in headers])

    ws.freeze_panes = "A2"
    final_row = max(1, len(rows) + 1)
    final_col = len(headers)
    ws.auto_filter.ref = f"A1:{get_column_letter(final_col)}{final_row}"

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_num in range(2, final_row + 1):
            val = ws.cell(row=row_num, column=col_idx).value
            if val is None:
                continue
            max_len = max(max_len, len(str(val)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 42)


def save_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    wb_out = Workbook()
    default_sheet = wb_out.active
    wb_out.remove(default_sheet)

    empirical_ws = wb_out.create_sheet("empirical_candidates")
    regression_ws = wb_out.create_sheet("regression_candidates")

    write_rows_to_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_rows_to_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)
    wb_out.save(output_path)


def iter_input_files(folder: Path) -> Iterable[Path]:
    for item in sorted(folder.iterdir(), key=lambda p: p.name.lower()):
        if not item.is_file():
            print(f"Skipped {item.name}: not a file")
            continue
        if item.name.startswith("~"):
            print(f"Skipped {item.name}: temporary workbook")
            continue
        if item.suffix.lower() != ".xlsx":
            print(f"Skipped {item.name}: not an .xlsx file")
            continue
        yield item


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()
    output_path.mkdir(parents=True, exist_ok=True)

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"Input directory does not exist: {input_path}")

    output_file = get_unique_output_path(input_path=input_path, output_path=output_path)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in iter_input_files(input_path):
            print(f"Processing {file_path.name}")
            wb: Optional[xw.Book] = None
            try:
                meta = parse_model_meta(file_path.name)
                wb = app.books.open(str(file_path), update_links=False)

                empirical_rows.extend(extract_empirical_rows(wb, file_path, meta))
                regression_rows.extend(extract_regression_rows(wb, file_path, meta))
                files_processed += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error ({exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        try:
            app.quit()
        except Exception:
            pass

    save_output_workbook(
        output_path=output_file,
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
    )

    print(f"Output file: {output_file}")
    print(f"Files processed: {files_processed}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

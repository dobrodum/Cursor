#!/usr/bin/env python3
from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure these two paths before running.
input_dir = Path("./input")
output_dir = Path("./output")

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

DAY_BY_PERIOD = {
    "early": 5,
    "mid": 15,
    "late": 25,
}


def as_2d(values: Any) -> List[List[Any]]:
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


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    if not text:
        return ""
    text = re.sub(r"[\r\n\t]+", " ", text)
    text = re.sub(r"[_/\-]+", " ", text)
    text = re.sub(r"[^a-z0-9 %]+", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def to_number(value: Any) -> Optional[float]:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        stripped = value.strip()
        if not stripped:
            return None
        is_percent = stripped.endswith("%")
        cleaned = stripped[:-1] if is_percent else stripped
        cleaned = cleaned.replace(",", "")
        try:
            number = float(cleaned)
            return number / 100.0 if is_percent else number
        except ValueError:
            return None
    return None


def clean_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    number = to_number(value)
    if number is not None:
        rounded = round(number)
        if abs(number - rounded) < 1e-10:
            return int(rounded)
        return float(number)
    if isinstance(value, str):
        text = value.strip()
        return text if text else None
    return value


def to_int(value: Any) -> Optional[int]:
    number = to_number(value)
    if number is None:
        return None
    return int(round(number))


def matrix_value(
    matrix: List[List[Any]],
    top_row: int,
    left_col: int,
    row: int,
    col: int,
) -> Any:
    row_idx = row - top_row
    col_idx = col - left_col
    if row_idx < 0 or row_idx >= len(matrix):
        return None
    row_values = matrix[row_idx]
    if col_idx < 0 or col_idx >= len(row_values):
        return None
    return row_values[col_idx]


def find_anchor_cell(
    matrix: List[List[Any]],
    top_row: int,
    left_col: int,
    target: str = "max",
) -> Optional[Tuple[int, int]]:
    target_normalized = normalize_text(target)
    for r_idx, row_values in enumerate(matrix):
        for c_idx, value in enumerate(row_values):
            if normalize_text(value) == target_normalized:
                return top_row + r_idx, left_col + c_idx
    return None


def find_col_in_row(
    matrix: List[List[Any]],
    top_row: int,
    left_col: int,
    row: int,
    min_col: int,
    max_col: int,
    keyword_groups: Sequence[Sequence[str]],
) -> Optional[int]:
    for col in range(min_col, max_col + 1):
        cell_text = normalize_text(matrix_value(matrix, top_row, left_col, row, col))
        if not cell_text:
            continue
        for group in keyword_groups:
            if all(keyword in cell_text for keyword in group):
                return col
    return None


def resolve_col_near_anchor(
    matrix: List[List[Any]],
    top_row: int,
    left_col: int,
    anchor_row: int,
    min_col: int,
    max_col: int,
    keyword_groups: Sequence[Sequence[str]],
    fallback_col: int,
) -> int:
    for row in (anchor_row, anchor_row - 1, anchor_row + 1):
        col = find_col_in_row(
            matrix=matrix,
            top_row=top_row,
            left_col=left_col,
            row=row,
            min_col=min_col,
            max_col=max_col,
            keyword_groups=keyword_groups,
        )
        if col is not None:
            return col
    return fallback_col


def find_col_anywhere(
    matrix: List[List[Any]],
    top_row: int,
    left_col: int,
    keyword_groups: Sequence[Sequence[str]],
    exclude_tokens: Sequence[str] = (),
) -> Optional[int]:
    exclude_normalized = [normalize_text(token) for token in exclude_tokens if token]
    for r_idx, row_values in enumerate(matrix):
        for c_idx, value in enumerate(row_values):
            text = normalize_text(value)
            if not text:
                continue
            if any(token in text for token in exclude_normalized):
                continue
            for group in keyword_groups:
                if all(keyword in text for keyword in group):
                    return left_col + c_idx
    return None


def collect_numeric_rows_for_col(
    matrix: List[List[Any]],
    top_row: int,
    left_col: int,
    col: int,
    max_row: Optional[int] = None,
) -> List[Tuple[int, float]]:
    out: List[Tuple[int, float]] = []
    for r_idx, _ in enumerate(matrix):
        row_number = top_row + r_idx
        if max_row is not None and row_number > max_row:
            break
        value = matrix_value(matrix, top_row, left_col, row_number, col)
        number = to_number(value)
        if number is not None:
            out.append((row_number, number))
    return out


def parse_month(month_token: str) -> Optional[int]:
    token = month_token.strip()
    if not token:
        return None
    for fmt in ("%b", "%B"):
        try:
            return datetime.strptime(token, fmt).month
        except ValueError:
            continue
    if len(token) >= 3:
        try:
            return datetime.strptime(token[:3], "%b").month
        except ValueError:
            return None
    return None


def parse_file_label(file_path: Path) -> Dict[str, Optional[str]]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = parts[1].upper() if len(parts) >= 2 and parts[1] else "UNKNOWN"
    period_source = parts[2] if len(parts) >= 3 else stem
    period_token = period_source.split("_")[0]

    period_match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]+)\s*(\d{4})",
        period_token,
        flags=re.IGNORECASE,
    )

    model_period = "Unknown_0000"
    model_date: Optional[str] = None

    if period_match:
        period_part = period_match.group(1).title()
        month_part = period_match.group(2)
        year_part = int(period_match.group(3))
        month_num = parse_month(month_part)
        if month_num is not None:
            month_short = datetime(year_part, month_num, 1).strftime("%b")
            model_period = f"{period_part}{month_short}_{year_part}"
            day = DAY_BY_PERIOD[period_part.lower()]
            model_date = date(year_part, month_num, day).isoformat()
        else:
            model_period = f"{period_part}{month_part[:3].title()}_{year_part}"

    model = f"{ticker}_{model_period}"
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def close_without_saving(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        wb.close()
    except Exception:
        pass


def next_output_path(input_path: Path, out_dir: Path) -> Path:
    input_folder_name = input_path.resolve().name
    base_name = f"{input_folder_name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    idx = 1
    while True:
        candidate = out_dir / f"{base_name}.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def extract_empirical_rows(
    wb: xw.Book,
    metadata: Dict[str, Optional[str]],
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        return []

    used = sheet.used_range
    matrix = as_2d(used.value)
    if not matrix:
        return []

    top_row = used.row
    left_col = used.column

    anchor = find_anchor_cell(matrix, top_row, left_col, target="max")
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    search_min_col = max(left_col, anchor_col - 24)
    search_max_col = anchor_col + 24

    num_quarters_col = resolve_col_near_anchor(
        matrix,
        top_row,
        left_col,
        anchor_row,
        search_min_col,
        search_max_col,
        keyword_groups=(("num", "quarter"), ("quarters", "used"), ("n", "quarter")),
        fallback_col=anchor_col - 6,
    )
    last_quarter_col = resolve_col_near_anchor(
        matrix,
        top_row,
        left_col,
        anchor_row,
        search_min_col,
        search_max_col,
        keyword_groups=(("last", "quarter"), ("quarter", "used")),
        fallback_col=anchor_col - 5,
    )
    avg_penetration_col = resolve_col_near_anchor(
        matrix,
        top_row,
        left_col,
        anchor_row,
        search_min_col,
        search_max_col,
        keyword_groups=(("avg", "penetration"), ("average", "penetration")),
        fallback_col=anchor_col - 4,
    )
    forecast_col = resolve_col_near_anchor(
        matrix,
        top_row,
        left_col,
        anchor_row,
        search_min_col,
        search_max_col,
        keyword_groups=(
            ("estimated", "sold"),
            ("est", "sold"),
            ("forecast", "total"),
            ("tot", "fcst"),
        ),
        fallback_col=anchor_col - 3,
    )
    actual_col = resolve_col_near_anchor(
        matrix,
        top_row,
        left_col,
        anchor_row,
        search_min_col,
        search_max_col,
        keyword_groups=(("reported", "sales"), ("actual", "sales")),
        fallback_col=anchor_col - 2,
    )
    forecast_max_col = resolve_col_near_anchor(
        matrix,
        top_row,
        left_col,
        anchor_row,
        search_min_col,
        search_max_col,
        keyword_groups=(("max",),),
        fallback_col=anchor_col,
    )
    forecast_min_col = resolve_col_near_anchor(
        matrix,
        top_row,
        left_col,
        anchor_row,
        search_min_col,
        search_max_col,
        keyword_groups=(("min",),),
        fallback_col=anchor_col + 1,
    )
    quarterly_sales_col = resolve_col_near_anchor(
        matrix,
        top_row,
        left_col,
        anchor_row,
        search_min_col,
        search_max_col,
        keyword_groups=(("quarterly", "sales"), ("qtr", "sales")),
        fallback_col=anchor_col - 7,
    )
    growth_rate_col = resolve_col_near_anchor(
        matrix,
        top_row,
        left_col,
        anchor_row,
        search_min_col,
        search_max_col,
        keyword_groups=(("growth", "rate"),),
        fallback_col=anchor_col + 2,
    )
    sales_captured_col = resolve_col_near_anchor(
        matrix,
        top_row,
        left_col,
        anchor_row,
        search_min_col,
        search_max_col,
        keyword_groups=(("sales", "captured"), ("captured", "db")),
        fallback_col=anchor_col + 3,
    )

    penetration_source_col = find_col_anywhere(
        matrix,
        top_row,
        left_col,
        keyword_groups=(("penetration",),),
        exclude_tokens=("avg", "average"),
    )
    if penetration_source_col is None:
        penetration_source_col = anchor_col - 11

    penetration_values = collect_numeric_rows_for_col(
        matrix=matrix,
        top_row=top_row,
        left_col=left_col,
        col=penetration_source_col,
        max_row=anchor_row - 1,
    )

    avg_pen_by_n: Dict[int, Optional[float]] = {}
    if penetration_values:
        first_pen_row = penetration_values[0][0]
        last_pen_row = penetration_values[-1][0]
        max_n = min(10, last_pen_row - first_pen_row + 1)
        temp_col = used.last_cell.column + 5
        temp_start_row = anchor_row + 1

        for n in range(1, max_n + 1):
            start_row = max(first_pen_row, last_pen_row - n + 1)
            sheet.range((temp_start_row + n - 1, temp_col)).formula2 = (
                f"=AVERAGE(R{start_row}C{penetration_source_col}:"
                f"R{last_pen_row}C{penetration_source_col})"
            )
        wb.app.calculate()

        avg_values = as_2d(
            sheet.range(
                (temp_start_row, temp_col),
                (temp_start_row + max_n - 1, temp_col),
            ).value
        )
        for n in range(1, max_n + 1):
            avg_pen_by_n[n] = to_number(avg_values[n - 1][0])

    rows: List[Dict[str, Any]] = []
    for i in range(10):
        row = anchor_row + 1 + i
        num_quarters_used = to_int(
            matrix_value(matrix, top_row, left_col, row, num_quarters_col)
        )
        if num_quarters_used is None:
            num_quarters_used = i + 1

        last_quarter_used = clean_value(
            matrix_value(matrix, top_row, left_col, row, last_quarter_col)
        )
        avg_penetration_pct = to_number(
            matrix_value(matrix, top_row, left_col, row, avg_penetration_col)
        )
        if avg_penetration_pct is None:
            avg_penetration_pct = avg_pen_by_n.get(num_quarters_used)

        forecast_value = clean_value(
            matrix_value(matrix, top_row, left_col, row, forecast_col)
        )
        actual_value = clean_value(
            matrix_value(matrix, top_row, left_col, row, actual_col)
        )
        forecast_max = to_number(
            matrix_value(matrix, top_row, left_col, row, forecast_max_col)
        )
        forecast_min = to_number(
            matrix_value(matrix, top_row, left_col, row, forecast_min_col)
        )
        quarterly_sales = clean_value(
            matrix_value(matrix, top_row, left_col, row, quarterly_sales_col)
        )
        growth_rate_pct = clean_value(
            matrix_value(matrix, top_row, left_col, row, growth_rate_col)
        )
        sales_captured_pct = clean_value(
            matrix_value(matrix, top_row, left_col, row, sales_captured_col)
        )

        reported_sales = actual_value
        range_width = (
            (forecast_max - forecast_min)
            if forecast_max is not None and forecast_min is not None
            else None
        )

        row_has_data = any(
            value is not None
            for value in (
                last_quarter_used,
                avg_penetration_pct,
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                quarterly_sales,
                growth_rate_pct,
                sales_captured_pct,
            )
        )
        if not row_has_data:
            continue

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
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
                "sales_captured_in_db_pct": sales_captured_pct,
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_rows(
    wb: xw.Book,
    metadata: Dict[str, Optional[str]],
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        return []

    used = sheet.used_range
    matrix = as_2d(used.value)
    if not matrix:
        return []

    top_row = used.row
    left_col = used.column

    anchor = find_anchor_cell(matrix, top_row, left_col, target="max")
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    search_min_col = max(left_col, anchor_col - 24)
    search_max_col = anchor_col + 24

    num_quarters_col = resolve_col_near_anchor(
        matrix,
        top_row,
        left_col,
        anchor_row,
        search_min_col,
        search_max_col,
        keyword_groups=(("num", "quarter"), ("quarters", "used"), ("n", "quarter")),
        fallback_col=anchor_col - 4,
    )
    forecast_col = resolve_col_near_anchor(
        matrix,
        top_row,
        left_col,
        anchor_row,
        search_min_col,
        search_max_col,
        keyword_groups=(("tot", "fcst"), ("forecast", "without"), ("forecast", "w/o")),
        fallback_col=anchor_col - 1,
    )
    forecast_max_col = resolve_col_near_anchor(
        matrix,
        top_row,
        left_col,
        anchor_row,
        search_min_col,
        search_max_col,
        keyword_groups=(("max",),),
        fallback_col=anchor_col,
    )
    forecast_min_col = resolve_col_near_anchor(
        matrix,
        top_row,
        left_col,
        anchor_row,
        search_min_col,
        search_max_col,
        keyword_groups=(("min",),),
        fallback_col=anchor_col + 1,
    )
    intercept_col = resolve_col_near_anchor(
        matrix,
        top_row,
        left_col,
        anchor_row,
        search_min_col,
        search_max_col,
        keyword_groups=(("intercept",),),
        fallback_col=anchor_col - 3,
    )
    slope_col = resolve_col_near_anchor(
        matrix,
        top_row,
        left_col,
        anchor_row,
        search_min_col,
        search_max_col,
        keyword_groups=(("slope",),),
        fallback_col=anchor_col - 2,
    )
    actual_col = find_col_in_row(
        matrix,
        top_row,
        left_col,
        anchor_row,
        search_min_col,
        search_max_col,
        keyword_groups=(("actual",), ("reported", "sales")),
    )

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    xy_rows: List[int] = []
    for r in range(top_row, top_row + len(matrix)):
        x_val = to_number(matrix_value(matrix, top_row, left_col, r, x_col))
        y_val = to_number(matrix_value(matrix, top_row, left_col, r, y_col))
        if x_val is not None and y_val is not None and r < anchor_row:
            xy_rows.append(r)

    if len(xy_rows) < 2:
        xy_rows = []
        for r in range(top_row, top_row + len(matrix)):
            x_val = to_number(matrix_value(matrix, top_row, left_col, r, x_col))
            y_val = to_number(matrix_value(matrix, top_row, left_col, r, y_col))
            if x_val is not None and y_val is not None:
                xy_rows.append(r)

    intercept_by_n: Dict[int, Optional[float]] = {}
    slope_by_n: Dict[int, Optional[float]] = {}
    next_x_value: Optional[float] = None

    if len(xy_rows) >= 2:
        first_xy_row = xy_rows[0]
        last_xy_row = xy_rows[-1]
        next_x_raw = matrix_value(matrix, top_row, left_col, last_xy_row, x_col)
        last_x = to_number(next_x_raw)
        if last_x is not None:
            next_x_value = last_x + 1.0

        max_n = min(10, last_xy_row - first_xy_row + 1)
        temp_col = used.last_cell.column + 5
        temp_start_row = anchor_row + 1

        for n in range(1, max_n + 1):
            start_row = max(first_xy_row, last_xy_row - n + 1)
            sheet.range((temp_start_row + n - 1, temp_col)).formula2 = (
                f"=INTERCEPT(R{start_row}C{y_col}:R{last_xy_row}C{y_col},"
                f"R{start_row}C{x_col}:R{last_xy_row}C{x_col})"
            )
            sheet.range((temp_start_row + n - 1, temp_col + 1)).formula2 = (
                f"=SLOPE(R{start_row}C{y_col}:R{last_xy_row}C{y_col},"
                f"R{start_row}C{x_col}:R{last_xy_row}C{x_col})"
            )
        wb.app.calculate()

        coeff_values = as_2d(
            sheet.range(
                (temp_start_row, temp_col),
                (temp_start_row + max_n - 1, temp_col + 1),
            ).value
        )
        for n in range(1, max_n + 1):
            intercept_by_n[n] = to_number(coeff_values[n - 1][0])
            slope_by_n[n] = to_number(coeff_values[n - 1][1])

    rows: List[Dict[str, Any]] = []
    prev_signature: Optional[Tuple[Any, ...]] = None

    for i in range(10):
        row = anchor_row + 1 + i
        num_quarters_used = to_int(
            matrix_value(matrix, top_row, left_col, row, num_quarters_col)
        )
        if num_quarters_used is None:
            num_quarters_used = i + 1

        intercept_value = to_number(
            matrix_value(matrix, top_row, left_col, row, intercept_col)
        )
        slope_value = to_number(matrix_value(matrix, top_row, left_col, row, slope_col))

        if intercept_value is None:
            intercept_value = intercept_by_n.get(num_quarters_used)
        if slope_value is None:
            slope_value = slope_by_n.get(num_quarters_used)

        forecast_value = clean_value(
            matrix_value(matrix, top_row, left_col, row, forecast_col)
        )
        if (
            forecast_value is None
            and intercept_value is not None
            and slope_value is not None
            and next_x_value is not None
        ):
            forecast_value = intercept_value + slope_value * next_x_value

        actual_value = (
            clean_value(matrix_value(matrix, top_row, left_col, row, actual_col))
            if actual_col is not None
            else None
        )
        forecast_max = to_number(
            matrix_value(matrix, top_row, left_col, row, forecast_max_col)
        )
        forecast_min = to_number(
            matrix_value(matrix, top_row, left_col, row, forecast_min_col)
        )

        range_width = (
            (forecast_max - forecast_min)
            if forecast_max is not None and forecast_min is not None
            else None
        )

        row_has_data = any(
            value is not None
            for value in (
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                intercept_value,
                slope_value,
            )
        )
        if not row_has_data:
            continue

        signature = (
            num_quarters_used,
            round(intercept_value, 10) if intercept_value is not None else None,
            round(slope_value, 10) if slope_value is not None else None,
            round(forecast_value, 10) if isinstance(forecast_value, (int, float)) else forecast_value,
            round(forecast_max, 10) if forecast_max is not None else None,
            round(forecast_min, 10) if forecast_min is not None else None,
        )
        if signature == prev_signature:
            continue
        prev_signature = signature

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept_value,
                "slope": slope_value,
                "source_file": source_file,
            }
        )

    return rows


def style_sheet(ws) -> None:
    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx in range(1, ws.max_column + 1):
        max_len = 0
        for row_idx in range(1, ws.max_row + 1):
            cell_value = ws.cell(row=row_idx, column=col_idx).value
            if cell_value is None:
                continue
            max_len = max(max_len, len(str(cell_value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    out_wb = Workbook()
    default_sheet = out_wb.active
    out_wb.remove(default_sheet)

    empirical_ws = out_wb.create_sheet("empirical_candidates")
    regression_ws = out_wb.create_sheet("regression_candidates")

    empirical_ws.append(EMPIRICAL_COLUMNS)
    for row in empirical_rows:
        empirical_ws.append([row.get(col) for col in EMPIRICAL_COLUMNS])

    regression_ws.append(REGRESSION_COLUMNS)
    for row in regression_rows:
        regression_ws.append([row.get(col) for col in REGRESSION_COLUMNS])

    style_sheet(empirical_ws)
    style_sheet(regression_ws)

    out_wb.save(output_path)


def iter_source_files(folder: Path) -> Iterable[Path]:
    for file_path in sorted(folder.iterdir()):
        if not file_path.is_file():
            print(f"Skipped {file_path.name}: not a file")
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped {file_path.name}: temporary workbook")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped {file_path.name}: not an .xlsx file")
            continue
        yield file_path


def main() -> None:
    src_dir = Path(input_dir)
    dst_dir = Path(output_dir)
    dst_dir.mkdir(parents=True, exist_ok=True)

    if not src_dir.exists():
        print(f"Input directory does not exist: {src_dir}")
        return

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in iter_source_files(src_dir):
            print(f"Processing {file_path.name}")
            try:
                wb = app.books.open(str(file_path), update_links=False)
            except Exception as exc:
                print(f"Skipped {file_path.name}: failed to open ({exc})")
                continue

            try:
                metadata = parse_file_label(file_path)
                empirical_rows.extend(
                    extract_empirical_rows(wb=wb, metadata=metadata, source_file=file_path.name)
                )
                regression_rows.extend(
                    extract_regression_rows(wb=wb, metadata=metadata, source_file=file_path.name)
                )
                files_processed += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: extraction failed ({exc})")
            finally:
                close_without_saving(wb)
    finally:
        app.quit()

    output_path = next_output_path(src_dir, dst_dir)
    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

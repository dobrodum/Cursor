#!/usr/bin/env python3
"""Extract empirical and regression candidates from Excel model workbooks."""

from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw

# -----------------------------
# Configure input/output paths.
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

PHASE_DAY = {"early": 5, "mid": 15, "late": 25}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[_\-/]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text


def as_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "")
        if not cleaned:
            return None
        if cleaned.endswith("%"):
            cleaned = cleaned[:-1]
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def to_rows(value: Any) -> List[List[Any]]:
    if value is None:
        return []
    if isinstance(value, tuple):
        value = list(value)
    if not isinstance(value, list):
        return [[value]]
    if not value:
        return []
    first = value[0]
    if isinstance(first, tuple):
        value = [list(v) if isinstance(v, tuple) else v for v in value]
        first = value[0]
    if isinstance(first, list):
        return value
    return [value]


def to_list(value: Any) -> List[Any]:
    if value is None:
        return []
    if isinstance(value, tuple):
        value = list(value)
    if isinstance(value, list):
        return value
    return [value]


def parse_month_number(month_token: str) -> Optional[int]:
    cleaned = month_token.strip()
    for fmt in ("%b", "%B"):
        try:
            return datetime.strptime(cleaned, fmt).month
        except ValueError:
            continue
    return None


def parse_file_label(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = parts[1] if len(parts) >= 2 else ""
    period_source = parts[2] if len(parts) >= 3 else stem
    period_source = period_source.split("_")[0]

    model_period = ""
    model_date = ""

    match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(\d{4})",
        period_source,
        flags=re.IGNORECASE,
    )
    if match:
        phase_raw, month_raw, year_raw = match.groups()
        month_num = parse_month_number(month_raw)
        if month_num is not None:
            phase = phase_raw[0].upper() + phase_raw[1:].lower()
            year = int(year_raw)
            day = PHASE_DAY[phase.lower()]
            month_abbrev = datetime(year, month_num, day).strftime("%b")
            model_period = f"{phase}{month_abbrev}_{year}"
            model_date = date(year, month_num, day).isoformat()

    if not ticker:
        ticker_guess = re.search(r"\b[A-Z]{1,6}\b", stem)
        ticker = ticker_guess.group(0) if ticker_guess else "UNKNOWN"

    if not model_period:
        fallback = re.sub(r"[^A-Za-z0-9]", "", period_source)
        model_period = fallback or "UnknownPeriod"

    model = f"{ticker}_{model_period}"
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def get_output_path(input_path: Path, output_path: Path) -> Path:
    base_name = f"{input_path.name}_PARAM.xlsx"
    candidate = output_path / base_name
    if not candidate.exists():
        return candidate

    i = 1
    while True:
        candidate = output_path / f"{input_path.name}_PARAM.{i}.xlsx"
        if not candidate.exists():
            return candidate
        i += 1


def get_sheet(wb: xw.Book, target_name: str) -> Optional[xw.Sheet]:
    target = target_name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def close_source_workbook(wb: xw.Book) -> None:
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
        try:
            wb.close()
        except Exception:
            pass


def set_formula2(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def find_anchor_cell(sheet: xw.Sheet, anchor_text: str = "max") -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    values = to_rows(used.value)
    if not values:
        return None

    start_row = used.row
    start_col = used.column
    needle = anchor_text.strip().lower()

    best: Optional[Tuple[int, int, int, int]] = None
    for r_idx, row_values in enumerate(values):
        for c_idx, cell_value in enumerate(row_values):
            if not (isinstance(cell_value, str) and cell_value.strip().lower() == needle):
                continue

            score = 0
            right_value = row_values[c_idx + 1] if c_idx + 1 < len(row_values) else None
            if isinstance(right_value, str) and right_value.strip().lower() == "min":
                score += 3

            left_value = row_values[c_idx - 1] if c_idx - 1 >= 0 else None
            if isinstance(left_value, str):
                left_norm = normalize_text(left_value)
                if "forecast" in left_norm or "fcst" in left_norm:
                    score += 1

            abs_row = start_row + r_idx
            abs_col = start_col + c_idx
            candidate = (score, -abs_row, -abs_col, abs_row)
            if best is None or candidate > best:
                best = candidate

    if best is None:
        return None

    return (best[3], -best[2])


def build_header_entries(
    sheet: xw.Sheet,
    header_row: int,
    start_col: int,
    end_col: int,
) -> List[Tuple[int, str]]:
    values = to_list(sheet.range((header_row, start_col), (header_row, end_col)).value)
    entries: List[Tuple[int, str]] = []
    for idx, value in enumerate(values):
        text = normalize_text(value)
        if text:
            entries.append((start_col + idx, text))
    return entries


def find_offset_by_patterns(
    header_entries: Sequence[Tuple[int, str]],
    anchor_col: int,
    patterns: Sequence[Sequence[str]],
) -> Optional[int]:
    for tokens in patterns:
        for col, text in header_entries:
            if all(token in text for token in tokens):
                return col - anchor_col
    return None


def read_block(
    sheet: xw.Sheet,
    start_row: int,
    end_row: int,
    start_col: int,
    end_col: int,
) -> List[List[Any]]:
    return to_rows(sheet.range((start_row, start_col), (end_row, end_col)).value)


def get_block_value(
    block: Sequence[Sequence[Any]],
    block_start_col: int,
    row_index: int,
    absolute_col: int,
) -> Any:
    if row_index < 0 or row_index >= len(block):
        return None
    row_values = block[row_index]
    col_index = absolute_col - block_start_col
    if col_index < 0 or col_index >= len(row_values):
        return None
    return row_values[col_index]


def collect_numeric_rows(sheet: xw.Sheet, col: int, start_row: int, end_row: int) -> List[int]:
    if end_row < start_row:
        return []
    values = to_list(sheet.range((start_row, col), (end_row, col)).value)
    rows: List[int] = []
    for idx, value in enumerate(values):
        if as_float(value) is not None:
            rows.append(start_row + idx)
    return rows


def collect_xy_history(
    sheet: xw.Sheet,
    x_col: int,
    y_col: int,
    start_row: int,
    end_row: int,
) -> List[Tuple[int, float, float]]:
    if end_row < start_row:
        return []
    x_values = to_list(sheet.range((start_row, x_col), (end_row, x_col)).value)
    y_values = to_list(sheet.range((start_row, y_col), (end_row, y_col)).value)
    history: List[Tuple[int, float, float]] = []
    for idx in range(min(len(x_values), len(y_values))):
        x_num = as_float(x_values[idx])
        y_num = as_float(y_values[idx])
        if x_num is None or y_num is None:
            continue
        history.append((start_row + idx, x_num, y_num))
    return history


def extract_empirical_rows(
    wb: xw.Book,
    meta: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = get_sheet(wb, "Empirical Model")
    if sheet is None:
        return []

    anchor = find_anchor_cell(sheet, anchor_text="max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    data_start_row = anchor_row + 1
    data_end_row = data_start_row + N_QUARTERS - 1

    start_col = max(1, anchor_col - 16)
    end_col = anchor_col + 10

    header_entries = build_header_entries(sheet, anchor_row, start_col, end_col)
    data_block = read_block(sheet, data_start_row, data_end_row, start_col, end_col)

    offsets: Dict[str, Optional[int]] = {
        "num_quarters_used": find_offset_by_patterns(
            header_entries,
            anchor_col,
            (("num", "quarter"), ("n", "quarter")),
        ),
        "last_quarter_used": find_offset_by_patterns(
            header_entries,
            anchor_col,
            (("last", "quarter"),),
        ),
        "forecast_value": find_offset_by_patterns(
            header_entries,
            anchor_col,
            (("estimated", "total", "sold"), ("forecast", "value"), ("tot", "fcst")),
        ),
        "actual_value": find_offset_by_patterns(
            header_entries,
            anchor_col,
            (("reported", "sales"), ("actual", "value"), ("actual",)),
        ),
        "forecast_min": find_offset_by_patterns(
            header_entries,
            anchor_col,
            (("min",),),
        ),
        "quarterly_sales": find_offset_by_patterns(
            header_entries,
            anchor_col,
            (("quarterly", "sales"),),
        ),
        "reported_sales": find_offset_by_patterns(
            header_entries,
            anchor_col,
            (("reported", "sales"),),
        ),
        "growth_rate_pct": find_offset_by_patterns(
            header_entries,
            anchor_col,
            (("growth", "rate"), ("growth",)),
        ),
        "sales_captured_in_db_pct": find_offset_by_patterns(
            header_entries,
            anchor_col,
            (("captured", "db"), ("db", "%"), ("sales", "captured")),
        ),
        "avg_penetration_existing": find_offset_by_patterns(
            header_entries,
            anchor_col,
            (("avg", "penetration"), ("average", "penetration")),
        ),
    }

    if offsets["forecast_min"] is None:
        offsets["forecast_min"] = 1
    if offsets["forecast_value"] is None:
        offsets["forecast_value"] = -1
    if offsets["actual_value"] is None:
        offsets["actual_value"] = -2

    def value_at(row_index: int, offset: Optional[int]) -> Any:
        if offset is None:
            return None
        absolute_col = anchor_col + offset
        block_value = get_block_value(data_block, start_col, row_index, absolute_col)
        if block_value is not None:
            return block_value
        absolute_row = data_start_row + row_index
        return sheet.range((absolute_row, absolute_col)).value

    penetration_col: Optional[int] = None
    if offsets["sales_captured_in_db_pct"] is not None:
        penetration_col = anchor_col + offsets["sales_captured_in_db_pct"]  # type: ignore[arg-type]
    if penetration_col is None:
        penetration_offset = find_offset_by_patterns(
            header_entries,
            anchor_col,
            (("penetration",),),
        )
        if penetration_offset is not None:
            penetration_col = anchor_col + penetration_offset

    temp_col = end_col + 2
    formula_rows = 0
    penetration_history_rows: List[int] = []
    if penetration_col is not None:
        history_start = max(1, anchor_row - 300)
        penetration_history_rows = collect_numeric_rows(sheet, penetration_col, history_start, anchor_row - 1)

    for i in range(N_QUARTERS):
        if not penetration_history_rows:
            break
        n_quarters = i + 1
        start_idx = max(0, len(penetration_history_rows) - n_quarters)
        start_row = penetration_history_rows[start_idx]
        end_row = penetration_history_rows[-1]
        calc_cell = sheet.range((data_start_row + i, temp_col))
        set_formula2(calc_cell, f"=AVERAGE(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})")
        formula_rows += 1

    if formula_rows:
        wb.app.calculate()

    formula_values = to_list(sheet.range((data_start_row, temp_col), (data_end_row, temp_col)).value)
    rows: List[Dict[str, Any]] = []

    for i in range(N_QUARTERS):
        forecast_max = as_float(value_at(i, 0))
        forecast_min = as_float(value_at(i, offsets["forecast_min"]))
        forecast_value = as_float(value_at(i, offsets["forecast_value"]))
        reported_sales = as_float(value_at(i, offsets["reported_sales"]))
        actual_value = as_float(value_at(i, offsets["actual_value"]))
        quarterly_sales = as_float(value_at(i, offsets["quarterly_sales"]))
        growth_rate_pct = as_float(value_at(i, offsets["growth_rate_pct"]))
        sales_captured_in_db_pct = as_float(value_at(i, offsets["sales_captured_in_db_pct"]))
        existing_avg_pen = as_float(value_at(i, offsets["avg_penetration_existing"]))
        num_quarters_raw = value_at(i, offsets["num_quarters_used"])
        last_quarter_used = value_at(i, offsets["last_quarter_used"])

        if actual_value is None:
            actual_value = reported_sales

        avg_pen_formula = as_float(formula_values[i]) if i < len(formula_values) else None
        avg_penetration_pct = avg_pen_formula
        if avg_penetration_pct is None:
            avg_penetration_pct = existing_avg_pen
        if avg_penetration_pct is None:
            avg_penetration_pct = sales_captured_in_db_pct

        if all(
            value is None
            for value in (
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                quarterly_sales,
                reported_sales,
            )
        ):
            continue

        num_quarters_num = as_float(num_quarters_raw)
        num_quarters_used: Any
        if num_quarters_num is None:
            num_quarters_used = i + 1
        elif float(num_quarters_num).is_integer():
            num_quarters_used = int(num_quarters_num)
        else:
            num_quarters_used = num_quarters_num

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        rows.append(
            {
                "model": meta["model"],
                "ticker": meta["ticker"],
                "model_period": meta["model_period"],
                "model_date": meta["model_date"],
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


def round_sig(value: Any) -> Any:
    if isinstance(value, float):
        return round(value, 10)
    return value


def extract_regression_rows(
    wb: xw.Book,
    meta: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = get_sheet(wb, "Regression Model")
    if sheet is None:
        return []

    anchor = find_anchor_cell(sheet, anchor_text="max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    history_start = max(1, anchor_row - 300)
    xy_history = collect_xy_history(sheet, x_col, y_col, history_start, anchor_row - 1)
    if len(xy_history) < 2:
        return []

    data_start_row = anchor_row + 1
    data_end_row = data_start_row + N_QUARTERS - 1
    start_col = max(1, anchor_col - 14)
    end_col = anchor_col + 5

    header_entries = build_header_entries(sheet, anchor_row, start_col, end_col)
    data_block = read_block(sheet, data_start_row, data_end_row, start_col, end_col)

    offsets: Dict[str, Optional[int]] = {
        "num_quarters_used": find_offset_by_patterns(
            header_entries,
            anchor_col,
            (("num", "quarter"), ("n", "quarter")),
        ),
        "forecast_total_without_sa": find_offset_by_patterns(
            header_entries,
            anchor_col,
            (("tot", "fcst", "w", "o", "sa"), ("tot", "fcst"), ("forecast", "without", "sa")),
        ),
        "actual_value": find_offset_by_patterns(
            header_entries,
            anchor_col,
            (("actual",), ("reported", "sales")),
        ),
        "forecast_min": find_offset_by_patterns(
            header_entries,
            anchor_col,
            (("min",),),
        ),
    }

    if offsets["forecast_min"] is None:
        offsets["forecast_min"] = 1
    if offsets["forecast_total_without_sa"] is None:
        offsets["forecast_total_without_sa"] = -1

    def value_at(row_index: int, offset: Optional[int]) -> Any:
        if offset is None:
            return None
        absolute_col = anchor_col + offset
        block_value = get_block_value(data_block, start_col, row_index, absolute_col)
        if block_value is not None:
            return block_value
        absolute_row = data_start_row + row_index
        return sheet.range((absolute_row, absolute_col)).value

    temp_intercept_col = end_col + 2
    temp_slope_col = end_col + 3
    temp_forecast_col = end_col + 4
    latest_x = xy_history[-1][1]

    for i in range(N_QUARTERS):
        n_quarters = i + 1
        if len(xy_history) < n_quarters:
            continue

        start_row = xy_history[-n_quarters][0]
        end_row = xy_history[-1][0]
        calc_row = data_start_row + i

        intercept_cell = sheet.range((calc_row, temp_intercept_col))
        slope_cell = sheet.range((calc_row, temp_slope_col))
        forecast_cell = sheet.range((calc_row, temp_forecast_col))

        set_formula2(
            intercept_cell,
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})",
        )
        set_formula2(
            slope_cell,
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})",
        )
        set_formula2(
            forecast_cell,
            f"=R{calc_row}C{temp_intercept_col}+R{calc_row}C{temp_slope_col}*{latest_x}",
        )

    wb.app.calculate()

    intercept_values = to_list(
        sheet.range((data_start_row, temp_intercept_col), (data_end_row, temp_intercept_col)).value
    )
    slope_values = to_list(
        sheet.range((data_start_row, temp_slope_col), (data_end_row, temp_slope_col)).value
    )
    forecast_formula_values = to_list(
        sheet.range((data_start_row, temp_forecast_col), (data_end_row, temp_forecast_col)).value
    )

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for i in range(N_QUARTERS):
        n_quarters = i + 1
        if len(xy_history) < n_quarters:
            continue

        num_quarters_raw = value_at(i, offsets["num_quarters_used"])
        num_quarters_num = as_float(num_quarters_raw)
        if num_quarters_num is None:
            num_quarters_used: Any = n_quarters
        elif num_quarters_num.is_integer():
            num_quarters_used = int(num_quarters_num)
        else:
            num_quarters_used = num_quarters_num

        forecast_value = as_float(value_at(i, offsets["forecast_total_without_sa"]))
        if forecast_value is None and i < len(forecast_formula_values):
            forecast_value = as_float(forecast_formula_values[i])

        actual_value = as_float(value_at(i, offsets["actual_value"]))
        forecast_max = as_float(value_at(i, 0))
        forecast_min = as_float(value_at(i, offsets["forecast_min"]))
        intercept = as_float(intercept_values[i]) if i < len(intercept_values) else None
        slope = as_float(slope_values[i]) if i < len(slope_values) else None

        if all(value is None for value in (forecast_value, forecast_max, forecast_min, intercept, slope)):
            continue

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        signature = (
            round_sig(num_quarters_used),
            round_sig(forecast_value),
            round_sig(forecast_max),
            round_sig(forecast_min),
            round_sig(intercept),
            round_sig(slope),
        )
        if previous_signature == signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": meta["model"],
                "ticker": meta["ticker"],
                "model_period": meta["model_period"],
                "model_date": meta["model_date"],
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


def infer_column_widths(
    columns: Sequence[str],
    rows: Sequence[Dict[str, Any]],
    sample_size: int = 200,
) -> List[float]:
    widths: List[float] = []
    sample_rows = rows[:sample_size]
    for col in columns:
        max_len = len(col)
        for row in sample_rows:
            value = row.get(col)
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        widths.append(float(min(max(12, max_len + 2), 48)))
    return widths


def format_sheet(sheet: xw.Sheet, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    sheet.range((1, 1)).value = list(columns)
    if rows:
        data = [[row.get(col) for col in columns] for row in rows]
        sheet.range((2, 1)).value = data

    sheet.range((1, 1), (1, len(columns))).api.Font.Bold = True

    last_row = max(2, len(rows) + 1)
    sheet.range((1, 1), (last_row, len(columns))).api.AutoFilter()

    try:
        sheet.activate()
        active_window = sheet.api.Application.ActiveWindow
        active_window.SplitRow = 1
        active_window.FreezePanes = True
    except Exception:
        pass

    widths = infer_column_widths(columns, rows)
    for col_idx, width in enumerate(widths, start=1):
        sheet.range((1, col_idx)).column_width = width


def process_source_workbook(
    app: xw.App,
    file_path: Path,
) -> Tuple[bool, List[Dict[str, Any]], List[Dict[str, Any]]]:
    wb: Optional[xw.Book] = None
    try:
        wb = app.books.open(str(file_path), update_links=False)
        meta = parse_file_label(file_path.name)
        empirical_rows = extract_empirical_rows(wb, meta, file_path.name)
        regression_rows = extract_regression_rows(wb, meta, file_path.name)
        return True, empirical_rows, regression_rows
    except Exception as exc:
        print(f"Skipped {file_path.name}: {exc}")
        return False, [], []
    finally:
        if wb is not None:
            close_source_workbook(wb)


def write_output_workbook(
    app: xw.App,
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    out_wb = app.books.add()
    try:
        empirical_sheet = out_wb.sheets[0]
        empirical_sheet.name = "empirical_candidates"
        regression_sheet = out_wb.sheets.add("regression_candidates", after=empirical_sheet)

        format_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)
        format_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)

        out_wb.save(str(output_path))
    finally:
        try:
            out_wb.close()
        except Exception:
            pass


def iter_source_files(input_path: Path) -> Iterable[Path]:
    for file_path in sorted(input_path.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped {file_path.name}: not an .xlsx file")
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped {file_path.name}: temporary Excel file")
            continue
        yield file_path


def main() -> None:
    input_path = Path(input_dir).expanduser()
    output_path = Path(output_dir).expanduser()

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"Input directory does not exist: {input_path}")

    output_path.mkdir(parents=True, exist_ok=True)
    output_file = get_output_path(input_path, output_path)

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.enable_events = False
    previous_calc_mode = app.calculation
    app.calculation = "manual"

    files_processed = 0
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []

    try:
        for file_path in iter_source_files(input_path):
            if output_path == input_path and file_path.name.startswith(f"{input_path.name}_PARAM"):
                print(f"Skipped {file_path.name}: output workbook pattern")
                continue

            ok, empirical_part, regression_part = process_source_workbook(app, file_path)
            if not ok:
                continue

            empirical_rows.extend(empirical_part)
            regression_rows.extend(regression_part)
            files_processed += 1
            print(f"Processed {file_path.name}")

        write_output_workbook(app, output_file, empirical_rows, regression_rows)
    finally:
        try:
            app.calculation = previous_calc_mode
        except Exception:
            pass
        app.quit()

    print(f"Output path: {output_file}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

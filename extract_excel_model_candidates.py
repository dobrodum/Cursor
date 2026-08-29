from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# Configure these two paths before running.
input_dir = "./input"
output_dir = "./output"


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


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip()).lower()


def normalize_header(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", " ", str(value).lower()).strip()


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def first_non_blank(*values: Any) -> Any:
    for value in values:
        if not is_blank(value):
            return value
    return None


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
            return None
        if text.endswith("%"):
            try:
                return float(text[:-1]) / 100.0
            except ValueError:
                return None
        try:
            return float(text)
        except ValueError:
            return None
    return None


def to_int(value: Any, fallback: int) -> int:
    num = to_float(value)
    if num is None:
        return fallback
    return int(round(num))


def ensure_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values
    return [values]


def find_max_anchor(sheet: xw.Sheet) -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    values = ensure_2d(used.value)
    if not values:
        return None

    base_row = used.row
    base_col = used.column
    for r_idx, row_values in enumerate(values):
        for c_idx, cell_value in enumerate(row_values):
            if normalize_text(cell_value) == "max":
                return base_row + r_idx, base_col + c_idx
    return None


def canonical_header(value: Any) -> Optional[str]:
    text = normalize_header(value)
    if not text:
        return None

    if text == "max":
        return "forecast_max"
    if text == "min":
        return "forecast_min"
    if "num quarter" in text:
        return "num_quarters_used"
    if "last quarter" in text:
        return "last_quarter_used"
    if "avg penetration" in text:
        return "avg_penetration_pct"
    if "penetration" in text and "avg" not in text:
        return "penetration_series"
    if "quarterly sales" in text:
        return "quarterly_sales"
    if "reported sales" in text:
        return "reported_sales"
    if "growth rate" in text:
        return "growth_rate_pct"
    if "sales captured" in text and "db" in text:
        return "sales_captured_in_db_pct"
    if "estimated total sold" in text:
        return "forecast_value"
    if "forecast total sold" in text:
        return "forecast_value"
    if "actual" in text and "sales" in text:
        return "actual_value"
    if "tot fcst" in text and "sa" in text:
        return "forecast_total_without_sa"
    if "without sa" in text and "forecast" in text:
        return "forecast_total_without_sa"
    if text == "intercept":
        return "intercept"
    if text == "slope":
        return "slope"
    return None


def build_header_offsets(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    search_left: int = 24,
    search_right: int = 24,
) -> Dict[str, int]:
    start_col = max(1, anchor_col - search_left)
    end_col = anchor_col + search_right
    row_values = sheet.range((anchor_row, start_col), (anchor_row, end_col)).value
    headers = row_values if isinstance(row_values, list) else [row_values]

    offsets: Dict[str, int] = {"forecast_max": 0}
    for idx, header_value in enumerate(headers):
        key = canonical_header(header_value)
        if key and key not in offsets:
            absolute_col = start_col + idx
            offsets[key] = absolute_col - anchor_col
    return offsets


def get_sheet_case_insensitive(workbook: xw.Book, name: str) -> Optional[xw.Sheet]:
    needle = name.strip().lower()
    for sheet in workbook.sheets:
        if sheet.name.strip().lower() == needle:
            return sheet
    return None


def read_rows_block(
    sheet: xw.Sheet,
    row_start: int,
    row_end: int,
    absolute_cols: Sequence[int],
) -> List[Dict[int, Any]]:
    if row_end < row_start or not absolute_cols:
        return []

    min_col = min(absolute_cols)
    max_col = max(absolute_cols)
    matrix = ensure_2d(sheet.range((row_start, min_col), (row_end, max_col)).value)

    rows: List[Dict[int, Any]] = []
    for row_values in matrix:
        row_map: Dict[int, Any] = {}
        for col in absolute_cols:
            row_map[col] = row_values[col - min_col] if (col - min_col) < len(row_values) else None
        rows.append(row_map)
    return rows


def set_formula2_r1c1(target_cell: xw.Range, formula_r1c1: str) -> None:
    try:
        target_cell.api.Formula2R1C1 = formula_r1c1
    except Exception:
        try:
            target_cell.formula2 = formula_r1c1
        except Exception:
            target_cell.formula = formula_r1c1


def calculate_avg_penetration_series(
    sheet: xw.Sheet,
    penetration_values: List[float],
    n_quarters: int,
) -> Dict[int, Any]:
    if not penetration_values:
        return {}

    values = penetration_values[-max(len(penetration_values), n_quarters) :]
    value_count = len(values)
    if value_count == 0:
        return {}

    # Far-right scratch area; writes are temporary and source workbook is closed without saving.
    scratch_col = 16370
    sheet.range((1, scratch_col), (value_count, scratch_col)).value = [[v] for v in values]

    for n in range(1, n_quarters + 1):
        result_cell = sheet.range((n, scratch_col + 1))
        if n <= value_count:
            start_row = value_count - n + 1
            end_row = value_count
            formula = (
                f'=IFERROR(AVERAGE(R{start_row}C{scratch_col}:R{end_row}C{scratch_col}),"")'
            )
            set_formula2_r1c1(result_cell, formula)
        else:
            result_cell.value = None

    sheet.book.app.calculate()
    raw_results = sheet.range((1, scratch_col + 1), (n_quarters, scratch_col + 1)).value
    if not isinstance(raw_results, list):
        raw_results = [raw_results]

    results = {idx + 1: raw_results[idx] for idx in range(len(raw_results))}
    sheet.range((1, scratch_col), (max(value_count, n_quarters), scratch_col + 1)).clear_contents()
    return results


def calculate_regression_series(
    sheet: xw.Sheet,
    x_values: List[float],
    y_values: List[float],
    max_n: int,
) -> Dict[int, Dict[str, Any]]:
    pair_count = min(len(x_values), len(y_values))
    if pair_count < 2:
        return {}

    x_vals = x_values[-pair_count:]
    y_vals = y_values[-pair_count:]

    scratch_col = 16360
    sheet.range((1, scratch_col), (pair_count, scratch_col + 1)).value = [
        [x_vals[i], y_vals[i]] for i in range(pair_count)
    ]

    for n in range(1, max_n + 1):
        intercept_cell = sheet.range((n, scratch_col + 2))
        slope_cell = sheet.range((n, scratch_col + 3))
        if 2 <= n <= pair_count:
            start_row = pair_count - n + 1
            end_row = pair_count
            intercept_formula = (
                f'=IFERROR(INTERCEPT(R{start_row}C{scratch_col + 1}:R{end_row}C{scratch_col + 1},'
                f'R{start_row}C{scratch_col}:R{end_row}C{scratch_col}),"")'
            )
            slope_formula = (
                f'=IFERROR(SLOPE(R{start_row}C{scratch_col + 1}:R{end_row}C{scratch_col + 1},'
                f'R{start_row}C{scratch_col}:R{end_row}C{scratch_col}),"")'
            )
            set_formula2_r1c1(intercept_cell, intercept_formula)
            set_formula2_r1c1(slope_cell, slope_formula)
        else:
            intercept_cell.value = None
            slope_cell.value = None

    sheet.book.app.calculate()
    intercepts = sheet.range((1, scratch_col + 2), (max_n, scratch_col + 2)).value
    slopes = sheet.range((1, scratch_col + 3), (max_n, scratch_col + 3)).value
    if not isinstance(intercepts, list):
        intercepts = [intercepts]
    if not isinstance(slopes, list):
        slopes = [slopes]

    results: Dict[int, Dict[str, Any]] = {}
    for n in range(1, max_n + 1):
        results[n] = {
            "intercept": intercepts[n - 1] if n - 1 < len(intercepts) else None,
            "slope": slopes[n - 1] if n - 1 < len(slopes) else None,
        }

    sheet.range((1, scratch_col), (max(pair_count, max_n), scratch_col + 3)).clear_contents()
    return results


def calc_range_width(max_value: Any, min_value: Any) -> Optional[float]:
    max_num = to_float(max_value)
    min_num = to_float(min_value)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


def parse_file_labels(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    parts = stem.split(" - ")

    ticker = parts[1].strip() if len(parts) >= 2 and parts[1].strip() else "UNKNOWN"
    period_chunk = parts[2] if len(parts) >= 3 else stem
    period_token = period_chunk.split("_")[0].strip()

    match = re.search(
        r"(?i)\b(early|mid|late)\s*"
        r"(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|"
        r"jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|"
        r"nov(?:ember)?|dec(?:ember)?)\s*(\d{4})\b",
        period_token,
    )

    if not match:
        model_period = "Unknown_0000"
        model_date = ""
    else:
        period_label = match.group(1).title()
        month_text = match.group(2).title()
        year = int(match.group(3))

        month_num = datetime.strptime(month_text[:3], "%b").month
        month_short = datetime(year, month_num, 1).strftime("%b")

        day_map = {"early": 5, "mid": 15, "late": 25}
        day = day_map[period_label.lower()]

        model_period = f"{period_label}{month_short}_{year}"
        model_date = f"{year:04d}-{month_num:02d}-{day:02d}"

    model = f"{ticker}_{model_period}"
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "source_file": file_path.name,
    }


def extract_empirical_candidates(sheet: xw.Sheet, meta: Dict[str, str]) -> List[Dict[str, Any]]:
    anchor = find_max_anchor(sheet)
    if anchor is None:
        print(f"  skipped empirical extraction: 'max' anchor not found in sheet '{sheet.name}'")
        return []

    anchor_row, anchor_col = anchor
    n_quarters = 10
    offsets = build_header_offsets(sheet, anchor_row, anchor_col)

    key_defaults: Dict[str, Optional[int]] = {
        "num_quarters_used": -5,
        "last_quarter_used": -4,
        "forecast_value": -2,
        "actual_value": -3,
        "forecast_max": 0,
        "forecast_min": 1,
        "avg_penetration_pct": -6,
        "quarterly_sales": -7,
        "reported_sales": -3,
        "growth_rate_pct": -8,
        "sales_captured_in_db_pct": -9,
    }

    resolved_offsets: Dict[str, Optional[int]] = {}
    for key, fallback in key_defaults.items():
        resolved_offsets[key] = offsets.get(key, fallback)

    absolute_cols = sorted(
        {anchor_col + off for off in resolved_offsets.values() if off is not None and anchor_col + off > 0}
    )
    rows = read_rows_block(sheet, anchor_row + 1, anchor_row + n_quarters, absolute_cols)

    penetration_offset = offsets.get("penetration_series")
    if penetration_offset is None:
        penetration_offset = resolved_offsets.get("avg_penetration_pct")
    penetration_col = anchor_col + penetration_offset if penetration_offset is not None else None

    penetration_history: List[float] = []
    if penetration_col is not None and penetration_col > 0 and anchor_row > 1:
        hist_start = max(1, anchor_row - 48)
        hist_values = sheet.range((hist_start, penetration_col), (anchor_row - 1, penetration_col)).value
        if not isinstance(hist_values, list):
            hist_values = [hist_values]
        for value in hist_values:
            numeric = to_float(value)
            if numeric is not None:
                penetration_history.append(numeric)

    avg_by_quarter = calculate_avg_penetration_series(sheet, penetration_history, n_quarters)

    output_rows: List[Dict[str, Any]] = []
    for idx, row_map in enumerate(rows, start=1):
        def get_value(key: str) -> Any:
            offset = resolved_offsets.get(key)
            if offset is None:
                return None
            return row_map.get(anchor_col + offset)

        num_quarters_used = to_int(get_value("num_quarters_used"), fallback=idx)
        forecast_value = first_non_blank(get_value("forecast_value"), get_value("quarterly_sales"))
        actual_value = first_non_blank(get_value("actual_value"), get_value("reported_sales"))
        forecast_max = get_value("forecast_max")
        forecast_min = get_value("forecast_min")
        avg_penetration = first_non_blank(get_value("avg_penetration_pct"), avg_by_quarter.get(idx))
        reported_sales = first_non_blank(get_value("reported_sales"), actual_value)

        has_signal = any(
            not is_blank(v)
            for v in [forecast_value, actual_value, forecast_max, forecast_min, avg_penetration]
        )
        if not has_signal:
            continue

        output_rows.append(
            {
                "model": meta["model"],
                "ticker": meta["ticker"],
                "model_period": meta["model_period"],
                "model_date": meta["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": get_value("last_quarter_used"),
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": calc_range_width(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": get_value("quarterly_sales"),
                "reported_sales": reported_sales,
                "growth_rate_pct": get_value("growth_rate_pct"),
                "sales_captured_in_db_pct": get_value("sales_captured_in_db_pct"),
                "source_file": meta["source_file"],
            }
        )

    return output_rows


def extract_regression_candidates(sheet: xw.Sheet, meta: Dict[str, str]) -> List[Dict[str, Any]]:
    anchor = find_max_anchor(sheet)
    if anchor is None:
        print(f"  skipped regression extraction: 'max' anchor not found in sheet '{sheet.name}'")
        return []

    anchor_row, anchor_col = anchor
    max_rows = 10
    offsets = build_header_offsets(sheet, anchor_row, anchor_col)

    key_defaults: Dict[str, Optional[int]] = {
        "num_quarters_used": -5,
        "forecast_total_without_sa": -2,
        "actual_value": -3,
        "forecast_max": 0,
        "forecast_min": 1,
        "intercept": 2,
        "slope": 3,
    }

    resolved_offsets: Dict[str, Optional[int]] = {}
    for key, fallback in key_defaults.items():
        resolved_offsets[key] = offsets.get(key, fallback)

    absolute_cols = sorted(
        {anchor_col + off for off in resolved_offsets.values() if off is not None and anchor_col + off > 0}
    )
    table_rows = read_rows_block(sheet, anchor_row + 1, anchor_row + max_rows, absolute_cols)

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    x_values: List[float] = []
    y_values: List[float] = []
    if anchor_row > 1 and x_col > 0 and y_col > 0:
        data_start = max(1, anchor_row - 96)
        pair_data = sheet.range((data_start, min(x_col, y_col)), (anchor_row - 1, max(x_col, y_col))).value
        matrix = ensure_2d(pair_data)
        for row in matrix:
            x_raw = row[x_col - min(x_col, y_col)] if row else None
            y_raw = row[y_col - min(x_col, y_col)] if row else None
            x_num = to_float(x_raw)
            y_num = to_float(y_raw)
            if x_num is not None and y_num is not None:
                x_values.append(x_num)
                y_values.append(y_num)

    regression_by_quarter = calculate_regression_series(sheet, x_values, y_values, max_rows)

    output_rows: List[Dict[str, Any]] = []
    for idx, row_map in enumerate(table_rows, start=1):
        def get_value(key: str) -> Any:
            offset = resolved_offsets.get(key)
            if offset is None:
                return None
            return row_map.get(anchor_col + offset)

        num_quarters_used = to_int(get_value("num_quarters_used"), fallback=idx)
        calc_stats = regression_by_quarter.get(num_quarters_used, regression_by_quarter.get(idx, {}))
        intercept = first_non_blank(get_value("intercept"), calc_stats.get("intercept"))
        slope = first_non_blank(get_value("slope"), calc_stats.get("slope"))
        forecast_value = get_value("forecast_total_without_sa")
        actual_value = get_value("actual_value")
        forecast_max = get_value("forecast_max")
        forecast_min = get_value("forecast_min")

        has_signal = any(
            not is_blank(v)
            for v in [forecast_value, forecast_max, forecast_min, intercept, slope]
        )
        if not has_signal:
            continue

        output_rows.append(
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
                "range_width": calc_range_width(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": meta["source_file"],
            }
        )

    # Remove duplicate trailing row if it matches the previous one.
    if len(output_rows) >= 2:
        prev = output_rows[-2]
        last = output_rows[-1]
        check_fields = ["forecast_value", "forecast_max", "forecast_min", "intercept", "slope"]
        if all(first_non_blank(prev.get(field)) == first_non_blank(last.get(field)) for field in check_fields):
            output_rows.pop()

    return output_rows


def close_without_saving(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        workbook.api.Close(SaveChanges=0)
        return
    except Exception:
        pass

    try:
        workbook.close()
    except Exception:
        pass


def write_output_workbook(
    destination: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    sheet_map = [
        ("empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows),
        ("regression_candidates", REGRESSION_COLUMNS, regression_rows),
    ]

    for sheet_name, columns, rows in sheet_map:
        ws = wb.create_sheet(sheet_name)
        ws.append(columns)
        for row in rows:
            ws.append([row.get(col) for col in columns])

        for header_cell in ws[1]:
            header_cell.font = Font(bold=True)

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for col_idx, column_name in enumerate(columns, start=1):
            max_len = len(column_name)
            for row_idx in range(2, ws.max_row + 1):
                value = ws.cell(row=row_idx, column=col_idx).value
                if value is None:
                    continue
                max_len = max(max_len, len(str(value)))
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 48)

    wb.save(destination)


def next_output_path(input_path: Path, output_path: Path) -> Path:
    base_name = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        numbered = output_path / f"{base_name}.{suffix}.xlsx"
        if not numbered.exists():
            return numbered
        suffix += 1


def iter_source_files(path: Path) -> List[Path]:
    files = sorted(path.iterdir(), key=lambda p: p.name.lower())
    return [file_path for file_path in files if file_path.is_file()]


def main() -> None:
    source_dir = Path(input_dir).expanduser().resolve()
    target_dir = Path(output_dir).expanduser().resolve()
    target_dir.mkdir(parents=True, exist_ok=True)

    if not source_dir.exists():
        print(f"Input directory not found: {source_dir}")
        return

    source_files = iter_source_files(source_dir)
    output_file = next_output_path(source_dir, target_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in source_files:
            name = file_path.name
            if name.startswith("~"):
                print(f"Skipping file: {name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipping file: {name} (not .xlsx)")
                continue

            print(f"Processing file: {name}")
            workbook: Optional[xw.Book] = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                meta = parse_file_labels(file_path)

                empirical_sheet = get_sheet_case_insensitive(workbook, "Empirical Model")
                if empirical_sheet is None:
                    print("  skipped empirical extraction: sheet 'Empirical Model' not found")
                else:
                    empirical_rows.extend(extract_empirical_candidates(empirical_sheet, meta))

                regression_sheet = get_sheet_case_insensitive(workbook, "Regression Model")
                if regression_sheet is None:
                    print("  skipped regression extraction: sheet 'Regression Model' not found")
                else:
                    regression_rows.extend(extract_regression_candidates(regression_sheet, meta))

                processed_files += 1
            except Exception as exc:
                print(f"Skipping file: {name} (error: {exc})")
            finally:
                if workbook is not None:
                    close_without_saving(workbook)
    finally:
        app.quit()

    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"Output path: {output_file}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

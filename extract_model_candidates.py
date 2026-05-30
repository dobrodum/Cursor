from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# ===== User-configurable paths =====
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


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip().lower())


def safe_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def safe_int(value: Any) -> Optional[int]:
    fval = safe_float(value)
    if fval is None:
        return None
    return int(round(fval))


def clean_value(value: Any) -> Any:
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def month_from_token(month_token: str) -> Optional[int]:
    token = month_token[:3].title()
    try:
        return datetime.strptime(token, "%b").month
    except ValueError:
        return None


def parse_file_labels(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = parts[1] if len(parts) >= 2 and parts[1] else "UNKNOWN"
    period_token = parts[2] if len(parts) >= 3 else ""
    period_token = period_token.split("_")[0]

    model_period = "UNKNOWN_PERIOD"
    model_date = ""

    match = re.search(r"(Early|Mid|Late)([A-Za-z]+)(\d{4})", period_token, flags=re.IGNORECASE)
    if match:
        period_bucket = match.group(1).title()
        month_text = match.group(2).title()
        year = int(match.group(3))
        month_number = month_from_token(month_text)
        day_map = {"Early": 5, "Mid": 15, "Late": 25}
        day = day_map.get(period_bucket, 15)
        if month_number is not None:
            model_period = f"{period_bucket}{month_text[:3]}_{year}"
            model_date = date(year, month_number, day).isoformat()
    model = f"{ticker}_{model_period}"

    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def safe_close_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.close(SaveChanges=False)
        return
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


def output_workbook_path(input_folder: Path, output_folder: Path) -> Path:
    output_folder.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_folder.name}_PARAM"
    candidate = output_folder / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    idx = 1
    while True:
        numbered = output_folder / f"{base_name}.{idx}.xlsx"
        if not numbered.exists():
            return numbered
        idx += 1


def flatten_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, tuple):
        values = list(values)
    if not isinstance(values, list):
        return [[values]]
    if values and isinstance(values[0], tuple):
        values = [list(row) for row in values]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def find_anchor_cell(sheet: xw.Sheet, anchor_label: str = "max") -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    data = flatten_2d(used.value)
    for r_idx, row in enumerate(data):
        for c_idx, value in enumerate(row):
            if normalize_text(value) == anchor_label:
                return used.row + r_idx, used.column + c_idx
    return None


def sheet_bounds(sheet: xw.Sheet) -> Tuple[int, int, int, int]:
    used = sheet.used_range
    start_row = used.row
    start_col = used.column
    end_row = start_row + used.rows.count - 1
    end_col = start_col + used.columns.count - 1
    return start_row, start_col, end_row, end_col


def row_values(sheet: xw.Sheet, row: int, col_start: int, col_end: int) -> List[Any]:
    values = sheet.range((row, col_start), (row, col_end)).value
    if isinstance(values, tuple):
        return list(values)
    if isinstance(values, list):
        return values
    return [values]


def build_header_map(sheet: xw.Sheet, header_row: int, col_start: int, col_end: int) -> Dict[int, str]:
    values = row_values(sheet, header_row, col_start, col_end)
    mapping: Dict[int, str] = {}
    for offset, value in enumerate(values):
        mapping[col_start + offset] = normalize_text(value)
    return mapping


def find_column_by_keywords(
    header_map: Dict[int, str], keyword_options: Iterable[Tuple[str, ...]], anchor_col: int
) -> Optional[int]:
    scored: List[Tuple[int, int]] = []
    for col, label in header_map.items():
        if not label:
            continue
        for keywords in keyword_options:
            if all(keyword in label for keyword in keywords):
                scored.append((abs(col - anchor_col), col))
                break
    if not scored:
        return None
    scored.sort()
    return scored[0][1]


def percent_to_ratio(value: Optional[float]) -> Optional[float]:
    if value is None:
        return None
    if abs(value) > 1:
        return value / 100.0
    return value


def values_match(v1: Optional[float], v2: Optional[float], tolerance: float = 1e-10) -> bool:
    if v1 is None and v2 is None:
        return True
    if v1 is None or v2 is None:
        return False
    return abs(v1 - v2) <= tolerance


def process_empirical_sheet(
    wb: xw.Book,
    file_path: Path,
    labels: Dict[str, str],
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        return []

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    start_row, start_col, end_row, end_col = sheet_bounds(sheet)
    header_map = build_header_map(sheet, anchor_row, start_col, end_col)

    min_col = find_column_by_keywords(header_map, [("min",)], anchor_col) or (anchor_col + 1)
    num_q_col = find_column_by_keywords(
        header_map,
        [("num", "quarter"), ("quarters", "used"), ("n", "quarters"), ("qtrs",)],
        anchor_col,
    ) or (anchor_col - 8)
    last_q_col = find_column_by_keywords(
        header_map,
        [("last", "quarter"), ("last", "qtr"), ("quarter", "used")],
        anchor_col,
    )
    forecast_col = find_column_by_keywords(
        header_map,
        [("estimated", "total"), ("tot", "fcst"), ("forecast",), ("total", "sold")],
        anchor_col,
    )
    actual_col = find_column_by_keywords(
        header_map,
        [("reported", "sales"), ("actual", "sales"), ("actual",)],
        anchor_col,
    )
    avg_pen_col = find_column_by_keywords(
        header_map,
        [("avg", "penetration"), ("average", "penetration")],
        anchor_col,
    )
    quarterly_sales_col = find_column_by_keywords(
        header_map,
        [("quarterly", "sales"), ("sales", "captured"), ("db", "sales")],
        anchor_col,
    )
    growth_col = find_column_by_keywords(
        header_map,
        [("growth", "rate"), ("growth",)],
        anchor_col,
    )
    captured_col = find_column_by_keywords(
        header_map,
        [("captured", "db"), ("sales", "captured", "db"), ("penetration", "pct"), ("penetration",)],
        anchor_col,
    )
    reported_sales_col = actual_col

    n_quarters = 10
    data_start = anchor_row + 1
    data_end = min(end_row, data_start + n_quarters - 1)

    temp_avg_col = end_col + 2
    formula_rows: List[Tuple[int, int]] = []

    if captured_col is not None:
        for row in range(data_start, data_end + 1):
            num_quarters_used = safe_int(sheet.range((row, num_q_col)).value)
            if num_quarters_used is None or num_quarters_used <= 0:
                num_quarters_used = row - data_start + 1
            num_quarters_used = min(num_quarters_used, n_quarters)

            start_formula_row = max(data_start, row - num_quarters_used + 1)
            formula = (
                f"=AVERAGE(R{start_formula_row}C{captured_col}:R{row}C{captured_col})"
            )
            sheet.range((row, temp_avg_col)).formula2 = formula
            formula_rows.append((row, num_quarters_used))

        if formula_rows:
            wb.app.calculate()

    results: List[Dict[str, Any]] = []
    for row in range(data_start, data_end + 1):
        num_quarters_used = safe_int(sheet.range((row, num_q_col)).value)
        if num_quarters_used is None or num_quarters_used <= 0:
            num_quarters_used = row - data_start + 1

        forecast_max = safe_float(sheet.range((row, anchor_col)).value)
        forecast_min = safe_float(sheet.range((row, min_col)).value) if min_col is not None else None

        avg_penetration_pct = safe_float(sheet.range((row, temp_avg_col)).value)
        if avg_penetration_pct is None and avg_pen_col is not None:
            avg_penetration_pct = safe_float(sheet.range((row, avg_pen_col)).value)

        quarterly_sales = (
            safe_float(sheet.range((row, quarterly_sales_col)).value)
            if quarterly_sales_col is not None
            else None
        )
        reported_sales = (
            safe_float(sheet.range((row, reported_sales_col)).value)
            if reported_sales_col is not None
            else None
        )
        growth_rate_pct = safe_float(sheet.range((row, growth_col)).value) if growth_col is not None else None
        sales_captured_in_db_pct = (
            safe_float(sheet.range((row, captured_col)).value)
            if captured_col is not None
            else None
        )
        last_quarter_used = (
            sheet.range((row, last_q_col)).value if last_q_col is not None else None
        )

        forecast_value = (
            safe_float(sheet.range((row, forecast_col)).value) if forecast_col is not None else None
        )
        actual_value = safe_float(sheet.range((row, actual_col)).value) if actual_col is not None else None

        if forecast_value is None and quarterly_sales is not None and avg_penetration_pct is not None:
            ratio = percent_to_ratio(avg_penetration_pct)
            if ratio not in (None, 0):
                forecast_value = quarterly_sales / ratio

        if actual_value is None:
            actual_value = reported_sales

        if forecast_max is None and forecast_value is not None:
            forecast_max = forecast_value
        if forecast_min is None and forecast_value is not None:
            forecast_min = forecast_value

        if (
            forecast_value is None
            and actual_value is None
            and forecast_max is None
            and forecast_min is None
            and avg_penetration_pct is None
        ):
            continue

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        result = {
            "model": labels["model"],
            "ticker": labels["ticker"],
            "model_period": labels["model_period"],
            "model_date": labels["model_date"],
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
            "source_file": file_path.name,
        }
        results.append({k: clean_value(v) for k, v in result.items()})

    return results


def collect_numeric_pairs(sheet: xw.Sheet, x_col: int, y_col: int) -> List[Tuple[int, float, float]]:
    start_row, _, end_row, _ = sheet_bounds(sheet)
    pairs: List[Tuple[int, float, float]] = []
    for row in range(start_row, end_row + 1):
        x_val = safe_float(sheet.range((row, x_col)).value)
        y_val = safe_float(sheet.range((row, y_col)).value)
        if x_val is None or y_val is None:
            continue
        pairs.append((row, x_val, y_val))
    return pairs


def process_regression_sheet(
    wb: xw.Book,
    file_path: Path,
    labels: Dict[str, str],
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        return []

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    start_row, start_col, end_row, end_col = sheet_bounds(sheet)
    header_map = build_header_map(sheet, anchor_row, start_col, end_col)

    min_col = find_column_by_keywords(header_map, [("min",)], anchor_col) or (anchor_col + 1)
    num_q_col = find_column_by_keywords(
        header_map,
        [("num", "quarter"), ("quarters", "used"), ("n", "quarters"), ("qtrs",)],
        anchor_col,
    )
    fcst_col = find_column_by_keywords(
        header_map,
        [("tot", "fcst", "w/o", "sa"), ("forecast", "without", "sa"), ("tot", "fcst")],
        anchor_col,
    )
    actual_col = find_column_by_keywords(
        header_map,
        [("actual", "sales"), ("reported", "sales"), ("actual",)],
        anchor_col,
    )

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if x_col < 1 or y_col < 1:
        return []
    pairs = collect_numeric_pairs(sheet, x_col, y_col)
    if len(pairs) < 2:
        return []

    n_quarters_max = min(10, len(pairs))
    n_values = list(range(2, n_quarters_max + 1))
    if not n_values:
        return []

    intercept_col = end_col + 2
    slope_col = end_col + 3
    temp_base_row = max(anchor_row + 1, start_row)

    for idx, n_quarters in enumerate(n_values):
        subset = pairs[-n_quarters:]
        start_data_row = subset[0][0]
        end_data_row = subset[-1][0]

        target_row = temp_base_row + idx
        intercept_formula = (
            f"=INTERCEPT(R{start_data_row}C{y_col}:R{end_data_row}C{y_col},"
            f"R{start_data_row}C{x_col}:R{end_data_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_data_row}C{y_col}:R{end_data_row}C{y_col},"
            f"R{start_data_row}C{x_col}:R{end_data_row}C{x_col})"
        )

        sheet.range((target_row, intercept_col)).formula2 = intercept_formula
        sheet.range((target_row, slope_col)).formula2 = slope_formula

    wb.app.calculate()

    results: List[Dict[str, Any]] = []
    for idx, n_quarters in enumerate(n_values):
        subset = pairs[-n_quarters:]
        target_row = temp_base_row + idx
        source_row = anchor_row + idx + 1

        intercept = safe_float(sheet.range((target_row, intercept_col)).value)
        slope = safe_float(sheet.range((target_row, slope_col)).value)
        latest_x = subset[-1][1]

        forecast_total_without_sa = (
            safe_float(sheet.range((source_row, fcst_col)).value) if fcst_col is not None else None
        )
        if forecast_total_without_sa is None and intercept is not None and slope is not None:
            forecast_total_without_sa = intercept + (slope * latest_x)

        forecast_max = safe_float(sheet.range((source_row, anchor_col)).value)
        forecast_min = safe_float(sheet.range((source_row, min_col)).value) if min_col is not None else None
        actual_value = safe_float(sheet.range((source_row, actual_col)).value) if actual_col is not None else None

        if forecast_max is None and forecast_total_without_sa is not None:
            forecast_max = forecast_total_without_sa
        if forecast_min is None and forecast_total_without_sa is not None:
            forecast_min = forecast_total_without_sa

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        num_quarters_used = (
            safe_int(sheet.range((source_row, num_q_col)).value) if num_q_col is not None else None
        )
        if num_quarters_used is None or num_quarters_used <= 0:
            num_quarters_used = n_quarters

        result = {
            "model": labels["model"],
            "ticker": labels["ticker"],
            "model_period": labels["model_period"],
            "model_date": labels["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_total_without_sa,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": file_path.name,
        }

        if results:
            prev = results[-1]
            duplicate = (
                values_match(safe_float(prev["intercept"]), intercept)
                and values_match(safe_float(prev["slope"]), slope)
                and values_match(safe_float(prev["forecast_value"]), forecast_total_without_sa)
                and values_match(safe_float(prev["forecast_max"]), forecast_max)
                and values_match(safe_float(prev["forecast_min"]), forecast_min)
            )
            if duplicate:
                continue

        results.append({k: clean_value(v) for k, v in result.items()})

    return results


def set_column_widths(ws, headers: List[str], rows: List[Dict[str, Any]]) -> None:
    for idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row in rows:
            value = row.get(header)
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(idx)].width = min(max_len + 2, 40)


def write_output_workbook(
    out_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    ws_emp = wb.create_sheet("empirical_candidates")
    ws_reg = wb.create_sheet("regression_candidates")

    ws_emp.append(EMPIRICAL_COLUMNS)
    for row in empirical_rows:
        ws_emp.append([row.get(col) for col in EMPIRICAL_COLUMNS])

    ws_reg.append(REGRESSION_COLUMNS)
    for row in regression_rows:
        ws_reg.append([row.get(col) for col in REGRESSION_COLUMNS])

    for ws, headers, rows in (
        (ws_emp, EMPIRICAL_COLUMNS, empirical_rows),
        (ws_reg, REGRESSION_COLUMNS, regression_rows),
    ):
        for cell in ws[1]:
            cell.font = Font(bold=True)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        set_column_widths(ws, headers, rows)

    wb.save(out_path)


def collect_input_files(input_folder: Path, input_folder_name: str) -> Tuple[List[Path], int]:
    files: List[Path] = []
    skipped = 0

    output_name_pattern = re.compile(
        rf"^{re.escape(input_folder_name)}_PARAM(?:\.\d+)?\.xlsx$", flags=re.IGNORECASE
    )

    for path in sorted(input_folder.iterdir()):
        if not path.is_file():
            continue
        if path.name.startswith("~"):
            print(f"SKIPPED {path.name}: temporary Excel file")
            skipped += 1
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"SKIPPED {path.name}: not .xlsx")
            skipped += 1
            continue
        if output_name_pattern.match(path.name):
            print(f"SKIPPED {path.name}: output workbook pattern")
            skipped += 1
            continue
        files.append(path)

    return files, skipped


def main() -> None:
    in_dir = input_dir.resolve()
    out_dir = output_dir.resolve()

    if not in_dir.exists() or not in_dir.is_dir():
        raise FileNotFoundError(f"input_dir not found or not a directory: {in_dir}")

    source_files, skipped_count = collect_input_files(in_dir, in_dir.name)
    out_path = output_workbook_path(in_dir, out_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_count = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    original_calculation = app.calculation
    app.calculation = "manual"

    try:
        for file_path in source_files:
            print(f"PROCESSING {file_path.name}")
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                labels = parse_file_labels(file_path.name)

                empirical_rows.extend(process_empirical_sheet(wb, file_path, labels))
                regression_rows.extend(process_regression_sheet(wb, file_path, labels))
                processed_count += 1
            except Exception as exc:
                print(f"SKIPPED {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.calculation = original_calculation
        app.quit()

    write_output_workbook(out_path, empirical_rows, regression_rows)

    print(f"OUTPUT {out_path}")
    print(f"FILES_PROCESSED {processed_count}")
    print(f"FILES_SKIPPED {skipped_count}")
    print(f"EMPIRICAL_ROWS {len(empirical_rows)}")
    print(f"REGRESSION_ROWS {len(regression_rows)}")


if __name__ == "__main__":
    main()

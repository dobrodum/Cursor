#!/usr/bin/env python3
from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

try:
    import xlwings as xw
except ImportError as exc:  # pragma: no cover - dependency/runtime guard
    raise SystemExit(
        "xlwings is required for this script. Install it with `pip install xlwings`."
    ) from exc

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -----------------------------------------------------------------------------
# Configure these two paths before running.
# -----------------------------------------------------------------------------
input_dir = "./input"
output_dir = "./output"

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


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return text.strip()


def reshape_values(values: Any, nrows: int, ncols: int) -> List[List[Any]]:
    if nrows <= 0 or ncols <= 0:
        return []

    if nrows == 1 and ncols == 1:
        return [[values]]

    if nrows == 1:
        if isinstance(values, list):
            return [values]
        return [[values]]

    if ncols == 1:
        if isinstance(values, list):
            return [[item] for item in values]
        return [[values]]

    if isinstance(values, list):
        return values

    return [[values]]


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if text.endswith("%"):
        text = text[:-1]
    try:
        return float(text)
    except ValueError:
        return None


def to_int(value: Any) -> Optional[int]:
    numeric = to_float(value)
    if numeric is None:
        return None
    return int(round(numeric))


def get_month_number(month_text: str) -> Optional[int]:
    cleaned = month_text.strip()
    if not cleaned:
        return None
    for fmt in ("%b", "%B"):
        try:
            return datetime.strptime(cleaned[:3], "%b").month if fmt == "%B" else datetime.strptime(cleaned, fmt).month
        except ValueError:
            continue
    try:
        return datetime.strptime(cleaned[:3], "%b").month
    except ValueError:
        return None


def parse_file_metadata(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = parts[1] if len(parts) > 1 and parts[1] else "UNKNOWN"

    period_token = ""
    if len(parts) > 2:
        period_token = parts[2].split("_")[0].strip()

    if not period_token:
        match = re.search(r"(Early|Mid|Late)[A-Za-z]{3,9}\d{4}", stem, re.IGNORECASE)
        period_token = match.group(0) if match else ""

    match = re.match(r"^(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})$", period_token, re.IGNORECASE)
    if not match:
        model_period = "UNKNOWN_PERIOD"
        model_date = ""
        model = f"{ticker}_{model_period}"
        return {
            "ticker": ticker,
            "model_period": model_period,
            "model_date": model_date,
            "model": model,
        }

    period_bucket = match.group(1).capitalize()
    month_text = match.group(2)
    year = int(match.group(3))

    month_num = get_month_number(month_text)
    if month_num is None:
        model_period = "UNKNOWN_PERIOD"
        model_date = ""
    else:
        month_abbr = datetime(year, month_num, 1).strftime("%b")
        day_lookup = {"Early": 5, "Mid": 15, "Late": 25}
        model_period = f"{period_bucket}{month_abbr}_{year}"
        model_date = date(year, month_num, day_lookup[period_bucket]).isoformat()

    model = f"{ticker}_{model_period}"
    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def next_output_path(in_dir: Path, out_dir: Path) -> Path:
    folder_name = in_dir.name
    base_name = f"{folder_name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    index = 1
    while True:
        candidate = out_dir / f"{base_name}.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def build_sheet_cache(ws: xw.Sheet) -> Dict[str, Any]:
    used = ws.used_range
    start_row = used.row
    start_col = used.column
    nrows = max(int(used.rows.count), 1)
    ncols = max(int(used.columns.count), 1)
    end_row = start_row + nrows - 1
    end_col = start_col + ncols - 1

    values = reshape_values(used.value, nrows, ncols)
    text_index: Dict[str, List[Tuple[int, int]]] = {}

    for r_idx, row_values in enumerate(values):
        for c_idx, value in enumerate(row_values):
            key = normalize_text(value)
            if not key:
                continue
            abs_row = start_row + r_idx
            abs_col = start_col + c_idx
            text_index.setdefault(key, []).append((abs_row, abs_col))

    return {
        "start_row": start_row,
        "start_col": start_col,
        "end_row": end_row,
        "end_col": end_col,
        "nrows": nrows,
        "ncols": ncols,
        "values": values,
        "text_index": text_index,
    }


def get_row_values(cache: Dict[str, Any], row_number: int) -> List[Any]:
    row_offset = row_number - cache["start_row"]
    if row_offset < 0 or row_offset >= cache["nrows"]:
        return []
    return cache["values"][row_offset]


def find_anchor_max(cache: Dict[str, Any]) -> Optional[Tuple[int, int]]:
    candidates = cache["text_index"].get("max", [])
    if candidates:
        return sorted(candidates)[0]

    for key, coords in cache["text_index"].items():
        if key.endswith(" max") or key.startswith("max "):
            return sorted(coords)[0]
    return None


def build_header_map(cache: Dict[str, Any], header_row: int) -> Dict[str, int]:
    row_values = get_row_values(cache, header_row)
    header_map: Dict[str, int] = {}
    for idx, value in enumerate(row_values):
        key = normalize_text(value)
        if not key:
            continue
        column = cache["start_col"] + idx
        header_map.setdefault(key, column)
    return header_map


def resolve_column(header_map: Dict[str, int], aliases: List[str], fallback: int) -> int:
    normalized_aliases = [normalize_text(alias) for alias in aliases if alias]
    for alias in normalized_aliases:
        if alias in header_map:
            return header_map[alias]

    for alias in normalized_aliases:
        for key, column in header_map.items():
            if alias in key:
                return column
    return max(fallback, 1)


def block_value(
    block: List[List[Any]],
    row_idx: int,
    col: int,
    block_start_col: int,
    block_end_col: int,
) -> Any:
    if row_idx < 0 or row_idx >= len(block):
        return None
    if col < block_start_col or col > block_end_col:
        return None
    col_idx = col - block_start_col
    row_values = block[row_idx]
    if col_idx < 0 or col_idx >= len(row_values):
        return None
    return row_values[col_idx]


def close_workbook_safely(workbook: xw.Book) -> None:
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
        workbook.close()
    except Exception as exc:
        print(f"  WARN unable to close workbook safely: {exc}")


def extract_empirical_rows(
    ws: xw.Sheet,
    wb: xw.Book,
    file_meta: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    cache = build_sheet_cache(ws)
    anchor = find_anchor_max(cache)
    if anchor is None:
        print("  SKIP Empirical Model: max anchor not found")
        return []

    anchor_row, anchor_col = anchor
    header_map = build_header_map(cache, anchor_row)

    col_num_quarters = resolve_column(
        header_map, ["num quarters used", "num quarters", "n quarters"], anchor_col - 12
    )
    col_last_quarter = resolve_column(
        header_map, ["last quarter used", "last quarter"], anchor_col - 11
    )
    col_forecast_value = resolve_column(
        header_map,
        ["estimated total sold", "forecast value", "tot fcst", "forecast"],
        anchor_col - 1,
    )
    col_actual_value = resolve_column(
        header_map, ["reported sales", "actual value", "actual"], anchor_col - 2
    )
    col_forecast_max = resolve_column(header_map, ["max"], anchor_col)
    col_forecast_min = resolve_column(header_map, ["min"], anchor_col + 1)
    col_avg_pen = resolve_column(
        header_map,
        ["avg penetration %", "avg penetration pct", "average penetration"],
        anchor_col - 4,
    )
    col_quarterly_sales = resolve_column(header_map, ["quarterly sales"], anchor_col - 8)
    col_reported_sales = resolve_column(header_map, ["reported sales"], col_actual_value)
    col_growth_rate = resolve_column(
        header_map, ["growth rate %", "growth rate pct", "growth rate"], anchor_col - 6
    )
    col_sales_captured = resolve_column(
        header_map,
        ["sales captured in db %", "sales captured in db pct", "captured in db"],
        anchor_col - 5,
    )

    data_start_row = anchor_row + 1
    data_end_row = data_start_row + N_QUARTERS - 1

    block_start_col = min(
        col_num_quarters,
        col_last_quarter,
        col_forecast_value,
        col_actual_value,
        col_forecast_max,
        col_forecast_min,
        col_avg_pen,
        col_quarterly_sales,
        col_reported_sales,
        col_growth_rate,
        col_sales_captured,
    )
    block_end_col = max(
        col_num_quarters,
        col_last_quarter,
        col_forecast_value,
        col_actual_value,
        col_forecast_max,
        col_forecast_min,
        col_avg_pen,
        col_quarterly_sales,
        col_reported_sales,
        col_growth_rate,
        col_sales_captured,
    )

    row_block = ws.range((data_start_row, block_start_col), (data_end_row, block_end_col)).value
    row_block_2d = reshape_values(row_block, N_QUARTERS, block_end_col - block_start_col + 1)

    helper_col = cache["end_col"] + 2
    helper_start_row = cache["end_row"] + 2
    avg_pen_values: List[Optional[float]] = [None] * N_QUARTERS

    if col_sales_captured >= 1:
        avg_pen_formulas = []
        for i in range(N_QUARTERS):
            current_row = data_start_row + i
            start_row = max(data_start_row, current_row - i)
            avg_pen_formulas.append(
                [f"=AVERAGE(R{start_row}C{col_sales_captured}:R{current_row}C{col_sales_captured})"]
            )

        avg_pen_rng = ws.range(
            (helper_start_row, helper_col), (helper_start_row + N_QUARTERS - 1, helper_col)
        )
        avg_pen_rng.formula2 = avg_pen_formulas
        wb.app.calculate()
        avg_pen_calc = reshape_values(avg_pen_rng.value, N_QUARTERS, 1)
        avg_pen_values = [to_float(avg_pen_calc[i][0]) for i in range(N_QUARTERS)]

    rows: List[Dict[str, Any]] = []
    for i in range(N_QUARTERS):
        num_quarters_used = to_int(
            block_value(row_block_2d, i, col_num_quarters, block_start_col, block_end_col)
        )
        if num_quarters_used is None:
            num_quarters_used = i + 1

        last_quarter_used = block_value(
            row_block_2d, i, col_last_quarter, block_start_col, block_end_col
        )
        forecast_value = to_float(
            block_value(row_block_2d, i, col_forecast_value, block_start_col, block_end_col)
        )
        actual_value = to_float(
            block_value(row_block_2d, i, col_actual_value, block_start_col, block_end_col)
        )
        forecast_max = to_float(
            block_value(row_block_2d, i, col_forecast_max, block_start_col, block_end_col)
        )
        forecast_min = to_float(
            block_value(row_block_2d, i, col_forecast_min, block_start_col, block_end_col)
        )
        avg_pen = avg_pen_values[i]
        if avg_pen is None:
            avg_pen = to_float(
                block_value(row_block_2d, i, col_avg_pen, block_start_col, block_end_col)
            )
        quarterly_sales = to_float(
            block_value(row_block_2d, i, col_quarterly_sales, block_start_col, block_end_col)
        )
        reported_sales = to_float(
            block_value(row_block_2d, i, col_reported_sales, block_start_col, block_end_col)
        )
        growth_rate_pct = to_float(
            block_value(row_block_2d, i, col_growth_rate, block_start_col, block_end_col)
        )
        sales_captured_in_db_pct = to_float(
            block_value(row_block_2d, i, col_sales_captured, block_start_col, block_end_col)
        )

        range_width = (
            (forecast_max - forecast_min)
            if forecast_max is not None and forecast_min is not None
            else None
        )

        if all(
            value is None
            for value in (
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                avg_pen,
                quarterly_sales,
                reported_sales,
            )
        ):
            continue

        row = {
            "model": file_meta["model"],
            "ticker": file_meta["ticker"],
            "model_period": file_meta["model_period"],
            "model_date": file_meta["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_pen,
            "num_quarters_used": num_quarters_used,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "avg_penetration_pct": avg_pen,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate_pct,
            "sales_captured_in_db_pct": sales_captured_in_db_pct,
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def almost_equal(left: Any, right: Any, tol: float = 1e-9) -> bool:
    left_f = to_float(left)
    right_f = to_float(right)
    if left_f is None or right_f is None:
        return left == right
    return abs(left_f - right_f) <= tol


def extract_regression_rows(
    ws: xw.Sheet,
    wb: xw.Book,
    file_meta: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    cache = build_sheet_cache(ws)
    anchor = find_anchor_max(cache)
    if anchor is None:
        print("  SKIP Regression Model: max anchor not found")
        return []

    anchor_row, anchor_col = anchor
    header_map = build_header_map(cache, anchor_row)

    y_col = max(anchor_col - 7, 1)
    x_col = max(anchor_col - 11, 1)

    col_num_quarters = resolve_column(
        header_map, ["num quarters used", "num quarters", "n quarters"], anchor_col - 12
    )
    col_forecast_total = resolve_column(
        header_map,
        [
            "tot fcst w o sa",
            "tot fcst wo sa",
            "tot fcst w/o sa",
            "forecast total without sa",
            "forecast value",
        ],
        anchor_col - 1,
    )
    col_actual_value = resolve_column(
        header_map, ["actual value", "reported sales", "actual"], anchor_col - 2
    )
    col_forecast_max = resolve_column(header_map, ["max"], anchor_col)
    col_forecast_min = resolve_column(header_map, ["min"], anchor_col + 1)

    data_start_row = anchor_row + 1
    data_end_row = max(cache["end_row"], data_start_row + N_QUARTERS - 1)

    xy_start_col = min(x_col, y_col)
    xy_end_col = max(x_col, y_col)
    xy_values = ws.range((data_start_row, xy_start_col), (data_end_row, xy_end_col)).value
    xy_block = reshape_values(xy_values, data_end_row - data_start_row + 1, xy_end_col - xy_start_col + 1)

    available_rows: List[int] = []
    for idx, row_values in enumerate(xy_block):
        x_val = to_float(row_values[x_col - xy_start_col]) if (x_col - xy_start_col) < len(row_values) else None
        y_val = to_float(row_values[y_col - xy_start_col]) if (y_col - xy_start_col) < len(row_values) else None
        if x_val is None or y_val is None:
            continue
        available_rows.append(data_start_row + idx)
        if len(available_rows) >= N_QUARTERS:
            break

    if not available_rows:
        available_rows = [data_start_row + i for i in range(N_QUARTERS)]

    helper_col = cache["end_col"] + 3
    helper_start_row = cache["end_row"] + 2

    intercept_formulas = []
    slope_formulas = []
    for i in range(N_QUARTERS):
        n_rows = min(i + 1, len(available_rows))
        start_row = available_rows[0]
        end_row = available_rows[n_rows - 1]

        if n_rows < 2:
            intercept_formulas.append([""])
            slope_formulas.append([""])
            continue

        intercept_formulas.append(
            [
                f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},"
                f"R{start_row}C{x_col}:R{end_row}C{x_col})"
            ]
        )
        slope_formulas.append(
            [
                f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},"
                f"R{start_row}C{x_col}:R{end_row}C{x_col})"
            ]
        )

    intercept_rng = ws.range(
        (helper_start_row, helper_col), (helper_start_row + N_QUARTERS - 1, helper_col)
    )
    slope_rng = ws.range(
        (helper_start_row, helper_col + 1),
        (helper_start_row + N_QUARTERS - 1, helper_col + 1),
    )

    intercept_rng.formula2 = intercept_formulas
    slope_rng.formula2 = slope_formulas
    wb.app.calculate()

    intercept_values = reshape_values(intercept_rng.value, N_QUARTERS, 1)
    slope_values = reshape_values(slope_rng.value, N_QUARTERS, 1)

    block_start_col = min(
        col_num_quarters, col_forecast_total, col_actual_value, col_forecast_max, col_forecast_min
    )
    block_end_col = max(
        col_num_quarters, col_forecast_total, col_actual_value, col_forecast_max, col_forecast_min
    )
    row_block_values = ws.range(
        (data_start_row, block_start_col),
        (data_start_row + N_QUARTERS - 1, block_end_col),
    ).value
    row_block_2d = reshape_values(row_block_values, N_QUARTERS, block_end_col - block_start_col + 1)

    rows: List[Dict[str, Any]] = []
    for i in range(N_QUARTERS):
        num_quarters_used = to_int(
            block_value(row_block_2d, i, col_num_quarters, block_start_col, block_end_col)
        )
        if num_quarters_used is None:
            num_quarters_used = i + 1

        forecast_value = to_float(
            block_value(row_block_2d, i, col_forecast_total, block_start_col, block_end_col)
        )
        actual_value = to_float(
            block_value(row_block_2d, i, col_actual_value, block_start_col, block_end_col)
        )
        forecast_max = to_float(
            block_value(row_block_2d, i, col_forecast_max, block_start_col, block_end_col)
        )
        forecast_min = to_float(
            block_value(row_block_2d, i, col_forecast_min, block_start_col, block_end_col)
        )
        intercept = to_float(intercept_values[i][0])
        slope = to_float(slope_values[i][0])
        range_width = (
            (forecast_max - forecast_min)
            if forecast_max is not None and forecast_min is not None
            else None
        )

        if all(
            value is None
            for value in (forecast_value, forecast_max, forecast_min, intercept, slope)
        ):
            continue

        current = {
            "model": file_meta["model"],
            "ticker": file_meta["ticker"],
            "model_period": file_meta["model_period"],
            "model_date": file_meta["model_date"],
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

        if i == N_QUARTERS - 1 and rows:
            prev = rows[-1]
            duplicate_final = all(
                almost_equal(current[key], prev.get(key))
                for key in (
                    "forecast_value",
                    "forecast_max",
                    "forecast_min",
                    "intercept",
                    "slope",
                )
            )
            if duplicate_final:
                continue

        rows.append(current)

    return rows


def write_output_sheet(
    workbook: Workbook,
    title: str,
    columns: List[str],
    rows: List[Dict[str, Any]],
) -> None:
    ws = workbook.create_sheet(title=title)
    ws.append(columns)
    for row in rows:
        ws.append([row.get(col) for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row in rows:
            cell_value = row.get(column_name)
            text = "" if cell_value is None else str(cell_value)
            if len(text) > max_len:
                max_len = len(text)
        ws.column_dimensions[get_column_letter(idx)].width = min(max(max_len + 2, 12), 42)


def process_all_workbooks(input_path: Path, output_path: Path) -> None:
    output_path.mkdir(parents=True, exist_ok=True)
    target_output = next_output_path(input_path, output_path)

    output_pattern = re.compile(
        rf"^{re.escape(input_path.name)}_PARAM(?:\.\d+)?\.xlsx$", re.IGNORECASE
    )

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(input_path.iterdir()):
            if not file_path.is_file():
                continue

            if file_path.suffix.lower() != ".xlsx":
                print(f"SKIP {file_path.name}: not .xlsx")
                continue

            if file_path.name.startswith("~"):
                print(f"SKIP {file_path.name}: temp file")
                continue

            if output_pattern.match(file_path.name):
                print(f"SKIP {file_path.name}: generated output file")
                continue

            print(f"PROCESS {file_path.name}")

            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                files_processed += 1
                file_meta = parse_file_metadata(file_path.name)

                if "Empirical Model" in [sheet.name for sheet in wb.sheets]:
                    empirical_rows.extend(
                        extract_empirical_rows(
                            ws=wb.sheets["Empirical Model"],
                            wb=wb,
                            file_meta=file_meta,
                            source_file=file_path.name,
                        )
                    )
                else:
                    print("  SKIP Empirical Model: sheet missing")

                if "Regression Model" in [sheet.name for sheet in wb.sheets]:
                    regression_rows.extend(
                        extract_regression_rows(
                            ws=wb.sheets["Regression Model"],
                            wb=wb,
                            file_meta=file_meta,
                            source_file=file_path.name,
                        )
                    )
                else:
                    print("  SKIP Regression Model: sheet missing")
            except Exception as exc:
                print(f"  ERROR {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    close_workbook_safely(wb)
    finally:
        app.quit()

    output_wb = Workbook()
    output_wb.remove(output_wb.active)
    write_output_sheet(output_wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_output_sheet(output_wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    output_wb.save(target_output)

    print(f"Output path: {target_output}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


def main() -> None:
    in_path = Path(input_dir).expanduser().resolve()
    out_path = Path(output_dir).expanduser().resolve()

    if not in_path.exists():
        raise SystemExit(f"Input folder does not exist: {in_path}")
    if not in_path.is_dir():
        raise SystemExit(f"Input path is not a folder: {in_path}")

    process_all_workbooks(in_path, out_path)


if __name__ == "__main__":
    main()

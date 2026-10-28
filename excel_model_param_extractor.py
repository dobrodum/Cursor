from __future__ import annotations

import calendar
import re
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import xlwings as xw


# -------------------------------
# Inputs
# -------------------------------
input_dir = r"/path/to/input/folder"
output_dir = r"/path/to/output/folder"


# -------------------------------
# Constants
# -------------------------------
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

# Anchor-relative offsets from the "max" cell.
EMPIRICAL_OFFSETS = {
    "num_quarters_used": -9,
    "last_quarter_used": -8,
    "forecast_value": -1,
    "actual_value": -2,
    "forecast_max": 0,
    "forecast_min": 1,
    "quarterly_sales": -3,
    "reported_sales": -2,
    "growth_rate_pct": -4,
    "sales_captured_in_db_pct": -5,
}

REGRESSION_OFFSETS = {
    "num_quarters_used": -9,
    "forecast_value": -1,
    "actual_value": -2,
    "forecast_max": 0,
    "forecast_min": 1,
}

X_COL_OFFSET = -11
Y_COL_OFFSET = -7

MONTH_MAP = {name.lower(): idx for idx, name in enumerate(calendar.month_abbr) if name}


def to_float(value: Any) -> float | None:
    if value in (None, ""):
        return None
    if isinstance(value, bool):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def to_matrix(values: Any, rows: int, cols: int) -> list[list[Any]]:
    if rows <= 0 or cols <= 0:
        return []
    if rows == 1 and cols == 1:
        return [[values]]
    if rows == 1:
        if not isinstance(values, list):
            return [[values] + [None] * (cols - 1)]
        row = list(values)
        if len(row) < cols:
            row.extend([None] * (cols - len(row)))
        return [row[:cols]]
    if cols == 1:
        if not isinstance(values, list):
            return [[values]] + [[None] for _ in range(rows - 1)]
        return [[v] for v in list(values)[:rows]] + [[None] for _ in range(max(0, rows - len(values)))]
    if not isinstance(values, list):
        return [[None for _ in range(cols)] for _ in range(rows)]
    return values


def read_block(sheet: xw.Sheet, top: int, left: int, bottom: int, right: int) -> dict[str, Any]:
    top = max(1, top)
    left = max(1, left)
    if bottom < top or right < left:
        return {"top": top, "left": left, "values": []}
    rows = bottom - top + 1
    cols = right - left + 1
    values = sheet.range((top, left), (bottom, right)).value
    matrix = to_matrix(values, rows, cols)
    return {"top": top, "left": left, "values": matrix}


def block_get(block: dict[str, Any], row: int, col: int) -> Any:
    values = block["values"]
    if not values:
        return None
    r = row - block["top"]
    c = col - block["left"]
    if r < 0 or c < 0:
        return None
    if r >= len(values):
        return None
    if c >= len(values[r]):
        return None
    return values[r][c]


def normalize_column(values: Any, n_rows: int) -> list[Any]:
    if n_rows <= 0:
        return []
    if isinstance(values, list):
        if values and isinstance(values[0], list):
            out = [(row[0] if row else None) for row in values]
        else:
            out = list(values)
    else:
        out = [values]
    if len(out) < n_rows:
        out.extend([None] * (n_rows - len(out)))
    return out[:n_rows]


def set_formula2_range(cell_range: xw.Range, formulas: list[list[str]]) -> None:
    try:
        cell_range.formula2 = formulas
    except Exception:
        cell_range.formula = formulas


def safe_close_workbook(wb: xw.Book) -> None:
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

    try:
        wb.api.Close(SaveChanges=False)
    except Exception:
        wb.close()


def find_max_anchor(sheet: xw.Sheet) -> tuple[int, int]:
    used = sheet.used_range
    values = used.value

    if values is None:
        raise ValueError("sheet is empty")

    if not isinstance(values, list):
        matrix = [[values]]
    elif values and not isinstance(values[0], list):
        matrix = [values]
    else:
        matrix = values

    base_row = used.row
    base_col = used.column

    for r_idx, row in enumerate(matrix):
        for c_idx, cell_value in enumerate(row):
            if isinstance(cell_value, str) and cell_value.strip().lower() == "max":
                return base_row + r_idx, base_col + c_idx

    raise ValueError("no 'max' anchor found")


def parse_file_labels(filename: str) -> dict[str, str]:
    stem = Path(filename).stem
    parts = [p.strip() for p in stem.split(" - ")]

    ticker = ""
    if len(parts) >= 2 and parts[1]:
        ticker = parts[1].strip().upper()
    else:
        m_ticker = re.search(r"\b[A-Z]{1,6}\b", stem.upper())
        ticker = m_ticker.group(0) if m_ticker else "UNKNOWN"

    period_match = re.search(r"(Early|Mid|Late)([A-Za-z]{3})(\d{4})", stem, flags=re.IGNORECASE)

    model_period = "UNKNOWN_PERIOD"
    model_date = ""
    if period_match:
        phase_raw, month_raw, year = period_match.groups()
        phase = phase_raw[0].upper() + phase_raw[1:].lower()
        month_abbr = month_raw[0].upper() + month_raw[1:].lower()
        month_num = MONTH_MAP.get(month_abbr.lower())
        day_map = {"early": 5, "mid": 15, "late": 25}
        day = day_map.get(phase.lower(), 15)
        model_period = f"{phase}{month_abbr}_{year}"
        if month_num:
            model_date = f"{year}-{month_num:02d}-{day:02d}"

    model = f"{ticker}_{model_period}" if ticker else model_period
    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def build_output_path(input_path: Path, output_path: Path) -> Path:
    base_name = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    i = 1
    while True:
        candidate = output_path / f"{base_name}.{i}.xlsx"
        if not candidate.exists():
            return candidate
        i += 1


def collect_source_files(input_path: Path, output_prefix: str) -> list[Path]:
    source_files: list[Path] = []
    for file_path in sorted(input_path.iterdir()):
        if not file_path.is_file():
            continue

        name = file_path.name
        if name.startswith("~"):
            print(f"SKIPPED: {name} (temporary file)")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"SKIPPED: {name} (not .xlsx)")
            continue
        if name.lower().startswith(output_prefix.lower()):
            print(f"SKIPPED: {name} (existing PARAM output workbook)")
            continue

        source_files.append(file_path)
    return source_files


def fallback_spread(growth_rate: Any) -> float:
    g = to_float(growth_rate)
    if g is None:
        return 0.05
    g = abs(g)
    if g == 0:
        return 0.05
    return g


def extract_empirical_rows(wb: xw.Book, file_meta: dict[str, str], source_file: str) -> list[dict[str, Any]]:
    sheet_name = "Empirical Model"
    try:
        sheet = wb.sheets[sheet_name]
    except Exception:
        print(f"SKIPPED: {source_file} ({sheet_name} sheet not found)")
        return []

    try:
        anchor_row, anchor_col = find_max_anchor(sheet)
    except ValueError as exc:
        print(f"SKIPPED: {source_file} ({sheet_name}: {exc})")
        return []

    history_end_row = anchor_row - 1
    x_col = anchor_col + X_COL_OFFSET
    y_col = anchor_col + Y_COL_OFFSET

    data_block = read_block(
        sheet,
        anchor_row + 1,
        anchor_col - 12,
        anchor_row + N_QUARTERS,
        anchor_col + 2,
    )

    temp_avg_col = anchor_col + 25
    avg_formulas: list[list[str]] = []
    for n_quarters in range(1, N_QUARTERS + 1):
        start_row = history_end_row - n_quarters + 1
        if start_row < 1 or x_col < 1:
            avg_formulas.append([""])
            continue
        avg_formulas.append([f"=AVERAGE(R{start_row}C{x_col}:R{history_end_row}C{x_col})"])

    avg_range = sheet.range((anchor_row + 1, temp_avg_col), (anchor_row + N_QUARTERS, temp_avg_col))
    set_formula2_range(avg_range, avg_formulas)
    wb.app.calculate()
    avg_pen_values = normalize_column(avg_range.value, N_QUARTERS)

    latest_y = to_float(sheet.range((history_end_row, y_col)).value) if history_end_row >= 1 and y_col >= 1 else None
    prev_y = to_float(sheet.range((history_end_row - 1, y_col)).value) if history_end_row - 1 >= 1 and y_col >= 1 else None
    fallback_growth_rate = None
    if latest_y is not None and prev_y not in (None, 0):
        fallback_growth_rate = (latest_y / prev_y) - 1.0

    rows: list[dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        n_quarters = idx + 1
        row_idx = anchor_row + n_quarters

        num_quarters_used = block_get(data_block, row_idx, anchor_col + EMPIRICAL_OFFSETS["num_quarters_used"])
        if to_float(num_quarters_used) is not None:
            num_quarters_used = int(to_float(num_quarters_used))
        elif num_quarters_used in (None, ""):
            num_quarters_used = n_quarters

        last_quarter_used = block_get(data_block, row_idx, anchor_col + EMPIRICAL_OFFSETS["last_quarter_used"])
        forecast_value = block_get(data_block, row_idx, anchor_col + EMPIRICAL_OFFSETS["forecast_value"])
        actual_value = block_get(data_block, row_idx, anchor_col + EMPIRICAL_OFFSETS["actual_value"])
        forecast_max = block_get(data_block, row_idx, anchor_col + EMPIRICAL_OFFSETS["forecast_max"])
        forecast_min = block_get(data_block, row_idx, anchor_col + EMPIRICAL_OFFSETS["forecast_min"])
        quarterly_sales = block_get(data_block, row_idx, anchor_col + EMPIRICAL_OFFSETS["quarterly_sales"])
        reported_sales = block_get(data_block, row_idx, anchor_col + EMPIRICAL_OFFSETS["reported_sales"])
        growth_rate_pct = block_get(data_block, row_idx, anchor_col + EMPIRICAL_OFFSETS["growth_rate_pct"])
        sales_captured_in_db_pct = block_get(data_block, row_idx, anchor_col + EMPIRICAL_OFFSETS["sales_captured_in_db_pct"])
        avg_penetration_pct = avg_pen_values[idx]

        if reported_sales in (None, ""):
            reported_sales = actual_value
        if quarterly_sales in (None, ""):
            quarterly_sales = reported_sales
        if growth_rate_pct in (None, ""):
            growth_rate_pct = fallback_growth_rate
        if sales_captured_in_db_pct in (None, ""):
            sales_captured_in_db_pct = avg_penetration_pct

        fv_num = to_float(forecast_value)
        rs_num = to_float(reported_sales)
        ap_num = to_float(avg_penetration_pct)
        if fv_num is None and rs_num not in (None, 0) and ap_num not in (None, 0):
            forecast_value = rs_num / ap_num

        if to_float(forecast_max) is None and to_float(forecast_value) is not None:
            spread = fallback_spread(growth_rate_pct)
            forecast_max = to_float(forecast_value) * (1.0 + spread)

        if to_float(forecast_min) is None and to_float(forecast_value) is not None:
            spread = fallback_spread(growth_rate_pct)
            forecast_min = to_float(forecast_value) * (1.0 - spread)

        max_num = to_float(forecast_max)
        min_num = to_float(forecast_min)
        range_width = (max_num - min_num) if max_num is not None and min_num is not None else None

        if all(
            v in (None, "")
            for v in (
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
                reported_sales,
            )
        ):
            continue

        rows.append(
            {
                "model": file_meta["model"],
                "ticker": file_meta["ticker"],
                "model_period": file_meta["model_period"],
                "model_date": file_meta["model_date"],
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


def extract_regression_rows(wb: xw.Book, file_meta: dict[str, str], source_file: str) -> list[dict[str, Any]]:
    sheet_name = "Regression Model"
    try:
        sheet = wb.sheets[sheet_name]
    except Exception:
        print(f"SKIPPED: {source_file} ({sheet_name} sheet not found)")
        return []

    try:
        anchor_row, anchor_col = find_max_anchor(sheet)
    except ValueError as exc:
        print(f"SKIPPED: {source_file} ({sheet_name}: {exc})")
        return []

    history_end_row = anchor_row - 1
    x_col = anchor_col - 11
    y_col = anchor_col - 7

    data_block = read_block(
        sheet,
        anchor_row + 1,
        anchor_col - 12,
        anchor_row + N_QUARTERS,
        anchor_col + 2,
    )

    temp_intercept_col = anchor_col + 25
    temp_slope_col = anchor_col + 26
    intercept_formulas: list[list[str]] = []
    slope_formulas: list[list[str]] = []
    for n_quarters in range(1, N_QUARTERS + 1):
        start_row = history_end_row - n_quarters + 1
        if start_row < 1 or x_col < 1 or y_col < 1:
            intercept_formulas.append([""])
            slope_formulas.append([""])
            continue
        intercept_formulas.append(
            [f"=INTERCEPT(R{start_row}C{y_col}:R{history_end_row}C{y_col},R{start_row}C{x_col}:R{history_end_row}C{x_col})"]
        )
        slope_formulas.append(
            [f"=SLOPE(R{start_row}C{y_col}:R{history_end_row}C{y_col},R{start_row}C{x_col}:R{history_end_row}C{x_col})"]
        )

    intercept_range = sheet.range((anchor_row + 1, temp_intercept_col), (anchor_row + N_QUARTERS, temp_intercept_col))
    slope_range = sheet.range((anchor_row + 1, temp_slope_col), (anchor_row + N_QUARTERS, temp_slope_col))

    set_formula2_range(intercept_range, intercept_formulas)
    set_formula2_range(slope_range, slope_formulas)
    wb.app.calculate()

    intercept_values = normalize_column(intercept_range.value, N_QUARTERS)
    slope_values = normalize_column(slope_range.value, N_QUARTERS)

    latest_x = to_float(sheet.range((history_end_row, x_col)).value) if history_end_row >= 1 and x_col >= 1 else None

    rows: list[dict[str, Any]] = []
    prev_key: tuple[Any, ...] | None = None

    for idx in range(N_QUARTERS):
        n_quarters = idx + 1
        row_idx = anchor_row + n_quarters

        num_quarters_used = block_get(data_block, row_idx, anchor_col + REGRESSION_OFFSETS["num_quarters_used"])
        if to_float(num_quarters_used) is not None:
            num_quarters_used = int(to_float(num_quarters_used))
        elif num_quarters_used in (None, ""):
            num_quarters_used = n_quarters

        forecast_value = block_get(data_block, row_idx, anchor_col + REGRESSION_OFFSETS["forecast_value"])
        actual_value = block_get(data_block, row_idx, anchor_col + REGRESSION_OFFSETS["actual_value"])
        forecast_max = block_get(data_block, row_idx, anchor_col + REGRESSION_OFFSETS["forecast_max"])
        forecast_min = block_get(data_block, row_idx, anchor_col + REGRESSION_OFFSETS["forecast_min"])

        intercept = intercept_values[idx]
        slope = slope_values[idx]

        i_num = to_float(intercept)
        s_num = to_float(slope)
        fv_num = to_float(forecast_value)

        if fv_num is None and i_num is not None and s_num is not None and latest_x is not None:
            forecast_value = i_num + (s_num * latest_x)

        fv_num = to_float(forecast_value)
        if to_float(forecast_max) is None and fv_num is not None:
            forecast_max = fv_num * 1.05
        if to_float(forecast_min) is None and fv_num is not None:
            forecast_min = fv_num * 0.95

        max_num = to_float(forecast_max)
        min_num = to_float(forecast_min)
        range_width = (max_num - min_num) if max_num is not None and min_num is not None else None

        dedupe_key = (
            to_float(num_quarters_used),
            round(i_num, 10) if i_num is not None else None,
            round(s_num, 10) if s_num is not None else None,
            round(to_float(forecast_value), 10) if to_float(forecast_value) is not None else None,
            round(max_num, 10) if max_num is not None else None,
            round(min_num, 10) if min_num is not None else None,
        )
        if dedupe_key == prev_key:
            continue
        prev_key = dedupe_key

        if all(v in (None, "") for v in (intercept, slope, forecast_value, forecast_max, forecast_min)):
            continue

        rows.append(
            {
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
        )

    return rows


def write_sheet(ws: openpyxl.worksheet.worksheet.Worksheet, headers: list[str], rows: list[dict[str, Any]]) -> None:
    ws.append(headers)
    for row in rows:
        ws.append([row.get(col) for col in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 45)


def write_output_workbook(output_path: Path, empirical_rows: list[dict[str, Any]], regression_rows: list[dict[str, Any]]) -> None:
    wb = openpyxl.Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    ws_empirical = wb.create_sheet("empirical_candidates")
    ws_regression = wb.create_sheet("regression_candidates")

    write_sheet(ws_empirical, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(ws_regression, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {input_path}")

    output_path.mkdir(parents=True, exist_ok=True)

    output_file = build_output_path(input_path, output_path)
    output_prefix = f"{input_path.name}_PARAM"
    source_files = collect_source_files(input_path, output_prefix)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in source_files:
            print(f"PROCESSING: {file_path.name}")
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                file_meta = parse_file_labels(file_path.name)

                empirical_rows.extend(extract_empirical_rows(wb, file_meta, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, file_meta, file_path.name))

                files_processed += 1
            except Exception as exc:
                print(f"SKIPPED: {file_path.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"OUTPUT: {output_file}")
    print(f"FILES_PROCESSED: {files_processed}")
    print(f"EMPIRICAL_ROWS: {len(empirical_rows)}")
    print(f"REGRESSION_ROWS: {len(regression_rows)}")


if __name__ == "__main__":
    main()

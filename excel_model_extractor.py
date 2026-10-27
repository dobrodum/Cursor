#!/usr/bin/env python3
"""Extract empirical and regression model candidates from Excel workbooks."""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Any

import xlwings as xw


# ---------- User-configurable paths ----------
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

EMP_DEFAULT_OFFSETS = {
    "num_quarters_used": -12,
    "last_quarter_used": -11,
    "forecast_value": -3,
    "actual_value": -2,
    "forecast_max": 0,
    "forecast_min": 1,
    "avg_penetration_pct": -8,
    "quarterly_sales": -7,
    "reported_sales": -6,
    "growth_rate_pct": -5,
    "sales_captured_in_db_pct": -4,
}

REG_DEFAULT_OFFSETS = {
    "num_quarters_used": -12,
    "forecast_value": -3,  # TOT FCST w/o SA default location
    "actual_value": -2,
    "forecast_max": 0,
    "forecast_min": 1,
}

HEADER_TOKENS = {
    "num_quarters_used": [
        ("num", "quarter"),
        ("quarters", "used"),
        ("num", "qtr"),
        ("n", "qtr"),
    ],
    "last_quarter_used": [
        ("last", "quarter"),
        ("last", "qtr"),
    ],
    "forecast_value": [
        ("tot", "fcst", "w", "o", "sa"),
        ("tot", "forecast", "without", "sa"),
        ("estimated", "total", "sold"),
        ("forecast", "total"),
    ],
    "actual_value": [
        ("actual",),
        ("reported", "sales"),
    ],
    "forecast_max": [("max",)],
    "forecast_min": [("min",)],
    "avg_penetration_pct": [
        ("avg", "penetration"),
        ("average", "penetration"),
    ],
    "quarterly_sales": [("quarterly", "sales")],
    "reported_sales": [("reported", "sales")],
    "growth_rate_pct": [("growth", "rate")],
    "sales_captured_in_db_pct": [
        ("sales", "captured", "in", "db"),
        ("captured", "in", "db"),
    ],
}

TIMING_DAY = {"early": 5, "mid": 15, "late": 25}


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    if not text:
        return ""
    text = text.replace("%", " pct ").replace("&", " and ").replace("/", " ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def to_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def to_int(value: Any) -> int | None:
    numeric = to_float(value)
    if numeric is None:
        return None
    rounded = int(round(numeric))
    if abs(numeric - rounded) < 1e-9:
        return rounded
    return None


def safe_cell_value(sheet: xw.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    try:
        return sheet.range((row, col)).value
    except Exception:
        return None


def set_r1c1_formula2(target_range: xw.Range, formula: str) -> None:
    # Formula2R1C1 avoids A1 column-letter conversion and keeps formula updates fast.
    try:
        target_range.api.Formula2R1C1 = formula
    except Exception:
        try:
            target_range.formula2 = formula
        except Exception:
            target_range.api.FormulaR1C1 = formula


def parse_filename_labels(file_path: Path) -> dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1] if len(parts) >= 2 and parts[1] else "UNKNOWN"

    token = ""
    if len(parts) >= 3 and parts[2]:
        token = re.split(r"[_\s]+", parts[2])[0]
    if not token:
        match = re.search(r"(Early|Mid|Late)[A-Za-z]{3,9}\d{4}", stem, flags=re.IGNORECASE)
        token = match.group(0) if match else ""

    timing = "Mid"
    month_abbr = "Jan"
    year = 1900
    model_date = "1900-01-15"

    token_match = re.match(r"^(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})$", token, flags=re.IGNORECASE)
    if token_match:
        timing = token_match.group(1).title()
        month_fragment = token_match.group(2)[:3].title()
        try:
            month_num = datetime.strptime(month_fragment, "%b").month
            month_abbr = month_fragment
            year = int(token_match.group(3))
            day = TIMING_DAY[timing.lower()]
            model_date = f"{year:04d}-{month_num:02d}-{day:02d}"
        except ValueError:
            pass

    model_period = f"{timing}{month_abbr}_{year:04d}"
    model = f"{ticker}_{model_period}"
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def unique_output_path(in_dir: Path, out_dir: Path) -> Path:
    base_name = f"{in_dir.name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = out_dir / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def collect_input_files(in_dir: Path) -> list[Path]:
    if not in_dir.exists():
        print(f"skipped: input directory not found -> {in_dir}")
        return []

    output_prefix = f"{in_dir.name}_PARAM"
    candidates: list[Path] = []
    for file_path in sorted(in_dir.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"skipped: {file_path.name} (temp file)")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"skipped: {file_path.name} (not .xlsx)")
            continue
        if file_path.stem.startswith(output_prefix):
            print(f"skipped: {file_path.name} (looks like prior output)")
            continue
        candidates.append(file_path)
    return candidates


def find_sheet_case_insensitive(wb: xw.Book, name: str) -> xw.Sheet | None:
    target = name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def find_max_anchor(sheet: xw.Sheet) -> tuple[int, int] | None:
    search_specs = [
        {"What": "max", "LookAt": 1, "MatchCase": False},
        {"What": "max", "LookAt": 2, "MatchCase": False},
        {"What": "MAX", "LookAt": 1, "MatchCase": False},
    ]
    for spec in search_specs:
        try:
            found = sheet.api.Cells.Find(**spec)
        except Exception:
            found = None
        if found is not None:
            return int(found.Row), int(found.Column)
    return None


def build_header_cells(sheet: xw.Sheet, anchor_row: int, anchor_col: int) -> list[tuple[str, int]]:
    row_start = max(1, anchor_row - 2)
    row_end = anchor_row + 2
    col_start = max(1, anchor_col - 25)
    col_end = anchor_col + 25

    values = sheet.range((row_start, col_start), (row_end, col_end)).value
    if not isinstance(values, list):
        values = [[values]]
    elif values and not isinstance(values[0], list):
        values = [values]

    headers: list[tuple[str, int]] = []
    for row_index, row_values in enumerate(values):
        for col_index, value in enumerate(row_values):
            normalized = normalize_label(value)
            if not normalized:
                continue
            r_off = (row_start + row_index) - anchor_row
            c_off = (col_start + col_index) - anchor_col
            distance = abs(r_off) + abs(c_off)
            headers.append((normalized, c_off if distance >= 0 else c_off))
    return headers


def infer_offset(
    header_cells: list[tuple[str, int]],
    token_groups: list[tuple[str, ...]],
    default_offset: int,
) -> int:
    best_offset = None
    best_score = None
    for label, c_off in header_cells:
        for token_group in token_groups:
            if all(token in label for token in token_group):
                score = abs(c_off)
                if best_score is None or score < best_score:
                    best_score = score
                    best_offset = c_off
    return best_offset if best_offset is not None else default_offset


def resolve_columns(
    anchor_col: int,
    header_cells: list[tuple[str, int]],
    defaults: dict[str, int],
) -> dict[str, int]:
    resolved: dict[str, int] = {}
    for field_name, default_offset in defaults.items():
        token_groups = HEADER_TOKENS.get(field_name, [])
        offset = infer_offset(header_cells, token_groups, default_offset) if token_groups else default_offset
        resolved[field_name] = anchor_col + offset
    return resolved


def find_num_quarter_rows(
    sheet: xw.Sheet,
    num_quarters_col: int,
    anchor_row: int,
    max_rows: int = 10,
) -> list[tuple[int, int]]:
    if num_quarters_col < 1:
        return []

    entries: list[tuple[int, int]] = []
    row_start = max(1, anchor_row - 60)
    row_end = anchor_row + 60
    for row in range(row_start, row_end + 1):
        num_quarters = to_int(safe_cell_value(sheet, row, num_quarters_col))
        if num_quarters is None or num_quarters < 1 or num_quarters > 100:
            continue
        entries.append((num_quarters, row))

    if not entries:
        return []

    best_by_n: dict[int, tuple[tuple[int, int], int]] = {}
    for n_quarters, row in entries:
        score = (abs(row - anchor_row), row)
        if n_quarters not in best_by_n or score < best_by_n[n_quarters][0]:
            best_by_n[n_quarters] = (score, row)

    selected = [(n_quarters, best_by_n[n_quarters][1]) for n_quarters in sorted(best_by_n)]
    return selected[:max_rows]


def collect_contiguous_series_rows(
    sheet: xw.Sheet,
    x_col: int,
    y_col: int,
    anchor_row: int,
) -> list[int]:
    if x_col < 1 or y_col < 1:
        return []

    row_start = max(1, anchor_row - 150)
    row_end = anchor_row + 150
    numeric_rows: list[int] = []
    for row in range(row_start, row_end + 1):
        x_value = to_float(safe_cell_value(sheet, row, x_col))
        y_value = to_float(safe_cell_value(sheet, row, y_col))
        if x_value is not None and y_value is not None:
            numeric_rows.append(row)

    if len(numeric_rows) < 2:
        return []

    blocks: list[list[int]] = []
    current = [numeric_rows[0]]
    for row in numeric_rows[1:]:
        if row == current[-1] + 1:
            current.append(row)
        else:
            blocks.append(current)
            current = [row]
    blocks.append(current)

    blocks = [block for block in blocks if len(block) >= 2]
    if not blocks:
        return []

    blocks.sort(
        key=lambda block: (
            -len(block),
            min(abs(anchor_row - block[0]), abs(anchor_row - block[-1])),
            block[0],
        )
    )
    return blocks[0]


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
    except Exception as close_error:
        print(f"warning: could not close workbook cleanly -> {close_error}")


def calculate_range_width(max_value: float | None, min_value: float | None) -> float | None:
    if max_value is None or min_value is None:
        return None
    return max_value - min_value


def process_empirical_sheet(
    wb: xw.Book,
    labels: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    sheet = find_sheet_case_insensitive(wb, "Empirical Model")
    if sheet is None:
        return []

    anchor = find_max_anchor(sheet)
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    headers = build_header_cells(sheet, anchor_row, anchor_col)
    columns = resolve_columns(anchor_col, headers, EMP_DEFAULT_OFFSETS)
    detected_quarter_rows = find_num_quarter_rows(
        sheet,
        columns["num_quarters_used"],
        anchor_row,
        max_rows=10,
    )
    detected_row_map = {n: row for n, row in detected_quarter_rows}
    quarter_rows = [(n, detected_row_map.get(n, anchor_row + n)) for n in range(1, 11)]

    avg_formula_col = anchor_col + 8
    can_compute_avg = columns["sales_captured_in_db_pct"] > 0
    formula_targets: list[tuple[int, int]] = []
    if can_compute_avg:
        for fallback_n, row in quarter_rows:
            num_quarters = to_int(safe_cell_value(sheet, row, columns["num_quarters_used"])) or fallback_n
            start_row = max(1, row - num_quarters + 1)
            formula = (
                f"=AVERAGE("
                f"R{start_row}C{columns['sales_captured_in_db_pct']}:"
                f"R{row}C{columns['sales_captured_in_db_pct']}"
                f")"
            )
            set_r1c1_formula2(sheet.range((row, avg_formula_col)), formula)
            formula_targets.append((row, avg_formula_col))

    if formula_targets:
        wb.app.calculate()

    rows: list[dict[str, Any]] = []
    for fallback_n, row in quarter_rows:
        num_quarters = to_int(safe_cell_value(sheet, row, columns["num_quarters_used"])) or fallback_n
        last_quarter_used = safe_cell_value(sheet, row, columns["last_quarter_used"])
        avg_pen_from_formula = to_float(safe_cell_value(sheet, row, avg_formula_col))
        avg_pen_from_sheet = to_float(safe_cell_value(sheet, row, columns["avg_penetration_pct"]))
        avg_penetration = avg_pen_from_formula if avg_pen_from_formula is not None else avg_pen_from_sheet

        quarterly_sales = to_float(safe_cell_value(sheet, row, columns["quarterly_sales"]))
        reported_sales = to_float(safe_cell_value(sheet, row, columns["reported_sales"]))
        growth_rate_pct = to_float(safe_cell_value(sheet, row, columns["growth_rate_pct"]))
        sales_captured_pct = to_float(safe_cell_value(sheet, row, columns["sales_captured_in_db_pct"]))

        forecast_value = to_float(safe_cell_value(sheet, row, columns["forecast_value"]))
        if forecast_value is None and quarterly_sales is not None and avg_penetration not in (None, 0):
            penetration_ratio = avg_penetration / 100.0 if avg_penetration > 1 else avg_penetration
            if penetration_ratio:
                forecast_value = quarterly_sales / penetration_ratio

        actual_value = to_float(safe_cell_value(sheet, row, columns["actual_value"]))
        if actual_value is None:
            actual_value = reported_sales

        forecast_max = to_float(safe_cell_value(sheet, row, columns["forecast_max"]))
        forecast_min = to_float(safe_cell_value(sheet, row, columns["forecast_min"]))
        range_width = calculate_range_width(forecast_max, forecast_min)

        if (
            forecast_value is None
            and forecast_max is None
            and forecast_min is None
            and avg_penetration is None
            and quarterly_sales is None
            and reported_sales is None
        ):
            continue

        rows.append(
            {
                "model": labels["model"],
                "ticker": labels["ticker"],
                "model_period": labels["model_period"],
                "model_date": labels["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_pct,
                "source_file": source_file,
            }
        )
    return rows


def rows_are_equal(row_a: dict[str, Any], row_b: dict[str, Any]) -> bool:
    compare_keys = [
        "num_quarters_used",
        "forecast_value",
        "forecast_max",
        "forecast_min",
        "intercept",
        "slope",
    ]
    for key in compare_keys:
        a_value = row_a.get(key)
        b_value = row_b.get(key)
        if a_value is None and b_value is None:
            continue
        if isinstance(a_value, (int, float)) and isinstance(b_value, (int, float)):
            if abs(float(a_value) - float(b_value)) > 1e-9:
                return False
            continue
        if a_value != b_value:
            return False
    return True


def process_regression_sheet(
    wb: xw.Book,
    labels: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    sheet = find_sheet_case_insensitive(wb, "Regression Model")
    if sheet is None:
        return []

    anchor = find_max_anchor(sheet)
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    headers = build_header_cells(sheet, anchor_row, anchor_col)
    columns = resolve_columns(anchor_col, headers, REG_DEFAULT_OFFSETS)
    quarter_rows = find_num_quarter_rows(sheet, columns["num_quarters_used"], anchor_row, max_rows=10)
    quarter_row_map = {n: row for n, row in quarter_rows}

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    series_rows = collect_contiguous_series_rows(sheet, x_col, y_col, anchor_row)
    if len(series_rows) < 2:
        return []

    n_limit = min(10, len(series_rows))
    intercept_col = anchor_col + 8
    slope_col = anchor_col + 9

    formula_rows: list[tuple[int, int]] = []
    for n_quarters in range(1, n_limit + 1):
        start_row = series_rows[-n_quarters]
        end_row = series_rows[-1]
        target_row = anchor_row + n_quarters + 1

        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )

        set_r1c1_formula2(sheet.range((target_row, intercept_col)), intercept_formula)
        set_r1c1_formula2(sheet.range((target_row, slope_col)), slope_formula)
        formula_rows.append((n_quarters, target_row))

    if formula_rows:
        wb.app.calculate()

    latest_x = to_float(safe_cell_value(sheet, series_rows[-1], x_col))
    rows: list[dict[str, Any]] = []
    for n_quarters, formula_row in formula_rows:
        intercept = to_float(safe_cell_value(sheet, formula_row, intercept_col))
        slope = to_float(safe_cell_value(sheet, formula_row, slope_col))

        table_row = quarter_row_map.get(n_quarters, anchor_row + n_quarters)
        forecast_value = to_float(safe_cell_value(sheet, table_row, columns["forecast_value"]))
        actual_value = to_float(safe_cell_value(sheet, table_row, columns["actual_value"]))
        forecast_max = to_float(safe_cell_value(sheet, table_row, columns["forecast_max"]))
        forecast_min = to_float(safe_cell_value(sheet, table_row, columns["forecast_min"]))

        if forecast_value is None and intercept is not None and slope is not None and latest_x is not None:
            forecast_value = intercept + (slope * latest_x)

        range_width = calculate_range_width(forecast_max, forecast_min)

        if (
            forecast_value is None
            and intercept is None
            and slope is None
            and forecast_max is None
            and forecast_min is None
        ):
            continue

        candidate = {
            "model": labels["model"],
            "ticker": labels["ticker"],
            "model_period": labels["model_period"],
            "model_date": labels["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n_quarters,
            "num_quarters_used": n_quarters,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        if rows and rows_are_equal(rows[-1], candidate):
            continue
        rows.append(candidate)

    return rows


def write_rows_to_sheet(
    sheet: xw.Sheet,
    columns: list[str],
    rows: list[dict[str, Any]],
) -> None:
    data = [columns] + [[row.get(column) for column in columns] for row in rows]
    sheet.range("A1").value = data
    sheet.range((1, 1), (1, len(columns))).api.Font.Bold = True
    sheet.range((1, 1), (max(1, len(data)), len(columns))).api.AutoFilter()

    try:
        sheet.activate()
        sheet.range("A2").select()
        sheet.book.app.api.ActiveWindow.FreezePanes = True
    except Exception:
        pass

    sample_rows = data[: min(len(data), 80)]
    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for sample_row in sample_rows[1:]:
            value = sample_row[col_idx - 1]
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        width = min(max(max_len + 2, 12), 42)
        sheet.range((1, col_idx), (1, col_idx)).column_width = width


def write_output_workbook(
    app: xw.App,
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    out_wb = app.books.add()
    try:
        while len(out_wb.sheets) < 2:
            out_wb.sheets.add(after=out_wb.sheets[-1])

        empirical_sheet = out_wb.sheets[0]
        regression_sheet = out_wb.sheets[1]
        empirical_sheet.name = "empirical_candidates"
        regression_sheet.name = "regression_candidates"

        for extra_sheet in list(out_wb.sheets)[2:]:
            extra_sheet.delete()

        write_rows_to_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)
        write_rows_to_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)
        out_wb.save(str(output_path))
    finally:
        safe_close_workbook(out_wb)


def run() -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    files = collect_input_files(input_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0
    output_path = unique_output_path(input_dir, output_dir)

    app: xw.App | None = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        try:
            app.api.EnableEvents = False
        except Exception:
            pass
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in files:
            wb: xw.Book | None = None
            try:
                labels = parse_filename_labels(file_path)
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(process_empirical_sheet(wb, labels, file_path.name))
                regression_rows.extend(process_regression_sheet(wb, labels, file_path.name))
                processed_files += 1
                print(f"processed: {file_path.name}")
            except Exception as file_error:
                print(f"skipped: {file_path.name} (error: {file_error})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)

        write_output_workbook(app, output_path, empirical_rows, regression_rows)
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass

    print(f"output_path: {output_path}")
    print(f"files_processed: {processed_files}")
    print(f"empirical_rows: {len(empirical_rows)}")
    print(f"regression_rows: {len(regression_rows)}")


if __name__ == "__main__":
    run()

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import math
import re
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font


# ---------------------------
# User-editable input/output
# ---------------------------
input_dir = Path("./input")
output_dir = Path("./output")


EMPIRICAL_HEADERS = [
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


REGRESSION_HEADERS = [
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


N_QUARTERS = 10
EMPIRICAL_QUARTERLY_COL_OFFSET = -11
EMPIRICAL_REPORTED_COL_OFFSET = -7
REGRESSION_X_COL_OFFSET = -11
REGRESSION_Y_COL_OFFSET = -7


MONTH_MAP = {
    "jan": 1,
    "feb": 2,
    "mar": 3,
    "apr": 4,
    "may": 5,
    "jun": 6,
    "jul": 7,
    "aug": 8,
    "sep": 9,
    "oct": 10,
    "nov": 11,
    "dec": 12,
}

PERIOD_DAY_MAP = {"early": 5, "mid": 15, "late": 25}


@dataclass
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


def log(msg: str) -> None:
    print(msg, flush=True)


def is_number(value: Any) -> bool:
    if value is None:
        return False
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return False
        return True
    return False


def to_float(value: Any) -> float | None:
    if not is_number(value):
        return None
    return float(value)


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    return re.sub(r"\s+", " ", text)


def format_model_date(year: int, month: int, day: int) -> str:
    return datetime(year=year, month=month, day=day).strftime("%Y-%m-%d")


def parse_file_label(file_path: Path) -> FileLabel:
    stem = file_path.stem
    parts = [p.strip() for p in stem.split(" - ")]

    ticker = parts[1] if len(parts) >= 2 else ""
    period_token = parts[2] if len(parts) >= 3 else ""
    period_match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3})\s*(\d{4})",
        period_token,
        flags=re.IGNORECASE,
    )

    if period_match:
        period_name = period_match.group(1).title()
        month_name = period_match.group(2).title()
        year = int(period_match.group(3))
        month_num = MONTH_MAP.get(month_name.lower())
        day = PERIOD_DAY_MAP[period_name.lower()]

        if month_num is not None:
            model_period = f"{period_name}{month_name}_{year}"
            model_date = format_model_date(year, month_num, day)
        else:
            model_period = period_token.replace(" ", "_")
            model_date = ""
    else:
        model_period = period_token.replace(" ", "_")
        model_date = ""

    if not ticker:
        ticker = parts[0].strip() if parts else stem
    model = f"{ticker}_{model_period}" if model_period else ticker
    return FileLabel(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def next_output_path(input_folder: Path, target_dir: Path) -> Path:
    base_name = f"{input_folder.name}_PARAM"
    output_file = target_dir / f"{base_name}.xlsx"
    if not output_file.exists():
        return output_file

    i = 1
    while True:
        candidate = target_dir / f"{base_name}.{i}.xlsx"
        if not candidate.exists():
            return candidate
        i += 1


def ensure_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values
    return [values]


@dataclass
class SheetGrid:
    values: list[list[Any]]
    min_row: int
    min_col: int
    max_row: int
    max_col: int

    def in_bounds(self, row: int, col: int) -> bool:
        return self.min_row <= row <= self.max_row and self.min_col <= col <= self.max_col

    def get(self, row: int, col: int) -> Any:
        if not self.in_bounds(row, col):
            return None
        return self.values[row - self.min_row][col - self.min_col]


def load_sheet_grid(sheet: xw.Sheet) -> SheetGrid:
    used = sheet.used_range
    values = ensure_2d(used.value)
    min_row, min_col = used.row, used.column
    max_row = min_row + max(len(values), 1) - 1
    max_col = min_col + max((len(r) for r in values), default=1) - 1
    return SheetGrid(values=values, min_row=min_row, min_col=min_col, max_row=max_row, max_col=max_col)


def find_anchor_cell(grid: SheetGrid, target: str = "max") -> tuple[int, int] | None:
    target_norm = normalize_text(target)
    for r in range(grid.min_row, grid.max_row + 1):
        for c in range(grid.min_col, grid.max_col + 1):
            if normalize_text(grid.get(r, c)) == target_norm:
                return r, c
    return None


def find_candidate_cols(grid: SheetGrid, anchor_row: int, anchor_col: int, cap: int = N_QUARTERS) -> list[int]:
    candidate_cols: list[int] = []
    found_numeric = False
    for c in range(anchor_col + 1, grid.max_col + 1):
        value = grid.get(anchor_row, c)
        if is_number(value):
            candidate_cols.append(c)
            found_numeric = True
            if len(candidate_cols) >= cap:
                break
        elif found_numeric:
            break
    return candidate_cols


def find_row_with_keywords(
    grid: SheetGrid,
    anchor_row: int,
    anchor_col: int,
    keywords: tuple[str, ...],
    row_window: int = 120,
    col_window_left: int = 20,
    col_window_right: int = 2,
) -> int | None:
    start_row = max(grid.min_row, anchor_row - row_window)
    end_row = min(grid.max_row, anchor_row + row_window)
    start_col = max(grid.min_col, anchor_col - col_window_left)
    end_col = min(grid.max_col, anchor_col + col_window_right)

    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            txt = normalize_text(grid.get(r, c))
            if any(k in txt for k in keywords):
                return r
    return None


def map_num_quarters_by_col(grid: SheetGrid, anchor_row: int, candidate_cols: list[int]) -> dict[int, int]:
    if not candidate_cols:
        return {}
    best_row = None
    best_score = -1
    search_start = max(grid.min_row, anchor_row - 20)
    search_end = min(grid.max_row, anchor_row + 5)

    for r in range(search_start, search_end + 1):
        score = 0
        for c in candidate_cols:
            value = grid.get(r, c)
            if is_number(value):
                v = int(round(float(value)))
                if 1 <= v <= 40:
                    score += 1
        if score > best_score:
            best_score = score
            best_row = r

    mapping: dict[int, int] = {}
    if best_row is not None and best_score > 0:
        for idx, c in enumerate(candidate_cols, start=1):
            value = grid.get(best_row, c)
            if is_number(value):
                mapping[c] = int(round(float(value)))
            else:
                mapping[c] = idx
    else:
        for idx, c in enumerate(candidate_cols, start=1):
            mapping[c] = idx
    return mapping


def collect_recent_numeric_rows(
    grid: SheetGrid,
    x_col: int,
    y_col: int,
    end_row: int,
    max_points: int = 60,
) -> list[int]:
    rows: list[int] = []
    r = min(end_row, grid.max_row)
    while r >= grid.min_row and len(rows) < max_points:
        xv = grid.get(r, x_col)
        yv = grid.get(r, y_col)
        if is_number(xv) and is_number(yv):
            rows.append(r)
        elif rows:
            break
        r -= 1
    rows.reverse()
    return rows


def safe_close_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
    except TypeError:
        try:
            wb.api.Close(SaveChanges=False)
        except Exception:
            wb.close()
    except Exception:
        try:
            wb.api.Close(SaveChanges=False)
        except Exception:
            pass


def set_formula2(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        # Fallback for Excel builds that don't expose Formula2.
        cell.formula = formula_r1c1


def first_non_empty_label(grid: SheetGrid, row: int, search_until_col: int) -> str:
    end_col = min(search_until_col, grid.max_col)
    for c in range(grid.min_col, end_col + 1):
        value = grid.get(row, c)
        if value not in (None, ""):
            return str(value)
    return ""


def get_metric_value(grid: SheetGrid, row: int | None, col: int) -> float | None:
    if row is None:
        return None
    return to_float(grid.get(row, col))


def build_col_for_quarters(candidate_cols: list[int], num_quarters_map: dict[int, int]) -> dict[int, int]:
    out: dict[int, int] = {}
    for idx, c in enumerate(candidate_cols, start=1):
        q = num_quarters_map.get(c, idx)
        out[q] = c
    return out


def extract_empirical_rows(
    wb: xw.Book,
    sheet: xw.Sheet,
    file_label: FileLabel,
    source_file: str,
) -> list[dict[str, Any]]:
    grid = load_sheet_grid(sheet)
    anchor = find_anchor_cell(grid, "max")
    if anchor is None:
        log(f"  empirical skipped: no 'max' anchor on sheet '{sheet.name}'")
        return []

    anchor_row, anchor_col = anchor
    candidate_cols = find_candidate_cols(grid, anchor_row, anchor_col, cap=N_QUARTERS)
    num_quarters_map = map_num_quarters_by_col(grid, anchor_row, candidate_cols)
    col_for_quarters = build_col_for_quarters(candidate_cols, num_quarters_map)

    min_row = find_row_with_keywords(grid, anchor_row, anchor_col, ("min",))
    forecast_row = find_row_with_keywords(
        grid,
        anchor_row,
        anchor_col,
        ("estimated total sold", "est total sold", "forecast", "tot fcst"),
    )
    actual_row = find_row_with_keywords(grid, anchor_row, anchor_col, ("reported sales", "actual"))
    growth_row = find_row_with_keywords(grid, anchor_row, anchor_col, ("growth",))
    captured_row = find_row_with_keywords(grid, anchor_row, anchor_col, ("sales captured", "captured in db"))

    quarterly_col = anchor_col + EMPIRICAL_QUARTERLY_COL_OFFSET
    reported_col = anchor_col + EMPIRICAL_REPORTED_COL_OFFSET
    data_rows = collect_recent_numeric_rows(grid, quarterly_col, reported_col, anchor_row - 1, max_points=80)
    if not data_rows:
        log(f"  empirical skipped: no numeric history found using anchor offsets in '{sheet.name}'")
        return []

    helper_col = max(grid.max_col + 5, anchor_col + 5)
    helper_row = anchor_row
    helper_cell = sheet.range((helper_row, helper_col))

    out: list[dict[str, Any]] = []
    max_n = min(N_QUARTERS, len(data_rows))
    for n in range(1, max_n + 1):
        start_row = data_rows[-n]
        end_row = data_rows[-1]

        # R1C1 formula2 to compute average penetration for the last N quarters.
        avg_formula = (
            f"=AVERAGE((R{start_row}C{quarterly_col}:R{end_row}C{quarterly_col})/"
            f"(R{start_row}C{reported_col}:R{end_row}C{reported_col}))"
        )
        set_formula2(helper_cell, avg_formula)
        wb.app.calculate()
        avg_penetration = to_float(helper_cell.value)

        quarterly_sales_vals = [
            to_float(grid.get(r, quarterly_col)) for r in range(start_row, end_row + 1)
        ]
        reported_sales_vals = [
            to_float(grid.get(r, reported_col)) for r in range(start_row, end_row + 1)
        ]
        quarterly_sales_sum = sum(v for v in quarterly_sales_vals if v is not None)
        reported_sales_sum = sum(v for v in reported_sales_vals if v is not None)
        last_quarter_used = first_non_empty_label(grid, start_row, search_until_col=anchor_col - 1)

        metric_col = col_for_quarters.get(n)
        forecast_max = get_metric_value(grid, anchor_row, metric_col) if metric_col else None
        forecast_min = get_metric_value(grid, min_row, metric_col) if metric_col else None
        forecast_value = get_metric_value(grid, forecast_row, metric_col) if metric_col else None
        actual_value = get_metric_value(grid, actual_row, metric_col) if metric_col else None
        growth_rate_pct = get_metric_value(grid, growth_row, metric_col) if metric_col else None
        sales_captured = get_metric_value(grid, captured_row, metric_col) if metric_col else None
        range_width = (
            forecast_max - forecast_min if forecast_max is not None and forecast_min is not None else None
        )

        out.append(
            {
                "model": file_label.model,
                "ticker": file_label.ticker,
                "model_period": file_label.model_period,
                "model_date": file_label.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": n,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales_sum,
                "reported_sales": reported_sales_sum,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured,
                "source_file": source_file,
            }
        )

    helper_cell.value = None
    return out


def extract_regression_rows(
    wb: xw.Book,
    sheet: xw.Sheet,
    file_label: FileLabel,
    source_file: str,
) -> list[dict[str, Any]]:
    grid = load_sheet_grid(sheet)
    anchor = find_anchor_cell(grid, "max")
    if anchor is None:
        log(f"  regression skipped: no 'max' anchor on sheet '{sheet.name}'")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col + REGRESSION_Y_COL_OFFSET
    x_col = anchor_col + REGRESSION_X_COL_OFFSET
    data_rows = collect_recent_numeric_rows(grid, x_col, y_col, anchor_row - 1, max_points=80)
    if not data_rows:
        log(f"  regression skipped: no numeric X/Y history found in '{sheet.name}'")
        return []

    candidate_cols = find_candidate_cols(grid, anchor_row, anchor_col, cap=N_QUARTERS)
    num_quarters_map = map_num_quarters_by_col(grid, anchor_row, candidate_cols)
    col_for_quarters = build_col_for_quarters(candidate_cols, num_quarters_map)

    min_row = find_row_with_keywords(grid, anchor_row, anchor_col, ("min",))
    forecast_row = find_row_with_keywords(
        grid,
        anchor_row,
        anchor_col,
        ("tot fcst w/o sa", "tot fcst", "forecast"),
    )
    actual_row = find_row_with_keywords(grid, anchor_row, anchor_col, ("actual", "reported sales"))

    helper_col = max(grid.max_col + 7, anchor_col + 7)
    intercept_cell = sheet.range((anchor_row, helper_col))
    slope_cell = sheet.range((anchor_row + 1, helper_col))

    out: list[dict[str, Any]] = []
    prev_signature: tuple[Any, ...] | None = None
    max_n = min(N_QUARTERS, len(data_rows))
    for n in range(1, max_n + 1):
        start_row = data_rows[-n]
        end_row = data_rows[-1]

        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )

        set_formula2(intercept_cell, intercept_formula)
        set_formula2(slope_cell, slope_formula)
        wb.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)

        metric_col = col_for_quarters.get(n)
        forecast_value = get_metric_value(grid, forecast_row, metric_col) if metric_col else None
        actual_value = get_metric_value(grid, actual_row, metric_col) if metric_col else None
        forecast_max = get_metric_value(grid, anchor_row, metric_col) if metric_col else None
        forecast_min = get_metric_value(grid, min_row, metric_col) if metric_col else None
        range_width = (
            forecast_max - forecast_min if forecast_max is not None and forecast_min is not None else None
        )

        row = {
            "model": file_label.model,
            "ticker": file_label.ticker,
            "model_period": file_label.model_period,
            "model_date": file_label.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n,
            "num_quarters_used": n,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        signature = (
            row["forecast_value"],
            row["forecast_max"],
            row["forecast_min"],
            row["intercept"],
            row["slope"],
        )
        if signature != prev_signature:
            out.append(row)
        prev_signature = signature

    intercept_cell.value = None
    slope_cell.value = None
    return out


def write_sheet(ws, headers: list[str], rows: list[dict[str, Any]]) -> None:
    ws.append(headers)
    for row in rows:
        ws.append([row.get(h, "") for h in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for idx, header in enumerate(headers, start=1):
        values = [str(header)]
        for r in range(2, ws.max_row + 1):
            v = ws.cell(row=r, column=idx).value
            if v is None:
                continue
            values.append(str(v))
        max_len = max(len(v) for v in values) if values else len(header)
        ws.column_dimensions[ws.cell(row=1, column=idx).column_letter].width = min(max(max_len + 2, 12), 42)


def save_output(
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
    output_path: Path,
) -> None:
    wb = Workbook()
    default_ws = wb.active
    wb.remove(default_ws)

    empirical_ws = wb.create_sheet("empirical_candidates")
    write_sheet(empirical_ws, EMPIRICAL_HEADERS, empirical_rows)

    regression_ws = wb.create_sheet("regression_candidates")
    write_sheet(regression_ws, REGRESSION_HEADERS, regression_rows)

    wb.save(output_path)


def get_workbook_sheet(wb: xw.Book, sheet_name: str) -> xw.Sheet | None:
    for sh in wb.sheets:
        if sh.name.strip().lower() == sheet_name.strip().lower():
            return sh
    return None


def process_workbooks(input_path: Path, output_path: Path) -> tuple[int, int, int]:
    files_processed = 0
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []

    xlsx_files = sorted(input_path.iterdir(), key=lambda p: p.name.lower())
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.calculation = "manual"

    try:
        for file_path in xlsx_files:
            if file_path.is_dir():
                continue
            if file_path.name.startswith("~"):
                log(f"skipped {file_path.name}: temp file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                log(f"skipped {file_path.name}: not .xlsx")
                continue

            log(f"processing {file_path.name}")
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                file_label = parse_file_label(file_path)

                empirical_sheet = get_workbook_sheet(wb, "Empirical Model")
                if empirical_sheet is None:
                    log(f"  empirical skipped: sheet 'Empirical Model' not found")
                else:
                    empirical_rows.extend(
                        extract_empirical_rows(wb, empirical_sheet, file_label, file_path.name)
                    )

                regression_sheet = get_workbook_sheet(wb, "Regression Model")
                if regression_sheet is None:
                    log(f"  regression skipped: sheet 'Regression Model' not found")
                else:
                    regression_rows.extend(
                        extract_regression_rows(wb, regression_sheet, file_label, file_path.name)
                    )

                files_processed += 1
            except Exception as exc:
                log(f"skipped {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    save_output(empirical_rows, regression_rows, output_path)
    return files_processed, len(empirical_rows), len(regression_rows)


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"input_dir not found: {input_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = next_output_path(input_dir.resolve(), output_dir.resolve())

    files_processed, empirical_count, regression_count = process_workbooks(input_dir, output_path)

    log(f"output path: {output_path}")
    log(f"number of files processed: {files_processed}")
    log(f"number of empirical rows: {empirical_count}")
    log(f"number of regression rows: {regression_count}")


if __name__ == "__main__":
    main()

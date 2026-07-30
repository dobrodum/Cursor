#!/usr/bin/env python3
"""
Build one consolidated PARAM workbook from all .xlsx source models.

The script opens each source workbook only once, extracts candidates from both:
  - Empirical Model
  - Regression Model
and writes two sheets to one output workbook:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Iterable

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# User-configurable paths
# ---------------------------------------------------------------------------
input_dir = Path("./input")
output_dir = Path("./output")


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
N_QUARTERS = 10
PERIOD_DAY_MAP = {"early": 5, "mid": 15, "late": 25}
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
class ModelMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetContext:
    values: list[list[Any]]
    top_row: int
    left_col: int
    last_row: int
    last_col: int


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def as_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if isinstance(values, (list, tuple)):
        outer = list(values)
        if not outer:
            return []
        if not isinstance(outer[0], (list, tuple)):
            return [outer]
        return [list(row) if isinstance(row, (list, tuple)) else [row] for row in outer]
    return [[values]]


def to_number(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    text = text.replace(",", "")
    if text.endswith("%"):
        text = text[:-1]
    try:
        return float(text)
    except ValueError:
        return None


def is_numeric(value: Any) -> bool:
    return to_number(value) is not None


def safe_strip(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text if text else None


def model_meta_from_filename(file_path: Path) -> ModelMeta:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = (parts[1] if len(parts) >= 2 else "UNKNOWN").strip().upper()

    # Example: MidJan2026 -> MidJan_2026 and 2026-01-15
    match = re.search(r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(\d{4})", stem, flags=re.IGNORECASE)
    if not match:
        return ModelMeta(
            model=f"{ticker}_unknown_period",
            ticker=ticker,
            model_period="unknown_period",
            model_date="",
        )

    phase_raw, month_raw, year_raw = match.groups()
    phase = phase_raw.title()
    month_abbrev = month_raw[:3].title()
    month_num = MONTH_MAP.get(month_abbrev.lower())
    year = int(year_raw)
    day = PERIOD_DAY_MAP[phase.lower()]

    model_period = f"{phase}{month_abbrev}_{year}"
    if month_num is None:
        model_date = ""
    else:
        model_date = date(year, month_num, day).isoformat()

    return ModelMeta(
        model=f"{ticker}_{model_period}",
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def get_output_path(input_folder: Path, out_folder: Path) -> Path:
    out_folder.mkdir(parents=True, exist_ok=True)
    base = f"{input_folder.name}_PARAM"
    first_candidate = out_folder / f"{base}.xlsx"
    if not first_candidate.exists():
        return first_candidate

    counter = 1
    while True:
        candidate = out_folder / f"{base}.{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


def close_workbook_no_save(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    # Safe fallback when close(save=False) is unsupported
    try:
        wb.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        wb.close()
    except Exception:
        pass


def read_sheet_context(sheet: xw.Sheet) -> SheetContext:
    used = sheet.used_range
    values = as_2d(used.value)
    if not values:
        return SheetContext(values=[], top_row=1, left_col=1, last_row=1, last_col=1)

    top_row = used.row
    left_col = used.column
    width = max(len(row) for row in values) if values else 1
    last_row = top_row + len(values) - 1
    last_col = left_col + width - 1
    return SheetContext(values=values, top_row=top_row, left_col=left_col, last_row=last_row, last_col=last_col)


def ctx_value(ctx: SheetContext, row: int, col: int) -> Any:
    r_idx = row - ctx.top_row
    c_idx = col - ctx.left_col
    if r_idx < 0 or c_idx < 0 or r_idx >= len(ctx.values):
        return None
    row_vals = ctx.values[r_idx]
    if c_idx >= len(row_vals):
        return None
    return row_vals[c_idx]


def find_anchor(ctx: SheetContext, keyword: str = "max") -> tuple[int, int]:
    target = normalize_text(keyword)
    for r_offset, row_values in enumerate(ctx.values):
        for c_offset, value in enumerate(row_values):
            if normalize_text(value) == target:
                return ctx.top_row + r_offset, ctx.left_col + c_offset
    raise ValueError(f"Could not find '{keyword}' anchor")


def find_column_by_phrases(
    ctx: SheetContext,
    phrases: Iterable[str],
    anchor_row: int,
    row_min: int | None = None,
    row_max: int | None = None,
) -> int | None:
    phrase_set = [normalize_text(p) for p in phrases]
    if not phrase_set:
        return None

    min_row = row_min if row_min is not None else ctx.top_row
    max_row = row_max if row_max is not None else ctx.last_row

    best: tuple[int, int] | None = None
    for row in range(min_row, max_row + 1):
        for col in range(ctx.left_col, ctx.last_col + 1):
            text = normalize_text(ctx_value(ctx, row, col))
            if not text:
                continue
            if any(phrase in text for phrase in phrase_set):
                score = abs(row - anchor_row) * 100 + abs(col - ctx.left_col)
                if best is None or score < best[0]:
                    best = (score, col)
    return best[1] if best else None


def find_contiguous_numeric_range_upwards(ctx: SheetContext, col: int, from_row: int) -> tuple[int, int] | None:
    row = from_row
    while row >= ctx.top_row and not is_numeric(ctx_value(ctx, row, col)):
        row -= 1
    if row < ctx.top_row:
        return None

    end_row = row
    start_row = end_row
    while start_row - 1 >= ctx.top_row and is_numeric(ctx_value(ctx, start_row - 1, col)):
        start_row -= 1
    return start_row, end_row


def collect_summary_rows(
    ctx: SheetContext,
    start_row: int,
    preferred_col: int | None,
    fallback_col: int | None,
    limit: int,
) -> list[int]:
    rows: list[int] = []
    blank_streak = 0
    row = start_row
    while row <= ctx.last_row and len(rows) < limit:
        preferred = ctx_value(ctx, row, preferred_col) if preferred_col else None
        fallback = ctx_value(ctx, row, fallback_col) if fallback_col else None
        if is_numeric(preferred) or is_numeric(fallback):
            rows.append(row)
            blank_streak = 0
        elif rows:
            blank_streak += 1
            if blank_streak >= 2:
                break
        row += 1

    if rows:
        return rows[:limit]

    # Fallback: assume the table starts immediately below the anchor
    max_row = min(ctx.last_row, start_row + limit - 1)
    return list(range(start_row, max_row + 1))


def build_last_n_avg_formula_r1c1(start_row: int, end_row: int, col: int, n: int) -> str:
    rng = f"R{start_row}C{col}:R{end_row}C{col}"
    return f"=AVERAGE(INDEX({rng},ROWS({rng})-{n}+1):INDEX({rng},ROWS({rng})))"


def set_formula2_r1c1(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        # Fallback for environments that require explicit Formula2R1C1
        cell.api.Formula2R1C1 = formula


def extract_empirical_rows(wb: xw.Book, meta: ModelMeta, source_file: str) -> list[dict[str, Any]]:
    if "Empirical Model" not in [s.name for s in wb.sheets]:
        return []

    sheet = wb.sheets["Empirical Model"]
    ctx = read_sheet_context(sheet)
    anchor_row, anchor_col = find_anchor(ctx, "max")

    min_col = find_column_by_phrases(ctx, ["min"], anchor_row, row_min=anchor_row - 2, row_max=anchor_row + 2) or (anchor_col + 1)
    forecast_col = find_column_by_phrases(
        ctx,
        ["estimated total sold", "forecast value", "forecast total"],
        anchor_row,
        row_min=anchor_row - 3,
        row_max=anchor_row + 3,
    ) or (anchor_col - 1)
    actual_col = find_column_by_phrases(
        ctx,
        ["reported sales", "actual value", "actual sales"],
        anchor_row,
        row_min=anchor_row - 3,
        row_max=anchor_row + 3,
    ) or (anchor_col - 2)
    num_quarters_col = find_column_by_phrases(
        ctx,
        ["num quarters used", "quarters used", "n quarters", "num quarters"],
        anchor_row,
        row_min=anchor_row - 3,
        row_max=anchor_row + 3,
    ) or (anchor_col - 4)
    last_quarter_col = find_column_by_phrases(
        ctx,
        ["last quarter used", "last quarter"],
        anchor_row,
        row_min=anchor_row - 3,
        row_max=anchor_row + 3,
    ) or (anchor_col - 3)
    quarterly_sales_col = find_column_by_phrases(
        ctx,
        ["quarterly sales"],
        anchor_row,
        row_min=anchor_row - 4,
        row_max=anchor_row + 4,
    ) or (anchor_col - 6)
    growth_rate_col = find_column_by_phrases(
        ctx,
        ["growth rate", "growth"],
        anchor_row,
        row_min=anchor_row - 4,
        row_max=anchor_row + 4,
    ) or (anchor_col - 5)
    sales_captured_col = find_column_by_phrases(
        ctx,
        ["sales captured in db", "captured in db", "captured"],
        anchor_row,
        row_min=anchor_row - 4,
        row_max=anchor_row + 4,
    ) or (anchor_col - 7)
    penetration_col = find_column_by_phrases(
        ctx,
        ["penetration pct", "penetration", "penetration %"],
        anchor_row,
    )

    summary_rows = collect_summary_rows(
        ctx=ctx,
        start_row=anchor_row + 1,
        preferred_col=num_quarters_col,
        fallback_col=anchor_col,
        limit=N_QUARTERS,
    )

    avg_pen_values: list[float | None] = [None] * N_QUARTERS
    if penetration_col is not None:
        numeric_range = find_contiguous_numeric_range_upwards(ctx, penetration_col, anchor_row - 1)
        if numeric_range:
            range_start, range_end = numeric_range
            available = range_end - range_start + 1

            helper_col = ctx.last_col + 2
            helper_start_row = ctx.last_row + 2
            formula_cells: list[xw.Range] = []
            for idx in range(N_QUARTERS):
                n = idx + 1
                if n > available:
                    continue
                cell = sheet.cells(helper_start_row + idx, helper_col)
                formula = build_last_n_avg_formula_r1c1(range_start, range_end, penetration_col, n)
                set_formula2_r1c1(cell, formula)
                formula_cells.append(cell)

            if formula_cells:
                wb.app.calculate()
                for idx in range(N_QUARTERS):
                    n = idx + 1
                    if n > available:
                        continue
                    avg_pen_values[idx] = to_number(sheet.cells(helper_start_row + idx, helper_col).value)

    rows: list[dict[str, Any]] = []
    for idx, row_num in enumerate(summary_rows[:N_QUARTERS]):
        n_default = idx + 1
        n_used = to_number(ctx_value(ctx, row_num, num_quarters_col))
        n_used_int = int(n_used) if n_used is not None else n_default

        forecast_value = to_number(ctx_value(ctx, row_num, forecast_col))
        actual_value = to_number(ctx_value(ctx, row_num, actual_col))
        forecast_max = to_number(ctx_value(ctx, row_num, anchor_col))
        forecast_min = to_number(ctx_value(ctx, row_num, min_col))
        range_width = (forecast_max - forecast_min) if (forecast_max is not None and forecast_min is not None) else None

        avg_penetration_pct = avg_pen_values[idx] if idx < len(avg_pen_values) else None
        last_quarter_used = safe_strip(ctx_value(ctx, row_num, last_quarter_col))
        quarterly_sales = to_number(ctx_value(ctx, row_num, quarterly_sales_col))
        reported_sales = to_number(ctx_value(ctx, row_num, actual_col))
        growth_rate_pct = to_number(ctx_value(ctx, row_num, growth_rate_col))
        sales_captured = to_number(ctx_value(ctx, row_num, sales_captured_col))

        if all(
            value is None
            for value in (
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
            )
        ):
            continue

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": n_used_int,
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
                "sales_captured_in_db_pct": sales_captured,
                "source_file": source_file,
            }
        )

    return rows


def build_intercept_formula_r1c1(y_start: int, y_end: int, y_col: int, x_start: int, x_end: int, x_col: int) -> str:
    y_rng = f"R{y_start}C{y_col}:R{y_end}C{y_col}"
    x_rng = f"R{x_start}C{x_col}:R{x_end}C{x_col}"
    return f"=INTERCEPT({y_rng},{x_rng})"


def build_slope_formula_r1c1(y_start: int, y_end: int, y_col: int, x_start: int, x_end: int, x_col: int) -> str:
    y_rng = f"R{y_start}C{y_col}:R{y_end}C{y_col}"
    x_rng = f"R{x_start}C{x_col}:R{x_end}C{x_col}"
    return f"=SLOPE({y_rng},{x_rng})"


def extract_regression_rows(wb: xw.Book, meta: ModelMeta, source_file: str) -> list[dict[str, Any]]:
    if "Regression Model" not in [s.name for s in wb.sheets]:
        return []

    sheet = wb.sheets["Regression Model"]
    ctx = read_sheet_context(sheet)
    anchor_row, anchor_col = find_anchor(ctx, "max")

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    min_col = find_column_by_phrases(ctx, ["min"], anchor_row, row_min=anchor_row - 2, row_max=anchor_row + 2) or (anchor_col + 1)
    forecast_col = find_column_by_phrases(
        ctx,
        ["tot fcst w/o sa", "total forecast w/o sa", "forecast total w/o sa"],
        anchor_row,
        row_min=anchor_row - 3,
        row_max=anchor_row + 3,
    ) or (anchor_col - 1)
    num_quarters_col = find_column_by_phrases(
        ctx,
        ["num quarters used", "quarters used", "n quarters", "num quarters"],
        anchor_row,
        row_min=anchor_row - 3,
        row_max=anchor_row + 3,
    ) or (anchor_col - 3)
    actual_col = find_column_by_phrases(
        ctx,
        ["actual", "reported sales", "actual value"],
        anchor_row,
        row_min=anchor_row - 3,
        row_max=anchor_row + 3,
    )

    history_rows: list[int] = []
    for row in range(ctx.top_row, anchor_row):
        x_val = to_number(ctx_value(ctx, row, x_col))
        y_val = to_number(ctx_value(ctx, row, y_col))
        if x_val is not None and y_val is not None:
            history_rows.append(row)

    if len(history_rows) < 2:
        return []

    max_loop = min(N_QUARTERS, len(history_rows))
    helper_col = ctx.last_col + 2
    helper_row = ctx.last_row + 2

    intercept_vals: list[float | None] = [None] * max_loop
    slope_vals: list[float | None] = [None] * max_loop

    for idx in range(max_loop):
        n = idx + 1
        sub_rows = history_rows[-n:]
        y_start, y_end = sub_rows[0], sub_rows[-1]
        x_start, x_end = sub_rows[0], sub_rows[-1]

        int_cell = sheet.cells(helper_row + idx, helper_col)
        slope_cell = sheet.cells(helper_row + idx, helper_col + 1)
        int_formula = build_intercept_formula_r1c1(y_start, y_end, y_col, x_start, x_end, x_col)
        slope_formula = build_slope_formula_r1c1(y_start, y_end, y_col, x_start, x_end, x_col)
        set_formula2_r1c1(int_cell, int_formula)
        set_formula2_r1c1(slope_cell, slope_formula)

    wb.app.calculate()

    for idx in range(max_loop):
        intercept_vals[idx] = to_number(sheet.cells(helper_row + idx, helper_col).value)
        slope_vals[idx] = to_number(sheet.cells(helper_row + idx, helper_col + 1).value)

    summary_rows = collect_summary_rows(
        ctx=ctx,
        start_row=anchor_row + 1,
        preferred_col=num_quarters_col,
        fallback_col=anchor_col,
        limit=max_loop,
    )

    latest_x = to_number(ctx_value(ctx, history_rows[-1], x_col))

    rows: list[dict[str, Any]] = []
    prev_signature: tuple[Any, ...] | None = None
    for idx in range(max_loop):
        table_row = summary_rows[idx] if idx < len(summary_rows) else None
        n_default = idx + 1
        n_used = to_number(ctx_value(ctx, table_row, num_quarters_col)) if table_row is not None else None
        n_used_int = int(n_used) if n_used is not None else n_default

        intercept = intercept_vals[idx]
        slope = slope_vals[idx]
        forecast_value = to_number(ctx_value(ctx, table_row, forecast_col)) if table_row is not None else None
        if forecast_value is None and intercept is not None and slope is not None and latest_x is not None:
            forecast_value = intercept + (slope * latest_x)

        actual_value = to_number(ctx_value(ctx, table_row, actual_col)) if (table_row is not None and actual_col is not None) else None
        forecast_max = to_number(ctx_value(ctx, table_row, anchor_col)) if table_row is not None else None
        forecast_min = to_number(ctx_value(ctx, table_row, min_col)) if table_row is not None else None
        range_width = (forecast_max - forecast_min) if (forecast_max is not None and forecast_min is not None) else None

        signature = (
            n_used_int,
            round(forecast_value, 8) if forecast_value is not None else None,
            round(forecast_max, 8) if forecast_max is not None else None,
            round(forecast_min, 8) if forecast_min is not None else None,
            round(intercept, 8) if intercept is not None else None,
            round(slope, 8) if slope is not None else None,
        )

        # Prevent duplicate trailing row from repeated formula output.
        if prev_signature is not None and signature == prev_signature:
            continue
        prev_signature = signature

        if all(v is None for v in (forecast_value, forecast_max, forecast_min, intercept, slope)):
            continue

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n_used_int,
                "num_quarters_used": n_used_int,
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


def write_sheet(
    wb: Workbook,
    sheet_name: str,
    columns: list[str],
    rows: list[dict[str, Any]],
) -> None:
    ws = wb.create_sheet(title=sheet_name)
    ws.append(columns)

    for col_idx, col_name in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = Font(bold=True)

    for row in rows:
        ws.append([row.get(column) for column in columns])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for data_row in range(2, ws.max_row + 1):
            value = ws.cell(row=data_row, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 42)


def write_output_workbook(
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
    path: Path,
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)
    write_sheet(wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    wb.save(path)


def run() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"Input directory does not exist: {input_dir}")

    output_path = get_output_path(input_dir, output_dir)
    files = sorted(path for path in input_dir.iterdir() if path.is_file())

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        try:
            app.calculation = "manual"
        except Exception:
            pass

        empirical_rows: list[dict[str, Any]] = []
        regression_rows: list[dict[str, Any]] = []
        files_processed = 0

        for file_path in files:
            if file_path.name.startswith("~"):
                print(f"Skipping {file_path.name}: temporary Excel file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipping {file_path.name}: not an .xlsx file")
                continue

            print(f"Processing {file_path.name}")
            wb: xw.Book | None = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                meta = model_meta_from_filename(file_path)

                empirical_rows.extend(extract_empirical_rows(wb, meta, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, meta, file_path.name))
                files_processed += 1
            except Exception as exc:
                print(f"Skipping {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    close_workbook_no_save(wb)

        write_output_workbook(empirical_rows, regression_rows, output_path)

        print(f"Output path: {output_path}")
        print(f"Number of files processed: {files_processed}")
        print(f"Number of empirical rows: {len(empirical_rows)}")
        print(f"Number of regression rows: {len(regression_rows)}")
    finally:
        app.quit()


if __name__ == "__main__":
    run()

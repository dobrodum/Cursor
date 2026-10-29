#!/usr/bin/env python3
from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
import re
from typing import Any, Iterable

import openpyxl
from openpyxl.styles import Font
import xlwings as xw

# ---------------------------------------------------------------------------
# Configure paths here
# ---------------------------------------------------------------------------
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")

# ---------------------------------------------------------------------------
# Output schema
# ---------------------------------------------------------------------------
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

N_QUARTERS = 10


@dataclass
class FileMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetContext:
    sheet: Any
    first_row: int
    first_col: int
    last_row: int
    last_col: int
    values_2d: list[list[Any]]
    anchor_row: int
    anchor_col: int
    header_map: dict[str, int]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    return re.sub(r"[^a-z0-9]+", "", text)


def to_2d(values: Any, rows: int, cols: int) -> list[list[Any]]:
    if rows <= 0 or cols <= 0:
        return []
    if rows == 1 and cols == 1:
        return [[values]]
    if rows == 1:
        return [list(values)]
    if cols == 1:
        return [[v] for v in values]
    return [list(r) for r in values]


def month_str_to_num(month_abbrev: str) -> int:
    mapping = {
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
    return mapping[month_abbrev.lower()]


def parse_filename_metadata(file_name: str) -> FileMeta:
    stem = Path(file_name).stem
    pattern = re.compile(
        r"^\s*.+?\s-\s(?P<ticker>[A-Za-z0-9]+)\s-\s(?P<period>(Early|Mid|Late)[A-Za-z]{3}\d{4})",
        re.IGNORECASE,
    )
    match = pattern.match(stem)
    if not match:
        ticker = stem.split(" - ")[1].strip().upper() if " - " in stem else "UNKNOWN"
        model_period = "Unknown_0000"
        return FileMeta(
            model=f"{ticker}_{model_period}",
            ticker=ticker,
            model_period=model_period,
            model_date="",
        )

    ticker = match.group("ticker").upper()
    period = match.group("period")
    timing = re.match(r"(Early|Mid|Late)", period, re.IGNORECASE).group(1).title()
    month_abbrev = period[len(timing) : len(timing) + 3].title()
    year = period[-4:]
    model_period = f"{timing}{month_abbrev}_{year}"

    day_map = {"Early": 5, "Mid": 15, "Late": 25}
    model_date = date(int(year), month_str_to_num(month_abbrev), day_map[timing]).isoformat()
    model = f"{ticker}_{model_period}"
    return FileMeta(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def next_output_path(input_folder: Path, out_folder: Path) -> Path:
    out_folder.mkdir(parents=True, exist_ok=True)
    base_stem = f"{input_folder.resolve().name}_PARAM"
    candidate = out_folder / f"{base_stem}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = out_folder / f"{base_stem}.{suffix}.xlsx"
        suffix += 1
    return candidate


def safe_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.replace(",", "").replace("%", "").strip()
        if not cleaned:
            return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def difference(a: Any, b: Any) -> float | None:
    a_float = safe_float(a)
    b_float = safe_float(b)
    if a_float is None or b_float is None:
        return None
    return a_float - b_float


def get_sheet_if_exists(wb: Any, sheet_name: str) -> Any | None:
    target = sheet_name.strip().lower()
    for sh in wb.sheets:
        if sh.name.strip().lower() == target:
            return sh
    return None


def find_anchor(values_2d: list[list[Any]], first_row: int, first_col: int) -> tuple[int, int] | None:
    candidates: list[tuple[int, int, int, int]] = []
    for r_idx, row in enumerate(values_2d):
        for c_idx, value in enumerate(row):
            if normalize_text(value) != "max":
                continue
            score = 0
            if c_idx + 1 < len(row) and normalize_text(row[c_idx + 1]) == "min":
                score += 2
            if r_idx + 1 < len(values_2d):
                below = values_2d[r_idx + 1][c_idx]
                if isinstance(below, (int, float)):
                    score += 1
            candidates.append((score, r_idx, c_idx, -c_idx))

    if not candidates:
        return None

    candidates.sort(reverse=True)
    _, r_idx, c_idx, _ = candidates[0]
    return first_row + r_idx, first_col + c_idx


def build_header_map(sheet: Any, header_row: int, first_col: int, last_col: int) -> dict[str, int]:
    row_values = sheet.range((header_row, first_col), (header_row, last_col)).value
    if first_col == last_col:
        row_values = [row_values]
    header_map: dict[str, int] = {}
    for offset, value in enumerate(row_values):
        key = normalize_text(value)
        if key:
            header_map[key] = first_col + offset
    return header_map


def resolve_col(header_map: dict[str, int], aliases: Iterable[str], default_col: int | None) -> int | None:
    normalized_aliases = [normalize_text(a) for a in aliases]

    for alias in normalized_aliases:
        if alias in header_map:
            return header_map[alias]

    for header_key, col in header_map.items():
        for alias in normalized_aliases:
            if alias and (alias in header_key or header_key in alias):
                return col

    return default_col


def build_sheet_context(sheet: Any) -> SheetContext | None:
    used = sheet.used_range
    rows_count = used.rows.count
    cols_count = used.columns.count
    if rows_count == 0 or cols_count == 0:
        return None

    first_row = used.row
    first_col = used.column
    last_row = first_row + rows_count - 1
    last_col = first_col + cols_count - 1

    values_2d = to_2d(used.value, rows_count, cols_count)
    anchor = find_anchor(values_2d, first_row, first_col)
    if anchor is None:
        return None
    anchor_row, anchor_col = anchor

    header_map = build_header_map(sheet, anchor_row, first_col, last_col)
    return SheetContext(
        sheet=sheet,
        first_row=first_row,
        first_col=first_col,
        last_row=last_row,
        last_col=last_col,
        values_2d=values_2d,
        anchor_row=anchor_row,
        anchor_col=anchor_col,
        header_map=header_map,
    )


def write_formula2_r1c1(cell: Any, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        cell.api.Formula2R1C1 = formula_r1c1


def close_workbook_without_save(wb: Any) -> None:
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
        wb.saved = True
    except Exception:
        pass
    wb.close()


def cell_value(sheet: Any, row: int, col: int | None) -> Any:
    if col is None or col <= 0:
        return None
    return sheet.cells(row, col).value


def compute_empirical_rows(ctx: SheetContext, meta: FileMeta, source_file: str) -> list[dict[str, Any]]:
    sh = ctx.sheet
    anchor_row = ctx.anchor_row
    anchor_col = ctx.anchor_col
    headers = ctx.header_map

    num_q_col = resolve_col(headers, ["num_quarters_used", "num quarters used", "quarters used"], anchor_col - 6)
    last_q_col = resolve_col(headers, ["last_quarter_used", "last quarter used", "last quarter"], anchor_col - 5)
    avg_pen_col = resolve_col(
        headers,
        ["avg_penetration_pct", "avg penetration pct", "average penetration", "avg penetration"],
        anchor_col - 4,
    )
    forecast_col = resolve_col(
        headers,
        ["estimated total sold", "forecast value", "tot fcst", "forecast"],
        anchor_col - 1,
    )
    actual_col = resolve_col(headers, ["actual value", "reported sales", "actual sales"], anchor_col - 2)
    max_col = anchor_col
    min_col = resolve_col(headers, ["forecast min", "min"], anchor_col + 1)
    quarterly_sales_col = resolve_col(headers, ["quarterly sales", "qtr sales"], anchor_col - 3)
    reported_sales_col = resolve_col(headers, ["reported sales", "actual sales"], actual_col)
    growth_col = resolve_col(headers, ["growth rate pct", "growth rate", "growth"], anchor_col + 2)
    capture_col = resolve_col(
        headers,
        ["sales captured in db pct", "sales captured in db", "captured in db"],
        anchor_col + 3,
    )
    penetration_source_col = resolve_col(
        headers,
        ["penetration pct", "penetration", "actual penetration"],
        avg_pen_col,
    )

    start_row = anchor_row + 1
    end_row = min(start_row + N_QUARTERS - 1, ctx.last_row)
    candidate_rows = list(range(start_row, end_row + 1))
    if not candidate_rows:
        return []

    helper_col = ctx.last_col + 5
    hist_end = anchor_row - 1

    for idx, row in enumerate(candidate_rows, start=1):
        num_q_raw = cell_value(sh, row, num_q_col)
        num_q = int(num_q_raw) if isinstance(num_q_raw, (int, float)) and num_q_raw > 0 else idx
        hist_start = max(ctx.first_row, hist_end - num_q + 1)
        formula = f"=AVERAGE(R{hist_start}C{penetration_source_col}:R{hist_end}C{penetration_source_col})"
        write_formula2_r1c1(sh.cells(row, helper_col), formula)

    sh.book.app.calculate()

    rows: list[dict[str, Any]] = []
    for idx, row in enumerate(candidate_rows, start=1):
        num_q_raw = cell_value(sh, row, num_q_col)
        num_q = int(num_q_raw) if isinstance(num_q_raw, (int, float)) and num_q_raw > 0 else idx

        avg_penetration = cell_value(sh, row, helper_col)
        if avg_penetration in (None, ""):
            avg_penetration = cell_value(sh, row, avg_pen_col)

        forecast_value = cell_value(sh, row, forecast_col)
        actual_value = cell_value(sh, row, actual_col)
        forecast_max = cell_value(sh, row, max_col)
        forecast_min = cell_value(sh, row, min_col)

        if all(v in (None, "") for v in [avg_penetration, forecast_value, forecast_max, forecast_min]):
            continue

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_q,
                "last_quarter_used": cell_value(sh, row, last_q_col),
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": difference(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": cell_value(sh, row, quarterly_sales_col),
                "reported_sales": cell_value(sh, row, reported_sales_col),
                "growth_rate_pct": cell_value(sh, row, growth_col),
                "sales_captured_in_db_pct": cell_value(sh, row, capture_col),
                "source_file": source_file,
            }
        )

    return rows


def compute_regression_rows(ctx: SheetContext, meta: FileMeta, source_file: str) -> list[dict[str, Any]]:
    sh = ctx.sheet
    anchor_row = ctx.anchor_row
    anchor_col = ctx.anchor_col
    headers = ctx.header_map

    num_q_col = resolve_col(headers, ["num_quarters_used", "num quarters used", "quarters used"], anchor_col - 6)
    forecast_col = resolve_col(
        headers,
        ["tot fcst w/o sa", "forecast total without sa", "tot fcst", "forecast value", "forecast"],
        anchor_col - 1,
    )
    actual_col = resolve_col(headers, ["actual value", "reported sales", "actual sales"], None)
    max_col = anchor_col
    min_col = resolve_col(headers, ["forecast min", "min"], anchor_col + 1)

    # Required by the specification.
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    start_row = anchor_row + 1
    end_row = min(start_row + N_QUARTERS - 1, ctx.last_row)
    candidate_rows = list(range(start_row, end_row + 1))
    if not candidate_rows:
        return []

    helper_col = ctx.last_col + 10
    hist_end = anchor_row - 1

    for idx, row in enumerate(candidate_rows, start=1):
        num_q_raw = cell_value(sh, row, num_q_col)
        num_q = int(num_q_raw) if isinstance(num_q_raw, (int, float)) and num_q_raw > 0 else idx
        hist_start = max(ctx.first_row, hist_end - num_q + 1)

        intercept_formula = (
            f"=INTERCEPT(R{hist_start}C{y_col}:R{hist_end}C{y_col},R{hist_start}C{x_col}:R{hist_end}C{x_col})"
        )
        slope_formula = f"=SLOPE(R{hist_start}C{y_col}:R{hist_end}C{y_col},R{hist_start}C{x_col}:R{hist_end}C{x_col})"
        write_formula2_r1c1(sh.cells(row, helper_col), intercept_formula)
        write_formula2_r1c1(sh.cells(row, helper_col + 1), slope_formula)

    sh.book.app.calculate()

    rows: list[dict[str, Any]] = []
    prev_signature: tuple[Any, ...] | None = None

    for idx, row in enumerate(candidate_rows, start=1):
        num_q_raw = cell_value(sh, row, num_q_col)
        num_q = int(num_q_raw) if isinstance(num_q_raw, (int, float)) and num_q_raw > 0 else idx

        intercept = cell_value(sh, row, helper_col)
        slope = cell_value(sh, row, helper_col + 1)
        forecast_value = cell_value(sh, row, forecast_col)
        forecast_max = cell_value(sh, row, max_col)
        forecast_min = cell_value(sh, row, min_col)
        actual_value = cell_value(sh, row, actual_col)

        if all(v in (None, "") for v in [intercept, slope, forecast_value, forecast_max, forecast_min]):
            continue

        signature = (num_q, intercept, slope, forecast_value, forecast_max, forecast_min)
        if prev_signature is not None and signature == prev_signature:
            continue
        prev_signature = signature

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_q,
                "num_quarters_used": num_q,
                "forecast_value": forecast_value,
                "actual_value": actual_value if actual_value is not None else "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": difference(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def set_column_widths(ws: openpyxl.worksheet.worksheet.Worksheet, max_width: int = 42) -> None:
    for col_cells in ws.columns:
        first_cell = col_cells[0]
        col_letter = first_cell.column_letter
        longest = 0
        for cell in col_cells:
            if cell.value is None:
                continue
            longest = max(longest, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max(12, longest + 2), max_width)


def write_table_sheet(
    wb: openpyxl.Workbook,
    sheet_name: str,
    columns: list[str],
    rows: list[dict[str, Any]],
) -> None:
    ws = wb.create_sheet(sheet_name)
    ws.append(columns)
    for item in rows:
        ws.append([item.get(col) for col in columns])

    for header_cell in ws[1]:
        header_cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    if ws.max_row >= 1 and ws.max_column >= 1:
        ws.auto_filter.ref = ws.dimensions
    set_column_widths(ws)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = openpyxl.Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)
    write_table_sheet(wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_table_sheet(wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    wb.save(output_path)


def should_skip_file(path: Path) -> tuple[bool, str]:
    if not path.is_file():
        return True, "not a file"
    if path.name.startswith("~"):
        return True, "temporary file"
    if path.suffix.lower() != ".xlsx":
        return True, "not .xlsx"
    return False, ""


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"Input directory does not exist: {input_dir}")

    output_path = next_output_path(input_dir, output_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        app.calculation = "manual"
    except Exception:
        pass

    try:
        for file_path in sorted(input_dir.iterdir()):
            skip, reason = should_skip_file(file_path)
            if skip:
                print(f"SKIPPED: {file_path.name} ({reason})")
                continue

            print(f"PROCESSING: {file_path.name}")
            meta = parse_filename_metadata(file_path.name)

            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_sheet = get_sheet_if_exists(wb, "Empirical Model")
                regression_sheet = get_sheet_if_exists(wb, "Regression Model")

                if empirical_sheet is None and regression_sheet is None:
                    print(f"SKIPPED: {file_path.name} (missing both required model sheets)")
                    continue

                if empirical_sheet is not None:
                    ctx = build_sheet_context(empirical_sheet)
                    if ctx is None:
                        print(f"SKIPPED: {file_path.name} (could not locate 'max' anchor in Empirical Model)")
                    else:
                        empirical_rows.extend(compute_empirical_rows(ctx, meta, file_path.name))

                if regression_sheet is not None:
                    ctx = build_sheet_context(regression_sheet)
                    if ctx is None:
                        print(f"SKIPPED: {file_path.name} (could not locate 'max' anchor in Regression Model)")
                    else:
                        regression_rows.extend(compute_regression_rows(ctx, meta, file_path.name))

                processed_files += 1
            except Exception as exc:
                print(f"SKIPPED: {file_path.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    close_workbook_without_save(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"OUTPUT: {output_path}")
    print(f"FILES_PROCESSED: {processed_files}")
    print(f"EMPIRICAL_ROWS: {len(empirical_rows)}")
    print(f"REGRESSION_ROWS: {len(regression_rows)}")


if __name__ == "__main__":
    main()

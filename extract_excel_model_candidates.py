#!/usr/bin/env python3
"""Extract empirical and regression model candidates from Excel workbooks."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
import re
from typing import Any

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import xlwings as xw


# ----------------------------
# Configure folders here
# ----------------------------
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

EMPIRICAL_DEFAULT_OFFSETS = {
    "num_quarters_used": -8,
    "last_quarter_used": -7,
    "avg_penetration_pct": -6,
    "quarterly_sales": -5,
    "reported_sales": -4,
    "growth_rate_pct": -3,
    "sales_captured_in_db_pct": -2,
    "forecast_value": -1,
    "forecast_max": 0,
    "forecast_min": 1,
}

REGRESSION_DEFAULT_OFFSETS = {
    "num_quarters_used": -4,
    "forecast_value": -1,  # TOT FCST w/o SA
    "forecast_max": 0,
    "forecast_min": 1,
}

MODEL_DAY = {"early": 5, "mid": 15, "late": 25}
MONTH_NUM = {
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

EMPIRICAL_HEADER_ALIASES = {
    "num_quarters_used": ["num quarters", "quarters used", "n quarters", "num_q"],
    "last_quarter_used": ["last quarter"],
    "avg_penetration_pct": ["avg penetration", "average penetration"],
    "quarterly_sales": ["quarterly sales"],
    "reported_sales": ["reported sales"],
    "growth_rate_pct": ["growth rate"],
    "sales_captured_in_db_pct": ["sales captured", "captured in db"],
    "forecast_value": ["estimated total sold", "forecast value", "estimate total"],
    "forecast_max": ["max"],
    "forecast_min": ["min"],
}

REGRESSION_HEADER_ALIASES = {
    "num_quarters_used": ["num quarters", "quarters used", "n quarters"],
    "forecast_value": ["tot fcst w/o sa", "total fcst w/o sa", "forecast value"],
    "forecast_max": ["max"],
    "forecast_min": ["min"],
    "actual_value": ["actual value", "reported sales"],
}


@dataclass
class ModelLabel:
    ticker: str
    model_period: str
    model_date: str
    model: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(text.split())


def to_number(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    text = text.replace(",", "")
    was_percent = text.endswith("%")
    if was_percent:
        text = text[:-1]
    try:
        parsed = float(text)
        if was_percent:
            return parsed / 100.0
        return parsed
    except ValueError:
        return None


def to_int(value: Any) -> int | None:
    num = to_number(value)
    if num is None:
        return None
    return int(round(num))


def calc_range_width(max_value: Any, min_value: Any) -> float | None:
    max_num = to_number(max_value)
    min_num = to_number(min_value)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


def parse_model_label(file_path: Path) -> ModelLabel:
    stem = file_path.stem
    parts = [p.strip() for p in stem.split(" - ") if p.strip()]

    ticker = "UNKNOWN"
    if len(parts) >= 2:
        parsed = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper()
        if parsed:
            ticker = parsed

    period_token = ""
    if len(parts) >= 3:
        period_token = re.split(r"[_\s-]", parts[2], maxsplit=1)[0]
    period_match = re.match(
        r"^(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(20\d{2})$",
        period_token,
        flags=re.IGNORECASE,
    )

    if period_match:
        cadence = period_match.group(1).lower()
        cadence_title = cadence.capitalize()
        month_abbrev = period_match.group(2).lower()
        month_title = month_abbrev.capitalize()
        year = int(period_match.group(3))
        model_period = f"{cadence_title}{month_title}_{year}"
        day = MODEL_DAY[cadence]
        month = MONTH_NUM[month_abbrev]
        model_date = date(year, month, day).isoformat()
    else:
        model_period = "Unknown_0000"
        model_date = ""

    model = f"{ticker}_{model_period}"
    return ModelLabel(
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
        model=model,
    )


def output_path_for_run(source_dir: Path, target_dir: Path) -> Path:
    target_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{source_dir.name}_PARAM"
    candidate = target_dir / f"{base_name}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = target_dir / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def list_input_files(source_dir: Path) -> tuple[list[Path], int]:
    candidates: list[Path] = []
    skipped = 0
    for path in sorted(source_dir.iterdir()):
        if not path.is_file():
            print(f"Skipping {path.name}: not a file")
            skipped += 1
            continue
        if path.name.startswith("~"):
            print(f"Skipping {path.name}: temporary file")
            skipped += 1
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"Skipping {path.name}: not .xlsx")
            skipped += 1
            continue
        candidates.append(path)
    return candidates, skipped


def safe_close_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.close(save_changes=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.api.Close(SaveChanges=False)  # type: ignore[attr-defined]
        return
    except Exception:
        pass

    try:
        wb.close()
    except Exception:
        pass


def get_sheet_case_insensitive(wb: xw.Book, sheet_name: str) -> xw.Sheet | None:
    target = sheet_name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def read_used_grid(ws: xw.Sheet) -> tuple[int, int, list[list[Any]]]:
    used = ws.used_range
    start_row = used.row
    start_col = used.column
    raw = used.value
    if raw is None:
        return start_row, start_col, []
    if isinstance(raw, list):
        if raw and isinstance(raw[0], list):
            return start_row, start_col, raw
        return start_row, start_col, [raw]
    return start_row, start_col, [[raw]]


def find_max_anchor(start_row: int, start_col: int, grid: list[list[Any]]) -> tuple[int, int] | None:
    for row_offset, row_values in enumerate(grid):
        for col_offset, value in enumerate(row_values):
            if normalize_text(value) == "max":
                return start_row + row_offset, start_col + col_offset
    return None


def detect_header_columns(
    start_row: int,
    start_col: int,
    grid: list[list[Any]],
    anchor_row: int,
    anchor_col: int,
    aliases: dict[str, list[str]],
    radius: int = 40,
) -> dict[str, int]:
    row_index = anchor_row - start_row
    if row_index < 0 or row_index >= len(grid):
        return {}

    header_row = grid[row_index]
    mapping: dict[str, int] = {}
    for col_offset, value in enumerate(header_row):
        col = start_col + col_offset
        if abs(col - anchor_col) > radius:
            continue
        normalized = normalize_text(value)
        if not normalized:
            continue
        for field, patterns in aliases.items():
            if field in mapping:
                continue
            if any(pattern in normalized for pattern in patterns):
                mapping[field] = col
    return mapping


def resolve_cols(
    anchor_col: int,
    detected: dict[str, int],
    defaults: dict[str, int],
) -> dict[str, int]:
    resolved: dict[str, int] = {}
    for key, offset in defaults.items():
        resolved[key] = detected.get(key, anchor_col + offset)
    return resolved


def set_formula2(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        cell.formula = formula_r1c1


def is_effectively_empty(values: list[Any]) -> bool:
    return all(value is None or value == "" for value in values)


def build_empirical_rows(wb: xw.Book, ws: xw.Sheet, labels: ModelLabel, source_file: str) -> list[dict[str, Any]]:
    start_row, start_col, grid = read_used_grid(ws)
    anchor = find_max_anchor(start_row, start_col, grid)
    if anchor is None:
        print(f"Skipping empirical extraction in {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    detected = detect_header_columns(
        start_row=start_row,
        start_col=start_col,
        grid=grid,
        anchor_row=anchor_row,
        anchor_col=anchor_col,
        aliases=EMPIRICAL_HEADER_ALIASES,
    )
    cols = resolve_cols(anchor_col=anchor_col, detected=detected, defaults=EMPIRICAL_DEFAULT_OFFSETS)
    helper_col = anchor_col + 6

    quarter_end_col = max(1, anchor_col - 11)

    # Write average-penetration formulas once, then calculate once.
    formula_rows: list[int] = []
    for n_quarters in range(1, 11):
        row = anchor_row + n_quarters
        quarter_start_col = max(1, quarter_end_col - (n_quarters - 1))
        formula = f"=AVERAGE(R{row}C{quarter_start_col}:R{row}C{quarter_end_col})"
        set_formula2(ws.range((row, helper_col)), formula)
        formula_rows.append(row)

    if formula_rows:
        wb.app.calculate()

    rows: list[dict[str, Any]] = []
    for n_quarters in range(1, 11):
        row = anchor_row + n_quarters
        row_values_for_blank_check = [
            ws.range((row, cols["forecast_value"])).value,
            ws.range((row, cols["forecast_max"])).value,
            ws.range((row, cols["forecast_min"])).value,
            ws.range((row, cols["reported_sales"])).value,
            ws.range((row, cols["avg_penetration_pct"])).value,
        ]
        if is_effectively_empty(row_values_for_blank_check):
            continue

        avg_pen_cell = ws.range((row, cols["avg_penetration_pct"])).value
        avg_pen_formula = ws.range((row, helper_col)).value
        avg_penetration = avg_pen_cell if avg_pen_cell not in (None, "") else avg_pen_formula

        num_quarters_used = ws.range((row, cols["num_quarters_used"])).value
        if num_quarters_used in (None, ""):
            num_quarters_used = n_quarters

        forecast_max = ws.range((row, cols["forecast_max"])).value
        forecast_min = ws.range((row, cols["forecast_min"])).value
        reported_sales = ws.range((row, cols["reported_sales"])).value

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": ws.range((row, cols["last_quarter_used"])).value,
                "forecast_value": ws.range((row, cols["forecast_value"])).value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": calc_range_width(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": ws.range((row, cols["quarterly_sales"])).value,
                "reported_sales": reported_sales,
                "growth_rate_pct": ws.range((row, cols["growth_rate_pct"])).value,
                "sales_captured_in_db_pct": ws.range((row, cols["sales_captured_in_db_pct"])).value,
                "source_file": source_file,
            }
        )
    return rows


def build_regression_rows(wb: xw.Book, ws: xw.Sheet, labels: ModelLabel, source_file: str) -> list[dict[str, Any]]:
    start_row, start_col, grid = read_used_grid(ws)
    anchor = find_max_anchor(start_row, start_col, grid)
    if anchor is None:
        print(f"Skipping regression extraction in {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    detected = detect_header_columns(
        start_row=start_row,
        start_col=start_col,
        grid=grid,
        anchor_row=anchor_row,
        anchor_col=anchor_col,
        aliases=REGRESSION_HEADER_ALIASES,
    )
    cols = resolve_cols(anchor_col=anchor_col, detected=detected, defaults=REGRESSION_DEFAULT_OFFSETS)
    # Keep regression actuals blank unless an actual-value header is explicitly found.
    actual_col = detected.get("actual_value")

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    intercept_col = anchor_col + 4
    slope_col = anchor_col + 5

    row_bounds = range(anchor_row + 1, anchor_row + 51)
    rows_to_calculate: list[tuple[int, int]] = []
    blank_streak = 0
    for row in row_bounds:
        num_quarters_used = ws.range((row, cols["num_quarters_used"])).value
        forecast_value = ws.range((row, cols["forecast_value"])).value
        forecast_max = ws.range((row, cols["forecast_max"])).value
        forecast_min = ws.range((row, cols["forecast_min"])).value
        if is_effectively_empty([num_quarters_used, forecast_value, forecast_max, forecast_min]):
            blank_streak += 1
            if blank_streak >= 2:
                break
            continue
        blank_streak = 0

        n_quarters = to_int(num_quarters_used) or 1
        start_fit_row = max(1, row - n_quarters + 1)
        intercept_formula = (
            f"=INTERCEPT(R{start_fit_row}C{y_col}:R{row}C{y_col},"
            f"R{start_fit_row}C{x_col}:R{row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_fit_row}C{y_col}:R{row}C{y_col},"
            f"R{start_fit_row}C{x_col}:R{row}C{x_col})"
        )
        set_formula2(ws.range((row, intercept_col)), intercept_formula)
        set_formula2(ws.range((row, slope_col)), slope_formula)
        rows_to_calculate.append((row, n_quarters))

    if rows_to_calculate:
        wb.app.calculate()

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None
    for row, n_quarters in rows_to_calculate:
        forecast_value = ws.range((row, cols["forecast_value"])).value
        forecast_max = ws.range((row, cols["forecast_max"])).value
        forecast_min = ws.range((row, cols["forecast_min"])).value
        intercept = ws.range((row, intercept_col)).value
        slope = ws.range((row, slope_col)).value
        actual_value = ws.range((row, actual_col)).value if actual_col is not None else None

        signature = (
            to_int(ws.range((row, cols["num_quarters_used"])).value) or n_quarters,
            round(to_number(forecast_value) or 0.0, 10),
            round(to_number(forecast_max) or 0.0, 10),
            round(to_number(forecast_min) or 0.0, 10),
            round(to_number(intercept) or 0.0, 10),
            round(to_number(slope) or 0.0, 10),
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        num_quarters_used = ws.range((row, cols["num_quarters_used"])).value
        if num_quarters_used in (None, ""):
            num_quarters_used = n_quarters

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
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
                "source_file": source_file,
            }
        )
    return rows


def style_output_sheet(ws: openpyxl.worksheet.worksheet.Worksheet) -> None:
    if ws.max_row == 0:
        return
    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx in range(1, ws.max_column + 1):
        max_len = 0
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=col_idx, max_col=col_idx):
            value = row[0].value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        width = max(12, min(max_len + 2, 60))
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def write_output_workbook(
    path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = openpyxl.Workbook()
    ws_emp = wb.active
    ws_emp.title = "empirical_candidates"
    ws_reg = wb.create_sheet("regression_candidates")

    ws_emp.append(EMPIRICAL_HEADERS)
    for row in empirical_rows:
        ws_emp.append([row.get(col) for col in EMPIRICAL_HEADERS])

    ws_reg.append(REGRESSION_HEADERS)
    for row in regression_rows:
        ws_reg.append([row.get(col) for col in REGRESSION_HEADERS])

    style_output_sheet(ws_emp)
    style_output_sheet(ws_reg)
    wb.save(path)


def process_file(
    app: xw.App,
    file_path: Path,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    labels = parse_model_label(file_path)
    wb = app.books.open(str(file_path), update_links=False)
    try:
        empirical_rows: list[dict[str, Any]] = []
        regression_rows: list[dict[str, Any]] = []

        empirical_sheet = get_sheet_case_insensitive(wb, "Empirical Model")
        if empirical_sheet is None:
            print(f"Skipping empirical for {file_path.name}: sheet not found")
        else:
            empirical_rows = build_empirical_rows(
                wb=wb,
                ws=empirical_sheet,
                labels=labels,
                source_file=file_path.name,
            )

        regression_sheet = get_sheet_case_insensitive(wb, "Regression Model")
        if regression_sheet is None:
            print(f"Skipping regression for {file_path.name}: sheet not found")
        else:
            regression_rows = build_regression_rows(
                wb=wb,
                ws=regression_sheet,
                labels=labels,
                source_file=file_path.name,
            )

        return empirical_rows, regression_rows
    finally:
        safe_close_workbook(wb)


def run() -> None:
    source_dir = input_dir.expanduser().resolve()
    target_dir = output_dir.expanduser().resolve()

    if not source_dir.exists():
        raise SystemExit(f"Input directory not found: {source_dir}")

    files, skipped_count = list_input_files(source_dir)
    out_path = output_path_for_run(source_dir, target_dir)

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
        for file_path in files:
            print(f"Processing {file_path.name}")
            try:
                empirical, regression = process_file(app, file_path)
                empirical_rows.extend(empirical)
                regression_rows.extend(regression)
                processed_files += 1
            except Exception as exc:  # pylint: disable=broad-except
                print(f"Skipping {file_path.name}: {exc}")
    finally:
        app.quit()

    write_output_workbook(
        path=out_path,
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
    )

    print(f"Output file: {out_path}")
    print(f"Files processed: {processed_files}")
    print(f"Files skipped: {skipped_count + max(0, len(files) - processed_files)}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run()

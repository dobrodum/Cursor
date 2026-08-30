#!/usr/bin/env python3
"""Extract empirical and regression model candidates from .xlsx files."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
import math
from pathlib import Path
import re
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# Configure these two paths before running.
input_dir = Path("./input")
output_dir = Path("./output")

N_QUARTERS = 10

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

PERIOD_DAY = {"early": 5, "mid": 15, "late": 25}
MONTH_NUMBER = {
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

FILENAME_RE = re.compile(
    r"^.*-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*"
    r"(?P<period>Early|Mid|Late)(?P<month>[A-Za-z]{3})(?P<year>\d{4})_Send$",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class ModelMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if isinstance(value, str):
        return value.strip().lower()
    return ""


def to_float(value: Any) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return None
        return float(value)
    raw = str(value).strip().replace(",", "")
    if not raw:
        return None
    try:
        parsed = float(raw)
        if math.isnan(parsed) or math.isinf(parsed):
            return None
        return parsed
    except ValueError:
        return None


def clean_for_excel(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
        return ""
    return value


def parse_model_meta(file_path: Path) -> ModelMeta:
    stem = file_path.stem
    matched = FILENAME_RE.match(stem)
    if not matched:
        ticker_guess = "UNKNOWN"
        parts = [part.strip() for part in stem.split("-") if part.strip()]
        if len(parts) >= 2:
            ticker_guess = re.sub(r"[^A-Za-z0-9]", "", parts[-2]).upper() or "UNKNOWN"
        model_period = "Unknown_0000"
        return ModelMeta(
            model=f"{ticker_guess}_{model_period}",
            ticker=ticker_guess,
            model_period=model_period,
            model_date="",
        )

    ticker = matched.group("ticker").upper()
    period = matched.group("period").title()
    month_abbr = matched.group("month").title()
    year = int(matched.group("year"))

    month_num = MONTH_NUMBER.get(month_abbr.lower())
    day_num = PERIOD_DAY[period.lower()]

    if month_num is None:
        model_period = f"{period}{month_abbr}_{year}"
        model_date = ""
    else:
        model_period = f"{period}{month_abbr}_{year}"
        model_date = date(year, month_num, day_num).isoformat()

    return ModelMeta(
        model=f"{ticker}_{model_period}",
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def ensure_output_path(input_folder: Path, out_folder: Path) -> Path:
    out_folder.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_folder.name}_PARAM"
    candidate = out_folder / f"{base_name}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = out_folder / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def get_sheet_ci(wb: xw.Book, sheet_name: str) -> xw.Sheet | None:
    wanted = sheet_name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == wanted:
            return sheet
    return None


def as_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def find_anchor_max(sheet: xw.Sheet) -> xw.Range | None:
    try:
        found = sheet.api.Cells.Find(
            What="max",
            LookIn=-4163,  # xlValues
            LookAt=1,  # xlWhole
            SearchOrder=1,  # xlByRows
            SearchDirection=1,  # xlNext
            MatchCase=False,
        )
        if found is not None:
            return sheet.range((int(found.Row), int(found.Column)))
    except Exception:
        pass

    used = sheet.used_range
    values = as_2d(used.value)
    start_row = used.row
    start_col = used.column
    for r_idx, row_values in enumerate(values, start=start_row):
        for c_idx, cell_value in enumerate(row_values, start=start_col):
            if normalize_text(cell_value) == "max":
                return sheet.range((r_idx, c_idx))
    return None


def collect_numeric_rows(sheet: xw.Sheet, col: int, end_row: int, max_scan: int = 300) -> list[tuple[int, float]]:
    rows: list[tuple[int, float]] = []
    blank_streak = 0
    row = end_row
    while row >= 1 and (end_row - row) < max_scan:
        val = to_float(sheet.range((row, col)).value)
        if val is None:
            blank_streak += 1
            if rows and blank_streak >= 2:
                break
        else:
            blank_streak = 0
            rows.append((row, val))
        row -= 1
    rows.reverse()
    return rows


def collect_numeric_pairs(
    sheet: xw.Sheet, x_col: int, y_col: int, end_row: int, max_scan: int = 300
) -> list[tuple[int, float, float]]:
    rows: list[tuple[int, float, float]] = []
    blank_streak = 0
    row = end_row
    while row >= 1 and (end_row - row) < max_scan:
        x_val = to_float(sheet.range((row, x_col)).value)
        y_val = to_float(sheet.range((row, y_col)).value)
        if x_val is None or y_val is None:
            blank_streak += 1
            if rows and blank_streak >= 2:
                break
        else:
            blank_streak = 0
            rows.append((row, x_val, y_val))
        row -= 1
    rows.reverse()
    return rows


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
    formula = formula_r1c1 if formula_r1c1.startswith("=") else f"={formula_r1c1}"
    try:
        cell.api.Formula2R1C1 = formula
        return
    except Exception:
        pass
    try:
        cell.formula2 = formula
        return
    except Exception:
        pass
    cell.api.FormulaR1C1 = formula


def close_source_workbook(wb: xw.Book) -> None:
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
    except Exception as exc:
        print(f"Warning: close without save fallback failed for {wb.name}: {exc}")


def extract_empirical_candidates(
    wb: xw.Book, meta: ModelMeta, source_file: str
) -> list[dict[str, Any]]:
    sheet = get_sheet_ci(wb, "Empirical Model")
    if sheet is None:
        print(f"Skipped empirical in {source_file}: sheet 'Empirical Model' not found")
        return []

    anchor = find_anchor_max(sheet)
    if anchor is None:
        print(f"Skipped empirical in {source_file}: anchor 'max' not found")
        return []

    anchor_row = anchor.row
    anchor_col = anchor.column

    # Anchor-based offsets; keep these near each other for easy tuning if layout shifts.
    penetration_col = anchor_col - 9
    quarterly_sales_col = anchor_col - 11
    reported_sales_col = anchor_col - 10
    growth_rate_col = anchor_col - 8
    sales_captured_col = anchor_col - 7
    quarter_label_col = anchor_col - 12

    if min(
        penetration_col,
        quarterly_sales_col,
        reported_sales_col,
        growth_rate_col,
        sales_captured_col,
        quarter_label_col,
    ) < 1:
        print(f"Skipped empirical in {source_file}: anchor offsets fall outside sheet")
        return []

    penetration_history = collect_numeric_rows(sheet, penetration_col, anchor_row - 1)
    if not penetration_history:
        print(f"Skipped empirical in {source_file}: no penetration history found")
        return []

    rows: list[dict[str, Any]] = []
    helper_avg_cell = sheet.range((anchor_row, anchor_col + 40))
    helper_fcst_cell = sheet.range((anchor_row, anchor_col + 41))

    max_quarters = min(N_QUARTERS, len(penetration_history))
    for n_quarters in range(1, max_quarters + 1):
        start_row = penetration_history[-n_quarters][0]
        end_row = penetration_history[-1][0]

        set_formula2_r1c1(
            helper_avg_cell,
            f"AVERAGE(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})",
        )
        set_formula2_r1c1(
            helper_fcst_cell,
            (
                f'IFERROR(R{end_row}C{quarterly_sales_col}'
                f"/R{helper_avg_cell.row}C{helper_avg_cell.column},\"\")"
            ),
        )
        wb.app.calculate()

        avg_penetration_pct = to_float(helper_avg_cell.value)
        quarterly_sales = to_float(sheet.range((end_row, quarterly_sales_col)).value)
        reported_sales = to_float(sheet.range((end_row, reported_sales_col)).value)
        growth_rate_pct = to_float(sheet.range((end_row, growth_rate_col)).value)
        sales_captured_in_db_pct = to_float(sheet.range((end_row, sales_captured_col)).value)

        forecast_value = to_float(helper_fcst_cell.value)
        if forecast_value is None and quarterly_sales is not None and avg_penetration_pct:
            if avg_penetration_pct != 0:
                forecast_value = quarterly_sales / avg_penetration_pct

        forecast_max = to_float(sheet.range((anchor_row + n_quarters, anchor_col)).value)
        forecast_min = to_float(sheet.range((anchor_row + n_quarters, anchor_col + 1)).value)
        if (
            (forecast_max is None or forecast_min is None)
            and forecast_value is not None
            and growth_rate_pct is not None
        ):
            movement = abs(growth_rate_pct)
            forecast_max = forecast_value * (1 + movement)
            forecast_min = forecast_value * (1 - movement)

        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )
        last_quarter_used = sheet.range((end_row, quarter_label_col)).value

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": n_quarters,
                "last_quarter_used": last_quarter_used if last_quarter_used is not None else "",
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
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


def extract_regression_candidates(
    wb: xw.Book, meta: ModelMeta, source_file: str
) -> list[dict[str, Any]]:
    sheet = get_sheet_ci(wb, "Regression Model")
    if sheet is None:
        print(f"Skipped regression in {source_file}: sheet 'Regression Model' not found")
        return []

    anchor = find_anchor_max(sheet)
    if anchor is None:
        print(f"Skipped regression in {source_file}: anchor 'max' not found")
        return []

    anchor_row = anchor.row
    anchor_col = anchor.column
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    if min(y_col, x_col) < 1:
        print(f"Skipped regression in {source_file}: anchor offsets fall outside sheet")
        return []

    pairs = collect_numeric_pairs(sheet, x_col, y_col, anchor_row - 1)
    if len(pairs) < 2:
        print(f"Skipped regression in {source_file}: not enough historical x/y points")
        return []

    rows: list[dict[str, Any]] = []
    helper_intercept_cell = sheet.range((anchor_row, anchor_col + 40))
    helper_slope_cell = sheet.range((anchor_row, anchor_col + 41))
    helper_fcst_cell = sheet.range((anchor_row, anchor_col + 42))

    max_quarters = min(N_QUARTERS, len(pairs))
    previous_signature: tuple[Any, ...] | None = None
    for n_quarters in range(2, max_quarters + 1):
        start_row = pairs[-n_quarters][0]
        end_row = pairs[-1][0]

        set_formula2_r1c1(
            helper_intercept_cell,
            f"INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})",
        )
        set_formula2_r1c1(
            helper_slope_cell,
            f"SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})",
        )
        set_formula2_r1c1(
            helper_fcst_cell,
            (
                f"R{helper_intercept_cell.row}C{helper_intercept_cell.column}"
                f"+R{helper_slope_cell.row}C{helper_slope_cell.column}*R{end_row}C{x_col}"
            ),
        )
        wb.app.calculate()

        intercept = to_float(helper_intercept_cell.value)
        slope = to_float(helper_slope_cell.value)

        forecast_total_without_sa = to_float(sheet.range((anchor_row + n_quarters, anchor_col - 1)).value)
        if forecast_total_without_sa is None:
            forecast_total_without_sa = to_float(helper_fcst_cell.value)

        forecast_max = to_float(sheet.range((anchor_row + n_quarters, anchor_col)).value)
        forecast_min = to_float(sheet.range((anchor_row + n_quarters, anchor_col + 1)).value)
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        actual_candidate = to_float(sheet.range((anchor_row + n_quarters, anchor_col - 2)).value)
        actual_value: Any = actual_candidate if actual_candidate is not None else ""

        signature = (
            round(forecast_total_without_sa, 10) if forecast_total_without_sa is not None else None,
            round(forecast_max, 10) if forecast_max is not None else None,
            round(forecast_min, 10) if forecast_min is not None else None,
            round(intercept, 10) if intercept is not None else None,
            round(slope, 10) if slope is not None else None,
        )
        if signature == previous_signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n_quarters,
                "num_quarters_used": n_quarters,
                "forecast_value": forecast_total_without_sa,
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


def write_sheet(ws, headers: list[str], rows: list[dict[str, Any]]) -> None:
    ws.append(headers)
    for row in rows:
        ws.append([clean_for_excel(row.get(header, "")) for header in headers])

    for header_cell in ws[1]:
        header_cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row in rows:
            value = row.get(header, "")
            text = "" if value is None else str(value)
            if len(text) > max_len:
                max_len = len(text)
        ws.column_dimensions[get_column_letter(idx)].width = min(max(12, max_len + 2), 52)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    out_wb = Workbook()
    default_sheet = out_wb.active
    out_wb.remove(default_sheet)

    empirical_ws = out_wb.create_sheet("empirical_candidates")
    regression_ws = out_wb.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_HEADERS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_HEADERS, regression_rows)
    out_wb.save(output_path)


def build_candidates() -> None:
    if not input_dir.exists() or not input_dir.is_dir():
        print(f"Input folder does not exist: {input_dir}")
        return

    output_path = ensure_output_path(input_dir, output_dir)
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
            if not file_path.is_file():
                print(f"Skipped {file_path.name}: not a file")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file")
                continue

            print(f"Processing {file_path.name}")
            meta = parse_model_meta(file_path)
            wb: xw.Book | None = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(
                    extract_empirical_candidates(
                        wb=wb,
                        meta=meta,
                        source_file=file_path.name,
                    )
                )
                regression_rows.extend(
                    extract_regression_candidates(
                        wb=wb,
                        meta=meta,
                        source_file=file_path.name,
                    )
                )
                processed_files += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error ({exc})")
            finally:
                if wb is not None:
                    close_source_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(
        output_path=output_path,
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
    )

    print(f"Output path: {output_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    build_candidates()

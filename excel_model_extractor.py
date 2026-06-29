from __future__ import annotations

import calendar
import datetime as dt
import re
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -----------------------------
# User-configurable paths
# -----------------------------
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")


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


# Anchor-based offsets (relative to "max" anchor cell).
# Keep these centralized for easy adjustment if source model layout shifts.
EMPIRICAL_OFFSETS = {
    "first_data_row_offset": 1,
    "row_step": 1,
    "forecast_max_col_offset": 0,
    "forecast_min_col_offset": 1,
    "forecast_value_col_offset": -2,  # estimated total sold
    "actual_value_col_offset": -1,  # reported sales
    "last_quarter_col_offset": -11,
    "quarterly_sales_col_offset": -8,
    "reported_sales_col_offset": -1,
    "growth_rate_col_offset": -6,
    "sales_captured_col_offset": -5,
    "penetration_source_col_offset": -7,
    "temp_formula_col_offset": 6,
}

REGRESSION_OFFSETS = {
    "result_first_row_offset": 1,
    "result_row_step": 1,
    "forecast_total_without_sa_col_offset": -2,  # TOT FCST w/o SA
    "forecast_max_col_offset": 0,
    "forecast_min_col_offset": 1,
    "actual_value_col_offset": -1,  # optional
    "temp_intercept_col_offset": 5,
    "temp_slope_col_offset": 6,
}


def parse_model_metadata(file_name: str) -> dict[str, str] | None:
    """
    Example:
      MedMiner_Model - AORT - MidJan2026_Send.xlsx
    -> ticker=AORT, model_period=MidJan_2026, model_date=2026-01-15, model=AORT_MidJan_2026
    """
    name_wo_ext = Path(file_name).stem
    match = re.search(r"\s-\s(?P<ticker>[^-]+?)\s-\s(?P<label>[A-Za-z]+\d{4})", name_wo_ext)
    if not match:
        return None

    ticker = match.group("ticker").strip()
    period_label = match.group("label").strip()

    period_match = re.fullmatch(
        r"(?P<timing>Early|Mid|Late)(?P<month>[A-Za-z]{3})(?P<year>\d{4})",
        period_label,
        flags=re.IGNORECASE,
    )
    if not period_match:
        return None

    timing_raw = period_match.group("timing")
    month_raw = period_match.group("month")
    year_str = period_match.group("year")

    timing = timing_raw[0].upper() + timing_raw[1:].lower()
    month_abbr = month_raw[0].upper() + month_raw[1:].lower()
    year = int(year_str)

    month_num = list(calendar.month_abbr).index(month_abbr) if month_abbr in calendar.month_abbr else 0
    if month_num == 0:
        return None

    day_map = {"Early": 5, "Mid": 15, "Late": 25}
    model_day = day_map[timing]
    model_date = dt.date(year, month_num, model_day).isoformat()
    model_period = f"{timing}{month_abbr}_{year}"

    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": f"{ticker}_{model_period}",
    }


def to_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).strip().replace(",", ""))
    except (TypeError, ValueError):
        return None


def safe_read(sheet: xw.main.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    try:
        return sheet.range((row, col)).value
    except Exception:
        return None


def find_anchor(sheet: xw.main.Sheet, anchor_text: str = "max") -> tuple[int, int] | None:
    used = sheet.used_range
    values = used.value
    if values is None:
        return None

    if not isinstance(values, list):
        grid = [[values]]
    elif values and not isinstance(values[0], list):
        grid = [values]
    else:
        grid = values

    start_row = used.row
    start_col = used.column
    target = anchor_text.strip().lower()

    for r_idx, row_values in enumerate(grid):
        for c_idx, value in enumerate(row_values):
            if isinstance(value, str) and value.strip().lower() == target:
                return start_row + r_idx, start_col + c_idx
    return None


def set_formula2(cell: xw.main.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        # Safe compatibility fallback for older Excel versions.
        cell.formula = formula_r1c1


def safe_close_workbook(wb: xw.main.Book) -> None:
    # Never save source workbooks.
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.api.Close(SaveChanges=False)
    except Exception:
        try:
            wb.close()
        except Exception:
            pass


def process_empirical_sheet(
    wb: xw.main.Book,
    metadata: dict[str, str],
    source_file: str,
    n_quarters: int = 10,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"  Skipped empirical for {source_file}: missing sheet 'Empirical Model'")
        return rows

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"  Skipped empirical for {source_file}: 'max' anchor not found")
        return rows

    anchor_row, anchor_col = anchor
    offsets = EMPIRICAL_OFFSETS

    for i in range(n_quarters):
        num_quarters_used = i + 1
        data_row = anchor_row + offsets["first_data_row_offset"] + (i * offsets["row_step"])

        formula_col = anchor_col + offsets["temp_formula_col_offset"]
        formula_cell = sheet.range((data_row, formula_col))

        penetration_col = anchor_col + offsets["penetration_source_col_offset"]
        start_row = data_row - num_quarters_used + 1
        if start_row < 1:
            continue

        # R1C1 + formula2 for rolling average penetration %.
        avg_formula = (
            f"=AVERAGE(R{start_row}C{penetration_col}:R{data_row}C{penetration_col})"
        )
        set_formula2(formula_cell, avg_formula)
        wb.app.calculate()
        avg_penetration_pct = to_float(formula_cell.value)

        forecast_value = to_float(
            safe_read(sheet, data_row, anchor_col + offsets["forecast_value_col_offset"])
        )
        actual_value = to_float(
            safe_read(sheet, data_row, anchor_col + offsets["actual_value_col_offset"])
        )
        forecast_max = to_float(
            safe_read(sheet, data_row, anchor_col + offsets["forecast_max_col_offset"])
        )
        forecast_min = to_float(
            safe_read(sheet, data_row, anchor_col + offsets["forecast_min_col_offset"])
        )
        quarterly_sales = to_float(
            safe_read(sheet, data_row, anchor_col + offsets["quarterly_sales_col_offset"])
        )
        reported_sales = to_float(
            safe_read(sheet, data_row, anchor_col + offsets["reported_sales_col_offset"])
        )
        growth_rate_pct = to_float(
            safe_read(sheet, data_row, anchor_col + offsets["growth_rate_col_offset"])
        )
        sales_captured_pct = to_float(
            safe_read(sheet, data_row, anchor_col + offsets["sales_captured_col_offset"])
        )

        last_quarter_used = safe_read(
            sheet, data_row, anchor_col + offsets["last_quarter_col_offset"]
        )
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        # Skip empty rows to avoid output noise.
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
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
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
                "sales_captured_in_db_pct": sales_captured_pct,
                "source_file": source_file,
            }
        )

    return rows


def process_regression_sheet(
    wb: xw.main.Book,
    metadata: dict[str, str],
    source_file: str,
    n_quarters: int = 10,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"  Skipped regression for {source_file}: missing sheet 'Regression Model'")
        return rows

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"  Skipped regression for {source_file}: 'max' anchor not found")
        return rows

    anchor_row, anchor_col = anchor
    offsets = REGRESSION_OFFSETS
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    prev_signature: tuple[Any, ...] | None = None

    for i in range(n_quarters):
        num_quarters_used = i + 1
        data_end_row = anchor_row - 1
        data_start_row = data_end_row - num_quarters_used + 1
        if data_start_row < 1:
            continue

        temp_row = anchor_row + offsets["result_first_row_offset"] + (i * offsets["result_row_step"])
        intercept_cell = sheet.range((temp_row, anchor_col + offsets["temp_intercept_col_offset"]))
        slope_cell = sheet.range((temp_row, anchor_col + offsets["temp_slope_col_offset"]))

        intercept_formula = (
            f"=INTERCEPT(R{data_start_row}C{y_col}:R{data_end_row}C{y_col},"
            f"R{data_start_row}C{x_col}:R{data_end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{data_start_row}C{y_col}:R{data_end_row}C{y_col},"
            f"R{data_start_row}C{x_col}:R{data_end_row}C{x_col})"
        )

        # R1C1 + formula2 per requirement.
        set_formula2(intercept_cell, intercept_formula)
        set_formula2(slope_cell, slope_formula)
        wb.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)

        result_row = anchor_row + offsets["result_first_row_offset"] + (i * offsets["result_row_step"])

        forecast_total_without_sa = to_float(
            safe_read(
                sheet,
                result_row,
                anchor_col + offsets["forecast_total_without_sa_col_offset"],
            )
        )
        if forecast_total_without_sa is None and intercept is not None and slope is not None:
            latest_x = to_float(safe_read(sheet, data_end_row, x_col))
            if latest_x is not None:
                forecast_total_without_sa = intercept + (slope * latest_x)

        forecast_max = to_float(
            safe_read(sheet, result_row, anchor_col + offsets["forecast_max_col_offset"])
        )
        forecast_min = to_float(
            safe_read(sheet, result_row, anchor_col + offsets["forecast_min_col_offset"])
        )
        actual_value = to_float(
            safe_read(sheet, result_row, anchor_col + offsets["actual_value_col_offset"])
        )
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        if all(
            value is None
            for value in (
                forecast_total_without_sa,
                forecast_max,
                forecast_min,
                intercept,
                slope,
            )
        ):
            continue

        # Prevent duplicate final row.
        signature = (
            num_quarters_used,
            round(forecast_total_without_sa, 10) if forecast_total_without_sa is not None else None,
            round(forecast_max, 10) if forecast_max is not None else None,
            round(forecast_min, 10) if forecast_min is not None else None,
            round(intercept, 10) if intercept is not None else None,
            round(slope, 10) if slope is not None else None,
        )
        if signature == prev_signature:
            continue
        prev_signature = signature

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
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
                "source_file": source_file,
            }
        )

    return rows


def next_output_path(in_dir: Path, out_dir: Path) -> Path:
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


def write_sheet(ws, columns: list[str], rows: list[dict[str, Any]]) -> None:
    ws.append(columns)
    for row in rows:
        ws.append([row.get(col) for col in columns])

    # Formatting requirements.
    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws.freeze_panes = "A2"

    if ws.max_row >= 1 and ws.max_column >= 1:
        ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"

    for idx in range(1, ws.max_column + 1):
        col_letter = get_column_letter(idx)
        max_len = 0
        for cell in ws[col_letter]:
            val = "" if cell.value is None else str(cell.value)
            if len(val) > max_len:
                max_len = len(val)
        ws.column_dimensions[col_letter].width = max(12, min(max_len + 2, 48))


def iter_candidate_files(in_dir: Path) -> tuple[list[Path], int]:
    valid_files: list[Path] = []
    skipped = 0
    for file_path in sorted(in_dir.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped file: {file_path.name} (temporary file)")
            skipped += 1
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped file: {file_path.name} (not an .xlsx file)")
            skipped += 1
            continue
        valid_files.append(file_path)
    return valid_files, skipped


def main() -> None:
    in_dir = Path(input_dir)
    out_dir = Path(output_dir)

    if not in_dir.exists() or not in_dir.is_dir():
        raise FileNotFoundError(f"Input directory not found: {in_dir}")

    out_dir.mkdir(parents=True, exist_ok=True)
    output_path = next_output_path(in_dir, out_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []

    files, skipped_count = iter_candidate_files(in_dir)
    processed_count = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    original_calculation = app.calculation
    app.calculation = "manual"

    try:
        for file_path in files:
            metadata = parse_model_metadata(file_path.name)
            if metadata is None:
                print(f"Skipped file: {file_path.name} (could not parse ticker/period from filename)")
                skipped_count += 1
                continue

            print(f"Processing file: {file_path.name}")

            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(
                    process_empirical_sheet(wb=wb, metadata=metadata, source_file=file_path.name, n_quarters=10)
                )
                regression_rows.extend(
                    process_regression_sheet(wb=wb, metadata=metadata, source_file=file_path.name, n_quarters=10)
                )
                processed_count += 1
            except Exception as exc:
                print(f"Skipped file: {file_path.name} (processing error: {exc})")
                skipped_count += 1
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.calculation = original_calculation
        app.quit()

    out_wb = Workbook()
    empirical_ws = out_wb.active
    empirical_ws.title = "empirical_candidates"
    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)

    regression_ws = out_wb.create_sheet("regression_candidates")
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    out_wb.save(output_path)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_count}")
    print(f"Number of files skipped: {skipped_count}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

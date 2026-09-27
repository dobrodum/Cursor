from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# =============================
# User-configurable directories
# =============================
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

DAY_BY_PERIOD_PREFIX = {
    "early": 5,
    "mid": 15,
    "late": 25,
}

MONTH_TO_NUMBER = {
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


@dataclass
class FileLabel:
    ticker: str
    model_period: str
    model_date: str
    model: str


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[\s_\-/]+", " ", text)
    return text


def to_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    try:
        parsed = float(str(value).replace(",", ""))
        if math.isnan(parsed):
            return None
        return parsed
    except (TypeError, ValueError):
        return None


def float_equal(a: float | None, b: float | None, tol: float = 1e-9) -> bool:
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    return abs(a - b) <= tol


def parse_file_label(file_path: Path) -> FileLabel:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split("-")]
    ticker = "UNKNOWN"
    period_token_source = stem

    if len(parts) >= 2 and parts[1]:
        ticker = parts[1].upper()
    if len(parts) >= 3 and parts[2]:
        period_token_source = parts[2]

    period_match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3})\s*(\d{4})",
        period_token_source,
        flags=re.IGNORECASE,
    )

    if not period_match:
        # Fallback: attempt anywhere in full stem.
        period_match = re.search(
            r"(Early|Mid|Late)\s*([A-Za-z]{3})\s*(\d{4})",
            stem,
            flags=re.IGNORECASE,
        )

    if not period_match:
        model_period = "unknown_period"
        model_date = ""
        model = f"{ticker}_{model_period}"
        return FileLabel(
            ticker=ticker,
            model_period=model_period,
            model_date=model_date,
            model=model,
        )

    period_prefix = period_match.group(1).capitalize()
    month_abbrev = period_match.group(2).lower()
    year = int(period_match.group(3))

    month_num = MONTH_TO_NUMBER.get(month_abbrev)
    day_num = DAY_BY_PERIOD_PREFIX.get(period_prefix.lower())

    model_period = f"{period_prefix}{month_abbrev.capitalize()}_{year}"
    model_date = ""
    if month_num and day_num:
        model_date = date(year, month_num, day_num).isoformat()

    model = f"{ticker}_{model_period}"
    return FileLabel(
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
        model=model,
    )


def choose_output_path(input_folder: Path, out_folder: Path) -> Path:
    out_folder.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_folder.name}_PARAM.xlsx"
    base_path = out_folder / base_name
    if not base_path.exists():
        return base_path

    suffix = 1
    while True:
        candidate = out_folder / f"{input_folder.name}_PARAM.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def ensure_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def find_anchor_cell(sheet: xw.Sheet, target: str = "max") -> tuple[int, int] | None:
    used = sheet.used_range
    values = ensure_2d(used.value)
    if not values:
        return None

    start_row = used.row
    start_col = used.column
    target_norm = normalize_label(target)

    for r_idx, row_vals in enumerate(values):
        for c_idx, cell_val in enumerate(row_vals):
            if normalize_label(cell_val) == target_norm:
                return start_row + r_idx, start_col + c_idx
    return None


def build_header_column_map(sheet: xw.Sheet, header_row: int, last_col: int) -> dict[str, int]:
    if last_col < 1:
        return {}

    row_values = sheet.range((header_row, 1), (header_row, last_col)).value
    if row_values is None:
        return {}
    if not isinstance(row_values, list):
        row_values = [row_values]

    normalized_to_col: dict[str, int] = {}
    for idx, value in enumerate(row_values, start=1):
        key = normalize_label(value)
        if key:
            normalized_to_col[key] = idx
    return normalized_to_col


def find_column_by_alias(
    header_map: dict[str, int],
    aliases: list[str],
) -> int | None:
    for alias in aliases:
        col = header_map.get(normalize_label(alias))
        if col is not None:
            return col
    return None


def read_cell(sheet: xw.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    return sheet.cells(row, col).value


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
    """
    Prefer Formula2 R1C1 writes. Fallbacks keep the script usable across versions.
    """
    try:
        cell.api.Formula2R1C1 = formula_r1c1
        return
    except Exception:
        pass

    try:
        cell.formula2 = formula_r1c1
        return
    except Exception:
        pass

    cell.formula = formula_r1c1


def safe_close_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.api.Saved = True
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


def extract_empirical_rows(
    wb: xw.Book,
    labels: FileLabel,
    source_file_name: str,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []

    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"  skipped empirical in {source_file_name}: missing sheet 'Empirical Model'")
        return rows

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"  skipped empirical in {source_file_name}: missing 'max' anchor")
        return rows

    anchor_row, anchor_col = anchor
    used = sheet.used_range
    last_row = used.last_cell.row
    last_col = used.last_cell.column

    header_map = build_header_column_map(sheet, anchor_row, last_col)

    # Header lookup first, then anchor-based offsets as fallback.
    num_quarters_col = find_column_by_alias(
        header_map,
        ["num_quarters_used", "num quarters used", "num quarters", "quarters used"],
    )
    last_quarter_col = find_column_by_alias(
        header_map,
        ["last_quarter_used", "last quarter used", "last quarter"],
    )
    forecast_col = find_column_by_alias(
        header_map,
        ["estimated total sold", "forecast value", "forecast", "est total sold"],
    )
    actual_col = find_column_by_alias(
        header_map,
        ["reported sales", "actual value", "actual"],
    )
    max_col = find_column_by_alias(header_map, ["max", "forecast max"]) or anchor_col
    min_col = find_column_by_alias(header_map, ["min", "forecast min"]) or (anchor_col + 1)
    quarterly_sales_col = find_column_by_alias(
        header_map,
        ["quarterly sales", "quarterly_sales"],
    )
    growth_rate_col = find_column_by_alias(
        header_map,
        ["growth rate pct", "growth_rate_pct", "growth rate %"],
    )
    sales_captured_col = find_column_by_alias(
        header_map,
        [
            "sales captured in db pct",
            "sales_captured_in_db_pct",
            "sales captured in db %",
        ],
    )

    # Anchor fallback offsets if headers are not present.
    if num_quarters_col is None:
        num_quarters_col = anchor_col - 4
    if last_quarter_col is None:
        last_quarter_col = anchor_col - 3
    if forecast_col is None:
        forecast_col = anchor_col - 1
    if actual_col is None:
        actual_col = anchor_col + 2
    if quarterly_sales_col is None:
        quarterly_sales_col = anchor_col - 2
    if growth_rate_col is None:
        growth_rate_col = anchor_col + 3
    if sales_captured_col is None:
        sales_captured_col = anchor_col + 4

    # Historical penetration/support column for formula-based average.
    # Kept anchor-relative to avoid full-sheet rescans.
    penetration_history_col = anchor_col - 6
    history_end_row = max(anchor_row - 1, 1)

    n_quarters = 10
    data_rows = [anchor_row + i for i in range(1, n_quarters + 1)]

    # Scratch area for temporary formula2 writes.
    scratch_start_row = last_row + 2
    scratch_col = last_col + 2

    scratch_cells: list[xw.Range] = []
    for idx, n in enumerate(range(1, n_quarters + 1)):
        hist_start = max(history_end_row - n + 1, 1)
        scratch_cell = sheet.cells(scratch_start_row + idx, scratch_col)
        formula = (
            f'=IFERROR(AVERAGE(R{hist_start}C{penetration_history_col}:'
            f'R{history_end_row}C{penetration_history_col}),"")'
        )
        set_formula2_r1c1(scratch_cell, formula)
        scratch_cells.append(scratch_cell)

    wb.app.calculate()
    avg_penetration_values = [to_float(cell.value) for cell in scratch_cells]

    for idx, row_idx in enumerate(data_rows):
        num_quarters_used = to_float(read_cell(sheet, row_idx, num_quarters_col))
        if num_quarters_used is None:
            num_quarters_used = float(idx + 1)

        last_quarter_used = read_cell(sheet, row_idx, last_quarter_col)
        forecast_value = to_float(read_cell(sheet, row_idx, forecast_col))
        actual_value = to_float(read_cell(sheet, row_idx, actual_col))
        forecast_max = to_float(read_cell(sheet, row_idx, max_col))
        forecast_min = to_float(read_cell(sheet, row_idx, min_col))
        quarterly_sales = to_float(read_cell(sheet, row_idx, quarterly_sales_col))
        reported_sales = actual_value
        growth_rate_pct = to_float(read_cell(sheet, row_idx, growth_rate_col))
        sales_captured_pct = to_float(read_cell(sheet, row_idx, sales_captured_col))
        avg_penetration_pct = avg_penetration_values[idx]

        if forecast_value is None and quarterly_sales is not None and avg_penetration_pct not in (None, 0):
            forecast_value = quarterly_sales / avg_penetration_pct

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
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
                "source_file": source_file_name,
            }
        )

    return rows


def extract_regression_rows(
    wb: xw.Book,
    labels: FileLabel,
    source_file_name: str,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []

    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"  skipped regression in {source_file_name}: missing sheet 'Regression Model'")
        return rows

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"  skipped regression in {source_file_name}: missing 'max' anchor")
        return rows

    anchor_row, anchor_col = anchor
    used = sheet.used_range
    last_row = used.last_cell.row
    last_col = used.last_cell.column

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    history_end_row = max(anchor_row - 1, 1)
    next_x_row = history_end_row + 1

    header_map = build_header_column_map(sheet, anchor_row, last_col)
    num_quarters_col = find_column_by_alias(
        header_map,
        ["num_quarters_used", "num quarters used", "num quarters", "quarters used"],
    )
    forecast_col = find_column_by_alias(
        header_map,
        ["tot fcst w/o sa", "forecast value", "forecast", "tot forecast w/o sa"],
    )
    actual_col = find_column_by_alias(header_map, ["actual value", "actual", "reported sales"])
    max_col = find_column_by_alias(header_map, ["max", "forecast max"]) or anchor_col
    min_col = find_column_by_alias(header_map, ["min", "forecast min"]) or (anchor_col + 1)

    if num_quarters_col is None:
        num_quarters_col = anchor_col - 4
    if forecast_col is None:
        forecast_col = anchor_col - 1
    if actual_col is None:
        actual_col = anchor_col + 2

    n_quarters = 10
    data_rows = [anchor_row + i for i in range(1, n_quarters + 1)]

    scratch_start_row = last_row + 2
    intercept_col = last_col + 2
    slope_col = last_col + 3
    intercept_cells: list[xw.Range] = []
    slope_cells: list[xw.Range] = []

    for idx, n in enumerate(range(1, n_quarters + 1)):
        hist_start = max(history_end_row - n + 1, 1)
        intercept_cell = sheet.cells(scratch_start_row + idx, intercept_col)
        slope_cell = sheet.cells(scratch_start_row + idx, slope_col)

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{hist_start}C{y_col}:R{history_end_row}C{y_col},'
            f'R{hist_start}C{x_col}:R{history_end_row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{hist_start}C{y_col}:R{history_end_row}C{y_col},'
            f'R{hist_start}C{x_col}:R{history_end_row}C{x_col}),"")'
        )

        set_formula2_r1c1(intercept_cell, intercept_formula)
        set_formula2_r1c1(slope_cell, slope_formula)
        intercept_cells.append(intercept_cell)
        slope_cells.append(slope_cell)

    wb.app.calculate()
    intercept_values = [to_float(cell.value) for cell in intercept_cells]
    slope_values = [to_float(cell.value) for cell in slope_cells]

    next_x_value = to_float(read_cell(sheet, next_x_row, x_col))
    if next_x_value is None:
        current_x = to_float(read_cell(sheet, history_end_row, x_col))
        next_x_value = (current_x + 1) if current_x is not None else None

    previous_row: dict[str, Any] | None = None
    for idx, row_idx in enumerate(data_rows):
        num_quarters_used = to_float(read_cell(sheet, row_idx, num_quarters_col))
        if num_quarters_used is None:
            num_quarters_used = float(idx + 1)

        intercept = intercept_values[idx]
        slope = slope_values[idx]

        forecast_value = to_float(read_cell(sheet, row_idx, forecast_col))
        if forecast_value is None and intercept is not None and slope is not None and next_x_value is not None:
            forecast_value = intercept + (slope * next_x_value)

        actual_value = to_float(read_cell(sheet, row_idx, actual_col))
        forecast_max = to_float(read_cell(sheet, row_idx, max_col))
        forecast_min = to_float(read_cell(sheet, row_idx, min_col))

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        row_data = {
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
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file_name,
        }

        # Duplicate guard for final repeated rows with identical computed outputs.
        if previous_row is not None:
            duplicate = (
                float_equal(row_data["forecast_value"], previous_row["forecast_value"])
                and float_equal(row_data["forecast_max"], previous_row["forecast_max"])
                and float_equal(row_data["forecast_min"], previous_row["forecast_min"])
                and float_equal(row_data["intercept"], previous_row["intercept"])
                and float_equal(row_data["slope"], previous_row["slope"])
            )
            if duplicate:
                continue

        rows.append(row_data)
        previous_row = row_data

    return rows


def autosize_columns(ws) -> None:
    for col_idx in range(1, ws.max_column + 1):
        letter = get_column_letter(col_idx)
        max_len = 0
        for row_idx in range(1, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[letter].width = min(max(12, max_len + 2), 40)


def write_sheet(ws, columns: list[str], rows: list[dict[str, Any]]) -> None:
    ws.append(columns)
    for col_idx in range(1, len(columns) + 1):
        ws.cell(row=1, column=col_idx).font = Font(bold=True)

    for row in rows:
        ws.append([row.get(col) for col in columns])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    autosize_columns(ws)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = Workbook()
    ws_empirical = wb.active
    ws_empirical.title = "empirical_candidates"
    write_sheet(ws_empirical, EMPIRICAL_COLUMNS, empirical_rows)

    ws_regression = wb.create_sheet("regression_candidates")
    write_sheet(ws_regression, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def process_all_files(in_dir: Path, out_dir: Path) -> None:
    if not in_dir.exists():
        print(f"Input directory does not exist: {in_dir}")
        return

    output_path = choose_output_path(in_dir, out_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(in_dir.iterdir()):
            if not file_path.is_file():
                print(f"skipped: {file_path.name} (not a file)")
                continue
            if file_path.name.startswith("~"):
                print(f"skipped: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped: {file_path.name} (not .xlsx)")
                continue

            print(f"processing: {file_path.name}")
            labels = parse_file_label(file_path)

            wb: xw.Book | None = None
            try:
                wb = app.books.open(str(file_path), update_links=False)

                empirical = extract_empirical_rows(wb, labels, file_path.name)
                regression = extract_regression_rows(wb, labels, file_path.name)

                empirical_rows.extend(empirical)
                regression_rows.extend(regression)
                processed_files += 1

                print(
                    f"processed: {file_path.name} "
                    f"(empirical_rows={len(empirical)}, regression_rows={len(regression)})"
                )
            except Exception as exc:
                print(f"skipped: {file_path.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"output path: {output_path}")
    print(f"number of files processed: {processed_files}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


def main() -> None:
    in_dir = Path(input_dir).expanduser().resolve()
    out_dir = Path(output_dir).expanduser().resolve()
    process_all_files(in_dir, out_dir)


if __name__ == "__main__":
    main()

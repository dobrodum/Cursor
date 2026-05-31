from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Optional

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -----------------------------
# User-configurable paths
# -----------------------------
input_dir = Path("input")
output_dir = Path("output")


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
class FileLabel:
    ticker: str
    model_period: str
    model_date: str
    model: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", " ", str(value).strip().lower()).strip()


def coerce_number(value: Any) -> Optional[float]:
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(",", "").strip())
    except Exception:
        return None


def parse_file_label(file_name: str) -> Optional[FileLabel]:
    # Expected format example:
    # MedMiner_Model - AORT - MidJan2026_Send.xlsx
    match = re.search(
        r"-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*(?P<bucket>Early|Mid|Late)(?P<month>[A-Za-z]{3})(?P<year>\d{4})",
        file_name,
        flags=re.IGNORECASE,
    )
    if not match:
        return None

    ticker = match.group("ticker").upper()
    bucket = match.group("bucket").title()
    month_abbrev = match.group("month").title()
    year = int(match.group("year"))

    month_lookup = {
        "Jan": 1,
        "Feb": 2,
        "Mar": 3,
        "Apr": 4,
        "May": 5,
        "Jun": 6,
        "Jul": 7,
        "Aug": 8,
        "Sep": 9,
        "Oct": 10,
        "Nov": 11,
        "Dec": 12,
    }

    month = month_lookup.get(month_abbrev)
    if month is None:
        return None

    day_lookup = {"Early": 5, "Mid": 15, "Late": 25}
    day = day_lookup[bucket]

    model_period = f"{bucket}{month_abbrev}_{year}"
    model_date = date(year, month, day).isoformat()
    model = f"{ticker}_{model_period}"

    return FileLabel(
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
        model=model,
    )


def get_next_output_path(input_folder: Path, out_folder: Path) -> Path:
    out_folder.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_folder.name}_PARAM"
    candidate = out_folder / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    index = 1
    while True:
        candidate = out_folder / f"{base_name}.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def close_workbook_safely(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    # Fallbacks for versions where close(save=False) is unsupported.
    try:
        workbook.close(False)
        return
    except Exception:
        pass

    try:
        workbook.api.Close(SaveChanges=False)
    except Exception:
        # Last-resort close; if this fails too, the app quit in main() still
        # ensures no source workbook gets saved.
        workbook.close()


def get_sheet_by_name(workbook: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    for sheet in workbook.sheets:
        if sheet.name.strip().lower() == sheet_name.strip().lower():
            return sheet
    return None


def find_anchor(sheet: xw.Sheet, target_text: str = "max") -> Optional[tuple[int, int]]:
    used = sheet.used_range
    values = used.value

    # Normalize used range to 2D.
    if values is None:
        return None
    if not isinstance(values, list):
        values_2d: list[list[Any]] = [[values]]
    elif values and not isinstance(values[0], list):
        values_2d = [values]
    else:
        values_2d = values

    wanted = normalize_text(target_text)
    for r_idx, row in enumerate(values_2d):
        for c_idx, cell_value in enumerate(row):
            if normalize_text(cell_value) == wanted:
                return used.row + r_idx, used.column + c_idx
    return None


def build_header_offset_map(
    sheet: xw.Sheet,
    header_row: int,
    anchor_col: int,
    span: int = 30,
) -> dict[str, int]:
    min_col = max(1, anchor_col - span)
    max_col = anchor_col + span
    header_map: dict[str, int] = {}
    for col in range(min_col, max_col + 1):
        label = normalize_text(sheet.range((header_row, col)).value)
        if label and label not in header_map:
            header_map[label] = col - anchor_col
    return header_map


def first_matching_offset(
    header_offsets: dict[str, int],
    candidates: list[str],
) -> Optional[int]:
    for candidate in candidates:
        normalized = normalize_text(candidate)
        if normalized in header_offsets:
            return header_offsets[normalized]
    return None


def cell_value_with_offset(
    sheet: xw.Sheet,
    row: int,
    anchor_col: int,
    offset: Optional[int],
) -> Any:
    if offset is None:
        return None
    return sheet.range((row, anchor_col + offset)).value


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
    # Primary: Formula2R1C1 gives dynamic-array aware formula parsing in R1C1 mode.
    try:
        cell.api.Formula2R1C1 = formula_r1c1
        return
    except Exception:
        pass

    # Fallbacks for older Excel/xlwings environments.
    try:
        cell.api.FormulaR1C1 = formula_r1c1
    except Exception:
        cell.formula2 = formula_r1c1


def roughly_equal(a: Any, b: Any, tolerance: float = 1e-9) -> bool:
    a_num = coerce_number(a)
    b_num = coerce_number(b)
    if a_num is None or b_num is None:
        return a == b
    return abs(a_num - b_num) <= tolerance


def should_skip_duplicate_regression_row(
    previous_row: dict[str, Any],
    current_row: dict[str, Any],
) -> bool:
    compare_keys = [
        "num_quarters_used",
        "forecast_value",
        "forecast_max",
        "forecast_min",
        "intercept",
        "slope",
    ]
    for key in compare_keys:
        if not roughly_equal(previous_row.get(key), current_row.get(key)):
            return False
    return True


def extract_empirical_rows(
    workbook: xw.Book,
    meta: FileLabel,
    source_file: str,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    sheet = get_sheet_by_name(workbook, "Empirical Model")
    if sheet is None:
        return rows

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        return rows
    anchor_row, anchor_col = anchor

    header_offsets = build_header_offset_map(sheet, header_row=anchor_row, anchor_col=anchor_col)

    num_quarters_offset = first_matching_offset(
        header_offsets,
        ["num_quarters_used", "num quarters used", "quarters used", "n quarters", "n_quarters"],
    )
    last_quarter_offset = first_matching_offset(
        header_offsets,
        ["last_quarter_used", "last quarter used", "last quarter"],
    )
    forecast_value_offset = first_matching_offset(
        header_offsets,
        ["estimated total sold", "forecast_value", "tot fcst", "tot fcst w/o sa", "forecast"],
    )
    actual_value_offset = first_matching_offset(
        header_offsets,
        ["reported sales", "actual value", "actual sales", "actual"],
    )
    min_offset = first_matching_offset(header_offsets, ["min", "forecast_min"])
    avg_pen_offset = first_matching_offset(
        header_offsets,
        ["avg_penetration_pct", "avg penetration pct", "avg penetration"],
    )
    quarterly_sales_offset = first_matching_offset(
        header_offsets,
        ["quarterly_sales", "quarterly sales", "sales captured", "captured sales"],
    )
    reported_sales_offset = first_matching_offset(
        header_offsets,
        ["reported_sales", "reported sales"],
    )
    growth_rate_offset = first_matching_offset(
        header_offsets,
        ["growth_rate_pct", "growth rate pct", "growth rate"],
    )
    sales_captured_offset = first_matching_offset(
        header_offsets,
        [
            "sales_captured_in_db_pct",
            "sales captured in db pct",
            "captured in db",
            "captured in database pct",
        ],
    )

    if min_offset is None:
        min_offset = 1  # conventional "max" then "min" side-by-side

    # Use temporary R1C1 formulas for avg penetration whenever supporting columns exist.
    helper_col = sheet.used_range.last_cell.column + 1
    formula_rows: list[tuple[int, xw.Range]] = []
    if quarterly_sales_offset is not None and reported_sales_offset is not None:
        q_col = anchor_col + quarterly_sales_offset
        r_col = anchor_col + reported_sales_offset
        for n in range(1, N_QUARTERS + 1):
            row = anchor_row + n
            helper_cell = sheet.range((row, helper_col))
            formula = (
                f'=IFERROR(AVERAGE('
                f'INDEX(C{q_col},COUNTA(C{q_col})-{n}+1):INDEX(C{q_col},COUNTA(C{q_col}))'
                f'/'
                f'INDEX(C{r_col},COUNTA(C{r_col})-{n}+1):INDEX(C{r_col},COUNTA(C{r_col}))'
                f'),"")'
            )
            set_formula2_r1c1(helper_cell, formula)
            formula_rows.append((row, helper_cell))

        if formula_rows:
            workbook.app.calculate()

    for n in range(1, N_QUARTERS + 1):
        row = anchor_row + n
        num_quarters_used = cell_value_with_offset(sheet, row, anchor_col, num_quarters_offset)
        if num_quarters_used in (None, ""):
            num_quarters_used = n

        forecast_max = cell_value_with_offset(sheet, row, anchor_col, 0)
        forecast_min = cell_value_with_offset(sheet, row, anchor_col, min_offset)
        forecast_value = cell_value_with_offset(sheet, row, anchor_col, forecast_value_offset)
        actual_value = cell_value_with_offset(sheet, row, anchor_col, actual_value_offset)
        last_quarter_used = cell_value_with_offset(sheet, row, anchor_col, last_quarter_offset)
        quarterly_sales = cell_value_with_offset(sheet, row, anchor_col, quarterly_sales_offset)
        reported_sales = cell_value_with_offset(sheet, row, anchor_col, reported_sales_offset)
        growth_rate_pct = cell_value_with_offset(sheet, row, anchor_col, growth_rate_offset)
        sales_captured_in_db_pct = cell_value_with_offset(sheet, row, anchor_col, sales_captured_offset)

        avg_penetration_pct = (
            sheet.range((row, helper_col)).value
            if formula_rows
            else cell_value_with_offset(sheet, row, anchor_col, avg_pen_offset)
        )

        max_num = coerce_number(forecast_max)
        min_num = coerce_number(forecast_min)
        range_width = (max_num - min_num) if max_num is not None and min_num is not None else None

        if all(
            value in (None, "")
            for value in (forecast_value, forecast_max, forecast_min, avg_penetration_pct, actual_value)
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
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,  # estimated total sold
                "actual_value": actual_value,  # reported sales
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


def find_last_numeric(sheet: xw.Sheet, col: int, from_row: int, to_row: int = 1) -> Optional[float]:
    for row in range(from_row, to_row - 1, -1):
        value = coerce_number(sheet.range((row, col)).value)
        if value is not None:
            return value
    return None


def extract_regression_rows(
    workbook: xw.Book,
    meta: FileLabel,
    source_file: str,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    sheet = get_sheet_by_name(workbook, "Regression Model")
    if sheet is None:
        return rows

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        return rows
    anchor_row, anchor_col = anchor

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    header_offsets = build_header_offset_map(sheet, header_row=anchor_row, anchor_col=anchor_col)
    num_quarters_offset = first_matching_offset(
        header_offsets,
        ["num_quarters_used", "num quarters used", "quarters used", "n quarters", "n_quarters"],
    )
    forecast_value_offset = first_matching_offset(
        header_offsets,
        ["tot fcst w/o sa", "forecast_value", "tot fcst", "forecast total", "forecast"],
    )
    actual_value_offset = first_matching_offset(
        header_offsets,
        ["actual value", "actual sales", "actual"],
    )
    min_offset = first_matching_offset(header_offsets, ["min", "forecast_min"])
    intercept_offset = first_matching_offset(header_offsets, ["intercept"])
    slope_offset = first_matching_offset(header_offsets, ["slope"])

    if min_offset is None:
        min_offset = 1

    helper_intercept_col = sheet.used_range.last_cell.column + 1
    helper_slope_col = helper_intercept_col + 1
    formula_cells: list[tuple[xw.Range, xw.Range]] = []

    for n in range(1, N_QUARTERS + 1):
        row = anchor_row + n
        intercept_cell = sheet.range((row, helper_intercept_col))
        slope_cell = sheet.range((row, helper_slope_col))

        intercept_formula = (
            f'=IFERROR(INTERCEPT('
            f'INDEX(C{y_col},COUNTA(C{y_col})-{n}+1):INDEX(C{y_col},COUNTA(C{y_col})),'
            f'INDEX(C{x_col},COUNTA(C{x_col})-{n}+1):INDEX(C{x_col},COUNTA(C{x_col}))'
            f'),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE('
            f'INDEX(C{y_col},COUNTA(C{y_col})-{n}+1):INDEX(C{y_col},COUNTA(C{y_col})),'
            f'INDEX(C{x_col},COUNTA(C{x_col})-{n}+1):INDEX(C{x_col},COUNTA(C{x_col}))'
            f'),"")'
        )

        set_formula2_r1c1(intercept_cell, intercept_formula)
        set_formula2_r1c1(slope_cell, slope_formula)
        formula_cells.append((intercept_cell, slope_cell))

    if formula_cells:
        workbook.app.calculate()

    latest_x = find_last_numeric(sheet, x_col, from_row=anchor_row - 1)

    for n in range(1, N_QUARTERS + 1):
        row = anchor_row + n
        num_quarters_used = cell_value_with_offset(sheet, row, anchor_col, num_quarters_offset)
        if num_quarters_used in (None, ""):
            num_quarters_used = n

        forecast_max = cell_value_with_offset(sheet, row, anchor_col, 0)
        forecast_min = cell_value_with_offset(sheet, row, anchor_col, min_offset)
        forecast_value = cell_value_with_offset(sheet, row, anchor_col, forecast_value_offset)
        actual_value = cell_value_with_offset(sheet, row, anchor_col, actual_value_offset)

        intercept = sheet.range((row, helper_intercept_col)).value
        slope = sheet.range((row, helper_slope_col)).value
        if intercept in (None, "") and intercept_offset is not None:
            intercept = cell_value_with_offset(sheet, row, anchor_col, intercept_offset)
        if slope in (None, "") and slope_offset is not None:
            slope = cell_value_with_offset(sheet, row, anchor_col, slope_offset)

        if forecast_value in (None, ""):
            intercept_num = coerce_number(intercept)
            slope_num = coerce_number(slope)
            if intercept_num is not None and slope_num is not None and latest_x is not None:
                forecast_value = intercept_num + slope_num * latest_x

        max_num = coerce_number(forecast_max)
        min_num = coerce_number(forecast_min)
        range_width = (max_num - min_num) if max_num is not None and min_num is not None else None

        if all(value in (None, "") for value in (forecast_value, forecast_max, forecast_min, intercept, slope)):
            continue

        current_row = {
            "model": meta.model,
            "ticker": meta.ticker,
            "model_period": meta.model_period,
            "model_date": meta.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,  # TOT FCST w/o SA
            "actual_value": actual_value if actual_value not in ("", None) else None,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        if rows and should_skip_duplicate_regression_row(rows[-1], current_row):
            continue

        rows.append(current_row)

    return rows


def write_sheet(
    workbook: Workbook,
    sheet_name: str,
    columns: list[str],
    rows: list[dict[str, Any]],
) -> None:
    ws = workbook.create_sheet(title=sheet_name)
    ws.append(columns)
    for row in rows:
        ws.append([row.get(column) for column in columns])

    # Header formatting.
    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"

    if ws.max_row >= 1 and ws.max_column >= 1:
        ws.auto_filter.ref = ws.dimensions

    for col_idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        width = max(12, min(60, max_len + 2))
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = Workbook()
    wb.remove(wb.active)

    write_sheet(wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def iter_source_files(folder: Path) -> list[Path]:
    files: list[Path] = []
    for file_path in sorted(folder.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"Skipping {file_path.name}: temp file")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipping {file_path.name}: not an .xlsx file")
            continue
        files.append(file_path)
    return files


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"Input folder does not exist: {input_dir}")

    output_path = get_next_output_path(input_dir, output_dir)
    files = iter_source_files(input_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.calculation = "manual"

    try:
        for file_path in files:
            meta = parse_file_label(file_path.name)
            if meta is None:
                print(f"Skipping {file_path.name}: filename does not match expected label format")
                continue

            print(f"Processing {file_path.name}")
            workbook: Optional[xw.Book] = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_rows(workbook, meta, file_path.name))
                regression_rows.extend(extract_regression_rows(workbook, meta, file_path.name))
                processed_files += 1
            except Exception as exc:
                print(f"Skipping {file_path.name}: processing error ({exc})")
            finally:
                if workbook is not None:
                    close_workbook_safely(workbook)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output workbook: {output_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

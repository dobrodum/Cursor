from __future__ import annotations

import datetime as dt
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import xlwings as xw


# ----------------------------
# Configure folders here
# ----------------------------
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")


EMPIRICAL_COLUMNS: List[str] = [
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

REGRESSION_COLUMNS: List[str] = [
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


EMPIRICAL_OFFSETS = {
    "num_quarters_used": -8,
    "last_quarter_used": -7,
    "forecast_value": -6,
    "actual_value": -5,
    "avg_penetration_pct": -4,
    "quarterly_sales": -3,
    "reported_sales": -2,
    "growth_rate_pct": -1,
    "forecast_max": 0,
    "forecast_min": 1,
    "sales_captured_in_db_pct": 2,
}

REGRESSION_OFFSETS = {
    "num_quarters_used": -6,
    "forecast_value": -5,
    "actual_value": -4,
    "forecast_max": 0,
    "forecast_min": 1,
}


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

DAY_BY_PERIOD = {
    "early": 5,
    "mid": 15,
    "late": 25,
}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    try:
        return float(text)
    except ValueError:
        return None


def to_int(value: Any) -> Optional[int]:
    number = to_float(value)
    if number is None:
        return None
    return int(round(number))


def parse_file_labels(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = (parts[1] if len(parts) >= 2 else stem).upper()
    period_token_raw = parts[2] if len(parts) >= 3 else ""
    period_token = period_token_raw.split("_")[0].strip()

    match = re.search(
        r"(?i)(early|mid|late)\s*([a-z]{3,9})\s*(20\d{2})",
        period_token,
    )

    model_period = period_token or "unknown_period"
    model_date = ""

    if match:
        period_part = match.group(1).lower()
        month_raw = match.group(2).lower()
        month_abbrev = month_raw[:3]
        year = int(match.group(3))

        if month_abbrev in MONTH_MAP:
            month_num = MONTH_MAP[month_abbrev]
            day = DAY_BY_PERIOD[period_part]
            model_period = f"{period_part.title()}{month_abbrev.title()}_{year}"
            model_date = dt.date(year, month_num, day).isoformat()

    model = f"{ticker}_{model_period}"
    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def next_output_path(input_folder: Path, output_folder: Path) -> Path:
    output_folder.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_folder.name}_PARAM"
    candidate = output_folder / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    index = 1
    while True:
        candidate = output_folder / f"{base_name}.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def safe_close_source_workbook(book: xw.Book) -> None:
    try:
        book.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        book.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        book.close()
    except Exception as close_exc:
        print(f"warning: failed to close workbook safely: {close_exc}")


def find_anchor(sheet: xw.Sheet, target: str = "max") -> Optional[Tuple[int, int]]:
    found = None
    try:
        found = sheet.api.Cells.Find(What=target, MatchCase=False, LookAt=1)
    except Exception:
        found = None

    if found is None:
        try:
            found = sheet.api.Cells.Find(What=target, MatchCase=False, LookAt=2)
        except Exception:
            found = None

    if found is None:
        return None

    return int(found.Row), int(found.Column)


def build_header_map(
    sheet: xw.Sheet,
    header_row: int,
    min_col: int,
    max_col: int,
) -> Dict[int, str]:
    values = sheet.range((header_row, min_col), (header_row, max_col)).value
    if not isinstance(values, list):
        values = [values]

    header_map: Dict[int, str] = {}
    for offset, value in enumerate(values):
        header_map[min_col + offset] = normalize_text(value)
    return header_map


def find_col(
    headers: Dict[int, str],
    anchor_col: int,
    keyword_groups: Sequence[Sequence[str]],
    fallback_offset: Optional[int] = None,
) -> Optional[int]:
    matches: List[int] = []
    for col, header in headers.items():
        for group in keyword_groups:
            if all(keyword in header for keyword in group):
                matches.append(col)
                break

    if matches:
        return min(matches, key=lambda col: abs(col - anchor_col))

    if fallback_offset is not None:
        return anchor_col + fallback_offset
    return None


def cell_value(sheet: xw.Sheet, row: int, col: Optional[int]) -> Any:
    if col is None or col < 1 or row < 1:
        return None
    return sheet.cells(row, col).value


def has_row_data(
    sheet: xw.Sheet,
    row: int,
    cols: Sequence[Optional[int]],
) -> bool:
    for col in cols:
        value = cell_value(sheet, row, col)
        if value not in (None, ""):
            return True
    return False


def extract_empirical_rows(
    workbook: xw.Book,
    sheet: xw.Sheet,
    labels: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"skipped empirical extraction in {source_file} (anchor 'max' not found)")
        return []

    anchor_row, anchor_col = anchor
    used = sheet.used_range
    min_col = int(used.column)
    max_col = int(used.last_cell.column)
    header_map = build_header_map(sheet, anchor_row, min_col, max_col)

    cols = {
        "num_quarters_used": find_col(
            header_map,
            anchor_col,
            keyword_groups=[("num", "quarter"), ("quarters", "used"), ("n", "quarter")],
            fallback_offset=EMPIRICAL_OFFSETS["num_quarters_used"],
        ),
        "last_quarter_used": find_col(
            header_map,
            anchor_col,
            keyword_groups=[("last", "quarter"), ("quarter", "used")],
            fallback_offset=EMPIRICAL_OFFSETS["last_quarter_used"],
        ),
        "forecast_value": find_col(
            header_map,
            anchor_col,
            keyword_groups=[
                ("estimated", "total", "sold"),
                ("forecast", "value"),
                ("forecast",),
            ],
            fallback_offset=EMPIRICAL_OFFSETS["forecast_value"],
        ),
        "actual_value": find_col(
            header_map,
            anchor_col,
            keyword_groups=[("reported", "sales"), ("actual", "value"), ("actual",)],
            fallback_offset=EMPIRICAL_OFFSETS["actual_value"],
        ),
        "forecast_max": anchor_col,
        "forecast_min": find_col(
            header_map,
            anchor_col,
            keyword_groups=[("min",)],
            fallback_offset=EMPIRICAL_OFFSETS["forecast_min"],
        ),
        "avg_penetration_pct": find_col(
            header_map,
            anchor_col,
            keyword_groups=[("avg", "penetration"), ("average", "penetration")],
            fallback_offset=EMPIRICAL_OFFSETS["avg_penetration_pct"],
        ),
        "quarterly_sales": find_col(
            header_map,
            anchor_col,
            keyword_groups=[("quarterly", "sales"), ("quarter", "sales")],
            fallback_offset=EMPIRICAL_OFFSETS["quarterly_sales"],
        ),
        "reported_sales": find_col(
            header_map,
            anchor_col,
            keyword_groups=[("reported", "sales")],
            fallback_offset=EMPIRICAL_OFFSETS["reported_sales"],
        ),
        "growth_rate_pct": find_col(
            header_map,
            anchor_col,
            keyword_groups=[("growth", "rate"), ("growth",)],
            fallback_offset=EMPIRICAL_OFFSETS["growth_rate_pct"],
        ),
        "sales_captured_in_db_pct": find_col(
            header_map,
            anchor_col,
            keyword_groups=[
                ("captured", "db"),
                ("sales", "captured"),
            ],
            fallback_offset=EMPIRICAL_OFFSETS["sales_captured_in_db_pct"],
        ),
    }

    start_row = anchor_row + 1
    candidate_rows = [start_row + i for i in range(10)]
    temp_avg_col = max_col + 2
    formula_rows: List[int] = []

    for row in candidate_rows:
        if not has_row_data(
            sheet,
            row,
            [
                cols["num_quarters_used"],
                cols["forecast_value"],
                cols["actual_value"],
                cols["forecast_max"],
                cols["forecast_min"],
            ],
        ):
            continue

        num_quarters = to_int(cell_value(sheet, row, cols["num_quarters_used"])) or 0
        if num_quarters <= 0:
            continue

        if cols["reported_sales"] and cols["quarterly_sales"]:
            trailing_start = max(start_row, row - num_quarters + 1)
            formula = (
                f'=IFERROR(AVERAGE('
                f'R{trailing_start}C{cols["reported_sales"]}:R{row}C{cols["reported_sales"]}'
                f'/R{trailing_start}C{cols["quarterly_sales"]}:R{row}C{cols["quarterly_sales"]}'
                f'),"")'
            )
            sheet.cells(row, temp_avg_col).formula2 = formula
            formula_rows.append(row)

    if formula_rows:
        workbook.app.calculate()

    rows: List[Dict[str, Any]] = []
    empty_streak = 0
    for idx, row in enumerate(candidate_rows, start=1):
        if not has_row_data(
            sheet,
            row,
            [
                cols["num_quarters_used"],
                cols["forecast_value"],
                cols["actual_value"],
                cols["forecast_max"],
                cols["forecast_min"],
            ],
        ):
            empty_streak += 1
            if empty_streak >= 3:
                break
            continue

        empty_streak = 0
        num_quarters = to_int(cell_value(sheet, row, cols["num_quarters_used"])) or idx
        last_quarter = cell_value(sheet, row, cols["last_quarter_used"])
        forecast_value = cell_value(sheet, row, cols["forecast_value"])
        actual_value = cell_value(sheet, row, cols["actual_value"])
        forecast_max = cell_value(sheet, row, cols["forecast_max"])
        forecast_min = cell_value(sheet, row, cols["forecast_min"])
        quarterly_sales = cell_value(sheet, row, cols["quarterly_sales"])
        reported_sales = cell_value(sheet, row, cols["reported_sales"])
        growth_rate_pct = cell_value(sheet, row, cols["growth_rate_pct"])
        sales_captured_pct = cell_value(sheet, row, cols["sales_captured_in_db_pct"])

        avg_penetration_pct = (
            cell_value(sheet, row, temp_avg_col)
            if row in formula_rows
            else cell_value(sheet, row, cols["avg_penetration_pct"])
        )
        if avg_penetration_pct in (None, ""):
            avg_penetration_pct = cell_value(sheet, row, cols["avg_penetration_pct"])

        max_num = to_float(forecast_max)
        min_num = to_float(forecast_min)
        range_width = (max_num - min_num) if max_num is not None and min_num is not None else None

        rows.append(
            {
                "model": labels["model"],
                "ticker": labels["ticker"],
                "model_period": labels["model_period"],
                "model_date": labels["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters,
                "last_quarter_used": last_quarter,
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

    if formula_rows:
        sheet.range((start_row, temp_avg_col), (start_row + 9, temp_avg_col)).clear_contents()

    return rows


def find_trailing_numeric_rows(sheet: xw.Sheet, anchor_row: int, x_col: int, y_col: int) -> List[int]:
    rows: List[int] = []
    for row in range(anchor_row - 1, 0, -1):
        x_val = to_float(cell_value(sheet, row, x_col))
        y_val = to_float(cell_value(sheet, row, y_col))
        if x_val is not None and y_val is not None:
            rows.append(row)
        elif rows:
            break
    rows.reverse()
    return rows


def extract_regression_rows(
    workbook: xw.Book,
    sheet: xw.Sheet,
    labels: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"skipped regression extraction in {source_file} (anchor 'max' not found)")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    used = sheet.used_range
    min_col = int(used.column)
    max_col = int(used.last_cell.column)
    header_map = build_header_map(sheet, anchor_row, min_col, max_col)

    cols = {
        "num_quarters_used": find_col(
            header_map,
            anchor_col,
            keyword_groups=[("num", "quarter"), ("quarters", "used"), ("n", "quarter")],
            fallback_offset=REGRESSION_OFFSETS["num_quarters_used"],
        ),
        "forecast_value": find_col(
            header_map,
            anchor_col,
            keyword_groups=[
                ("tot", "fcst", "w", "o", "sa"),
                ("forecast", "without", "sa"),
                ("forecast", "value"),
                ("forecast",),
            ],
            fallback_offset=REGRESSION_OFFSETS["forecast_value"],
        ),
        "actual_value": find_col(
            header_map,
            anchor_col,
            keyword_groups=[("actual",), ("reported", "sales")],
            fallback_offset=REGRESSION_OFFSETS["actual_value"],
        ),
        "forecast_max": anchor_col,
        "forecast_min": find_col(
            header_map,
            anchor_col,
            keyword_groups=[("min",)],
            fallback_offset=REGRESSION_OFFSETS["forecast_min"],
        ),
    }

    history_rows = find_trailing_numeric_rows(sheet, anchor_row, x_col, y_col)
    if len(history_rows) < 2:
        print(f"skipped regression extraction in {source_file} (insufficient x/y history)")
        return []

    max_quarters = min(10, len(history_rows))
    candidate_pairs: List[Tuple[int, int]] = []
    for i in range(10):
        row = anchor_row + 1 + i
        num_quarters = to_int(cell_value(sheet, row, cols["num_quarters_used"]))
        if num_quarters is None:
            continue
        if num_quarters < 2 or num_quarters > max_quarters:
            continue
        candidate_pairs.append((row, num_quarters))

    if not candidate_pairs:
        candidate_pairs = [
            (anchor_row + i, num_quarters)
            for i, num_quarters in enumerate(range(2, max_quarters + 1), start=1)
        ]

    # Remove duplicate n values while preserving order.
    seen_n: set[int] = set()
    deduped_pairs: List[Tuple[int, int]] = []
    for row, n_quarters in candidate_pairs:
        if n_quarters in seen_n:
            continue
        seen_n.add(n_quarters)
        deduped_pairs.append((row, n_quarters))

    temp_intercept_col = max_col + 2
    temp_slope_col = max_col + 3
    temp_forecast_col = max_col + 4

    x_forecast = (
        to_float(cell_value(sheet, anchor_row, x_col))
        or to_float(cell_value(sheet, anchor_row + 1, x_col))
        or to_float(cell_value(sheet, history_rows[-1], x_col))
        or 0.0
    )

    calc_plan: List[Tuple[int, int, int, int]] = []
    for calc_idx, (sheet_row, n_quarters) in enumerate(deduped_pairs):
        sample_rows = history_rows[-n_quarters:]
        start_row = sample_rows[0]
        end_row = sample_rows[-1]
        formula_row = anchor_row + calc_idx + 1
        calc_plan.append((sheet_row, n_quarters, start_row, end_row))

        intercept_formula = (
            f'=IFERROR(INTERCEPT('
            f'R{start_row}C{y_col}:R{end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{end_row}C{x_col}'
            f'),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE('
            f'R{start_row}C{y_col}:R{end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{end_row}C{x_col}'
            f'),"")'
        )
        forecast_formula = (
            f'=IFERROR('
            f'R{formula_row}C{temp_intercept_col}'
            f'+R{formula_row}C{temp_slope_col}*{x_forecast}'
            f',"")'
        )

        sheet.cells(formula_row, temp_intercept_col).formula2 = intercept_formula
        sheet.cells(formula_row, temp_slope_col).formula2 = slope_formula
        sheet.cells(formula_row, temp_forecast_col).formula2 = forecast_formula

    if calc_plan:
        workbook.app.calculate()

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for calc_idx, (sheet_row, n_quarters, start_row, end_row) in enumerate(calc_plan):
        formula_row = anchor_row + calc_idx + 1
        intercept = cell_value(sheet, formula_row, temp_intercept_col)
        slope = cell_value(sheet, formula_row, temp_slope_col)
        calc_forecast_value = cell_value(sheet, formula_row, temp_forecast_col)

        sheet_forecast_value = cell_value(sheet, sheet_row, cols["forecast_value"])
        forecast_value = (
            sheet_forecast_value
            if sheet_forecast_value not in (None, "")
            else calc_forecast_value
        )

        sheet_max = cell_value(sheet, sheet_row, cols["forecast_max"])
        sheet_min = cell_value(sheet, sheet_row, cols["forecast_min"])

        if sheet_max in (None, "") or sheet_min in (None, ""):
            y_values = [
                to_float(cell_value(sheet, row, y_col))
                for row in range(start_row, end_row + 1)
            ]
            y_values = [value for value in y_values if value is not None]
            fallback_max = max(y_values) if y_values else None
            fallback_min = min(y_values) if y_values else None
            forecast_max = fallback_max if sheet_max in (None, "") else sheet_max
            forecast_min = fallback_min if sheet_min in (None, "") else sheet_min
        else:
            forecast_max = sheet_max
            forecast_min = sheet_min

        max_num = to_float(forecast_max)
        min_num = to_float(forecast_min)
        range_width = (max_num - min_num) if max_num is not None and min_num is not None else None
        actual_value = cell_value(sheet, sheet_row, cols["actual_value"])

        signature = (
            round(to_float(intercept), 10) if to_float(intercept) is not None else None,
            round(to_float(slope), 10) if to_float(slope) is not None else None,
            round(to_float(forecast_value), 10) if to_float(forecast_value) is not None else None,
            round(to_float(forecast_max), 10) if to_float(forecast_max) is not None else None,
            round(to_float(forecast_min), 10) if to_float(forecast_min) is not None else None,
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": labels["model"],
                "ticker": labels["ticker"],
                "model_period": labels["model_period"],
                "model_date": labels["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n_quarters,
                "num_quarters_used": n_quarters,
                "forecast_value": forecast_value,
                "actual_value": actual_value if actual_value not in (None, "") else "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    if calc_plan:
        first_formula_row = anchor_row + 1
        last_formula_row = anchor_row + len(calc_plan)
        sheet.range(
            (first_formula_row, temp_intercept_col),
            (last_formula_row, temp_forecast_col),
        ).clear_contents()

    return rows


def write_output_sheet(
    wb: openpyxl.Workbook,
    sheet_name: str,
    columns: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    ws = wb.create_sheet(title=sheet_name)
    ws.append(list(columns))

    for item in rows:
        ws.append([item.get(column, "") for column in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, col_name in enumerate(columns, start=1):
        max_length = len(col_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_length = max(max_length, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_length + 2, 48)


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"input_dir does not exist: {input_dir}")
    output_dir.mkdir(parents=True, exist_ok=True)

    output_path = next_output_path(input_dir, output_dir)
    source_files = sorted(input_dir.iterdir())

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    original_calc_mode = app.calculation
    app.calculation = "manual"

    try:
        for file_path in source_files:
            if not file_path.is_file():
                continue
            if file_path.name.startswith("~"):
                print(f"skipped: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped: {file_path.name} (not an .xlsx file)")
                continue

            workbook: Optional[xw.Book] = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                labels = parse_file_labels(file_path)

                empirical_sheet = workbook.sheets["Empirical Model"]
                regression_sheet = workbook.sheets["Regression Model"]

                empirical_rows.extend(
                    extract_empirical_rows(
                        workbook=workbook,
                        sheet=empirical_sheet,
                        labels=labels,
                        source_file=file_path.name,
                    )
                )
                regression_rows.extend(
                    extract_regression_rows(
                        workbook=workbook,
                        sheet=regression_sheet,
                        labels=labels,
                        source_file=file_path.name,
                    )
                )

                processed_files += 1
                print(f"processed: {file_path.name}")
            except KeyError as exc:
                print(f"skipped: {file_path.name} (missing required sheet: {exc})")
            except Exception as exc:
                print(f"skipped: {file_path.name} (error: {exc})")
            finally:
                if workbook is not None:
                    safe_close_source_workbook(workbook)
    finally:
        app.calculation = original_calc_mode
        app.quit()

    output_wb = openpyxl.Workbook()
    default_sheet = output_wb.active
    output_wb.remove(default_sheet)

    write_output_sheet(
        wb=output_wb,
        sheet_name="empirical_candidates",
        columns=EMPIRICAL_COLUMNS,
        rows=empirical_rows,
    )
    write_output_sheet(
        wb=output_wb,
        sheet_name="regression_candidates",
        columns=REGRESSION_COLUMNS,
        rows=regression_rows,
    )

    output_wb.save(output_path)

    print(f"output_path: {output_path}")
    print(f"files_processed: {processed_files}")
    print(f"empirical_rows: {len(empirical_rows)}")
    print(f"regression_rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

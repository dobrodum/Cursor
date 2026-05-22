from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
import re
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# ----------------------------
# Configure paths here
# ----------------------------
input_dir = Path("/workspace/input")
output_dir = Path("/workspace/output")


N_QUARTERS = 10

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

PHASE_DAY = {"early": 5, "mid": 15, "late": 25}
MONTH_TO_NUM = {
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
MONTH_ABBR = {
    1: "Jan",
    2: "Feb",
    3: "Mar",
    4: "Apr",
    5: "May",
    6: "Jun",
    7: "Jul",
    8: "Aug",
    9: "Sep",
    10: "Oct",
    11: "Nov",
    12: "Dec",
}


@dataclass
class SheetCache:
    sheet: xw.Sheet
    start_row: int
    start_col: int
    end_row: int
    end_col: int
    values: list[list[Any]]

    def get(self, row: int, col: int) -> Any:
        r_idx = row - self.start_row
        c_idx = col - self.start_col
        if 0 <= r_idx < len(self.values) and 0 <= c_idx < len(self.values[0]):
            return self.values[r_idx][c_idx]
        return self.sheet.cells(row, col).value


def to_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return [[None]]
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return [[None]]
    if not isinstance(values[0], list):
        return [values]
    return values


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"\s+", " ", text)
    return text


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
    is_pct = "%" in text
    text = text.replace("%", "").replace(",", "")
    try:
        number = float(text)
    except ValueError:
        return None
    if is_pct:
        return number / 100.0
    return number


def as_output_value(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, str):
        return value.strip()
    return value


def build_sheet_cache(sheet: xw.Sheet) -> SheetCache:
    used = sheet.used_range
    data = to_2d(used.value)
    max_cols = max(len(row) for row in data) if data else 1
    if max_cols == 0:
        max_cols = 1
    for row in data:
        if len(row) < max_cols:
            row.extend([None] * (max_cols - len(row)))

    start_row = used.row
    start_col = used.column
    end_row = start_row + len(data) - 1
    end_col = start_col + max_cols - 1
    return SheetCache(
        sheet=sheet,
        start_row=start_row,
        start_col=start_col,
        end_row=end_row,
        end_col=end_col,
        values=data,
    )


def find_anchor_cell(cache: SheetCache, anchor_text: str = "max") -> tuple[int, int] | None:
    target = normalize_text(anchor_text)
    for r_idx, row in enumerate(cache.values):
        for c_idx, value in enumerate(row):
            if normalize_text(value) == target:
                return cache.start_row + r_idx, cache.start_col + c_idx
    return None


def find_column_by_tokens(
    cache: SheetCache,
    header_row: int,
    col_start: int,
    col_end: int,
    token_sets: list[tuple[str, ...]],
) -> int | None:
    left = max(cache.start_col, col_start)
    right = min(cache.end_col, col_end)
    if left > right:
        return None

    for tokens in token_sets:
        for col in range(left, right + 1):
            text = normalize_text(cache.get(header_row, col))
            if text and all(token in text for token in tokens):
                return col
    return None


def resolve_column(
    cache: SheetCache,
    header_row: int,
    anchor_col: int,
    default_offset: int,
    token_sets: list[tuple[str, ...]],
) -> int:
    found = find_column_by_tokens(
        cache=cache,
        header_row=header_row,
        col_start=anchor_col - 30,
        col_end=anchor_col + 30,
        token_sets=token_sets,
    )
    if found is not None:
        return found
    return anchor_col + default_offset


def get_sheet_case_insensitive(wb: xw.Book, name: str) -> xw.Sheet | None:
    wanted = normalize_text(name)
    for sheet in wb.sheets:
        if normalize_text(sheet.name) == wanted:
            return sheet
    return None


def month_token_to_number(token: str) -> int | None:
    key = token.strip().lower()[:3]
    return MONTH_TO_NUM.get(key)


def parse_file_metadata(file_path: Path) -> dict[str, str]:
    stem = file_path.stem
    parts = [p.strip() for p in stem.split(" - ")]
    ticker = "UNKNOWN"
    if len(parts) >= 2 and parts[1]:
        ticker = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper() or "UNKNOWN"
    else:
        ticker_match = re.search(r"-\s*([A-Za-z0-9]+)\s*-", stem)
        if ticker_match:
            ticker = ticker_match.group(1).upper()

    period_match = re.search(r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*[_-]?(\d{4})", stem, re.IGNORECASE)
    if not period_match:
        model_period = "unknown_period"
        model_date = ""
    else:
        phase = period_match.group(1).title()
        month_token = period_match.group(2)
        year = int(period_match.group(3))
        month_num = month_token_to_number(month_token)
        if month_num is None:
            model_period = f"{phase}{month_token}_{year}"
            model_date = ""
        else:
            model_period = f"{phase}{MONTH_ABBR[month_num]}_{year}"
            day = PHASE_DAY[phase.lower()]
            model_date = date(year, month_num, day).isoformat()

    model = f"{ticker}_{model_period}"
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def safe_close_workbook(wb: xw.Book) -> None:
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
        wb.close()
    except Exception:
        pass


def collect_numeric_rows(cache: SheetCache, col: int, row_start: int, row_end: int) -> list[int]:
    rows: list[int] = []
    lower = max(cache.start_row, row_start)
    upper = min(cache.end_row, row_end)
    for row in range(lower, upper + 1):
        if to_number(cache.get(row, col)) is not None:
            rows.append(row)
    return rows


def rows_equal_for_regression(a: dict[str, Any], b: dict[str, Any]) -> bool:
    keys = ["num_quarters_used", "forecast_value", "forecast_max", "forecast_min", "intercept", "slope"]
    for key in keys:
        left = a.get(key)
        right = b.get(key)
        if isinstance(left, (int, float)) and isinstance(right, (int, float)):
            if abs(float(left) - float(right)) > 1e-9:
                return False
        else:
            if left != right:
                return False
    return True


def extract_empirical_rows(
    wb: xw.Book,
    sheet: xw.Sheet,
    metadata: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    cache = build_sheet_cache(sheet)
    anchor = find_anchor_cell(cache, "max")
    if anchor is None:
        print(f"Skipped empirical extraction in {source_file}: 'max' anchor not found.")
        return []

    anchor_row, anchor_col = anchor
    cols = {
        "num_quarters_used": resolve_column(cache, anchor_row, anchor_col, -8, [("num", "quarter"), ("quarters", "used")]),
        "last_quarter_used": resolve_column(cache, anchor_row, anchor_col, -7, [("last", "quarter"), ("quarter", "used")]),
        "sales_captured_in_db_pct": resolve_column(
            cache,
            anchor_row,
            anchor_col,
            -6,
            [("sales", "captured"), ("captured", "db"), ("penetration",)],
        ),
        "growth_rate_pct": resolve_column(cache, anchor_row, anchor_col, -5, [("growth", "rate"), ("growth",)]),
        "avg_penetration_pct": resolve_column(
            cache,
            anchor_row,
            anchor_col,
            -4,
            [("avg", "penetration"), ("average", "penetration"), ("penetration", "avg")],
        ),
        "quarterly_sales": resolve_column(cache, anchor_row, anchor_col, -3, [("quarterly", "sales"), ("qtr", "sales")]),
        "actual_value": resolve_column(cache, anchor_row, anchor_col, -2, [("actual",), ("reported", "sales")]),
        "forecast_value": resolve_column(
            cache,
            anchor_row,
            anchor_col,
            -1,
            [("estimated", "total"), ("forecast", "value"), ("forecast",)],
        ),
        "forecast_max": anchor_col,
        "forecast_min": resolve_column(cache, anchor_row, anchor_col, 1, [("min",)]),
        "reported_sales": resolve_column(cache, anchor_row, anchor_col, -2, [("reported", "sales"), ("actual",)]),
    }

    penetration_col = cols["sales_captured_in_db_pct"]
    penetration_history_rows = collect_numeric_rows(
        cache=cache,
        col=penetration_col,
        row_start=cache.start_row,
        row_end=anchor_row - 1,
    )

    helper_avg_col = cache.end_col + 2
    helper_min_col = cache.end_col + 3
    helper_max_col = cache.end_col + 4
    formula_rows: list[tuple[int, int]] = []
    for n in range(1, min(N_QUARTERS, len(penetration_history_rows)) + 1):
        start_row = penetration_history_rows[-n]
        end_row = penetration_history_rows[-1]
        target_row = anchor_row + n
        avg_formula = f"=AVERAGE(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})"
        min_formula = f"=MIN(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})"
        max_formula = f"=MAX(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})"
        sheet.cells(target_row, helper_avg_col).formula2 = avg_formula
        sheet.cells(target_row, helper_min_col).formula2 = min_formula
        sheet.cells(target_row, helper_max_col).formula2 = max_formula
        formula_rows.append((target_row, n))

    if formula_rows:
        wb.app.calculate()

    rows: list[dict[str, Any]] = []
    for n in range(1, N_QUARTERS + 1):
        row_idx = anchor_row + n

        num_quarters_used = to_number(cache.get(row_idx, cols["num_quarters_used"]))
        if num_quarters_used is None:
            num_quarters_used = float(n)

        avg_penetration = to_number(sheet.cells(row_idx, helper_avg_col).value)
        if avg_penetration is None:
            avg_penetration = to_number(cache.get(row_idx, cols["avg_penetration_pct"]))

        min_penetration = to_number(sheet.cells(row_idx, helper_min_col).value)
        max_penetration = to_number(sheet.cells(row_idx, helper_max_col).value)

        quarterly_sales = to_number(cache.get(row_idx, cols["quarterly_sales"]))
        forecast_value = to_number(cache.get(row_idx, cols["forecast_value"]))
        actual_value = to_number(cache.get(row_idx, cols["actual_value"]))
        reported_sales = to_number(cache.get(row_idx, cols["reported_sales"]))
        if reported_sales is None:
            reported_sales = actual_value

        forecast_max = to_number(cache.get(row_idx, cols["forecast_max"]))
        forecast_min = to_number(cache.get(row_idx, cols["forecast_min"]))

        if forecast_value is None and quarterly_sales is not None and avg_penetration not in (None, 0):
            forecast_value = quarterly_sales / avg_penetration
        if forecast_max is None and quarterly_sales is not None and min_penetration not in (None, 0):
            forecast_max = quarterly_sales / min_penetration
        if forecast_min is None and quarterly_sales is not None and max_penetration not in (None, 0):
            forecast_min = quarterly_sales / max_penetration

        growth_rate_pct = to_number(cache.get(row_idx, cols["growth_rate_pct"]))
        sales_captured_in_db_pct = to_number(cache.get(row_idx, cols["sales_captured_in_db_pct"]))
        last_quarter_used = as_output_value(cache.get(row_idx, cols["last_quarter_used"]))

        range_width: float | None = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        has_signal = any(
            value is not None and value != ""
            for value in [
                avg_penetration,
                quarterly_sales,
                forecast_value,
                forecast_max,
                forecast_min,
                actual_value,
            ]
        )
        if not has_signal and n > len(penetration_history_rows):
            continue

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": int(num_quarters_used),
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
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_rows(
    wb: xw.Book,
    sheet: xw.Sheet,
    metadata: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    cache = build_sheet_cache(sheet)
    anchor = find_anchor_cell(cache, "max")
    if anchor is None:
        print(f"Skipped regression extraction in {source_file}: 'max' anchor not found.")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    cols = {
        "num_quarters_used": resolve_column(cache, anchor_row, anchor_col, -12, [("num", "quarter"), ("quarters", "used")]),
        "forecast_value": resolve_column(
            cache,
            anchor_row,
            anchor_col,
            -1,
            [("tot", "fcst"), ("forecast", "w/o"), ("without", "sa"), ("forecast",)],
        ),
        "actual_value": resolve_column(cache, anchor_row, anchor_col, -2, [("actual",), ("reported", "sales")]),
        "forecast_max": anchor_col,
        "forecast_min": resolve_column(cache, anchor_row, anchor_col, 1, [("min",)]),
    }

    xy_rows: list[tuple[int, float, float]] = []
    for row in range(cache.start_row, anchor_row):
        x_val = to_number(cache.get(row, x_col))
        y_val = to_number(cache.get(row, y_col))
        if x_val is not None and y_val is not None:
            xy_rows.append((row, x_val, y_val))

    if len(xy_rows) < 2:
        print(f"Skipped regression extraction in {source_file}: insufficient x/y points.")
        return []

    helper_intercept_col = cache.end_col + 2
    helper_slope_col = cache.end_col + 3
    helper_forecast_col = cache.end_col + 4

    max_n = min(N_QUARTERS, len(xy_rows))
    formulas_written = False
    for n in range(1, max_n + 1):
        target_row = anchor_row + n
        start_row = xy_rows[-n][0]
        end_row = xy_rows[-1][0]
        next_x = to_number(cache.get(target_row, x_col))
        if next_x is None:
            next_x = xy_rows[-1][1] + 1.0

        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        forecast_formula = f"=R{target_row}C{helper_intercept_col}+R{target_row}C{helper_slope_col}*{next_x}"

        sheet.cells(target_row, helper_intercept_col).formula2 = intercept_formula
        sheet.cells(target_row, helper_slope_col).formula2 = slope_formula
        sheet.cells(target_row, helper_forecast_col).formula2 = forecast_formula
        formulas_written = True

    if formulas_written:
        wb.app.calculate()

    rows: list[dict[str, Any]] = []
    for n in range(1, max_n + 1):
        row_idx = anchor_row + n
        num_quarters_used = to_number(cache.get(row_idx, cols["num_quarters_used"]))
        if num_quarters_used is None:
            num_quarters_used = float(n)

        intercept = to_number(sheet.cells(row_idx, helper_intercept_col).value)
        slope = to_number(sheet.cells(row_idx, helper_slope_col).value)
        forecast_calc = to_number(sheet.cells(row_idx, helper_forecast_col).value)

        forecast_value = to_number(cache.get(row_idx, cols["forecast_value"]))
        if forecast_value is None:
            forecast_value = forecast_calc

        forecast_max = to_number(cache.get(row_idx, cols["forecast_max"]))
        forecast_min = to_number(cache.get(row_idx, cols["forecast_min"]))
        actual_value = to_number(cache.get(row_idx, cols["actual_value"]))

        if forecast_max is None and forecast_value is not None:
            forecast_max = forecast_value
        if forecast_min is None and forecast_value is not None:
            forecast_min = forecast_value

        range_width: float | None = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": int(num_quarters_used),
                "num_quarters_used": int(num_quarters_used),
                "forecast_value": forecast_value,
                "actual_value": actual_value if actual_value is not None else "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    if len(rows) >= 2 and rows_equal_for_regression(rows[-1], rows[-2]):
        rows.pop()

    return rows


def next_output_path(input_folder: Path, output_folder: Path) -> Path:
    folder_name = input_folder.resolve().name
    base_name = f"{folder_name}_PARAM.xlsx"
    candidate = output_folder / base_name
    if not candidate.exists():
        return candidate

    index = 1
    while True:
        candidate = output_folder / f"{folder_name}_PARAM.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def write_sheet(ws: Any, columns: list[str], rows: list[dict[str, Any]]) -> None:
    ws.append(columns)
    for item in rows:
        ws.append([as_output_value(item.get(col)) for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for idx, col_name in enumerate(columns, start=1):
        column_letter = get_column_letter(idx)
        max_len = len(col_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[column_letter].width = min(max(12, max_len + 2), 45)


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


def run() -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    if not input_dir.exists():
        print(f"Skipped run: input_dir does not exist -> {input_dir}")
        return

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_file_count = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.calculation = "manual"

    try:
        for file_path in sorted(input_dir.iterdir(), key=lambda p: p.name.lower()):
            if not file_path.is_file():
                print(f"Skipped {file_path.name}: not a file.")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary Excel file.")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file.")
                continue

            print(f"Processing file: {file_path.name}")
            wb: xw.Book | None = None
            metadata = parse_file_metadata(file_path)
            try:
                wb = app.books.open(str(file_path), update_links=False)

                empirical_sheet = get_sheet_case_insensitive(wb, "Empirical Model")
                if empirical_sheet is None:
                    print(f"Skipped empirical in {file_path.name}: sheet 'Empirical Model' not found.")
                else:
                    empirical_rows.extend(
                        extract_empirical_rows(
                            wb=wb,
                            sheet=empirical_sheet,
                            metadata=metadata,
                            source_file=file_path.name,
                        )
                    )

                regression_sheet = get_sheet_case_insensitive(wb, "Regression Model")
                if regression_sheet is None:
                    print(f"Skipped regression in {file_path.name}: sheet 'Regression Model' not found.")
                else:
                    regression_rows.extend(
                        extract_regression_rows(
                            wb=wb,
                            sheet=regression_sheet,
                            metadata=metadata,
                            source_file=file_path.name,
                        )
                    )

                processed_file_count += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error -> {exc}")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        try:
            app.quit()
        except Exception:
            pass

    output_path = next_output_path(input_dir, output_dir)
    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_file_count}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run()

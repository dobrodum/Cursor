from __future__ import annotations

import math
import re
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# -----------------------------
# Configure paths here
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


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return None
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "")
        if not cleaned:
            return None
        try:
            parsed = float(cleaned)
            if math.isnan(parsed) or math.isinf(parsed):
                return None
            return parsed
        except ValueError:
            return None
    return None


def month_number(month_token: str) -> Optional[int]:
    if not month_token:
        return None
    lookup = {
        "jan": 1,
        "january": 1,
        "feb": 2,
        "february": 2,
        "mar": 3,
        "march": 3,
        "apr": 4,
        "april": 4,
        "may": 5,
        "jun": 6,
        "june": 6,
        "jul": 7,
        "july": 7,
        "aug": 8,
        "august": 8,
        "sep": 9,
        "sept": 9,
        "september": 9,
        "oct": 10,
        "october": 10,
        "nov": 11,
        "november": 11,
        "dec": 12,
        "december": 12,
    }
    return lookup.get(month_token.strip().lower())


def parse_file_labels(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split("-")]

    ticker = "UNKNOWN"
    if len(parts) >= 2 and parts[1]:
        ticker = parts[1].upper()

    period_match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(\d{4})",
        stem,
        flags=re.IGNORECASE,
    )

    model_period = "UNKNOWN_PERIOD"
    model_date = ""
    if period_match:
        part_name = period_match.group(1).title()
        month_text = period_match.group(2).title()
        year_text = period_match.group(3)
        month_abbr = month_text[:3].title()
        model_period = f"{part_name}{month_abbr}_{year_text}"

        day_lookup = {"Early": 5, "Mid": 15, "Late": 25}
        month_num = month_number(month_text)
        if month_num:
            parsed_date = date(int(year_text), month_num, day_lookup[part_name])
            model_date = parsed_date.isoformat()

    model = f"{ticker}_{model_period}" if model_period != "UNKNOWN_PERIOD" else ticker

    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def list_input_files(folder: Path) -> List[Path]:
    return sorted([p for p in folder.iterdir() if p.is_file()], key=lambda p: p.name.lower())


def output_file_path(input_folder: Path, destination: Path) -> Path:
    base_name = f"{input_folder.name}_PARAM"
    candidate = destination / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    counter = 1
    while True:
        candidate = destination / f"{base_name}.{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


def normalize_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def find_anchor_cell(sheet: xw.Sheet, anchor_text: str = "max") -> Optional[tuple[int, int]]:
    used = sheet.used_range
    matrix = normalize_2d(used.value)
    if not matrix:
        return None

    base_row = used.row
    base_col = used.column
    target = anchor_text.strip().lower()

    for row_idx, row in enumerate(matrix):
        for col_idx, value in enumerate(row):
            if isinstance(value, str) and value.strip().lower() == target:
                return base_row + row_idx, base_col + col_idx
    return None


def column_values(values: Any) -> List[Any]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return [row[0] if row else None for row in values]
        return values
    return [values]


def value_label(value: Any, row_number: int) -> str:
    if value is None:
        return f"row_{row_number}"
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
    text = str(value).strip()
    return text if text else f"row_{row_number}"


def collect_history_rows(
    sheet: xw.Sheet,
    anchor_row: int,
    x_col: int,
    y_col: int,
    quarter_col: int,
    max_scan_rows: int = 200,
) -> List[Dict[str, Any]]:
    if anchor_row <= 1 or x_col < 1 or y_col < 1:
        return []

    start_row = max(1, anchor_row - max_scan_rows)
    end_row = anchor_row - 1
    if start_row > end_row:
        return []

    x_vals = column_values(sheet.range((start_row, x_col), (end_row, x_col)).value)
    y_vals = column_values(sheet.range((start_row, y_col), (end_row, y_col)).value)
    q_vals = column_values(sheet.range((start_row, quarter_col), (end_row, quarter_col)).value)

    history: List[Dict[str, Any]] = []
    for i, (x_raw, y_raw, q_raw) in enumerate(zip(x_vals, y_vals, q_vals)):
        x_value = to_float(x_raw)
        y_value = to_float(y_raw)
        if x_value is None or y_value is None:
            continue
        row_no = start_row + i
        history.append(
            {
                "row": row_no,
                "x": x_value,
                "y": y_value,
                "quarter": value_label(q_raw, row_no),
            }
        )
    return history


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
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
        wb.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        wb.close()
    except Exception:
        pass


def get_sheet_case_insensitive(wb: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    wanted = sheet_name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == wanted:
            return sheet
    return None


def process_empirical_sheet(
    wb: xw.Book,
    labels: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = get_sheet_case_insensitive(wb, "Empirical Model")
    if sheet is None:
        print(f"Skipped empirical in {source_file}: 'Empirical Model' sheet not found")
        return []

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"Skipped empirical in {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    x_col = anchor_col - 11
    y_col = anchor_col - 7
    quarter_col = x_col - 1 if x_col > 1 else x_col

    history = collect_history_rows(sheet, anchor_row, x_col, y_col, quarter_col)
    if len(history) < 2:
        print(f"Skipped empirical in {source_file}: not enough historical rows")
        return []

    target = history[-1]
    max_n = min(N_QUARTERS, len(history) - 1)
    scratch_row = anchor_row + 2
    scratch_col = anchor_col + 6

    rows: List[Dict[str, Any]] = []
    for n_quarters in range(1, max_n + 1):
        support = history[-(n_quarters + 1) : -1]
        if not support:
            continue

        first_row = support[0]["row"]
        last_row = support[-1]["row"]

        avg_cell = sheet.range((scratch_row, scratch_col))
        pen_max_cell = sheet.range((scratch_row, scratch_col + 1))
        pen_min_cell = sheet.range((scratch_row, scratch_col + 2))

        set_formula2_r1c1(
            avg_cell,
            f"=AVERAGE(R{first_row}C{x_col}:R{last_row}C{x_col}/R{first_row}C{y_col}:R{last_row}C{y_col})",
        )
        set_formula2_r1c1(
            pen_max_cell,
            f"=MAX(R{first_row}C{x_col}:R{last_row}C{x_col}/R{first_row}C{y_col}:R{last_row}C{y_col})",
        )
        set_formula2_r1c1(
            pen_min_cell,
            f"=MIN(R{first_row}C{x_col}:R{last_row}C{x_col}/R{first_row}C{y_col}:R{last_row}C{y_col})",
        )
        wb.app.calculate()

        avg_penetration = to_float(avg_cell.value)
        penetration_max = to_float(pen_max_cell.value)
        penetration_min = to_float(pen_min_cell.value)

        quarterly_sales = target["x"]
        reported_sales = target["y"]

        forecast_value = (
            quarterly_sales / avg_penetration
            if avg_penetration not in (None, 0)
            else None
        )
        forecast_max = (
            quarterly_sales / penetration_min
            if penetration_min not in (None, 0)
            else None
        )
        forecast_min = (
            quarterly_sales / penetration_max
            if penetration_max not in (None, 0)
            else None
        )
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )
        growth_rate_pct = (
            (forecast_value / reported_sales - 1.0)
            if forecast_value is not None and reported_sales not in (None, 0)
            else None
        )
        sales_captured_pct = (
            quarterly_sales / reported_sales
            if reported_sales not in (None, 0)
            else None
        )

        rows.append(
            {
                "model": labels["model"],
                "ticker": labels["ticker"],
                "model_period": labels["model_period"],
                "model_date": labels["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": n_quarters,
                "last_quarter_used": support[-1]["quarter"],
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_pct,
                "source_file": source_file,
            }
        )

    sheet.range((scratch_row, scratch_col), (scratch_row, scratch_col + 2)).value = [[None, None, None]]
    return rows


def nearly_equal(a: Optional[float], b: Optional[float], tol: float = 1e-9) -> bool:
    if a is None or b is None:
        return a is b
    return abs(a - b) <= tol * max(1.0, abs(a), abs(b))


def is_duplicate_regression_row(prev_row: Dict[str, Any], row: Dict[str, Any]) -> bool:
    checks = [
        nearly_equal(to_float(prev_row.get("forecast_value")), to_float(row.get("forecast_value"))),
        nearly_equal(to_float(prev_row.get("forecast_max")), to_float(row.get("forecast_max"))),
        nearly_equal(to_float(prev_row.get("forecast_min")), to_float(row.get("forecast_min"))),
        nearly_equal(to_float(prev_row.get("intercept")), to_float(row.get("intercept"))),
        nearly_equal(to_float(prev_row.get("slope")), to_float(row.get("slope"))),
    ]
    return all(checks)


def process_regression_sheet(
    wb: xw.Book,
    labels: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = get_sheet_case_insensitive(wb, "Regression Model")
    if sheet is None:
        print(f"Skipped regression in {source_file}: 'Regression Model' sheet not found")
        return []

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"Skipped regression in {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    quarter_col = x_col - 1 if x_col > 1 else x_col

    history = collect_history_rows(sheet, anchor_row, x_col, y_col, quarter_col)
    if len(history) < 3:
        print(f"Skipped regression in {source_file}: not enough historical rows")
        return []

    target = history[-1]
    max_n = min(N_QUARTERS, len(history) - 1)
    scratch_row = anchor_row + 2
    scratch_col = anchor_col + 6
    rows: List[Dict[str, Any]] = []

    for n_quarters in range(2, max_n + 1):
        support = history[-(n_quarters + 1) : -1]
        if len(support) < 2:
            continue

        first_row = support[0]["row"]
        last_row = support[-1]["row"]

        intercept_cell = sheet.range((scratch_row, scratch_col))
        slope_cell = sheet.range((scratch_row, scratch_col + 1))

        set_formula2_r1c1(
            intercept_cell,
            f"=INTERCEPT(R{first_row}C{y_col}:R{last_row}C{y_col},R{first_row}C{x_col}:R{last_row}C{x_col})",
        )
        set_formula2_r1c1(
            slope_cell,
            f"=SLOPE(R{first_row}C{y_col}:R{last_row}C{y_col},R{first_row}C{x_col}:R{last_row}C{x_col})",
        )
        wb.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)
        if intercept is None or slope is None:
            continue

        forecast_total_without_sa = intercept + slope * target["x"]
        actual_value = target["y"] if target.get("y") is not None else None

        residuals = [point["y"] - (intercept + slope * point["x"]) for point in support]
        if len(residuals) >= 2:
            mean_residual = sum(residuals) / len(residuals)
            variance = sum((res - mean_residual) ** 2 for res in residuals) / (len(residuals) - 1)
            stdev = math.sqrt(variance)
        else:
            stdev = 0.0

        forecast_max = forecast_total_without_sa + stdev
        forecast_min = forecast_total_without_sa - stdev
        range_width = forecast_max - forecast_min

        row = {
            "model": labels["model"],
            "ticker": labels["ticker"],
            "model_period": labels["model_period"],
            "model_date": labels["model_date"],
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

        if rows and is_duplicate_regression_row(rows[-1], row):
            continue

        rows.append(row)

    sheet.range((scratch_row, scratch_col), (scratch_row, scratch_col + 1)).value = [[None, None]]
    return rows


def format_sheet(worksheet, headers: List[str], rows: List[Dict[str, Any]]) -> None:
    worksheet.append(headers)
    for row in rows:
        worksheet.append([row.get(col) for col in headers])

    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for idx, header in enumerate(headers, start=1):
        values = [header]
        values.extend(row.get(header) for row in rows)
        max_len = max(len(str(v)) for v in values if v is not None) if values else 10
        worksheet.column_dimensions[get_column_letter(idx)].width = min(max(12, max_len + 2), 40)


def save_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    empirical_ws = wb.create_sheet("empirical_candidates")
    regression_ws = wb.create_sheet("regression_candidates")

    format_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    format_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def configure_excel_app(app: xw.App) -> None:
    app.visible = False
    app.display_alerts = False
    app.screen_updating = False
    try:
        app.api.EnableEvents = False
    except Exception:
        pass
    try:
        # xlCalculationManual
        app.api.Calculation = -4135
    except Exception:
        pass


def main() -> None:
    source_dir = input_dir.expanduser().resolve()
    destination_dir = output_dir.expanduser().resolve()

    if not source_dir.exists() or not source_dir.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {source_dir}")

    destination_dir.mkdir(parents=True, exist_ok=True)
    final_output_path = output_file_path(source_dir, destination_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app: Optional[xw.App] = None
    try:
        app = xw.App(visible=False, add_book=False)
        configure_excel_app(app)

        for file_path in list_input_files(source_dir):
            name = file_path.name

            if name.startswith("~"):
                print(f"Skipped {name}: temporary file")
                continue

            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {name}: not an .xlsx file")
                continue

            wb: Optional[xw.Book] = None
            try:
                print(f"Processing {name}")
                wb = app.books.open(str(file_path), update_links=False)
                labels = parse_file_labels(name)

                empirical_rows.extend(process_empirical_sheet(wb, labels, name))
                regression_rows.extend(process_regression_sheet(wb, labels, name))
                processed_files += 1
            except Exception as exc:
                print(f"Skipped {name}: processing error: {exc}")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass

    save_output_workbook(final_output_path, empirical_rows, regression_rows)

    print(f"Output path: {final_output_path}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# User-editable paths.
input_dir = Path("./input")
output_dir = Path("./output")

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

MONTH_MAP = {
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

DAY_BY_PHASE = {"early": 5, "mid": 15, "late": 25}


def normalize_2d(values: Any) -> list[list[Any]]:
    """Normalize xlwings range output into a 2D list."""
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def textify(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def safe_subtract(a: Any, b: Any) -> Any:
    if is_number(a) and is_number(b):
        return a - b
    return None


def as_int(value: Any, default: int) -> int:
    if value in (None, ""):
        return default
    if isinstance(value, bool):
        return default
    if isinstance(value, (int, float)):
        return int(value)
    text = textify(value)
    match = re.search(r"-?\d+", text)
    if match:
        try:
            return int(match.group(0))
        except ValueError:
            return default
    return default


def get_output_path(src_input_dir: Path, dst_output_dir: Path) -> Path:
    base_name = f"{src_input_dir.name}_PARAM"
    candidate = dst_output_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    i = 1
    while True:
        candidate = dst_output_dir / f"{base_name}.{i}.xlsx"
        if not candidate.exists():
            return candidate
        i += 1


def parse_model_labels(file_path: Path) -> dict[str, str]:
    """Parse ticker/model labels from names like:
    MedMiner_Model - AORT - MidJan2026_Send.xlsx
    """
    stem = file_path.stem
    parts = [p.strip() for p in stem.split(" - ")]

    ticker = parts[1] if len(parts) >= 2 else ""
    raw_period = parts[2] if len(parts) >= 3 else stem
    raw_period = re.sub(r"(?i)[\s_-]*send$", "", raw_period).strip(" _-")

    period_match = re.search(
        r"(?i)\b(Early|Mid|Late)\s*[_ -]*([A-Za-z]{3,9})\s*[_ -]*(\d{4})\b",
        raw_period,
    )

    model_period = ""
    model_date = ""
    if period_match:
        phase_raw, month_raw, year_raw = period_match.groups()
        phase = phase_raw.lower()
        month_num = MONTH_MAP.get(month_raw.lower())
        if month_num:
            month_abbr = date(2000, month_num, 1).strftime("%b")
            model_period = f"{phase_raw.title()}{month_abbr}_{year_raw}"
            model_date = date(
                int(year_raw), month_num, DAY_BY_PHASE.get(phase, 15)
            ).isoformat()

    if not model_period:
        normalized = re.sub(r"\s+", "_", raw_period).strip("_")
        model_period = normalized

    model = f"{ticker}_{model_period}" if ticker and model_period else stem
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def close_source_wb_no_save(wb: xw.Book) -> None:
    """Close source workbook without saving, with safe fallbacks."""
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

    # Last fallback; app has display alerts off, so this is still safe.
    try:
        wb.close()
    except Exception:
        pass


def find_max_anchor(sheet: xw.Sheet) -> tuple[int, int] | None:
    used = sheet.used_range
    values = normalize_2d(used.value)
    if not values:
        return None

    base_row = used.row
    base_col = used.column
    candidates: list[tuple[int, int, int]] = []

    for r_i, row in enumerate(values):
        for c_i, cell in enumerate(row):
            if textify(cell).lower() != "max":
                continue

            row_has_min_nearby = False
            for delta in range(1, 6):
                j = c_i + delta
                if j < len(row) and textify(row[j]).lower() == "min":
                    row_has_min_nearby = True
                    break

            score = 10 if row_has_min_nearby else 0
            if r_i + 1 < len(values):
                below = values[r_i + 1][c_i] if c_i < len(values[r_i + 1]) else None
                if is_number(below):
                    score += 1

            candidates.append((score, base_row + r_i, base_col + c_i))

    if not candidates:
        return None

    candidates.sort(key=lambda x: (-x[0], x[1], x[2]))
    _, row, col = candidates[0]
    return row, col


def find_column_by_patterns(
    row_values: list[Any], patterns: list[str], base_col: int
) -> int | None:
    for idx, cell in enumerate(row_values):
        text = textify(cell).lower()
        if not text:
            continue
        if any(re.search(pattern, text, flags=re.IGNORECASE) for pattern in patterns):
            return base_col + idx
    return None


def set_formula2(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def get_row_values(sheet: xw.Sheet, row_num: int, start_col: int, end_col: int) -> list[Any]:
    if end_col < start_col:
        return []
    values = sheet.range((row_num, start_col), (row_num, end_col)).value
    if values is None:
        return []
    if not isinstance(values, list):
        return [values]
    return values


def build_empirical_column_map(
    sheet: xw.Sheet, anchor_row: int, anchor_col: int
) -> dict[str, int]:
    row_values = get_row_values(sheet, anchor_row, max(1, anchor_col - 30), anchor_col + 30)
    start_col = max(1, anchor_col - 30)

    col_map: dict[str, int] = {}
    col_map["forecast_max"] = anchor_col
    col_map["forecast_min"] = find_column_by_patterns(row_values, [r"^\s*min\s*$"], start_col)
    col_map["forecast_value"] = find_column_by_patterns(
        row_values,
        [r"estimated\s*total\s*sold", r"\bforecast\b", r"\btot\s*fcst\b"],
        start_col,
    )
    col_map["actual_value"] = find_column_by_patterns(
        row_values,
        [r"reported\s*sales", r"actual\s*sales", r"^\s*actual\s*$"],
        start_col,
    )
    col_map["num_quarters_used"] = find_column_by_patterns(
        row_values,
        [r"num.*quarter", r"n.*quarter", r"quarters?\s*used"],
        start_col,
    )
    col_map["last_quarter_used"] = find_column_by_patterns(
        row_values,
        [r"last.*quarter", r"quarter\s*end"],
        start_col,
    )
    col_map["avg_penetration_pct"] = find_column_by_patterns(
        row_values,
        [r"avg.*penetration", r"average.*penetration", r"penetration.*avg"],
        start_col,
    )
    col_map["quarterly_sales"] = find_column_by_patterns(
        row_values,
        [r"quarterly\s*sales", r"qtr.*sales", r"quarter.*sold"],
        start_col,
    )
    col_map["reported_sales"] = find_column_by_patterns(
        row_values,
        [r"reported\s*sales"],
        start_col,
    )
    col_map["growth_rate_pct"] = find_column_by_patterns(
        row_values,
        [r"growth\s*rate", r"growth\s*%"],
        start_col,
    )
    col_map["sales_captured_in_db_pct"] = find_column_by_patterns(
        row_values,
        [r"sales.*captured.*db", r"captured.*db", r"db\s*%"],
        start_col,
    )

    # Offset-based defaults from max anchor if headers are missing.
    defaults = {
        "forecast_min": anchor_col + 1,
        "forecast_value": anchor_col - 1,
        "actual_value": anchor_col - 2,
        "num_quarters_used": anchor_col - 8,
        "last_quarter_used": anchor_col - 7,
        "avg_penetration_pct": anchor_col - 4,
        "quarterly_sales": anchor_col - 6,
        "reported_sales": anchor_col - 2,
        "growth_rate_pct": anchor_col - 3,
        "sales_captured_in_db_pct": anchor_col - 5,
    }
    for key, fallback_col in defaults.items():
        if col_map.get(key) is None:
            col_map[key] = max(1, fallback_col)

    return col_map


def build_regression_column_map(
    sheet: xw.Sheet, anchor_row: int, anchor_col: int
) -> dict[str, int]:
    row_values = get_row_values(sheet, anchor_row, max(1, anchor_col - 30), anchor_col + 30)
    start_col = max(1, anchor_col - 30)

    col_map: dict[str, int] = {}
    col_map["forecast_max"] = anchor_col
    col_map["forecast_min"] = find_column_by_patterns(row_values, [r"^\s*min\s*$"], start_col)
    col_map["num_quarters_used"] = find_column_by_patterns(
        row_values,
        [r"num.*quarter", r"n.*quarter", r"quarters?\s*used"],
        start_col,
    )
    col_map["forecast_value"] = find_column_by_patterns(
        row_values,
        [r"tot.*fcst.*w\/?\s*o.*sa", r"forecast.*without.*sa", r"tot.*fcst"],
        start_col,
    )
    col_map["actual_value"] = find_column_by_patterns(
        row_values,
        [r"actual", r"reported\s*sales"],
        start_col,
    )

    defaults = {
        "forecast_min": anchor_col + 1,
        "num_quarters_used": anchor_col - 8,
        "forecast_value": anchor_col - 1,
    }
    for key, fallback_col in defaults.items():
        if col_map.get(key) is None:
            col_map[key] = max(1, fallback_col)

    return col_map


def read_cell(sheet: xw.Sheet, row_num: int, col_num: int | None) -> Any:
    if col_num is None or col_num < 1 or row_num < 1:
        return None
    return sheet.range((row_num, col_num)).value


def extract_empirical_rows(
    wb: xw.Book,
    labels: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"  skipped empirical sheet: missing 'Empirical Model' in {source_file}")
        return []

    anchor = find_max_anchor(sheet)
    if anchor is None:
        print(f"  skipped empirical sheet: 'max' anchor not found in {source_file}")
        return []

    anchor_row, anchor_col = anchor
    col_map = build_empirical_column_map(sheet, anchor_row, anchor_col)
    helper_cell = sheet.range((anchor_row, anchor_col + 60))

    rows: list[dict[str, Any]] = []
    for n in range(1, N_QUARTERS + 1):
        row_num = anchor_row + n
        num_quarters_used = read_cell(sheet, row_num, col_map["num_quarters_used"])
        if num_quarters_used in (None, ""):
            num_quarters_used = n
        quarters_count = max(1, as_int(num_quarters_used, n))

        last_quarter_used = read_cell(sheet, row_num, col_map["last_quarter_used"])
        forecast_value = read_cell(sheet, row_num, col_map["forecast_value"])
        actual_value = read_cell(sheet, row_num, col_map["actual_value"])
        forecast_max = read_cell(sheet, row_num, col_map["forecast_max"])
        forecast_min = read_cell(sheet, row_num, col_map["forecast_min"])
        avg_penetration_pct = read_cell(sheet, row_num, col_map["avg_penetration_pct"])
        quarterly_sales = read_cell(sheet, row_num, col_map["quarterly_sales"])
        reported_sales = read_cell(sheet, row_num, col_map["reported_sales"])
        growth_rate_pct = read_cell(sheet, row_num, col_map["growth_rate_pct"])
        sales_captured_in_db_pct = read_cell(
            sheet, row_num, col_map["sales_captured_in_db_pct"]
        )

        # Keep the existing "formula-driven" style by calculating avg penetration
        # via temporary R1C1 formula updates against the nearby penetration column.
        penetration_col = col_map["avg_penetration_pct"]
        formula_start = max(anchor_row + 1, row_num - quarters_count + 1)
        formula_end = row_num
        avg_formula = (
            f"=AVERAGE(R{formula_start}C{penetration_col}:R{formula_end}C{penetration_col})"
        )
        set_formula2(helper_cell, avg_formula)
        wb.app.calculate()
        avg_penetration_calc = helper_cell.value

        if avg_penetration_pct in (None, ""):
            avg_penetration_pct = avg_penetration_calc

        range_width = safe_subtract(forecast_max, forecast_min)

        signal_values = [
            forecast_value,
            forecast_max,
            forecast_min,
            avg_penetration_pct,
            quarterly_sales,
        ]
        if all(v in (None, "") for v in signal_values):
            if n == 1:
                continue
            break

        row = {
            "model": labels["model"],
            "ticker": labels["ticker"],
            "model_period": labels["model_period"],
            "model_date": labels["model_date"],
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
            "sales_captured_in_db_pct": sales_captured_in_db_pct,
            "source_file": source_file,
        }
        rows.append(row)

    helper_cell.value = None
    return rows


def extract_regression_rows(
    wb: xw.Book,
    labels: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"  skipped regression sheet: missing 'Regression Model' in {source_file}")
        return []

    anchor = find_max_anchor(sheet)
    if anchor is None:
        print(f"  skipped regression sheet: 'max' anchor not found in {source_file}")
        return []

    anchor_row, anchor_col = anchor
    col_map = build_regression_column_map(sheet, anchor_row, anchor_col)

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    intercept_cell = sheet.range((anchor_row, anchor_col + 60))
    slope_cell = sheet.range((anchor_row, anchor_col + 61))

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None

    data_end_row = anchor_row - 1
    if data_end_row < 1:
        return []

    for n in range(1, N_QUARTERS + 1):
        data_start_row = max(1, data_end_row - n + 1)

        intercept_formula = (
            f"=INTERCEPT(R{data_start_row}C{y_col}:R{data_end_row}C{y_col},"
            f"R{data_start_row}C{x_col}:R{data_end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{data_start_row}C{y_col}:R{data_end_row}C{y_col},"
            f"R{data_start_row}C{x_col}:R{data_end_row}C{x_col})"
        )
        set_formula2(intercept_cell, intercept_formula)
        set_formula2(slope_cell, slope_formula)
        wb.app.calculate()

        intercept = intercept_cell.value
        slope = slope_cell.value

        row_num = anchor_row + n
        num_quarters_used = read_cell(sheet, row_num, col_map["num_quarters_used"])
        if num_quarters_used in (None, ""):
            num_quarters_used = n

        forecast_value = read_cell(sheet, row_num, col_map["forecast_value"])
        if forecast_value in (None, "") and is_number(intercept) and is_number(slope):
            x_next = read_cell(sheet, row_num, x_col)
            if x_next in (None, ""):
                x_next = read_cell(sheet, data_end_row + 1, x_col)
            if is_number(x_next):
                forecast_value = intercept + slope * x_next

        actual_value = read_cell(sheet, row_num, col_map.get("actual_value"))
        forecast_max = read_cell(sheet, row_num, col_map["forecast_max"])
        forecast_min = read_cell(sheet, row_num, col_map["forecast_min"])
        range_width = safe_subtract(forecast_max, forecast_min)

        signal_values = [forecast_value, forecast_max, forecast_min, intercept, slope]
        if all(v in (None, "") for v in signal_values):
            if n == 1:
                continue
            break

        current_signature = (
            num_quarters_used,
            forecast_value,
            forecast_max,
            forecast_min,
            intercept,
            slope,
        )

        # Existing regression logic can occasionally emit a duplicated terminal row.
        if n == N_QUARTERS and previous_signature == current_signature:
            continue

        row = {
            "model": labels["model"],
            "ticker": labels["ticker"],
            "model_period": labels["model_period"],
            "model_date": labels["model_date"],
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
            "source_file": source_file,
        }
        rows.append(row)
        previous_signature = current_signature

    intercept_cell.value = None
    slope_cell.value = None
    return rows


def write_output_sheet(ws, columns: list[str], rows: list[dict[str, Any]]) -> None:
    ws.append(columns)
    for col_idx in range(1, len(columns) + 1):
        ws.cell(row=1, column=col_idx).font = Font(bold=True)

    for row in rows:
        ws.append([row.get(col) for col in columns])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(1, ws.max_row)}"

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row_idx in range(2, ws.max_row + 1):
            cell_value = ws.cell(row=row_idx, column=col_idx).value
            if cell_value is None:
                continue
            max_len = max(max_len, len(str(cell_value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 42)


def iter_source_files(src_dir: Path) -> list[Path]:
    files: list[Path] = []
    for file_path in sorted(src_dir.iterdir()):
        if not file_path.is_file():
            print(f"skipped: {file_path.name} (not a file)")
            continue
        if file_path.name.startswith("~"):
            print(f"skipped: {file_path.name} (temp file)")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"skipped: {file_path.name} (not .xlsx)")
            continue
        files.append(file_path)
    return files


def main() -> None:
    src_dir = input_dir.expanduser().resolve()
    dst_dir = output_dir.expanduser().resolve()

    if not src_dir.exists():
        raise FileNotFoundError(f"Input folder not found: {src_dir}")

    dst_dir.mkdir(parents=True, exist_ok=True)
    output_path = get_output_path(src_dir, dst_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.calculation = "manual"

    try:
        for file_path in iter_source_files(src_dir):
            print(f"processing file: {file_path.name}")
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                labels = parse_model_labels(file_path)

                empirical = extract_empirical_rows(wb, labels, file_path.name)
                regression = extract_regression_rows(wb, labels, file_path.name)
                empirical_rows.extend(empirical)
                regression_rows.extend(regression)
                processed_files += 1
                print(f"processed file: {file_path.name}")
            except Exception as exc:
                print(f"  skipped: {file_path.name} (read error: {exc})")
            finally:
                if wb is not None:
                    close_source_wb_no_save(wb)
    finally:
        app.quit()

    out_wb = Workbook()
    first_ws = out_wb.active
    first_ws.title = "empirical_candidates"
    write_output_sheet(first_ws, EMPIRICAL_COLUMNS, empirical_rows)

    regression_ws = out_wb.create_sheet("regression_candidates")
    write_output_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    out_wb.save(output_path)

    print(f"output path: {output_path}")
    print(f"number of files processed: {processed_files}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

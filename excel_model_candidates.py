from __future__ import annotations

import datetime as dt
import math
import re
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# -----------------------------
# User-configurable directories
# -----------------------------
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


N_QUARTERS = 10


EARLY_MID_LATE_DAY = {"early": 5, "mid": 15, "late": 25}
MONTH_LOOKUP = {
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


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def clean_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, float) and math.isnan(value):
        return None
    if isinstance(value, str):
        trimmed = value.strip()
        return trimmed if trimmed else None
    return value


def to_float(value: Any) -> float | None:
    value = clean_value(value)
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(",", ""))
    except (TypeError, ValueError):
        return None


def to_int_if_whole(value: Any) -> Any:
    value = clean_value(value)
    if isinstance(value, float) and value.is_integer():
        return int(value)
    if isinstance(value, int):
        return value
    return value


def compute_range_width(max_value: Any, min_value: Any) -> float | None:
    max_f = to_float(max_value)
    min_f = to_float(min_value)
    if max_f is None or min_f is None:
        return None
    return max_f - min_f


def ensure_matrix(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, (list, tuple)):
        return [[values]]

    if len(values) == 0:
        return []

    first = values[0]
    if isinstance(first, (list, tuple)):
        return [list(row) for row in values]

    return [list(values)]


def ensure_column_list(values: Any, expected_len: int) -> list[Any]:
    if expected_len <= 0:
        return []

    if values is None:
        return [None] * expected_len

    if isinstance(values, (list, tuple)):
        if len(values) == 0:
            return [None] * expected_len
        if isinstance(values[0], (list, tuple)):
            flattened = [row[0] if row else None for row in values]
        else:
            flattened = list(values)
    else:
        flattened = [values]

    flattened = [clean_value(v) for v in flattened]

    if len(flattened) < expected_len:
        flattened.extend([None] * (expected_len - len(flattened)))
    elif len(flattened) > expected_len:
        flattened = flattened[:expected_len]

    return flattened


def parse_filename_labels(file_name: str) -> dict[str, str]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = parts[1] if len(parts) >= 2 else ""
    raw_period = parts[2] if len(parts) >= 3 else ""
    raw_period = raw_period.split("_")[0].strip()
    compact_period = re.sub(r"[^A-Za-z0-9]", "", raw_period)

    model_period = compact_period
    model_date = ""

    match = re.match(
        r"^(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})$",
        compact_period,
        flags=re.IGNORECASE,
    )
    if match:
        phase = match.group(1).lower()
        month_token = match.group(2).lower()
        year = int(match.group(3))

        month = MONTH_LOOKUP.get(month_token)
        if month is None:
            month = MONTH_LOOKUP.get(month_token[:3], 0)

        if month:
            month_abbr = dt.date(year, month, 1).strftime("%b")
            model_period = f"{phase.capitalize()}{month_abbr}_{year}"
            model_date = dt.date(year, month, EARLY_MID_LATE_DAY[phase]).isoformat()

    model = f"{ticker}_{model_period}" if ticker and model_period else ticker or model_period

    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def next_output_path(src_input_dir: Path, src_output_dir: Path) -> Path:
    base_name = f"{src_input_dir.name}_PARAM"
    candidate = src_output_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = src_output_dir / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def close_workbook_no_save(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.close(False)
        return
    except Exception:
        pass

    try:
        workbook.api.Close(SaveChanges=False)
    except Exception:
        pass


def set_formula2(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        # Fallback if Formula2 is unavailable in this Excel runtime.
        cell.formula = formula_r1c1


def get_sheet(wb: xw.Book, sheet_name: str) -> xw.Sheet | None:
    try:
        return wb.sheets[sheet_name]
    except Exception:
        return None


def snapshot_sheet(sheet: xw.Sheet) -> dict[str, Any]:
    used = sheet.used_range
    matrix = ensure_matrix(used.value)
    top_row = used.row
    left_col = used.column
    width = max((len(row) for row in matrix), default=0)
    height = len(matrix)
    right_col = left_col + width - 1 if width else left_col
    bottom_row = top_row + height - 1 if height else top_row
    return {
        "matrix": matrix,
        "top_row": top_row,
        "left_col": left_col,
        "right_col": right_col,
        "bottom_row": bottom_row,
        "width": width,
        "height": height,
    }


def find_max_anchor(snapshot: dict[str, Any]) -> tuple[int, int]:
    matrix = snapshot["matrix"]
    top_row = snapshot["top_row"]
    left_col = snapshot["left_col"]
    for r_idx, row in enumerate(matrix):
        for c_idx, value in enumerate(row):
            if normalize_text(value) == "max":
                return top_row + r_idx, left_col + c_idx
    raise ValueError("Could not find 'max' anchor.")


def get_snapshot_row(snapshot: dict[str, Any], row_number: int) -> list[Any]:
    matrix = snapshot["matrix"]
    top_row = snapshot["top_row"]
    idx = row_number - top_row
    if idx < 0 or idx >= len(matrix):
        return []
    return matrix[idx]


def build_header_map(snapshot: dict[str, Any], header_row: int) -> list[tuple[int, str]]:
    left_col = snapshot["left_col"]
    row_values = get_snapshot_row(snapshot, header_row)
    headers: list[tuple[int, str]] = []
    for offset, value in enumerate(row_values):
        text = normalize_text(value)
        if text:
            headers.append((left_col + offset, text))
    return headers


def find_column(
    headers: list[tuple[int, str]],
    candidates: list[tuple[tuple[str, ...], tuple[str, ...]]],
    fallback: int | None = None,
) -> int | None:
    for includes, excludes in candidates:
        for col, label in headers:
            if all(token in label for token in includes) and not any(
                token in label for token in excludes
            ):
                return col
    return fallback


def read_columns(
    sheet: xw.Sheet,
    start_row: int,
    end_row: int,
    columns: dict[str, int | None],
) -> dict[str, list[Any]]:
    row_count = max(0, end_row - start_row + 1)
    result: dict[str, list[Any]] = {}
    for key, col in columns.items():
        if col is None or row_count == 0:
            result[key] = [None] * row_count
            continue
        values = sheet.range((start_row, col), (end_row, col)).value
        result[key] = ensure_column_list(values, row_count)
    return result


def row_has_any_value(values: list[Any]) -> bool:
    for value in values:
        if clean_value(value) is not None:
            return True
    return False


def extract_empirical_rows(
    wb: xw.Book,
    file_name: str,
    labels: dict[str, str],
) -> list[dict[str, Any]]:
    sheet = get_sheet(wb, "Empirical Model")
    if sheet is None:
        return []

    snapshot = snapshot_sheet(sheet)
    anchor_row, anchor_col = find_max_anchor(snapshot)
    headers = build_header_map(snapshot, anchor_row)

    min_col = find_column(
        headers,
        candidates=[(("min",), tuple())],
        fallback=anchor_col + 1,
    )
    num_col = find_column(
        headers,
        candidates=[
            (("num", "quarter"), tuple()),
            (("quarters", "used"), tuple()),
        ],
        fallback=anchor_col - 6,
    )
    last_quarter_col = find_column(
        headers,
        candidates=[(("last", "quarter"), tuple())],
        fallback=anchor_col - 5,
    )
    forecast_col = find_column(
        headers,
        candidates=[
            (("estimated", "total", "sold"), tuple()),
            (("tot", "fcst"), ("max", "min")),
            (("forecast",), ("max", "min")),
        ],
        fallback=anchor_col - 3,
    )
    actual_col = find_column(
        headers,
        candidates=[
            (("reported", "sales"), tuple()),
            (("actual",), tuple()),
        ],
        fallback=anchor_col - 2,
    )
    avg_pen_col = find_column(
        headers,
        candidates=[(("avg", "penetration"), tuple())],
        fallback=anchor_col - 7,
    )
    quarterly_sales_col = find_column(
        headers,
        candidates=[(("quarterly", "sales"), tuple())],
        fallback=anchor_col - 9,
    )
    growth_rate_col = find_column(
        headers,
        candidates=[(("growth",), tuple())],
        fallback=anchor_col - 8,
    )
    sales_captured_col = find_column(
        headers,
        candidates=[
            (("captured", "db"), tuple()),
            (("captured",), tuple()),
        ],
        fallback=anchor_col - 6,
    )
    penetration_history_col = find_column(
        headers,
        candidates=[
            (("penetration",), ("avg",)),
            (("pen",), ("avg",)),
        ],
        fallback=avg_pen_col,
    )

    start_row = anchor_row + 1
    end_row = anchor_row + N_QUARTERS
    row_count = max(0, end_row - start_row + 1)

    scratch_col = max(snapshot["right_col"] + 2, anchor_col + 2)
    wrote_formulas = False
    if penetration_history_col is not None and penetration_history_col > 0:
        for idx in range(row_count):
            n_used = idx + 1
            source_start = anchor_row - n_used
            source_end = anchor_row - 1
            if source_start < 1 or source_end < 1:
                continue
            formula_cell = sheet.cells(start_row + idx, scratch_col)
            formula = (
                f"=AVERAGE(R{source_start}C{penetration_history_col}:"
                f"R{source_end}C{penetration_history_col})"
            )
            set_formula2(formula_cell, formula)
            wrote_formulas = True

    if wrote_formulas:
        wb.app.calculate()

    scratch_values = ensure_column_list(
        sheet.range((start_row, scratch_col), (end_row, scratch_col)).value,
        row_count,
    )

    col_values = read_columns(
        sheet,
        start_row=start_row,
        end_row=end_row,
        columns={
            "num_quarters_used": num_col,
            "last_quarter_used": last_quarter_col,
            "forecast_value": forecast_col,
            "actual_value": actual_col,
            "forecast_max": anchor_col,
            "forecast_min": min_col,
            "avg_penetration_pct": avg_pen_col,
            "quarterly_sales": quarterly_sales_col,
            "reported_sales": actual_col,
            "growth_rate_pct": growth_rate_col,
            "sales_captured_in_db_pct": sales_captured_col,
        },
    )

    rows: list[dict[str, Any]] = []
    for idx in range(row_count):
        avg_pen = clean_value(col_values["avg_penetration_pct"][idx])
        if avg_pen is None:
            avg_pen = clean_value(scratch_values[idx])

        row_data = {
            "model": labels["model"],
            "ticker": labels["ticker"],
            "model_period": labels["model_period"],
            "model_date": labels["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_pen,
            "num_quarters_used": to_int_if_whole(col_values["num_quarters_used"][idx])
            or (idx + 1),
            "last_quarter_used": clean_value(col_values["last_quarter_used"][idx]),
            "forecast_value": clean_value(col_values["forecast_value"][idx]),
            "actual_value": clean_value(col_values["actual_value"][idx]),
            "forecast_max": clean_value(col_values["forecast_max"][idx]),
            "forecast_min": clean_value(col_values["forecast_min"][idx]),
            "range_width": compute_range_width(
                col_values["forecast_max"][idx], col_values["forecast_min"][idx]
            ),
            "avg_penetration_pct": avg_pen,
            "quarterly_sales": clean_value(col_values["quarterly_sales"][idx]),
            "reported_sales": clean_value(col_values["reported_sales"][idx]),
            "growth_rate_pct": clean_value(col_values["growth_rate_pct"][idx]),
            "sales_captured_in_db_pct": clean_value(
                col_values["sales_captured_in_db_pct"][idx]
            ),
            "source_file": file_name,
        }

        significant_values = [
            row_data["forecast_value"],
            row_data["forecast_max"],
            row_data["forecast_min"],
            row_data["parameter_value"],
        ]
        if row_has_any_value(significant_values):
            rows.append(row_data)

    return rows


def signature_for_regression_row(row_data: dict[str, Any]) -> tuple[Any, ...]:
    keys = [
        "num_quarters_used",
        "forecast_value",
        "forecast_max",
        "forecast_min",
        "intercept",
        "slope",
    ]
    signature: list[Any] = []
    for key in keys:
        value = clean_value(row_data.get(key))
        num = to_float(value)
        if num is None:
            signature.append(value)
        else:
            signature.append(round(num, 8))
    return tuple(signature)


def extract_regression_rows(
    wb: xw.Book,
    file_name: str,
    labels: dict[str, str],
) -> list[dict[str, Any]]:
    sheet = get_sheet(wb, "Regression Model")
    if sheet is None:
        return []

    snapshot = snapshot_sheet(sheet)
    anchor_row, anchor_col = find_max_anchor(snapshot)
    headers = build_header_map(snapshot, anchor_row)

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    min_col = find_column(
        headers,
        candidates=[(("min",), tuple())],
        fallback=anchor_col + 1,
    )
    num_col = find_column(
        headers,
        candidates=[
            (("num", "quarter"), tuple()),
            (("quarters", "used"), tuple()),
        ],
        fallback=anchor_col - 6,
    )
    forecast_col = find_column(
        headers,
        candidates=[
            (("tot", "fcst", "w/o", "sa"), tuple()),
            (("tot", "fcst"), ("max", "min")),
            (("forecast",), ("max", "min")),
        ],
        fallback=anchor_col - 2,
    )
    actual_col = find_column(
        headers,
        candidates=[(("actual",), tuple())],
        fallback=None,
    )
    intercept_col = find_column(
        headers,
        candidates=[(("intercept",), tuple())],
        fallback=None,
    )
    slope_col = find_column(
        headers,
        candidates=[(("slope",), tuple())],
        fallback=None,
    )

    start_row = anchor_row + 1
    end_row = anchor_row + N_QUARTERS
    row_count = max(0, end_row - start_row + 1)

    scratch_start_col = max(snapshot["right_col"] + 2, anchor_col + 2)
    scratch_intercept_col = scratch_start_col
    scratch_slope_col = scratch_start_col + 1
    scratch_forecast_col = scratch_start_col + 2

    wrote_formulas = False
    for idx in range(row_count):
        n_used = idx + 1
        source_start = anchor_row - n_used
        source_end = anchor_row - 1
        if source_start < 1 or source_end < 1:
            continue

        row = start_row + idx

        intercept_formula = (
            f"=INTERCEPT(R{source_start}C{y_col}:R{source_end}C{y_col},"
            f"R{source_start}C{x_col}:R{source_end}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{source_start}C{y_col}:R{source_end}C{y_col},"
            f"R{source_start}C{x_col}:R{source_end}C{x_col})"
        )
        forecast_formula = (
            f"=R{row}C{scratch_slope_col}*R{anchor_row}C{x_col}+"
            f"R{row}C{scratch_intercept_col}"
        )

        set_formula2(sheet.cells(row, scratch_intercept_col), intercept_formula)
        set_formula2(sheet.cells(row, scratch_slope_col), slope_formula)
        set_formula2(sheet.cells(row, scratch_forecast_col), forecast_formula)
        wrote_formulas = True

    if wrote_formulas:
        wb.app.calculate()

    intercept_values = ensure_column_list(
        sheet.range((start_row, scratch_intercept_col), (end_row, scratch_intercept_col)).value,
        row_count,
    )
    slope_values = ensure_column_list(
        sheet.range((start_row, scratch_slope_col), (end_row, scratch_slope_col)).value,
        row_count,
    )
    forecast_values_scratch = ensure_column_list(
        sheet.range((start_row, scratch_forecast_col), (end_row, scratch_forecast_col)).value,
        row_count,
    )

    col_values = read_columns(
        sheet,
        start_row=start_row,
        end_row=end_row,
        columns={
            "num_quarters_used": num_col,
            "forecast_value": forecast_col,
            "actual_value": actual_col,
            "forecast_max": anchor_col,
            "forecast_min": min_col,
            "intercept": intercept_col,
            "slope": slope_col,
        },
    )

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None

    for idx in range(row_count):
        intercept = clean_value(col_values["intercept"][idx])
        slope = clean_value(col_values["slope"][idx])
        if intercept is None:
            intercept = clean_value(intercept_values[idx])
        if slope is None:
            slope = clean_value(slope_values[idx])

        forecast_value = clean_value(col_values["forecast_value"][idx])
        if forecast_value is None:
            forecast_value = clean_value(forecast_values_scratch[idx])

        row_data = {
            "model": labels["model"],
            "ticker": labels["ticker"],
            "model_period": labels["model_period"],
            "model_date": labels["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": to_int_if_whole(col_values["num_quarters_used"][idx])
            or (idx + 1),
            "num_quarters_used": to_int_if_whole(col_values["num_quarters_used"][idx])
            or (idx + 1),
            "forecast_value": forecast_value,
            "actual_value": clean_value(col_values["actual_value"][idx]),
            "forecast_max": clean_value(col_values["forecast_max"][idx]),
            "forecast_min": clean_value(col_values["forecast_min"][idx]),
            "range_width": compute_range_width(
                col_values["forecast_max"][idx], col_values["forecast_min"][idx]
            ),
            "intercept": intercept,
            "slope": slope,
            "source_file": file_name,
        }

        significant_values = [
            row_data["forecast_value"],
            row_data["forecast_max"],
            row_data["forecast_min"],
            row_data["intercept"],
            row_data["slope"],
        ]
        if not row_has_any_value(significant_values):
            continue

        signature = signature_for_regression_row(row_data)
        if idx == row_count - 1 and previous_signature is not None and signature == previous_signature:
            # Prevent duplicate final row.
            continue

        rows.append(row_data)
        previous_signature = signature

    return rows


def set_column_widths(ws: Any, rows: list[dict[str, Any]], columns: list[str]) -> None:
    for idx, column_name in enumerate(columns, start=1):
        width = len(column_name) + 2
        for row in rows:
            value = clean_value(row.get(column_name))
            if value is None:
                continue
            width = max(width, len(str(value)) + 2)
        ws.column_dimensions[get_column_letter(idx)].width = min(max(width, 12), 60)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = Workbook()
    wb.remove(wb.active)

    def populate_sheet(sheet_name: str, columns: list[str], rows: list[dict[str, Any]]) -> None:
        ws = wb.create_sheet(title=sheet_name)
        ws.append(columns)
        for row in rows:
            ws.append([row.get(col) for col in columns])

        bold_font = Font(bold=True)
        for cell in ws[1]:
            cell.font = bold_font

        ws.freeze_panes = "A2"
        last_col = get_column_letter(len(columns))
        ws.auto_filter.ref = f"A1:{last_col}{max(1, ws.max_row)}"
        set_column_widths(ws, rows, columns)

    populate_sheet("empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    populate_sheet("regression_candidates", REGRESSION_COLUMNS, regression_rows)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def main() -> None:
    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    output_path = next_output_path(input_dir, output_dir)

    entries = sorted(input_dir.iterdir(), key=lambda p: p.name.lower())
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_file_count = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for path in entries:
            if not path.is_file():
                print(f"skipped: {path.name} (reason: not a file)")
                continue

            if path.name.startswith("~"):
                print(f"skipped: {path.name} (reason: temp file)")
                continue

            if path.suffix.lower() != ".xlsx":
                print(f"skipped: {path.name} (reason: not .xlsx)")
                continue

            print(f"processed file: {path.name}")
            labels = parse_filename_labels(path.name)
            source_wb: xw.Book | None = None
            try:
                source_wb = app.books.open(str(path), update_links=False)
                empirical_rows.extend(extract_empirical_rows(source_wb, path.name, labels))
                regression_rows.extend(extract_regression_rows(source_wb, path.name, labels))
                processed_file_count += 1
            except Exception as exc:
                print(f"skipped: {path.name} (reason: {exc})")
            finally:
                if source_wb is not None:
                    close_workbook_no_save(source_wb)
    finally:
        try:
            app.quit()
        except Exception:
            pass

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"output path: {output_path.resolve()}")
    print(f"number of files processed: {processed_file_count}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

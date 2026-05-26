#!/usr/bin/env python3
"""
Extract empirical and regression parameter candidates from Excel model workbooks.

The script opens each source workbook exactly once, processes both model sheets
while the workbook is open, and closes the workbook without saving any changes.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Mapping, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Required top-level input/output variables.
input_dir = "./input"
output_dir = "./output"

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

DAY_BY_PERIOD = {"early": 5, "mid": 15, "late": 25}
PERIOD_RE = re.compile(
    r"(?P<period>Early|Mid|Late)(?P<month>[A-Za-z]{3,9})(?P<year>\d{4})",
    re.IGNORECASE,
)


@dataclass
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values
    return [values]


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and not value.strip():
        return True
    return False


def first_non_blank(*values: Any) -> Any:
    for value in values:
        if not is_blank(value):
            return value
    return None


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    if not text:
        return ""
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return text.strip()


def to_number(value: Any) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
            return None
        try:
            if text.endswith("%"):
                return float(text[:-1]) / 100.0
            return float(text)
        except ValueError:
            return None
    return None


def subtract_numbers(left: Any, right: Any) -> float | None:
    left_num = to_number(left)
    right_num = to_number(right)
    if left_num is None or right_num is None:
        return None
    return left_num - right_num


def parse_month(month_text: str) -> int | None:
    clean = month_text.strip()
    if not clean:
        return None
    for fmt in ("%b", "%B"):
        try:
            return datetime.strptime(clean.title(), fmt).month
        except ValueError:
            continue
    try:
        return datetime.strptime(clean[:3].title(), "%b").month
    except ValueError:
        return None


def parse_file_labels(file_name: str) -> FileLabels:
    stem = Path(file_name).stem
    parts = [chunk.strip() for chunk in stem.split(" - ") if chunk.strip()]
    ticker = (parts[1] if len(parts) > 1 else "UNKNOWN").upper()

    period_source = parts[2] if len(parts) > 2 else stem
    period_match = PERIOD_RE.search(period_source) or PERIOD_RE.search(stem)

    model_period = "unknown_period"
    model_date = ""
    if period_match:
        period_prefix = period_match.group("period").title()
        month_token = period_match.group("month")
        year = int(period_match.group("year"))
        month_number = parse_month(month_token)
        if month_number is not None:
            month_label = datetime(year, month_number, 1).strftime("%b")
            model_period = f"{period_prefix}{month_label}_{year}"
            day = DAY_BY_PERIOD[period_prefix.lower()]
            model_date = f"{year:04d}-{month_number:02d}-{day:02d}"

    model = f"{ticker}_{model_period}"
    return FileLabels(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def find_anchor_max(sheet: xw.Sheet) -> Tuple[int, int] | None:
    used = sheet.used_range
    values = normalize_2d(used.value)
    if not values:
        return None
    start_row = used.row
    start_col = used.column
    for row_offset, row_values in enumerate(values):
        for col_offset, value in enumerate(row_values):
            if normalize_label(value) == "max":
                return start_row + row_offset, start_col + col_offset
    return None


def build_header_offsets(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    radius: int = 24,
) -> Dict[str, int]:
    offsets: Dict[str, int] = {}
    left = max(1, anchor_col - radius)
    right = anchor_col + radius
    for row in (anchor_row - 1, anchor_row, anchor_row + 1):
        if row < 1:
            continue
        row_values = normalize_2d(sheet.range((row, left), (row, right)).value)
        if not row_values:
            continue
        for idx, value in enumerate(row_values[0]):
            label = normalize_label(value)
            if label and label not in offsets:
                offsets[label] = (left + idx) - anchor_col
    offsets.setdefault("max", 0)
    return offsets


def resolve_offset(
    header_offsets: Mapping[str, int],
    candidates: Sequence[str],
    default: int,
) -> int:
    for candidate in candidates:
        expected = normalize_label(candidate)
        for label, offset in header_offsets.items():
            if expected and expected in label:
                return offset
    return default


def read_rows_for_offsets(
    sheet: xw.Sheet,
    start_row: int,
    n_rows: int,
    anchor_col: int,
    offsets: Mapping[str, int],
) -> Dict[str, List[Any]]:
    absolute_cols = {name: anchor_col + offset for name, offset in offsets.items()}
    valid_cols = [col for col in absolute_cols.values() if col >= 1]
    data: Dict[str, List[Any]] = {name: [None] * n_rows for name in offsets}
    if not valid_cols:
        return data

    left = min(valid_cols)
    right = max(valid_cols)
    raw = normalize_2d(sheet.range((start_row, left), (start_row + n_rows - 1, right)).value)
    while len(raw) < n_rows:
        raw.append([None] * (right - left + 1))

    for field_name, absolute_col in absolute_cols.items():
        if absolute_col < 1:
            continue
        index = absolute_col - left
        values: List[Any] = []
        for row_values in raw[:n_rows]:
            values.append(row_values[index] if 0 <= index < len(row_values) else None)
        data[field_name] = values
    return data


def flatten_column(range_values: Any, n_rows: int) -> List[Any]:
    rows = normalize_2d(range_values)
    flattened = [(row[0] if row else None) for row in rows]
    while len(flattened) < n_rows:
        flattened.append(None)
    return flattened[:n_rows]


def as_quarter_count(value: Any, default: int) -> int | float:
    parsed = to_number(value)
    if parsed is None:
        return default
    rounded = round(parsed)
    if abs(parsed - rounded) < 1e-9:
        return int(rounded)
    return parsed


def set_formula2(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        # Formula fallback is only for hosts that do not expose formula2.
        cell.formula = formula


def safe_close_workbook(workbook: xw.Book) -> None:
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


def build_output_path(input_path: Path, output_path: Path) -> Path:
    base_name = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    index = 1
    while candidate.exists():
        candidate = output_path / f"{base_name}.{index}.xlsx"
        index += 1
    return candidate


def get_sheet(workbook: xw.Book, sheet_name: str) -> xw.Sheet | None:
    try:
        return workbook.sheets[sheet_name]
    except Exception:
        return None


def empirical_offsets_from_headers(header_offsets: Mapping[str, int]) -> Dict[str, int]:
    return {
        "num_quarters_used": resolve_offset(
            header_offsets,
            ("num quarters used", "n quarters", "quarters used"),
            -10,
        ),
        "last_quarter_used": resolve_offset(
            header_offsets,
            ("last quarter used", "last quarter", "last qtr"),
            -9,
        ),
        "forecast_value": resolve_offset(
            header_offsets,
            ("estimated total sold", "est total sold", "forecast value"),
            -8,
        ),
        "actual_value": resolve_offset(
            header_offsets,
            ("actual value", "reported sales", "actual sales"),
            -7,
        ),
        "forecast_max": 0,
        "forecast_min": resolve_offset(header_offsets, ("forecast min", "min"), 1),
        "quarterly_sales": resolve_offset(
            header_offsets,
            ("quarterly sales", "quarter sales"),
            -6,
        ),
        "reported_sales": resolve_offset(
            header_offsets,
            ("reported sales", "reported"),
            -5,
        ),
        "growth_rate_pct": resolve_offset(
            header_offsets,
            ("growth rate", "growth pct"),
            -4,
        ),
        "sales_captured_in_db_pct": resolve_offset(
            header_offsets,
            ("sales captured in db", "captured in db", "capture pct", "penetration"),
            -3,
        ),
    }


def regression_offsets_from_headers(header_offsets: Mapping[str, int]) -> Dict[str, int]:
    return {
        "num_quarters_used": resolve_offset(
            header_offsets,
            ("num quarters used", "n quarters", "quarters used"),
            -10,
        ),
        "forecast_value": resolve_offset(
            header_offsets,
            ("tot fcst w o sa", "tot fcst without sa", "forecast"),
            -8,
        ),
        "actual_value": resolve_offset(
            header_offsets,
            ("actual value", "reported sales", "actual sales"),
            -7,
        ),
        "forecast_max": 0,
        "forecast_min": resolve_offset(header_offsets, ("forecast min", "min"), 1),
    }


def extract_empirical_rows(
    workbook: xw.Book,
    labels: FileLabels,
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = get_sheet(workbook, "Empirical Model")
    if sheet is None:
        print(f"  skipped empirical: missing sheet 'Empirical Model'")
        return []

    anchor = find_anchor_max(sheet)
    if anchor is None:
        print("  skipped empirical: could not find 'max' anchor")
        return []

    anchor_row, anchor_col = anchor
    header_offsets = build_header_offsets(sheet, anchor_row, anchor_col)
    offsets = empirical_offsets_from_headers(header_offsets)
    start_row = anchor_row + 1

    values = read_rows_for_offsets(
        sheet=sheet,
        start_row=start_row,
        n_rows=N_QUARTERS,
        anchor_col=anchor_col,
        offsets=offsets,
    )

    formula_col = max(sheet.used_range.last_cell.column + 2, anchor_col + 2)
    sales_capture_col = anchor_col + offsets["sales_captured_in_db_pct"]
    sales_capture_rel = sales_capture_col - formula_col
    for index in range(N_QUARTERS):
        formula = (
            f'=IFERROR(AVERAGE(R[-{index}]C[{sales_capture_rel}]:RC[{sales_capture_rel}]),"")'
        )
        set_formula2(sheet.range((start_row + index, formula_col)), formula)

    workbook.app.calculate()
    avg_penetration_values = flatten_column(
        sheet.range((start_row, formula_col), (start_row + N_QUARTERS - 1, formula_col)).value,
        N_QUARTERS,
    )

    rows: List[Dict[str, Any]] = []
    for index in range(N_QUARTERS):
        forecast_max = values["forecast_max"][index]
        forecast_min = values["forecast_min"][index]
        forecast_value = values["forecast_value"][index]
        reported_sales = values["reported_sales"][index]
        actual_value = first_non_blank(reported_sales, values["actual_value"][index])
        quarterly_sales = values["quarterly_sales"][index]
        growth_rate_pct = values["growth_rate_pct"][index]
        sales_captured = values["sales_captured_in_db_pct"][index]
        avg_penetration = avg_penetration_values[index]

        key_values = [
            forecast_value,
            actual_value,
            forecast_max,
            forecast_min,
            avg_penetration,
            quarterly_sales,
            sales_captured,
        ]
        if all(is_blank(value) for value in key_values):
            continue

        row: Dict[str, Any] = {
            "model": labels.model,
            "ticker": labels.ticker,
            "model_period": labels.model_period,
            "model_date": labels.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration,
            "num_quarters_used": as_quarter_count(values["num_quarters_used"][index], index + 1),
            "last_quarter_used": values["last_quarter_used"][index],
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": subtract_numbers(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate_pct,
            "sales_captured_in_db_pct": sales_captured,
            "source_file": source_file,
        }
        rows.append(row)
    return rows


def regression_signature(row: Mapping[str, Any]) -> Tuple[Any, ...]:
    signature_fields = (
        "num_quarters_used",
        "intercept",
        "slope",
        "forecast_value",
        "forecast_max",
        "forecast_min",
    )
    normalized: List[Any] = []
    for field in signature_fields:
        value = row.get(field)
        if isinstance(value, float):
            normalized.append(round(value, 10))
        else:
            normalized.append(value)
    return tuple(normalized)


def extract_regression_rows(
    workbook: xw.Book,
    labels: FileLabels,
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = get_sheet(workbook, "Regression Model")
    if sheet is None:
        print(f"  skipped regression: missing sheet 'Regression Model'")
        return []

    anchor = find_anchor_max(sheet)
    if anchor is None:
        print("  skipped regression: could not find 'max' anchor")
        return []

    anchor_row, anchor_col = anchor
    header_offsets = build_header_offsets(sheet, anchor_row, anchor_col)
    offsets = regression_offsets_from_headers(header_offsets)
    start_row = anchor_row + 1

    values = read_rows_for_offsets(
        sheet=sheet,
        start_row=start_row,
        n_rows=N_QUARTERS,
        anchor_col=anchor_col,
        offsets=offsets,
    )

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    intercept_col = max(sheet.used_range.last_cell.column + 2, anchor_col + 2)
    slope_col = intercept_col + 1
    intercept_y_rel = y_col - intercept_col
    intercept_x_rel = x_col - intercept_col
    slope_y_rel = y_col - slope_col
    slope_x_rel = x_col - slope_col

    for index in range(N_QUARTERS):
        intercept_formula = (
            f'=IFERROR(INTERCEPT(R[-{index}]C[{intercept_y_rel}]:RC[{intercept_y_rel}],'
            f'R[-{index}]C[{intercept_x_rel}]:RC[{intercept_x_rel}]),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R[-{index}]C[{slope_y_rel}]:RC[{slope_y_rel}],'
            f'R[-{index}]C[{slope_x_rel}]:RC[{slope_x_rel}]),"")'
        )
        set_formula2(sheet.range((start_row + index, intercept_col)), intercept_formula)
        set_formula2(sheet.range((start_row + index, slope_col)), slope_formula)

    workbook.app.calculate()
    intercept_values = flatten_column(
        sheet.range((start_row, intercept_col), (start_row + N_QUARTERS - 1, intercept_col)).value,
        N_QUARTERS,
    )
    slope_values = flatten_column(
        sheet.range((start_row, slope_col), (start_row + N_QUARTERS - 1, slope_col)).value,
        N_QUARTERS,
    )

    rows: List[Dict[str, Any]] = []
    for index in range(N_QUARTERS):
        forecast_max = values["forecast_max"][index]
        forecast_min = values["forecast_min"][index]
        forecast_value = values["forecast_value"][index]
        actual_value = values["actual_value"][index]
        intercept = intercept_values[index]
        slope = slope_values[index]

        key_values = [forecast_value, forecast_max, forecast_min, intercept, slope]
        if all(is_blank(value) for value in key_values):
            continue

        row: Dict[str, Any] = {
            "model": labels.model,
            "ticker": labels.ticker,
            "model_period": labels.model_period,
            "model_date": labels.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": as_quarter_count(values["num_quarters_used"][index], index + 1),
            "num_quarters_used": as_quarter_count(values["num_quarters_used"][index], index + 1),
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": subtract_numbers(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        if index == N_QUARTERS - 1 and rows and regression_signature(row) == regression_signature(rows[-1]):
            print("  skipped regression final duplicate row")
            continue
        rows.append(row)
    return rows


def write_sheet(
    worksheet,
    columns: Sequence[str],
    rows: Sequence[Mapping[str, Any]],
) -> None:
    worksheet.append(list(columns))
    for row in rows:
        worksheet.append([row.get(column) for column in columns])

    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for col_idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row_idx in range(2, worksheet.max_row + 1):
            value = worksheet.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        worksheet.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 44)


def write_output_workbook(
    output_file: Path,
    empirical_rows: Sequence[Mapping[str, Any]],
    regression_rows: Sequence[Mapping[str, Any]],
) -> None:
    workbook = Workbook()
    empirical_sheet = workbook.active
    empirical_sheet.title = "empirical_candidates"
    regression_sheet = workbook.create_sheet("regression_candidates")

    write_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)
    workbook.save(output_file)


def collect_source_files(input_path: Path) -> List[Path]:
    selected: List[Path] = []
    for file_path in sorted(input_path.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped {file_path.name}: temporary file")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped {file_path.name}: not an .xlsx file")
            continue
        selected.append(file_path)
    return selected


def run_extraction(input_path: Path, output_path: Path) -> None:
    output_path.mkdir(parents=True, exist_ok=True)

    source_files = collect_source_files(input_path)
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        for file_path in source_files:
            print(f"Processing {file_path.name}")
            labels = parse_file_labels(file_path.name)
            workbook = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_rows(workbook, labels, file_path.name))
                regression_rows.extend(extract_regression_rows(workbook, labels, file_path.name))
                processed_files += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        app.quit()

    output_file = build_output_path(input_path, output_path)
    write_output_workbook(output_file, empirical_rows, regression_rows)
    print(f"Output path: {output_file}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists():
        print(f"Input directory not found: {input_path}")
        return
    if not input_path.is_dir():
        print(f"Input path is not a directory: {input_path}")
        return

    run_extraction(input_path, output_path)


if __name__ == "__main__":
    main()

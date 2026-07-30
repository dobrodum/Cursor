#!/usr/bin/env python3
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# User-configurable paths
# ---------------------------------------------------------------------------
input_dir = Path("input")
output_dir = Path("output")


N_QUARTERS = 10
EMPIRICAL_MODEL_SHEET = "Empirical Model"
REGRESSION_MODEL_SHEET = "Regression Model"
OUTPUT_SUFFIX = "_PARAM"

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


@dataclass(frozen=True)
class FileModelMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


def to_matrix(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def to_number(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def to_int(value: Any) -> Optional[int]:
    number = to_number(value)
    if number is None:
        return None
    rounded = int(round(number))
    if abs(number - rounded) > 1e-9:
        return None
    return rounded


def safe_month_number(month_fragment: str) -> Optional[int]:
    month_token = month_fragment[:3].title()
    try:
        return datetime.strptime(month_token, "%b").month
    except ValueError:
        return None


def parse_file_metadata(file_name: str) -> Optional[FileModelMetadata]:
    """
    Expected input pattern example:
      MedMiner_Model - AORT - MidJan2026_Send.xlsx
    """
    pattern = re.compile(
        r"^\s*.+?-\s*([A-Za-z0-9]+)\s*-\s*((Early|Mid|Late)([A-Za-z]{3,9})(\d{4}))",
        re.IGNORECASE,
    )
    match = pattern.match(file_name)
    if not match:
        return None

    ticker = match.group(1).upper()
    timing = match.group(3).title()
    month_text = match.group(4)
    year = int(match.group(5))

    month_number = safe_month_number(month_text)
    if month_number is None:
        return None

    day_by_timing = {"Early": 5, "Mid": 15, "Late": 25}
    model_day = day_by_timing[timing]
    model_date = date(year, month_number, model_day).isoformat()
    model_period = f"{timing}{month_text[:3].title()}_{year}"
    model = f"{ticker}_{model_period}"
    return FileModelMetadata(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def get_unique_output_path(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_path.name}{OUTPUT_SUFFIX}"
    candidate = output_path / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    version = 1
    while True:
        candidate = output_path / f"{base_name}.{version}.xlsx"
        if not candidate.exists():
            return candidate
        version += 1


def safe_close_workbook(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        workbook.close()
    except Exception:
        pass


def find_anchor_cell(sheet: xw.Sheet, anchor_text: str = "max") -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    matrix = to_matrix(used.value)
    if not matrix:
        return None

    wanted = anchor_text.strip().lower()
    for row_idx, row_values in enumerate(matrix):
        for col_idx, value in enumerate(row_values):
            if isinstance(value, str) and value.strip().lower() == wanted:
                return used.row + row_idx, used.column + col_idx
    return None


def get_cell_value(sheet: xw.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    try:
        return sheet.cells(row, col).value
    except Exception:
        return None


def set_formula2(sheet: xw.Sheet, row: int, col: int, formula: str) -> None:
    if row < 1 or col < 1:
        return
    sheet.cells(row, col).formula2 = formula


def choose_helper_columns(anchor_col: int) -> Tuple[int, int]:
    """
    Excel column limit is 16384 (XFD). Keep helper formulas near anchor and in-bounds.
    """
    if anchor_col + 3 <= 16384:
        return anchor_col + 2, anchor_col + 3
    return max(1, anchor_col - 3), max(1, anchor_col - 2)


def map_empirical_header(raw: Any) -> Optional[str]:
    label = normalize_text(raw)
    if not label:
        return None
    if label == "max":
        return "forecast_max"
    if label == "min":
        return "forecast_min"
    if "quarter" in label and ("num" in label or "used" in label):
        return "num_quarters_used"
    if "last quarter" in label:
        return "last_quarter_used"
    if "estimated total sold" in label or "est total sold" in label:
        return "forecast_value"
    if "reported sales" in label:
        return "reported_sales"
    if "quarterly sales" in label:
        return "quarterly_sales"
    if "growth" in label:
        return "growth_rate_pct"
    if "captured" in label and "db" in label:
        return "sales_captured_in_db_pct"
    if "avg penetration" in label:
        return "avg_penetration_pct"
    return None


def map_regression_header(raw: Any) -> Optional[str]:
    label = normalize_text(raw)
    if not label:
        return None
    if label == "max":
        return "forecast_max"
    if label == "min":
        return "forecast_min"
    if "quarter" in label and ("num" in label or "used" in label):
        return "num_quarters_used"
    if "tot fcst" in label and ("wo sa" in label or "w o sa" in label or "without sa" in label):
        return "forecast_value"
    if "forecast" in label and "sa" in label:
        return "forecast_value"
    if "actual" in label or "reported sales" in label:
        return "actual_value"
    return None


def build_header_lookup(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    mapper,
    window: int = 30,
) -> Dict[str, int]:
    lookup: Dict[str, int] = {}
    start_col = max(1, anchor_col - window)
    end_col = min(16384, anchor_col + window)
    for col in range(start_col, end_col + 1):
        mapped = mapper(get_cell_value(sheet, anchor_row, col))
        if mapped and mapped not in lookup:
            lookup[mapped] = col

    # Anchor-based fallback offsets (fast and stable if labels are inconsistent).
    lookup.setdefault("forecast_max", anchor_col)
    lookup.setdefault("forecast_min", anchor_col + 1)
    return lookup


def detect_quarter_columns(sheet: xw.Sheet, header_row: int, anchor_col: int) -> List[int]:
    quarter_like = re.compile(
        r"^(q[1-4]\s*\d{2,4}|\d{4}\s*q[1-4}|[a-z]{3}\s*\d{2,4}|\d{2,4}\s*[a-z]{3})$",
        re.IGNORECASE,
    )
    result: List[int] = []
    start_col = max(1, anchor_col - 30)
    for col in range(start_col, anchor_col):
        value = get_cell_value(sheet, header_row, col)
        text = normalize_text(value).replace(" ", "")
        if not text:
            continue
        if quarter_like.match(text):
            result.append(col)
    return result


def read_from_lookup(
    sheet: xw.Sheet,
    row: int,
    lookup: Dict[str, int],
    key: str,
    fallback_offset_col: Optional[int] = None,
) -> Any:
    if key in lookup:
        return get_cell_value(sheet, row, lookup[key])
    if fallback_offset_col is not None:
        return get_cell_value(sheet, row, fallback_offset_col)
    return None


def add_empirical_rows(
    workbook: xw.Book,
    sheet: xw.Sheet,
    metadata: FileModelMetadata,
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    lookup = build_header_lookup(sheet, anchor_row, anchor_col, map_empirical_header)
    quarter_cols = detect_quarter_columns(sheet, anchor_row, anchor_col)
    avg_helper_col, _ = choose_helper_columns(anchor_col)

    data_rows = [anchor_row + 1 + idx for idx in range(N_QUARTERS)]
    formula_rows: List[int] = []

    for idx, row in enumerate(data_rows):
        n_used = read_from_lookup(sheet, row, lookup, "num_quarters_used")
        n_quarters = to_int(n_used) or (idx + 1)
        n_quarters = max(1, min(n_quarters, N_QUARTERS))

        if quarter_cols:
            selected = quarter_cols[-n_quarters:]
            start_col = selected[0]
            end_col = selected[-1]
        else:
            # Anchor-based fallback when headers are not quarter-labelled.
            end_col = max(1, anchor_col - 1)
            start_col = max(1, end_col - n_quarters + 1)

        formula = f'=IFERROR(AVERAGE(R{row}C{start_col}:R{row}C{end_col}),"")'
        set_formula2(sheet, row, avg_helper_col, formula)
        formula_rows.append(row)

    if formula_rows:
        workbook.app.calculate()

    rows: List[Dict[str, Any]] = []
    for idx, row in enumerate(data_rows):
        n_used = read_from_lookup(
            sheet,
            row,
            lookup,
            "num_quarters_used",
            fallback_offset_col=anchor_col - 10,
        )
        num_quarters_used = to_int(n_used) or (idx + 1)

        last_quarter_used = read_from_lookup(
            sheet,
            row,
            lookup,
            "last_quarter_used",
            fallback_offset_col=anchor_col - 9,
        )
        forecast_value = read_from_lookup(
            sheet,
            row,
            lookup,
            "forecast_value",
            fallback_offset_col=anchor_col - 2,
        )
        reported_sales = read_from_lookup(
            sheet,
            row,
            lookup,
            "reported_sales",
            fallback_offset_col=anchor_col - 1,
        )
        forecast_max = read_from_lookup(sheet, row, lookup, "forecast_max", fallback_offset_col=anchor_col)
        forecast_min = read_from_lookup(
            sheet,
            row,
            lookup,
            "forecast_min",
            fallback_offset_col=anchor_col + 1,
        )
        quarterly_sales = read_from_lookup(
            sheet,
            row,
            lookup,
            "quarterly_sales",
            fallback_offset_col=anchor_col - 6,
        )
        growth_rate_pct = read_from_lookup(
            sheet,
            row,
            lookup,
            "growth_rate_pct",
            fallback_offset_col=anchor_col - 5,
        )
        captured_pct = read_from_lookup(
            sheet,
            row,
            lookup,
            "sales_captured_in_db_pct",
            fallback_offset_col=anchor_col - 4,
        )

        avg_penetration_pct = get_cell_value(sheet, row, avg_helper_col)
        range_width = None
        max_num = to_number(forecast_max)
        min_num = to_number(forecast_min)
        if max_num is not None and min_num is not None:
            range_width = max_num - min_num

        has_payload = any(
            value not in (None, "")
            for value in (forecast_value, reported_sales, forecast_max, forecast_min, avg_penetration_pct)
        )
        if not has_payload:
            continue

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": captured_pct,
                "source_file": source_file,
            }
        )

    return rows


def add_regression_rows(
    workbook: xw.Book,
    sheet: xw.Sheet,
    metadata: FileModelMetadata,
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    lookup = build_header_lookup(sheet, anchor_row, anchor_col, map_regression_header)
    intercept_helper_col, slope_helper_col = choose_helper_columns(anchor_col)

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if x_col < 1 or y_col < 1:
        return []

    historical_end_row = anchor_row - 1
    while historical_end_row > 1:
        x_value = to_number(get_cell_value(sheet, historical_end_row, x_col))
        y_value = to_number(get_cell_value(sheet, historical_end_row, y_col))
        if x_value is not None and y_value is not None:
            break
        historical_end_row -= 1

    if historical_end_row <= 1:
        return []

    data_rows = [anchor_row + 1 + idx for idx in range(N_QUARTERS)]
    for idx, output_row in enumerate(data_rows):
        n_quarters = idx + 1
        start_row = historical_end_row - n_quarters + 1
        if start_row < 1:
            continue

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{historical_end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{historical_end_row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{historical_end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{historical_end_row}C{x_col}),"")'
        )
        set_formula2(sheet, output_row, intercept_helper_col, intercept_formula)
        set_formula2(sheet, output_row, slope_helper_col, slope_formula)

    workbook.app.calculate()

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for idx, row in enumerate(data_rows):
        n_used = read_from_lookup(
            sheet,
            row,
            lookup,
            "num_quarters_used",
            fallback_offset_col=anchor_col - 10,
        )
        num_quarters_used = to_int(n_used) or (idx + 1)

        forecast_value = read_from_lookup(
            sheet,
            row,
            lookup,
            "forecast_value",
            fallback_offset_col=anchor_col - 2,
        )
        actual_value = read_from_lookup(
            sheet,
            row,
            lookup,
            "actual_value",
            fallback_offset_col=None,
        )
        forecast_max = read_from_lookup(sheet, row, lookup, "forecast_max", fallback_offset_col=anchor_col)
        forecast_min = read_from_lookup(
            sheet,
            row,
            lookup,
            "forecast_min",
            fallback_offset_col=anchor_col + 1,
        )
        intercept_value = get_cell_value(sheet, row, intercept_helper_col)
        slope_value = get_cell_value(sheet, row, slope_helper_col)

        max_num = to_number(forecast_max)
        min_num = to_number(forecast_min)
        range_width = (max_num - min_num) if max_num is not None and min_num is not None else None

        has_payload = any(
            value not in (None, "")
            for value in (forecast_value, forecast_max, forecast_min, intercept_value, slope_value)
        )
        if not has_payload:
            continue

        signature = (
            num_quarters_used,
            to_number(forecast_value),
            to_number(forecast_max),
            to_number(forecast_min),
            to_number(intercept_value),
            to_number(slope_value),
        )
        if previous_signature is not None and signature == previous_signature:
            # Avoid duplicate terminal row when workbook formulas return repeated output.
            continue
        previous_signature = signature

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept_value,
                "slope": slope_value,
                "source_file": source_file,
            }
        )

    return rows


def write_sheet(ws, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(col) for col in columns])

    for col_idx in range(1, len(columns) + 1):
        ws.cell(row=1, column=col_idx).font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{ws.max_row}"

    for col_idx, column_name in enumerate(columns, start=1):
        values = [ws.cell(row=r, column=col_idx).value for r in range(1, ws.max_row + 1)]
        max_len = max(len(str(column_name)), *(len(str(v)) for v in values if v is not None))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(60, max(12, max_len + 2))


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    out_wb = Workbook()
    default_ws = out_wb.active
    out_wb.remove(default_ws)

    empirical_ws = out_wb.create_sheet("empirical_candidates")
    regression_ws = out_wb.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)
    out_wb.save(output_path)


def main() -> None:
    input_path = input_dir.expanduser().resolve()
    output_path = output_dir.expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"Input directory does not exist: {input_path}")

    result_path = get_unique_output_path(input_path, output_path)

    processed_files = 0
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    original_calculation = app.calculation
    app.calculation = "manual"

    try:
        for file_path in sorted(input_path.iterdir()):
            if not file_path.is_file():
                print(f"skipped {file_path.name}: not a file")
                continue
            if file_path.name.startswith("~"):
                print(f"skipped {file_path.name}: temporary file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped {file_path.name}: not an .xlsx file")
                continue

            metadata = parse_file_metadata(file_path.name)
            if metadata is None:
                print(f"skipped {file_path.name}: filename did not match expected convention")
                continue

            print(f"processed file: {file_path.name}")
            workbook = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                processed_files += 1

                try:
                    empirical_sheet = workbook.sheets[EMPIRICAL_MODEL_SHEET]
                    empirical_rows.extend(
                        add_empirical_rows(workbook, empirical_sheet, metadata, file_path.name)
                    )
                except Exception as exc:
                    print(f"skipped empirical extraction for {file_path.name}: {exc}")

                try:
                    regression_sheet = workbook.sheets[REGRESSION_MODEL_SHEET]
                    regression_rows.extend(
                        add_regression_rows(workbook, regression_sheet, metadata, file_path.name)
                    )
                except Exception as exc:
                    print(f"skipped regression extraction for {file_path.name}: {exc}")
            except Exception as exc:
                print(f"skipped {file_path.name}: failed to open workbook ({exc})")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        app.calculation = original_calculation
        app.quit()

    write_output_workbook(result_path, empirical_rows, regression_rows)

    print(f"output path: {result_path}")
    print(f"files processed: {processed_files}")
    print(f"empirical rows: {len(empirical_rows)}")
    print(f"regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

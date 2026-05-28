#!/usr/bin/env python3
from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure these two paths before running.
input_dir = Path("./input").resolve()
output_dir = Path("./output").resolve()

EMPIRICAL_SHEET = "Empirical Model"
REGRESSION_SHEET = "Regression Model"

EMPIRICAL_HEADERS = [
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

REGRESSION_HEADERS = [
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

DAY_BY_PHASE = {"early": 5, "mid": 15, "late": 25}
MONTH_BY_ABBR = {
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


def normalize_text(value: str) -> str:
    text = value.strip().lower()
    text = re.sub(r"[^a-z0-9%]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def to_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, (list, tuple)):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], (list, tuple)):
        matrix: List[List[Any]] = []
        for row in values:
            if isinstance(row, (list, tuple)):
                matrix.append(list(row))
            else:
                matrix.append([row])
        return matrix
    return [list(values)]


def is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def as_float(value: Any) -> Optional[float]:
    if is_number(value):
        return float(value)
    return None


def compute_range_width(max_value: Any, min_value: Any) -> Optional[float]:
    max_float = as_float(max_value)
    min_float = as_float(min_value)
    if max_float is None or min_float is None:
        return None
    return max_float - min_float


def choose_best_match(
    cells: Sequence[Tuple[int, int]], near: Optional[Tuple[int, int]]
) -> Optional[Tuple[int, int]]:
    if not cells:
        return None
    if near is None:
        return cells[0]
    return min(cells, key=lambda rc: abs(rc[0] - near[0]) + abs(rc[1] - near[1]))


def build_sheet_cache(sheet: xw.Sheet) -> Dict[str, Any]:
    used = sheet.used_range
    matrix = to_2d(used.value)
    base_row = used.row
    base_col = used.column
    labels: Dict[str, List[Tuple[int, int]]] = {}

    for row_idx, row_values in enumerate(matrix):
        for col_idx, raw_value in enumerate(row_values):
            if isinstance(raw_value, str):
                normalized = normalize_text(raw_value)
                if normalized:
                    labels.setdefault(normalized, []).append(
                        (base_row + row_idx, base_col + col_idx)
                    )

    return {
        "base_row": base_row,
        "base_col": base_col,
        "matrix": matrix,
        "labels": labels,
    }


def find_label_cell(
    cache: Dict[str, Any],
    candidates: Iterable[str],
    near: Optional[Tuple[int, int]] = None,
) -> Optional[Tuple[int, int]]:
    labels: Dict[str, List[Tuple[int, int]]] = cache["labels"]
    normalized_candidates = [normalize_text(candidate) for candidate in candidates]

    exact_hits: List[Tuple[int, int]] = []
    for candidate in normalized_candidates:
        exact_hits.extend(labels.get(candidate, []))
    best_exact = choose_best_match(exact_hits, near)
    if best_exact is not None:
        return best_exact

    partial_hits: List[Tuple[int, int]] = []
    for label_text, cells in labels.items():
        if any(candidate and candidate in label_text for candidate in normalized_candidates):
            partial_hits.extend(cells)
    return choose_best_match(partial_hits, near)


def find_max_anchor(cache: Dict[str, Any]) -> Optional[Tuple[int, int]]:
    labels: Dict[str, List[Tuple[int, int]]] = cache["labels"]
    if "max" in labels:
        return labels["max"][0]

    for label_text, cells in labels.items():
        if re.search(r"\bmax\b", label_text):
            return cells[0]
    return None


def read_right_value(
    sheet: xw.Sheet,
    row: int,
    col: int,
    max_steps: int = 4,
    prefer_numeric: bool = False,
) -> Any:
    fallback = None
    for step in range(1, max_steps + 1):
        value = sheet.range((row, col + step)).value
        if value in (None, ""):
            continue
        if not prefer_numeric:
            return value
        if is_number(value):
            return value
        if fallback is None:
            fallback = value
    return fallback


def read_value_from_label(
    sheet: xw.Sheet,
    label_cell: Optional[Tuple[int, int]],
    prefer_numeric: bool = False,
) -> Any:
    if label_cell is None:
        return None
    return read_right_value(
        sheet=sheet,
        row=label_cell[0],
        col=label_cell[1],
        prefer_numeric=prefer_numeric,
    )


def contiguous_numeric_rows(
    sheet: xw.Sheet,
    value_col: int,
    anchor_row: int,
) -> List[int]:
    rows: List[int] = []
    row = anchor_row - 1
    while row >= 1:
        value = sheet.range((row, value_col)).value
        if is_number(value):
            rows.append(row)
            row -= 1
            continue
        if rows:
            break
        row -= 1
    rows.reverse()
    return rows


def contiguous_xy_rows(
    sheet: xw.Sheet,
    x_col: int,
    y_col: int,
    anchor_row: int,
) -> List[int]:
    rows: List[int] = []
    row = anchor_row - 1
    while row >= 1:
        x_val = sheet.range((row, x_col)).value
        y_val = sheet.range((row, y_col)).value
        if is_number(x_val) and is_number(y_val):
            rows.append(row)
            row -= 1
            continue
        if rows:
            break
        row -= 1
    rows.reverse()
    return rows


def safe_close_workbook(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.api.Close(False)
        return
    except Exception:
        pass

    try:
        workbook.close()
    except Exception as exc:
        print(f"Warning: unable to close workbook safely ({workbook.name}): {exc}")


def parse_file_labels(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1] if len(parts) >= 2 else parts[0]
    ticker = re.sub(r"[^A-Za-z0-9_]", "", ticker).upper()

    period_segment = parts[2] if len(parts) >= 3 else ""
    period_match = re.search(
        r"(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(\d{4})",
        period_segment,
        flags=re.IGNORECASE,
    )

    model_period = ""
    model_date = ""
    if period_match:
        phase_raw = period_match.group(1)
        month_raw = period_match.group(2)
        year_raw = int(period_match.group(3))

        phase_title = phase_raw.title()
        month_title = month_raw.title()
        model_period = f"{phase_title}{month_title}_{year_raw}"

        day = DAY_BY_PHASE[phase_raw.lower()]
        month = MONTH_BY_ABBR[month_raw.lower()]
        model_date = date(year_raw, month, day).isoformat()

    model = f"{ticker}_{model_period}" if model_period else ticker
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def pick_or_default_value_cell(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    label_cell: Optional[Tuple[int, int]],
    default_offset: Tuple[int, int],
) -> xw.Range:
    if label_cell is not None:
        return sheet.range((label_cell[0], label_cell[1] + 1))
    return sheet.range((anchor_row + default_offset[0], anchor_col + default_offset[1]))


def extract_empirical_rows(
    workbook: xw.Book,
    labels: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = workbook.sheets[EMPIRICAL_SHEET]
    except Exception:
        print(f"  - Skipped empirical extraction: '{EMPIRICAL_SHEET}' sheet not found")
        return []

    cache = build_sheet_cache(sheet)
    anchor = find_max_anchor(cache)
    if anchor is None:
        print("  - Skipped empirical extraction: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor

    avg_penetration_label = find_label_cell(
        cache, ["avg penetration", "average penetration"], near=anchor
    )
    estimated_total_label = find_label_cell(
        cache,
        ["estimated total sold", "estimated total", "total sold estimate"],
        near=anchor,
    )
    reported_sales_label = find_label_cell(
        cache, ["reported sales", "actual sales", "reported"], near=anchor
    )
    growth_rate_label = find_label_cell(
        cache, ["growth rate %", "growth rate pct", "growth %"], near=anchor
    )
    min_label = find_label_cell(cache, ["min"], near=anchor)
    if min_label is None:
        min_label = (anchor_row + 1, anchor_col)

    if estimated_total_label is None:
        estimated_total_label = (anchor_row - 2, anchor_col)
    if reported_sales_label is None:
        reported_sales_label = (anchor_row - 1, anchor_col)
    if growth_rate_label is None:
        growth_rate_label = (anchor_row - 3, anchor_col)

    quarter_header = find_label_cell(
        cache, ["quarter", "fiscal quarter", "quarter ended"], near=anchor
    )
    quarterly_sales_header = find_label_cell(
        cache, ["quarterly sales", "qtr sales", "quarter sales"], near=anchor
    )
    penetration_header = find_label_cell(
        cache,
        ["sales captured in db %", "penetration %", "penetration"],
        near=anchor,
    )

    penetration_col = penetration_header[1] if penetration_header else max(1, anchor_col - 6)
    quarter_col = quarter_header[1] if quarter_header else max(1, anchor_col - 11)
    quarterly_sales_col = (
        quarterly_sales_header[1] if quarterly_sales_header else max(1, anchor_col - 8)
    )

    series_rows = contiguous_numeric_rows(sheet, penetration_col, anchor_row)
    if not series_rows:
        print("  - Skipped empirical extraction: penetration history series not found")
        return []

    avg_penetration_input = pick_or_default_value_cell(
        sheet=sheet,
        anchor_row=anchor_row,
        anchor_col=anchor_col,
        label_cell=avg_penetration_label,
        default_offset=(0, 4),
    )

    max_quarters = min(10, len(series_rows))
    output_rows: List[Dict[str, Any]] = []

    for num_quarters_used in range(1, max_quarters + 1):
        selected_rows = series_rows[-num_quarters_used:]
        start_row = selected_rows[0]
        end_row = selected_rows[-1]

        avg_penetration_input.formula2 = (
            f"=AVERAGE(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})"
        )
        workbook.app.calculate()

        avg_penetration_pct = avg_penetration_input.value
        forecast_max = read_value_from_label(sheet, anchor, prefer_numeric=True)
        forecast_min = read_value_from_label(sheet, min_label, prefer_numeric=True)
        forecast_value = read_value_from_label(sheet, estimated_total_label)
        reported_sales_value = read_value_from_label(sheet, reported_sales_label)
        growth_rate_pct = read_value_from_label(sheet, growth_rate_label)
        quarterly_sales = sheet.range((end_row, quarterly_sales_col)).value
        last_quarter_used = sheet.range((start_row, quarter_col)).value
        sales_captured_in_db_pct = sheet.range((end_row, penetration_col)).value

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
            "actual_value": reported_sales_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": compute_range_width(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration_pct,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales_value,
            "growth_rate_pct": growth_rate_pct,
            "sales_captured_in_db_pct": sales_captured_in_db_pct,
            "source_file": source_file,
        }
        output_rows.append(row)

    return output_rows


def rounded_signature(values: Sequence[Any], precision: int = 10) -> Tuple[Any, ...]:
    signature: List[Any] = []
    for value in values:
        if is_number(value):
            signature.append(round(float(value), precision))
        else:
            signature.append(value)
    return tuple(signature)


def extract_regression_rows(
    workbook: xw.Book,
    labels: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = workbook.sheets[REGRESSION_SHEET]
    except Exception:
        print(f"  - Skipped regression extraction: '{REGRESSION_SHEET}' sheet not found")
        return []

    cache = build_sheet_cache(sheet)
    anchor = find_max_anchor(cache)
    if anchor is None:
        print("  - Skipped regression extraction: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if x_col < 1 or y_col < 1:
        print("  - Skipped regression extraction: invalid anchor offsets for x/y columns")
        return []

    intercept_label = find_label_cell(cache, ["intercept"], near=anchor)
    slope_label = find_label_cell(cache, ["slope"], near=anchor)
    total_fcst_label = find_label_cell(
        cache,
        ["tot fcst w/o sa", "tot fcst without sa", "total forecast without sa"],
        near=anchor,
    )
    actual_label = find_label_cell(cache, ["reported sales", "actual sales"], near=anchor)
    min_label = find_label_cell(cache, ["min"], near=anchor)
    if min_label is None:
        min_label = (anchor_row + 1, anchor_col)
    if total_fcst_label is None:
        total_fcst_label = (anchor_row - 2, anchor_col)

    intercept_cell = pick_or_default_value_cell(
        sheet=sheet,
        anchor_row=anchor_row,
        anchor_col=anchor_col,
        label_cell=intercept_label,
        default_offset=(0, 4),
    )
    slope_cell = pick_or_default_value_cell(
        sheet=sheet,
        anchor_row=anchor_row,
        anchor_col=anchor_col,
        label_cell=slope_label,
        default_offset=(1, 4),
    )

    series_rows = contiguous_xy_rows(sheet=sheet, x_col=x_col, y_col=y_col, anchor_row=anchor_row)
    if len(series_rows) < 2:
        print("  - Skipped regression extraction: insufficient x/y history rows")
        return []

    max_quarters = min(10, len(series_rows))
    output_rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for num_quarters_used in range(2, max_quarters + 1):
        selected_rows = series_rows[-num_quarters_used:]
        start_row = selected_rows[0]
        end_row = selected_rows[-1]

        intercept_cell.formula2 = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_cell.formula2 = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        workbook.app.calculate()

        intercept_value = intercept_cell.value
        slope_value = slope_cell.value
        forecast_value = read_value_from_label(sheet, total_fcst_label)
        actual_value = read_value_from_label(sheet, actual_label)
        forecast_max = read_value_from_label(sheet, anchor, prefer_numeric=True)
        forecast_min = read_value_from_label(sheet, min_label, prefer_numeric=True)

        current_signature = rounded_signature(
            [
                num_quarters_used,
                intercept_value,
                slope_value,
                forecast_value,
                forecast_max,
                forecast_min,
            ]
        )
        if current_signature == previous_signature:
            continue
        previous_signature = current_signature

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
            "range_width": compute_range_width(forecast_max, forecast_min),
            "intercept": intercept_value,
            "slope": slope_value,
            "source_file": source_file,
        }
        output_rows.append(row)

    return output_rows


def next_output_path(in_dir: Path, out_dir: Path) -> Path:
    base_name = f"{in_dir.name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    index = 1
    while candidate.exists():
        candidate = out_dir / f"{base_name}.{index}.xlsx"
        index += 1
    return candidate


def write_sheet_rows(
    worksheet,
    headers: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    worksheet.append(list(headers))
    for row in rows:
        worksheet.append([row.get(header) for header in headers])

    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for col_idx, header in enumerate(headers, start=1):
        longest = len(header)
        for row_idx in range(2, worksheet.max_row + 1):
            value = worksheet.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            longest = max(longest, len(str(value)))
        worksheet.column_dimensions[get_column_letter(col_idx)].width = min(max(longest + 2, 12), 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    empirical_ws = workbook.active
    empirical_ws.title = "empirical_candidates"
    regression_ws = workbook.create_sheet("regression_candidates")

    write_sheet_rows(empirical_ws, EMPIRICAL_HEADERS, empirical_rows)
    write_sheet_rows(regression_ws, REGRESSION_HEADERS, regression_rows)

    workbook.save(output_path)


def create_excel_app() -> xw.App:
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        app.enable_events = False
    except Exception:
        pass
    try:
        app.calculation = "manual"
    except Exception:
        pass
    return app


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"Input directory does not exist: {input_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = next_output_path(input_dir, output_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = create_excel_app()
    try:
        for file_path in sorted(input_dir.iterdir(), key=lambda p: p.name.lower()):
            if not file_path.is_file():
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped file: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped file: {file_path.name} (not .xlsx)")
                continue

            print(f"Processing file: {file_path.name}")
            workbook = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                file_labels = parse_file_labels(file_path)

                empirical_rows.extend(
                    extract_empirical_rows(
                        workbook=workbook,
                        labels=file_labels,
                        source_file=file_path.name,
                    )
                )
                regression_rows.extend(
                    extract_regression_rows(
                        workbook=workbook,
                        labels=file_labels,
                        source_file=file_path.name,
                    )
                )
                processed_files += 1
            except Exception as exc:
                print(f"  - Skipped file: {file_path.name} (processing error: {exc})")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        app.quit()

    write_output_workbook(
        output_path=output_path,
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
    )

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

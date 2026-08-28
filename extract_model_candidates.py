from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# =====================
# User-configurable I/O
# =====================
input_dir = "/path/to/input"
output_dir = "/path/to/output"

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
MONTH_BY_ABBREV = {
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
class FileMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetScan:
    used_row: int
    used_col: int
    values: List[List[Any]]
    labels: Dict[str, List[Tuple[int, int]]]
    anchor_max: Optional[Tuple[int, int]]


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    return " ".join(text.split())


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def to_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values
    return [[item] for item in values]


def parse_file_metadata(file_name: str) -> FileMetadata:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split("-")]

    ticker = ""
    if len(parts) >= 2:
        ticker = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper()

    raw_period_token = ""
    if len(parts) >= 3:
        raw_period_token = parts[2].split("_")[0].strip()

    period_match = re.search(
        r"(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})",
        raw_period_token,
        flags=re.IGNORECASE,
    )
    if not period_match:
        period_match = re.search(
            r"(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})",
            stem,
            flags=re.IGNORECASE,
        )

    model_period = raw_period_token
    model_date = ""
    if period_match:
        period_bucket = period_match.group(1).title()
        month_token = period_match.group(2).strip()
        month_abbrev = month_token[:3].title()
        year_text = period_match.group(3)

        month_num = MONTH_BY_ABBREV.get(month_abbrev.lower())
        if month_num is not None:
            model_period = f"{period_bucket}{month_abbrev}_{year_text}"
            day = DAY_BY_PERIOD[period_bucket.lower()]
            model_date = date(int(year_text), month_num, day).isoformat()

    model = "_".join(item for item in [ticker, model_period] if item)
    return FileMetadata(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def unique_output_path(input_path: Path, output_path: Path) -> Path:
    input_folder_name = input_path.name or input_path.resolve().name
    base_name = f"{input_folder_name}_PARAM"

    candidate = output_path / f"{base_name}.xlsx"
    idx = 1
    while candidate.exists():
        candidate = output_path / f"{base_name}.{idx}.xlsx"
        idx += 1
    return candidate


def safe_close_source_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.close(False)
        return
    except Exception:
        pass

    try:
        wb.app.api.DisplayAlerts = False
        wb.api.Close(SaveChanges=False)
    except Exception:
        pass


def set_formula2(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        # Fallback keeps script resilient on older Excel builds.
        cell.formula = formula_r1c1


def scan_sheet_once(sheet: xw.Sheet) -> SheetScan:
    used = sheet.used_range
    values_2d = to_2d(used.value)
    used_row = used.row
    used_col = used.column

    labels: Dict[str, List[Tuple[int, int]]] = {}
    anchor_max: Optional[Tuple[int, int]] = None

    for row_idx, row_vals in enumerate(values_2d):
        for col_idx, raw_val in enumerate(row_vals):
            if not isinstance(raw_val, str):
                continue
            label = normalize_label(raw_val)
            if not label:
                continue

            abs_row = used_row + row_idx
            abs_col = used_col + col_idx
            labels.setdefault(label, []).append((abs_row, abs_col))

            if anchor_max is None and label == "max":
                anchor_max = (abs_row, abs_col)

    return SheetScan(
        used_row=used_row,
        used_col=used_col,
        values=values_2d,
        labels=labels,
        anchor_max=anchor_max,
    )


def distance(a: Tuple[int, int], b: Tuple[int, int]) -> int:
    return abs(a[0] - b[0]) + abs(a[1] - b[1])


def find_label_cells(scan: SheetScan, patterns: Sequence[str]) -> List[Tuple[int, int]]:
    matches: List[Tuple[int, int]] = []
    lowered = [normalize_label(p) for p in patterns]
    for label, cells in scan.labels.items():
        if any(pattern in label for pattern in lowered):
            matches.extend(cells)
    return matches


def pick_adjacent_value_cell(
    sheet: xw.Sheet,
    label_cell: Tuple[int, int],
    prefer_numeric: bool = False,
) -> Tuple[int, int]:
    row, col = label_cell
    candidates = [
        (row, col + 1),
        (row + 1, col),
        (row, col - 1),
        (row - 1, col),
    ]

    fallback = label_cell
    for r, c in candidates:
        val = sheet.cells(r, c).value
        if val is None or val == "":
            continue
        if prefer_numeric and to_float(val) is None:
            continue
        return (r, c)
    return fallback


def locate_value_cell_by_labels(
    sheet: xw.Sheet,
    scan: SheetScan,
    patterns: Sequence[str],
    anchor: Optional[Tuple[int, int]],
    prefer_numeric: bool = False,
) -> Optional[Tuple[int, int]]:
    label_cells = find_label_cells(scan, patterns)
    if not label_cells:
        return None

    if anchor is not None:
        label_cells.sort(key=lambda cell: distance(cell, anchor))

    for label_cell in label_cells:
        return pick_adjacent_value_cell(sheet, label_cell, prefer_numeric=prefer_numeric)
    return None


def locate_max_value_cell(sheet: xw.Sheet, anchor: Tuple[int, int]) -> Tuple[int, int]:
    row, col = anchor
    candidates = [
        (row, col + 1),
        (row + 1, col),
        (row + 1, col + 1),
        (row, col),
    ]
    for r, c in candidates:
        if to_float(sheet.cells(r, c).value) is not None:
            return (r, c)
    return (row, col + 1)


def locate_min_value_cell(
    sheet: xw.Sheet,
    scan: SheetScan,
    anchor: Tuple[int, int],
) -> Tuple[int, int]:
    min_label_cells = find_label_cells(scan, ["min"])
    if min_label_cells:
        min_label_cells.sort(key=lambda cell: distance(cell, anchor))
        return pick_adjacent_value_cell(sheet, min_label_cells[0], prefer_numeric=True)

    row, col = anchor
    candidates = [
        (row + 1, col),
        (row + 1, col + 1),
        (row + 2, col),
        (row + 2, col + 1),
    ]
    for r, c in candidates:
        if to_float(sheet.cells(r, c).value) is not None:
            return (r, c)
    return (row + 1, col)


def collect_numeric_series_upward(
    sheet: xw.Sheet,
    end_row: int,
    col: int,
    hard_limit: int = 120,
) -> List[Tuple[int, float]]:
    series: List[Tuple[int, float]] = []
    blank_streak = 0

    for row in range(end_row, max(0, end_row - hard_limit), -1):
        value = sheet.cells(row, col).value
        number = to_float(value)

        if number is None:
            if series:
                blank_streak += 1
                if blank_streak >= 2:
                    break
            continue

        blank_streak = 0
        series.append((row, number))

    series.reverse()
    return series


def get_cell_value(sheet: xw.Sheet, coord: Optional[Tuple[int, int]]) -> Any:
    if coord is None:
        return ""
    row, col = coord
    return sheet.cells(row, col).value


def calc_range_width(max_val: Any, min_val: Any) -> Any:
    max_num = to_float(max_val)
    min_num = to_float(min_val)
    if max_num is None or min_num is None:
        return ""
    return max_num - min_num


def extract_empirical_rows(
    wb: xw.Book,
    metadata: FileMetadata,
    source_file: str,
) -> List[Dict[str, Any]]:
    if "Empirical Model" not in [s.name for s in wb.sheets]:
        print(f"Skipped empirical extraction for {source_file}: sheet 'Empirical Model' not found")
        return []

    sheet = wb.sheets["Empirical Model"]
    scan = scan_sheet_once(sheet)
    if scan.anchor_max is None:
        print(f"Skipped empirical extraction for {source_file}: 'max' anchor not found")
        return []

    anchor = scan.anchor_max
    anchor_row, anchor_col = anchor

    avg_pen_cell = locate_value_cell_by_labels(
        sheet,
        scan,
        patterns=["avg penetration", "average penetration"],
        anchor=anchor,
        prefer_numeric=False,
    )
    if avg_pen_cell is None:
        avg_pen_cell = (anchor_row - 1, max(1, anchor_col - 2))

    penetration_col = avg_pen_cell[1]
    penetration_series = collect_numeric_series_upward(
        sheet=sheet,
        end_row=avg_pen_cell[0] - 1,
        col=penetration_col,
    )

    if not penetration_series:
        print(f"Skipped empirical extraction for {source_file}: no penetration history found")
        return []

    max_cell = locate_max_value_cell(sheet, anchor)
    min_cell = locate_min_value_cell(sheet, scan, anchor)

    forecast_cell = locate_value_cell_by_labels(
        sheet,
        scan,
        patterns=["estimated total sold", "estimate total sold", "total sold estimate"],
        anchor=anchor,
        prefer_numeric=False,
    )
    if forecast_cell is None:
        forecast_cell = (anchor_row, max(1, anchor_col - 1))

    reported_sales_cell = locate_value_cell_by_labels(
        sheet,
        scan,
        patterns=["reported sales", "actual sales", "actual value"],
        anchor=anchor,
        prefer_numeric=False,
    )
    actual_cell = reported_sales_cell

    quarterly_sales_cell = locate_value_cell_by_labels(
        sheet,
        scan,
        patterns=["quarterly sales", "quarter sales"],
        anchor=anchor,
        prefer_numeric=False,
    )

    growth_rate_cell = locate_value_cell_by_labels(
        sheet,
        scan,
        patterns=["growth rate", "growth %"],
        anchor=anchor,
        prefer_numeric=False,
    )

    captured_pct_cell = locate_value_cell_by_labels(
        sheet,
        scan,
        patterns=["sales captured in db", "captured in db", "captured %"],
        anchor=anchor,
        prefer_numeric=False,
    )

    rows: List[Dict[str, Any]] = []
    loop_count = min(10, len(penetration_series))

    for n_quarters in range(1, loop_count + 1):
        start_row = penetration_series[-n_quarters][0]
        end_row = penetration_series[-1][0]

        avg_formula = f"=AVERAGE(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})"
        set_formula2(sheet.cells(*avg_pen_cell), avg_formula)
        wb.app.calculate()

        avg_penetration_pct = get_cell_value(sheet, avg_pen_cell)
        forecast_value = get_cell_value(sheet, forecast_cell)
        actual_value = get_cell_value(sheet, actual_cell)
        forecast_max = get_cell_value(sheet, max_cell)
        forecast_min = get_cell_value(sheet, min_cell)

        last_quarter_used = ""
        if penetration_col > 1:
            last_quarter_used = sheet.cells(start_row, penetration_col - 1).value

        row = {
            "model": metadata.model,
            "ticker": metadata.ticker,
            "model_period": metadata.model_period,
            "model_date": metadata.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration_pct,
            "num_quarters_used": n_quarters,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": calc_range_width(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration_pct,
            "quarterly_sales": get_cell_value(sheet, quarterly_sales_cell),
            "reported_sales": get_cell_value(sheet, reported_sales_cell),
            "growth_rate_pct": get_cell_value(sheet, growth_rate_cell),
            "sales_captured_in_db_pct": get_cell_value(sheet, captured_pct_cell),
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def extract_regression_rows(
    wb: xw.Book,
    metadata: FileMetadata,
    source_file: str,
) -> List[Dict[str, Any]]:
    if "Regression Model" not in [s.name for s in wb.sheets]:
        print(f"Skipped regression extraction for {source_file}: sheet 'Regression Model' not found")
        return []

    sheet = wb.sheets["Regression Model"]
    scan = scan_sheet_once(sheet)
    if scan.anchor_max is None:
        print(f"Skipped regression extraction for {source_file}: 'max' anchor not found")
        return []

    anchor = scan.anchor_max
    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    if x_col < 1 or y_col < 1:
        print(f"Skipped regression extraction for {source_file}: invalid x/y columns from anchor")
        return []

    top_row = scan.used_row
    bottom_row = anchor_row - 1
    if bottom_row < top_row:
        print(f"Skipped regression extraction for {source_file}: no data rows above anchor")
        return []

    x_values_raw = sheet.range((top_row, x_col), (bottom_row, x_col)).value
    y_values_raw = sheet.range((top_row, y_col), (bottom_row, y_col)).value

    x_values = x_values_raw if isinstance(x_values_raw, list) else [x_values_raw]
    y_values = y_values_raw if isinstance(y_values_raw, list) else [y_values_raw]

    data_rows: List[Tuple[int, float, float]] = []
    for idx, (x_val, y_val) in enumerate(zip(x_values, y_values)):
        x_num = to_float(x_val)
        y_num = to_float(y_val)
        if x_num is None or y_num is None:
            continue
        data_rows.append((top_row + idx, x_num, y_num))

    if len(data_rows) < 2:
        print(f"Skipped regression extraction for {source_file}: fewer than 2 valid x/y rows")
        return []

    max_cell = locate_max_value_cell(sheet, anchor)
    min_cell = locate_min_value_cell(sheet, scan, anchor)
    actual_cell = locate_value_cell_by_labels(
        sheet,
        scan,
        patterns=["actual value", "actual sales", "reported sales"],
        anchor=anchor,
        prefer_numeric=False,
    )

    intercept_cell = (anchor_row + 2, anchor_col + 2)
    slope_cell = (anchor_row + 3, anchor_col + 2)
    forecast_cell = (anchor_row + 4, anchor_col + 2)

    max_n = min(10, len(data_rows))
    previous_signature: Optional[Tuple[Any, ...]] = None
    rows: List[Dict[str, Any]] = []

    for n_quarters in range(2, max_n + 1):
        subset = data_rows[-n_quarters:]
        r1 = subset[0][0]
        r2 = subset[-1][0]

        intercept_formula = (
            f"=INTERCEPT(R{r1}C{y_col}:R{r2}C{y_col},R{r1}C{x_col}:R{r2}C{x_col})"
        )
        slope_formula = f"=SLOPE(R{r1}C{y_col}:R{r2}C{y_col},R{r1}C{x_col}:R{r2}C{x_col})"
        forecast_formula = (
            f"=R{intercept_cell[0]}C{intercept_cell[1]}+"
            f"(R{slope_cell[0]}C{slope_cell[1]}*R{r2}C{x_col})"
        )

        set_formula2(sheet.cells(*intercept_cell), intercept_formula)
        set_formula2(sheet.cells(*slope_cell), slope_formula)
        set_formula2(sheet.cells(*forecast_cell), forecast_formula)
        wb.app.calculate()

        intercept = get_cell_value(sheet, intercept_cell)
        slope = get_cell_value(sheet, slope_cell)
        forecast_value = get_cell_value(sheet, forecast_cell)
        forecast_max = get_cell_value(sheet, max_cell)
        forecast_min = get_cell_value(sheet, min_cell)

        if to_float(forecast_max) is None:
            forecast_max = max(point[2] for point in subset)
        if to_float(forecast_min) is None:
            forecast_min = min(point[2] for point in subset)

        signature = (
            round(to_float(forecast_value) or 0.0, 8),
            round(to_float(forecast_max) or 0.0, 8),
            round(to_float(forecast_min) or 0.0, 8),
            round(to_float(intercept) or 0.0, 8),
            round(to_float(slope) or 0.0, 8),
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        row = {
            "model": metadata.model,
            "ticker": metadata.ticker,
            "model_period": metadata.model_period,
            "model_date": metadata.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n_quarters,
            "num_quarters_used": n_quarters,
            "forecast_value": forecast_value,
            "actual_value": get_cell_value(sheet, actual_cell),
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": calc_range_width(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def apply_output_format(sheet, headers: Sequence[str]) -> None:
    for col_idx, header in enumerate(headers, start=1):
        sheet.cell(row=1, column=col_idx, value=header)
        sheet.cell(row=1, column=col_idx).font = Font(bold=True)

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    for col_idx in range(1, len(headers) + 1):
        max_len = len(headers[col_idx - 1])
        for row_idx in range(2, sheet.max_row + 1):
            cell_value = sheet.cell(row=row_idx, column=col_idx).value
            if cell_value is None:
                continue
            max_len = max(max_len, len(str(cell_value)))
        sheet.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 42)


def write_rows(sheet, headers: Sequence[str], rows: Iterable[Dict[str, Any]]) -> None:
    for row_dict in rows:
        sheet.append([row_dict.get(header, "") for header in headers])


def gather_input_files(input_path: Path) -> Tuple[List[Path], int]:
    process_list: List[Path] = []
    skipped = 0

    for file_path in sorted(input_path.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped file {file_path.name}: temporary file")
            skipped += 1
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped file {file_path.name}: not .xlsx")
            skipped += 1
            continue
        process_list.append(file_path)

    return process_list, skipped


def build_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    out_wb = Workbook()
    default_sheet = out_wb.active
    out_wb.remove(default_sheet)

    empirical_sheet = out_wb.create_sheet("empirical_candidates")
    regression_sheet = out_wb.create_sheet("regression_candidates")

    apply_output_format(empirical_sheet, EMPIRICAL_COLUMNS)
    apply_output_format(regression_sheet, REGRESSION_COLUMNS)

    write_rows(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)
    write_rows(regression_sheet, REGRESSION_COLUMNS, regression_rows)

    # Refresh filters now that rows are appended.
    empirical_sheet.auto_filter.ref = empirical_sheet.dimensions
    regression_sheet.auto_filter.ref = regression_sheet.dimensions

    # Recompute widths with data included.
    apply_output_format(empirical_sheet, EMPIRICAL_COLUMNS)
    apply_output_format(regression_sheet, REGRESSION_COLUMNS)

    out_wb.save(output_path)


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()
    output_path.mkdir(parents=True, exist_ok=True)

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"Input directory does not exist: {input_path}")

    source_files, _ = gather_input_files(input_path)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    previous_calc_mode = None
    try:
        previous_calc_mode = app.calculation
        app.calculation = "manual"
    except Exception:
        previous_calc_mode = None

    try:
        for file_path in source_files:
            print(f"Processing file {file_path.name}")
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                metadata = parse_file_metadata(file_path.name)
                empirical_rows.extend(extract_empirical_rows(wb, metadata, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, metadata, file_path.name))
                processed_files += 1
            except Exception as exc:
                print(f"Skipped file {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    safe_close_source_workbook(wb)
    finally:
        try:
            if previous_calc_mode is not None:
                app.calculation = previous_calc_mode
        except Exception:
            pass
        app.quit()

    final_output_path = unique_output_path(input_path, output_path)
    build_output_workbook(final_output_path, empirical_rows, regression_rows)

    print(f"Output path: {final_output_path}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

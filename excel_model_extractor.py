from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# User-editable paths
input_dir = "./input"
output_dir = "./output"

N_QUARTERS = 10

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

DAY_BY_PHASE = {
    "early": 5,
    "mid": 15,
    "late": 25,
}


@dataclass
class ModelMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetScan:
    values: List[List[Any]]
    base_row: int
    base_col: int
    last_row: int
    last_col: int
    label_positions: Dict[str, List[Tuple[int, int]]]
    max_anchor: Optional[Tuple[int, int]]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip().lower())


def ensure_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, (list, tuple)):
        return [[values]]
    if len(values) == 0:
        return []
    first = values[0]
    if isinstance(first, (list, tuple)):
        return [list(row) if isinstance(row, (list, tuple)) else [row] for row in values]
    return [list(values)]


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip().replace(",", "")
    if not text:
        return None
    if text.endswith("%"):
        try:
            return float(text[:-1]) / 100.0
        except ValueError:
            return None
    try:
        return float(text)
    except ValueError:
        return None


def safe_subtract(a: Any, b: Any) -> Optional[float]:
    a_num = to_float(a)
    b_num = to_float(b)
    if a_num is None or b_num is None:
        return None
    return a_num - b_num


def safe_col(col: Optional[int]) -> Optional[int]:
    if col is None:
        return None
    return max(1, int(col))


def round_sig(value: Any, digits: int = 8) -> Any:
    number = to_float(value)
    if number is None:
        return value
    return round(number, digits)


def parse_filename_metadata(file_path: Path) -> ModelMetadata:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ") if part.strip()]

    ticker = "UNKNOWN"
    if len(parts) >= 2:
        ticker_match = re.search(r"[A-Za-z]{1,10}", parts[1])
        if ticker_match:
            ticker = ticker_match.group(0).upper()
    else:
        fallback_ticker = re.search(r"\b([A-Z]{2,6})\b", stem)
        if fallback_ticker:
            ticker = fallback_ticker.group(1).upper()

    period_source = parts[2] if len(parts) >= 3 else stem
    period_source = period_source.split("_")[0]

    period_match = re.search(
        r"(?i)\b(early|mid|late)\s*[-_ ]*([A-Za-z]{3,9})\s*[-_ ]*(\d{4})\b",
        period_source,
    )
    if period_match is None:
        period_match = re.search(r"(?i)(early|mid|late)([A-Za-z]{3,9})(\d{4})", period_source)

    model_period = "Unknown_0000"
    model_date = ""
    if period_match:
        phase = period_match.group(1).lower()
        month_token = period_match.group(2)
        year = period_match.group(3)
        month_abbr = month_token[:3].title()
        try:
            month_number = datetime.strptime(month_abbr, "%b").month
            day = DAY_BY_PHASE[phase]
            model_period = f"{phase.capitalize()}{month_abbr}_{year}"
            model_date = f"{year}-{month_number:02d}-{day:02d}"
        except (ValueError, KeyError):
            model_period = "Unknown_0000"
            model_date = ""

    model = f"{ticker}_{model_period}"
    return ModelMetadata(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def next_output_path(out_dir: Path, input_folder_name: str) -> Path:
    base = out_dir / f"{input_folder_name}_PARAM.xlsx"
    if not base.exists():
        return base

    index = 1
    while True:
        candidate = out_dir / f"{input_folder_name}_PARAM.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def scan_sheet(sheet: xw.Sheet) -> SheetScan:
    used_range = sheet.used_range
    values = ensure_2d(used_range.value)
    base_row = used_range.row
    base_col = used_range.column

    max_cols = max((len(row) for row in values), default=1)
    last_row = base_row + max(len(values), 1) - 1
    last_col = base_col + max_cols - 1

    label_positions: Dict[str, List[Tuple[int, int]]] = {}
    max_anchor: Optional[Tuple[int, int]] = None

    for row_offset, row_values in enumerate(values):
        for col_offset, cell_value in enumerate(row_values):
            if isinstance(cell_value, str):
                normalized = normalize_text(cell_value)
                if not normalized:
                    continue
                abs_row = base_row + row_offset
                abs_col = base_col + col_offset
                label_positions.setdefault(normalized, []).append((abs_row, abs_col))
                if normalized == "max" and max_anchor is None:
                    max_anchor = (abs_row, abs_col)

    if max_anchor is None:
        for label, positions in label_positions.items():
            if re.search(r"\bmax\b", label):
                max_anchor = positions[0]
                break

    return SheetScan(
        values=values,
        base_row=base_row,
        base_col=base_col,
        last_row=last_row,
        last_col=last_col,
        label_positions=label_positions,
        max_anchor=max_anchor,
    )


def get_scan_value(scan: SheetScan, row: int, col: int) -> Any:
    row_idx = row - scan.base_row
    col_idx = col - scan.base_col

    if row_idx < 0 or col_idx < 0:
        return None
    if row_idx >= len(scan.values):
        return None
    row_values = scan.values[row_idx]
    if col_idx >= len(row_values):
        return None
    return row_values[col_idx]


def find_column(
    scan: SheetScan,
    keywords: Sequence[str],
    anchor_row: int,
    anchor_col: int,
    fallback: Optional[int],
) -> Optional[int]:
    normalized_keywords = [normalize_text(keyword) for keyword in keywords]
    best_col: Optional[int] = None
    best_score: Optional[Tuple[int, int]] = None

    for label, positions in scan.label_positions.items():
        if not any(keyword in label for keyword in normalized_keywords):
            continue
        for row, col in positions:
            score = (abs(row - anchor_row), abs(col - anchor_col))
            if best_score is None or score < best_score:
                best_score = score
                best_col = col

    return safe_col(best_col if best_col is not None else fallback)


def numeric_rows(scan: SheetScan, col: Optional[int], start_row: int) -> List[int]:
    if col is None:
        return []
    rows: List[int] = []
    for row in range(start_row, scan.last_row + 1):
        if to_float(get_scan_value(scan, row, col)) is not None:
            rows.append(row)
    return rows


def set_formula2_r1c1(target_range: xw.Range, formula_r1c1: str) -> None:
    try:
        target_range.formula2 = formula_r1c1
        return
    except Exception:
        pass

    try:
        target_range.api.Formula2R1C1 = formula_r1c1
        return
    except Exception:
        target_range.api.FormulaR1C1 = formula_r1c1


def close_workbook_safely(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.close(SaveChanges=False)
        return
    except Exception:
        pass

    workbook.close()


def read_formula_results(sheet: xw.Sheet, start_row: int, col: int, count: int) -> List[Any]:
    values: List[Any] = []
    for offset in range(count):
        values.append(sheet.range((start_row + offset, col)).value)
    return values


def build_empirical_rows(
    workbook: xw.Book,
    metadata: ModelMetadata,
    source_file: str,
) -> List[List[Any]]:
    if "Empirical Model" not in [sheet.name for sheet in workbook.sheets]:
        return []

    sheet = workbook.sheets["Empirical Model"]
    scan = scan_sheet(sheet)
    if scan.max_anchor is None:
        return []

    anchor_row, anchor_col = scan.max_anchor
    row_start = anchor_row + 1

    num_quarters_col = find_column(
        scan,
        ["num quarters used", "num quarters", "quarters used"],
        anchor_row,
        anchor_col,
        fallback=None,
    )
    penetration_col = find_column(
        scan,
        ["avg penetration", "penetration"],
        anchor_row,
        anchor_col,
        fallback=max(1, anchor_col - 6),
    )
    last_quarter_col = find_column(
        scan,
        ["last quarter used", "last quarter"],
        anchor_row,
        anchor_col,
        fallback=max(1, anchor_col - 5),
    )
    forecast_value_col = find_column(
        scan,
        ["estimated total sold", "forecast value", "forecast"],
        anchor_row,
        anchor_col,
        fallback=max(1, anchor_col - 1),
    )
    actual_value_col = find_column(
        scan,
        ["actual value", "reported sales", "actual"],
        anchor_row,
        anchor_col,
        fallback=max(1, anchor_col + 2),
    )
    forecast_max_col = anchor_col
    forecast_min_col = find_column(
        scan,
        ["forecast min", "min"],
        anchor_row,
        anchor_col,
        fallback=max(1, anchor_col + 1),
    )
    quarterly_sales_col = find_column(
        scan,
        ["quarterly sales"],
        anchor_row,
        anchor_col,
        fallback=max(1, anchor_col - 4),
    )
    reported_sales_col = find_column(
        scan,
        ["reported sales"],
        anchor_row,
        anchor_col,
        fallback=actual_value_col,
    )
    growth_rate_col = find_column(
        scan,
        ["growth rate"],
        anchor_row,
        anchor_col,
        fallback=max(1, anchor_col - 3),
    )
    sales_captured_col = find_column(
        scan,
        ["sales captured in db", "captured in db", "captured"],
        anchor_row,
        anchor_col,
        fallback=max(1, anchor_col - 2),
    )

    penetration_rows = numeric_rows(scan, penetration_col, row_start)
    data_end = penetration_rows[-1] if penetration_rows else row_start + N_QUARTERS - 1

    helper_col = scan.last_col + 2
    for n in range(1, N_QUARTERS + 1):
        formula_row = row_start + (n - 1)
        data_start = max(row_start, data_end - n + 1)
        formula = f"=AVERAGE(R{data_start}C{penetration_col}:R{data_end}C{penetration_col})"
        set_formula2_r1c1(sheet.range((formula_row, helper_col)), formula)

    workbook.app.calculate()
    avg_penetration_values = read_formula_results(sheet, row_start, helper_col, N_QUARTERS)

    rows: List[List[Any]] = []
    for n in range(1, N_QUARTERS + 1):
        row = row_start + (n - 1)

        num_quarters_used = get_scan_value(scan, row, num_quarters_col) if num_quarters_col else n
        if to_float(num_quarters_used) is None:
            num_quarters_used = n

        last_quarter_used = get_scan_value(scan, row, last_quarter_col)
        forecast_value = get_scan_value(scan, row, forecast_value_col)
        actual_value = get_scan_value(scan, row, actual_value_col)
        forecast_max = get_scan_value(scan, row, forecast_max_col)
        forecast_min = get_scan_value(scan, row, forecast_min_col)
        range_width = safe_subtract(forecast_max, forecast_min)
        avg_penetration = avg_penetration_values[n - 1]
        quarterly_sales = get_scan_value(scan, row, quarterly_sales_col)
        reported_sales = get_scan_value(scan, row, reported_sales_col)
        growth_rate = get_scan_value(scan, row, growth_rate_col)
        sales_captured = get_scan_value(scan, row, sales_captured_col)

        has_values = any(
            value not in (None, "")
            for value in [
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                avg_penetration,
                quarterly_sales,
            ]
        )
        if not has_values:
            continue

        rows.append(
            [
                metadata.model,
                metadata.ticker,
                metadata.model_period,
                metadata.model_date,
                "empirical",
                "avg_penetration_pct",
                avg_penetration,
                num_quarters_used,
                last_quarter_used,
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                range_width,
                avg_penetration,
                quarterly_sales,
                reported_sales,
                growth_rate,
                sales_captured,
                source_file,
            ]
        )

    return rows


def build_regression_rows(
    workbook: xw.Book,
    metadata: ModelMetadata,
    source_file: str,
) -> List[List[Any]]:
    if "Regression Model" not in [sheet.name for sheet in workbook.sheets]:
        return []

    sheet = workbook.sheets["Regression Model"]
    scan = scan_sheet(sheet)
    if scan.max_anchor is None:
        return []

    anchor_row, anchor_col = scan.max_anchor
    row_start = anchor_row + 1

    y_col = max(1, anchor_col - 7)
    x_col = max(1, anchor_col - 11)

    xy_rows: List[int] = []
    for row in range(row_start, scan.last_row + 1):
        x_value = get_scan_value(scan, row, x_col)
        y_value = get_scan_value(scan, row, y_col)
        if to_float(x_value) is not None and to_float(y_value) is not None:
            xy_rows.append(row)
    data_end = xy_rows[-1] if xy_rows else row_start + N_QUARTERS - 1

    helper_intercept_col = scan.last_col + 2
    helper_slope_col = helper_intercept_col + 1

    for n in range(1, N_QUARTERS + 1):
        formula_row = row_start + (n - 1)
        data_start = max(row_start, data_end - n + 1)
        intercept_formula = (
            f"=INTERCEPT(R{data_start}C{y_col}:R{data_end}C{y_col},"
            f"R{data_start}C{x_col}:R{data_end}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{data_start}C{y_col}:R{data_end}C{y_col},"
            f"R{data_start}C{x_col}:R{data_end}C{x_col})"
        )
        set_formula2_r1c1(sheet.range((formula_row, helper_intercept_col)), intercept_formula)
        set_formula2_r1c1(sheet.range((formula_row, helper_slope_col)), slope_formula)

    workbook.app.calculate()
    intercept_values = read_formula_results(sheet, row_start, helper_intercept_col, N_QUARTERS)
    slope_values = read_formula_results(sheet, row_start, helper_slope_col, N_QUARTERS)

    num_quarters_col = find_column(
        scan,
        ["num quarters used", "num quarters", "quarters used"],
        anchor_row,
        anchor_col,
        fallback=None,
    )
    forecast_value_col = find_column(
        scan,
        ["tot fcst w/o sa", "tot fcst without sa", "forecast value", "forecast"],
        anchor_row,
        anchor_col,
        fallback=max(1, anchor_col - 1),
    )
    actual_value_col = find_column(
        scan,
        ["actual value", "reported sales", "actual"],
        anchor_row,
        anchor_col,
        fallback=None,
    )
    forecast_max_col = anchor_col
    forecast_min_col = find_column(
        scan,
        ["forecast min", "min"],
        anchor_row,
        anchor_col,
        fallback=max(1, anchor_col + 1),
    )

    rows: List[List[Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for n in range(1, N_QUARTERS + 1):
        row = row_start + (n - 1)

        num_quarters_used = get_scan_value(scan, row, num_quarters_col) if num_quarters_col else n
        if to_float(num_quarters_used) is None:
            num_quarters_used = n

        forecast_value = get_scan_value(scan, row, forecast_value_col)
        actual_value = get_scan_value(scan, row, actual_value_col) if actual_value_col else None
        forecast_max = get_scan_value(scan, row, forecast_max_col)
        forecast_min = get_scan_value(scan, row, forecast_min_col)
        range_width = safe_subtract(forecast_max, forecast_min)
        intercept = intercept_values[n - 1]
        slope = slope_values[n - 1]

        has_values = any(
            value not in (None, "")
            for value in [forecast_value, forecast_max, forecast_min, intercept, slope]
        )
        if not has_values:
            continue

        signature = (
            round_sig(num_quarters_used),
            round_sig(forecast_value),
            round_sig(forecast_max),
            round_sig(forecast_min),
            round_sig(intercept),
            round_sig(slope),
        )
        if signature == previous_signature:
            continue
        previous_signature = signature

        rows.append(
            [
                metadata.model,
                metadata.ticker,
                metadata.model_period,
                metadata.model_date,
                "regression",
                "num_quarters_used",
                num_quarters_used,
                num_quarters_used,
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                range_width,
                intercept,
                slope,
                source_file,
            ]
        )

    return rows


def apply_table_formatting(sheet) -> None:
    sheet.freeze_panes = "A2"

    if sheet.max_row >= 1:
        for cell in sheet[1]:
            cell.font = Font(bold=True)

    if sheet.max_row > 1 and sheet.max_column > 0:
        sheet.auto_filter.ref = sheet.dimensions

    for col_idx in range(1, sheet.max_column + 1):
        max_len = 0
        for row_idx in range(1, sheet.max_row + 1):
            value = sheet.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        width = min(max(max_len + 2, 12), 60)
        sheet.column_dimensions[get_column_letter(col_idx)].width = width


def write_output_workbook(
    destination: Path,
    empirical_rows: Iterable[List[Any]],
    regression_rows: Iterable[List[Any]],
) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)

    empirical_sheet = workbook.create_sheet("empirical_candidates")
    empirical_sheet.append(EMPIRICAL_HEADERS)
    for row in empirical_rows:
        empirical_sheet.append(list(row))
    apply_table_formatting(empirical_sheet)

    regression_sheet = workbook.create_sheet("regression_candidates")
    regression_sheet.append(REGRESSION_HEADERS)
    for row in regression_rows:
        regression_sheet.append(list(row))
    apply_table_formatting(regression_sheet)

    workbook.save(destination)


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()
    output_path.mkdir(parents=True, exist_ok=True)

    if not input_path.exists():
        print(f"Skipped: {input_path} (input_dir does not exist)")
        return
    if not input_path.is_dir():
        print(f"Skipped: {input_path} (input_dir is not a folder)")
        return

    source_files: List[Path] = []
    for file_path in sorted(input_path.iterdir()):
        if not file_path.is_file():
            print(f"Skipped: {file_path.name} (not a file)")
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped: {file_path.name} (temporary file)")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped: {file_path.name} (not .xlsx)")
            continue
        source_files.append(file_path)

    destination = next_output_path(output_path, input_path.name)
    empirical_rows: List[List[Any]] = []
    regression_rows: List[List[Any]] = []
    processed_count = 0

    app: Optional[xw.App] = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        for file_path in source_files:
            print(f"Processing: {file_path.name}")
            workbook: Optional[xw.Book] = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                metadata = parse_filename_metadata(file_path)

                empirical_rows.extend(
                    build_empirical_rows(
                        workbook=workbook,
                        metadata=metadata,
                        source_file=file_path.name,
                    )
                )
                regression_rows.extend(
                    build_regression_rows(
                        workbook=workbook,
                        metadata=metadata,
                        source_file=file_path.name,
                    )
                )
                processed_count += 1
                print(f"Processed: {file_path.name}")
            except Exception as exc:
                print(f"Skipped: {file_path.name} (processing error: {exc})")
            finally:
                if workbook is not None:
                    close_workbook_safely(workbook)
    finally:
        if app is not None:
            app.quit()

    write_output_workbook(
        destination=destination,
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
    )

    print(f"Output path: {destination}")
    print(f"Files processed: {processed_count}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

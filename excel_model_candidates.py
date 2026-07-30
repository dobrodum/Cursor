from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
import re
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure these two paths before running.
input_dir = Path("input")
output_dir = Path("output")

N_QUARTERS = 10

EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"

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

PERIOD_DAY = {"early": 5, "mid": 15, "late": 25}

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


@dataclass(frozen=True)
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_header(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = text.replace("%", " pct ")
    text = text.replace("&", " and ")
    text = re.sub(r"[^a-z0-9]+", "_", text)
    return re.sub(r"_+", "_", text).strip("_")


def to_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def to_float(value: Any) -> Optional[float]:
    if is_number(value):
        return float(value)
    return None


def safe_subtract(left: Any, right: Any) -> Optional[float]:
    left_float = to_float(left)
    right_float = to_float(right)
    if left_float is None or right_float is None:
        return None
    return left_float - right_float


def parse_month(month_text: str) -> Tuple[str, int]:
    cleaned = re.sub(r"[^A-Za-z]", "", month_text).lower()
    if cleaned in MONTH_LOOKUP:
        month_number = MONTH_LOOKUP[cleaned]
        return cleaned[:3].title(), month_number
    prefix = cleaned[:3]
    if prefix in MONTH_LOOKUP:
        month_number = MONTH_LOOKUP[prefix]
        return prefix.title(), month_number
    raise ValueError(f"Unrecognized month token: {month_text}")


def parse_file_labels(file_path: Path) -> FileLabels:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = "UNKNOWN"
    if len(parts) >= 2 and parts[1]:
        ticker = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper() or "UNKNOWN"

    period_match = re.search(r"(Early|Mid|Late)([A-Za-z]+)(\d{4})", stem, flags=re.IGNORECASE)
    if not period_match:
        raise ValueError("Could not parse period token (Early/Mid/Late + Month + Year) from file name")

    period_band = period_match.group(1).title()
    month_token = period_match.group(2)
    year = int(period_match.group(3))

    month_abbrev, month_number = parse_month(month_token)
    day = PERIOD_DAY[period_band.lower()]
    model_period = f"{period_band}{month_abbrev}_{year}"
    model_date = date(year, month_number, day).isoformat()
    model = f"{ticker}_{model_period}"

    return FileLabels(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def next_output_path(in_dir: Path, out_dir: Path) -> Path:
    base = f"{in_dir.name}_PARAM"
    candidate = out_dir / f"{base}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = out_dir / f"{base}.{suffix}.xlsx"
        suffix += 1
    return candidate


def safe_close_workbook(wb: xw.Book) -> None:
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
        wb.api.Close(SaveChanges=False)
    except Exception as exc:
        print(f"Warning: workbook close fallback failed: {exc}")


def find_anchor(sheet: xw.Sheet, anchor_text: str = "max") -> Tuple[int, int]:
    used = sheet.used_range
    values = to_2d(used.value)
    anchor_lookup = anchor_text.strip().lower()

    for row_offset, row_values in enumerate(values):
        for col_offset, cell_value in enumerate(row_values):
            if isinstance(cell_value, str) and cell_value.strip().lower() == anchor_lookup:
                return used.row + row_offset, used.column + col_offset

    raise ValueError(f'Anchor text "{anchor_text}" not found in sheet "{sheet.name}"')


def build_header_map(sheet: xw.Sheet, header_row: int, start_col: int, end_col: int) -> Dict[str, int]:
    if start_col > end_col:
        return {}

    values = sheet.range((header_row, start_col), (header_row, end_col)).value
    if not isinstance(values, list):
        values = [values]

    headers: Dict[str, int] = {}
    for offset, value in enumerate(values):
        key = normalize_header(value)
        if key and key not in headers:
            headers[key] = start_col + offset
    return headers


def resolve_column(header_map: Dict[str, int], candidates: Sequence[str], fallback: int) -> int:
    for candidate in candidates:
        normalized = normalize_header(candidate)
        if normalized in header_map:
            return header_map[normalized]
    return fallback


def get_column_values(sheet: xw.Sheet, col_idx: int, start_row: int, end_row: int) -> List[Any]:
    if end_row < start_row:
        return []
    values = sheet.range((start_row, col_idx), (end_row, col_idx)).value
    if isinstance(values, list):
        return values
    return [values]


def get_numeric_rows(sheet: xw.Sheet, col_idx: int, start_row: int, end_row: int) -> List[int]:
    values = get_column_values(sheet, col_idx, start_row, end_row)
    numeric_rows: List[int] = []
    for i, value in enumerate(values, start=start_row):
        if is_number(value):
            numeric_rows.append(i)
    return numeric_rows


def set_r1c1_formula2(cell: xw.Range, formula: str) -> None:
    try:
        cell.api.Formula2R1C1 = formula
        return
    except Exception:
        pass

    try:
        cell.formula2 = formula
        return
    except Exception:
        pass

    cell.formula = formula


def try_sheet(wb: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    try:
        return wb.sheets[sheet_name]
    except Exception:
        return None


def process_empirical_sheet(
    wb: xw.Book,
    sheet: xw.Sheet,
    labels: FileLabels,
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor_row, anchor_col = find_anchor(sheet, "max")
    header_row = anchor_row
    header_map = build_header_map(sheet, header_row, max(1, anchor_col - 20), anchor_col + 20)

    forecast_max_col = resolve_column(header_map, ["max", "forecast_max"], anchor_col)
    forecast_min_col = resolve_column(header_map, ["min", "forecast_min"], anchor_col + 1)
    forecast_value_col = resolve_column(
        header_map,
        [
            "estimated_total_sold",
            "forecast_value",
            "tot_fcst_wo_sa",
            "tot_fcst_without_sa",
            "total_forecast_without_sa",
        ],
        anchor_col - 1,
    )
    actual_value_col = resolve_column(header_map, ["reported_sales", "actual_value", "actual_sales"], anchor_col - 2)
    quarterly_sales_col = resolve_column(header_map, ["quarterly_sales", "qtr_sales"], anchor_col - 3)
    growth_rate_col = resolve_column(header_map, ["growth_rate_pct", "growth_pct"], anchor_col - 4)
    sales_captured_col = resolve_column(
        header_map,
        ["sales_captured_in_db_pct", "captured_in_db_pct", "penetration_pct", "penetration"],
        anchor_col - 5,
    )
    quarter_label_col = resolve_column(header_map, ["last_quarter_used", "quarter", "quarter_used"], anchor_col - 6)
    penetration_source_col = resolve_column(
        header_map,
        ["penetration_pct", "avg_penetration_pct", "sales_captured_in_db_pct"],
        sales_captured_col,
    )
    num_quarters_col = resolve_column(header_map, ["num_quarters_used", "num_quarters"], anchor_col - 7)

    penetration_rows = get_numeric_rows(sheet, penetration_source_col, 1, max(1, anchor_row - 1))
    max_n = min(N_QUARTERS, len(penetration_rows))
    if max_n == 0:
        return []

    temp_avg_col = anchor_col + 25
    temp_start_row = anchor_row + 1

    for index in range(max_n):
        n_quarters = index + 1
        start_row = penetration_rows[-n_quarters]
        end_row = penetration_rows[-1]
        avg_formula = f"=AVERAGE(R{start_row}C{penetration_source_col}:R{end_row}C{penetration_source_col})"
        set_r1c1_formula2(sheet.range((temp_start_row + index, temp_avg_col)), avg_formula)

    wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    for index in range(max_n):
        row_num = anchor_row + 1 + index
        n_quarters = index + 1

        avg_penetration_pct = sheet.range((temp_start_row + index, temp_avg_col)).value
        num_quarters_used = sheet.range((row_num, num_quarters_col)).value
        if not is_number(num_quarters_used):
            num_quarters_used = n_quarters

        forecast_value = sheet.range((row_num, forecast_value_col)).value
        reported_sales = sheet.range((row_num, actual_value_col)).value
        forecast_max = sheet.range((row_num, forecast_max_col)).value
        forecast_min = sheet.range((row_num, forecast_min_col)).value
        quarterly_sales = sheet.range((row_num, quarterly_sales_col)).value
        growth_rate_pct = sheet.range((row_num, growth_rate_col)).value
        sales_captured_in_db_pct = sheet.range((row_num, sales_captured_col)).value
        last_quarter_used = sheet.range((row_num, quarter_label_col)).value
        if last_quarter_used in (None, "") and penetration_rows:
            last_quarter_used = sheet.range((penetration_rows[-1], quarter_label_col)).value

        range_width = safe_subtract(forecast_max, forecast_min)

        if all(
            value in (None, "")
            for value in (
                forecast_value,
                reported_sales,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
            )
        ):
            continue

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
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
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    return rows


def process_regression_sheet(
    wb: xw.Book,
    sheet: xw.Sheet,
    labels: FileLabels,
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor_row, anchor_col = find_anchor(sheet, "max")
    header_row = anchor_row
    header_map = build_header_map(sheet, header_row, max(1, anchor_col - 20), anchor_col + 20)

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    forecast_max_col = resolve_column(header_map, ["max", "forecast_max"], anchor_col)
    forecast_min_col = resolve_column(header_map, ["min", "forecast_min"], anchor_col + 1)
    forecast_value_col = resolve_column(
        header_map,
        ["tot_fcst_wo_sa", "tot_fcst_without_sa", "forecast_value", "total_forecast_without_sa"],
        anchor_col - 1,
    )
    actual_value_col = resolve_column(header_map, ["actual_value", "reported_sales", "actual_sales"], anchor_col - 2)
    num_quarters_col = resolve_column(header_map, ["num_quarters_used", "num_quarters"], anchor_col - 3)

    x_values = get_column_values(sheet, x_col, 1, max(1, anchor_row - 1))
    y_values = get_column_values(sheet, y_col, 1, max(1, anchor_row - 1))
    paired_rows: List[int] = []
    for row_idx, (x_val, y_val) in enumerate(zip(x_values, y_values), start=1):
        if is_number(x_val) and is_number(y_val):
            paired_rows.append(row_idx)

    max_n = min(N_QUARTERS, len(paired_rows))
    if max_n < 2:
        return []

    temp_base_row = anchor_row + 1
    temp_intercept_col = anchor_col + 25
    temp_slope_col = anchor_col + 26

    for index in range(max_n):
        n_quarters = index + 1
        start_row = paired_rows[-n_quarters]
        end_row = paired_rows[-1]
        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        set_r1c1_formula2(sheet.range((temp_base_row + index, temp_intercept_col)), intercept_formula)
        set_r1c1_formula2(sheet.range((temp_base_row + index, temp_slope_col)), slope_formula)

    wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None
    next_x_value = sheet.range((paired_rows[-1] + 1, x_col)).value if paired_rows else None

    for index in range(max_n):
        row_num = anchor_row + 1 + index
        n_quarters = index + 1

        intercept = sheet.range((temp_base_row + index, temp_intercept_col)).value
        slope = sheet.range((temp_base_row + index, temp_slope_col)).value

        forecast_value = sheet.range((row_num, forecast_value_col)).value
        if forecast_value in (None, "") and is_number(intercept) and is_number(slope) and is_number(next_x_value):
            forecast_value = float(intercept) + float(slope) * float(next_x_value)

        forecast_max = sheet.range((row_num, forecast_max_col)).value
        forecast_min = sheet.range((row_num, forecast_min_col)).value
        actual_value = sheet.range((row_num, actual_value_col)).value
        num_quarters_used = sheet.range((row_num, num_quarters_col)).value
        if not is_number(num_quarters_used):
            num_quarters_used = n_quarters

        range_width = safe_subtract(forecast_max, forecast_min)
        signature = (
            round(float(forecast_value), 8) if is_number(forecast_value) else forecast_value,
            round(float(forecast_max), 8) if is_number(forecast_max) else forecast_max,
            round(float(forecast_min), 8) if is_number(forecast_min) else forecast_min,
            round(float(intercept), 8) if is_number(intercept) else intercept,
            round(float(slope), 8) if is_number(slope) else slope,
        )

        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        if all(value in (None, "") for value in (forecast_value, forecast_max, forecast_min, intercept, slope)):
            continue

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
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
        )

    return rows


def append_rows(ws: Any, columns: Sequence[str], rows: Iterable[Dict[str, Any]]) -> None:
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(column) for column in columns])

    header_font = Font(bold=True)
    for cell in ws[1]:
        cell.font = header_font

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for cell in ws.iter_cols(min_col=col_idx, max_col=col_idx, min_row=2, values_only=True):
            for value in cell:
                if value is None:
                    continue
                max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    wb = Workbook()
    ws_empirical = wb.active
    ws_empirical.title = "empirical_candidates"
    append_rows(ws_empirical, EMPIRICAL_COLUMNS, empirical_rows)

    ws_regression = wb.create_sheet("regression_candidates")
    append_rows(ws_regression, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def process_workbook(app: xw.App, file_path: Path) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    labels = parse_file_labels(file_path)
    wb = app.books.open(str(file_path), update_links=False)

    try:
        empirical_rows: List[Dict[str, Any]] = []
        regression_rows: List[Dict[str, Any]] = []

        empirical_sheet = try_sheet(wb, EMPIRICAL_SHEET_NAME)
        if empirical_sheet is None:
            print(f"Skipped empirical extraction for {file_path.name}: missing sheet '{EMPIRICAL_SHEET_NAME}'")
        else:
            empirical_rows = process_empirical_sheet(wb, empirical_sheet, labels, file_path.name)

        regression_sheet = try_sheet(wb, REGRESSION_SHEET_NAME)
        if regression_sheet is None:
            print(f"Skipped regression extraction for {file_path.name}: missing sheet '{REGRESSION_SHEET_NAME}'")
        else:
            regression_rows = process_regression_sheet(wb, regression_sheet, labels, file_path.name)

        return empirical_rows, regression_rows
    finally:
        safe_close_workbook(wb)


def iter_source_files(in_dir: Path) -> Iterable[Path]:
    for file_path in sorted(in_dir.iterdir()):
        if not file_path.is_file():
            print(f"Skipped {file_path.name}: not a file")
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped {file_path.name}: temporary file")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped {file_path.name}: not an .xlsx file")
            continue
        yield file_path


def main() -> None:
    in_dir = input_dir.expanduser().resolve()
    out_dir = output_dir.expanduser().resolve()

    if not in_dir.exists():
        raise FileNotFoundError(f"Input directory does not exist: {in_dir}")

    out_dir.mkdir(parents=True, exist_ok=True)
    output_path = next_output_path(in_dir, out_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in iter_source_files(in_dir):
            print(f"Processing {file_path.name}")
            try:
                file_empirical_rows, file_regression_rows = process_workbook(app, file_path)
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
                continue

            empirical_rows.extend(file_empirical_rows)
            regression_rows.extend(file_regression_rows)
            files_processed += 1
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output written to: {output_path}")
    print(f"Files processed: {files_processed}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

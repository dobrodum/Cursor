from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# Update these paths before running.
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")

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

MONTHS = {
    "jan": 1,
    "feb": 2,
    "mar": 3,
    "apr": 4,
    "may": 5,
    "jun": 6,
    "jul": 7,
    "aug": 8,
    "sep": 9,
    "sept": 9,
    "oct": 10,
    "nov": 11,
    "dec": 12,
}

PERIOD_DAY = {"early": 5, "mid": 15, "late": 25}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    return re.sub(r"\s+", " ", text)


def ensure_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, (list, tuple)):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], (list, tuple)):
        return [list(row) if isinstance(row, (list, tuple)) else [row] for row in values]
    return [list(values)]


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "")
        if cleaned.endswith("%"):
            cleaned = cleaned[:-1]
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def clean_excel_value(value: Any) -> Any:
    if isinstance(value, str) and value.startswith("#"):
        return ""
    return value


def numeric_diff(a: Any, b: Any) -> Any:
    a_num = to_float(a)
    b_num = to_float(b)
    if a_num is None or b_num is None:
        return ""
    return a_num - b_num


def normalized_signature(values: Iterable[Any]) -> Tuple[Any, ...]:
    out: List[Any] = []
    for value in values:
        numeric = to_float(value)
        if numeric is not None:
            out.append(round(numeric, 10))
        else:
            out.append(normalize_text(value))
    return tuple(out)


def parse_file_labels(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    parts = [p.strip() for p in stem.split("-")]

    ticker = ""
    period_token = ""

    if len(parts) >= 3:
        ticker = re.sub(r"\s+", "", parts[1]).upper()
        period_token = re.sub(r"\s+", "", parts[2].split("_")[0])

    if not ticker or not period_token:
        match = re.search(
            r"-\s*([A-Za-z0-9]+)\s*-\s*((?:Early|Mid|Late)[A-Za-z]+\d{4})",
            stem,
            flags=re.IGNORECASE,
        )
        if match:
            ticker = ticker or match.group(1).upper()
            period_token = period_token or match.group(2)

    if not period_token:
        match = re.search(r"(Early|Mid|Late)[A-Za-z]+\d{4}", stem, flags=re.IGNORECASE)
        period_token = match.group(0) if match else ""

    model_period = period_token
    model_date = ""

    period_match = re.match(
        r"^(Early|Mid|Late)([A-Za-z]+)(\d{4})$",
        period_token,
        flags=re.IGNORECASE,
    )
    if period_match:
        part = period_match.group(1).title()
        month_token = period_match.group(2)
        year = int(period_match.group(3))
        month_abbrev = month_token[:3].title()
        month_num = MONTHS.get(month_abbrev.lower())

        model_period = f"{part}{month_abbrev}_{year}"
        if month_num is not None:
            day = PERIOD_DAY[part.lower()]
            model_date = date(year, month_num, day).isoformat()

    model = f"{ticker}_{model_period}" if ticker and model_period else stem

    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def build_output_path(input_folder: Path, output_folder: Path) -> Path:
    base_name = f"{input_folder.name}_PARAM"
    candidate = output_folder / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    counter = 1
    while True:
        candidate = output_folder / f"{base_name}.{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


def find_anchor(sheet: xw.Sheet, target: str = "max") -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    values = ensure_2d(used.value)
    if not values:
        return None

    start_row = used.row
    start_col = used.column
    target_norm = normalize_text(target)

    for row_idx, row in enumerate(values):
        for col_idx, value in enumerate(row):
            if normalize_text(value) == target_norm:
                return start_row + row_idx, start_col + col_idx
    return None


HEADER_KEYWORDS = {
    "max": ["max"],
    "min": ["min"],
    "num_quarters_used": ["num quarters", "quarters used", "n quarters", "# quarters", "num qtr"],
    "last_quarter_used": ["last quarter", "last qtr"],
    "forecast_total_without_sa": ["tot fcst w/o sa", "tot fcst wo sa", "total fcst", "forecast"],
    "reported_sales": ["reported sales", "actual sales", "sales reported"],
    "quarterly_sales": ["quarterly sales", "qtr sales"],
    "growth_rate_pct": ["growth rate", "growth %"],
    "sales_captured_in_db_pct": ["sales captured in db", "captured in db", "db pct", "penetration"],
}


def locate_anchor_offsets(sheet: xw.Sheet, anchor_row: int, anchor_col: int) -> Dict[str, int]:
    row_start = max(1, anchor_row - 2)
    row_end = anchor_row + 2
    col_start = max(1, anchor_col - 16)
    col_end = anchor_col + 8

    values = ensure_2d(sheet.range((row_start, col_start), (row_end, col_end)).value)
    offsets: Dict[str, Tuple[int, int]] = {}

    for row_idx, row in enumerate(values):
        for col_idx, value in enumerate(row):
            text = normalize_text(value)
            if not text:
                continue

            abs_row = row_start + row_idx
            abs_col = col_start + col_idx
            score = abs(abs_row - anchor_row) * 100 + abs(abs_col - anchor_col)
            offset = abs_col - anchor_col

            for key, patterns in HEADER_KEYWORDS.items():
                if any(pattern in text for pattern in patterns):
                    current = offsets.get(key)
                    if current is None or score < current[1]:
                        offsets[key] = (offset, score)
                    break

    return {key: value[0] for key, value in offsets.items()}


def read_cell(sheet: xw.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    return clean_excel_value(sheet.cells(row, col).value)


def write_formula2(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def close_workbook_without_save(wb: xw.Book) -> None:
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

    wb.close()


def extract_empirical_rows(
    sheet: xw.Sheet, metadata: Dict[str, str], source_file: str, app: xw.App
) -> List[Dict[str, Any]]:
    anchor = find_anchor(sheet, target="max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    offsets = locate_anchor_offsets(sheet, anchor_row, anchor_col)
    data_start_row = anchor_row + 1

    max_col = anchor_col + offsets.get("max", 0)
    min_col = anchor_col + offsets.get("min", 1)
    forecast_col = anchor_col + offsets.get("forecast_total_without_sa", -7)
    reported_col = anchor_col + offsets.get("reported_sales", -6)
    quarterly_sales_col = anchor_col + offsets.get("quarterly_sales", -8)
    growth_col = anchor_col + offsets.get("growth_rate_pct", -9)
    captured_col = anchor_col + offsets.get("sales_captured_in_db_pct", -11)
    last_quarter_col = anchor_col + offsets.get("last_quarter_used", -10)
    num_quarters_col = anchor_col + offsets["num_quarters_used"] if "num_quarters_used" in offsets else None

    avg_penetration_col = anchor_col + 6

    # Fill all formulas first, then calculate once.
    for i in range(N_QUARTERS):
        row = data_start_row + i
        formula = f'=IF(COUNTA(R{data_start_row}C{captured_col}:R{row}C{captured_col})=0,"",AVERAGE(R{data_start_row}C{captured_col}:R{row}C{captured_col}))'
        write_formula2(sheet.cells(row, avg_penetration_col), formula)

    app.calculate()

    rows: List[Dict[str, Any]] = []
    for i in range(N_QUARTERS):
        row = data_start_row + i
        num_quarters_used = read_cell(sheet, row, num_quarters_col) if num_quarters_col else (i + 1)
        last_quarter_used = read_cell(sheet, row, last_quarter_col)
        forecast_value = read_cell(sheet, row, forecast_col)
        actual_value = read_cell(sheet, row, reported_col)
        forecast_max = read_cell(sheet, row, max_col)
        forecast_min = read_cell(sheet, row, min_col)
        avg_penetration_pct = read_cell(sheet, row, avg_penetration_col)
        quarterly_sales = read_cell(sheet, row, quarterly_sales_col)
        reported_sales = read_cell(sheet, row, reported_col)
        growth_rate_pct = read_cell(sheet, row, growth_col)
        sales_captured_in_db_pct = read_cell(sheet, row, captured_col)

        has_data = any(
            value not in (None, "")
            for value in (
                forecast_value,
                forecast_max,
                forecast_min,
                quarterly_sales,
                reported_sales,
                sales_captured_in_db_pct,
            )
        )
        if not has_data:
            continue

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": numeric_diff(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_rows(
    sheet: xw.Sheet, metadata: Dict[str, str], source_file: str, app: xw.App
) -> List[Dict[str, Any]]:
    anchor = find_anchor(sheet, target="max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    offsets = locate_anchor_offsets(sheet, anchor_row, anchor_col)
    data_start_row = anchor_row + 1

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if "forecast_total_without_sa" in offsets:
        y_col = anchor_col + offsets["forecast_total_without_sa"]
    if "sales_captured_in_db_pct" in offsets:
        x_col = anchor_col + offsets["sales_captured_in_db_pct"]

    max_col = anchor_col + offsets.get("max", 0)
    min_col = anchor_col + offsets.get("min", 1)
    num_quarters_col = anchor_col + offsets["num_quarters_used"] if "num_quarters_used" in offsets else None
    actual_col = anchor_col + offsets["reported_sales"] if "reported_sales" in offsets else None

    intercept_col = anchor_col + 2
    slope_col = anchor_col + 3

    # Fill all formulas first, then calculate once.
    for i in range(N_QUARTERS):
        row = data_start_row + i
        intercept_formula = (
            f'=IF(COUNTA(R{data_start_row}C{x_col}:R{row}C{x_col})<2,"",'
            f"INTERCEPT(R{data_start_row}C{y_col}:R{row}C{y_col},R{data_start_row}C{x_col}:R{row}C{x_col}))"
        )
        slope_formula = (
            f'=IF(COUNTA(R{data_start_row}C{x_col}:R{row}C{x_col})<2,"",'
            f"SLOPE(R{data_start_row}C{y_col}:R{row}C{y_col},R{data_start_row}C{x_col}:R{row}C{x_col}))"
        )
        write_formula2(sheet.cells(row, intercept_col), intercept_formula)
        write_formula2(sheet.cells(row, slope_col), slope_formula)

    app.calculate()

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for i in range(N_QUARTERS):
        row = data_start_row + i
        num_quarters_used = read_cell(sheet, row, num_quarters_col) if num_quarters_col else (i + 1)
        forecast_value = read_cell(sheet, row, y_col)
        forecast_max = read_cell(sheet, row, max_col)
        forecast_min = read_cell(sheet, row, min_col)
        intercept = read_cell(sheet, row, intercept_col)
        slope = read_cell(sheet, row, slope_col)
        actual_value = read_cell(sheet, row, actual_col) if actual_col else ""

        has_data = any(
            value not in (None, "")
            for value in (forecast_value, forecast_max, forecast_min, intercept, slope)
        )
        if not has_data:
            continue

        signature = normalized_signature(
            (
                num_quarters_used,
                forecast_value,
                forecast_max,
                forecast_min,
                intercept,
                slope,
            )
        )
        if signature == previous_signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": numeric_diff(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def write_sheet(ws, columns: List[str], rows: List[Dict[str, Any]]) -> None:
    ws.append(columns)
    for row in rows:
        ws.append([row.get(column, "") for column in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(1, ws.max_row)}"

    for col_index, header in enumerate(columns, start=1):
        max_len = len(header)
        for row_index in range(2, ws.max_row + 1):
            value = ws.cell(row=row_index, column=col_index).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_index)].width = min(max(max_len + 2, 12), 42)


def safe_get_sheet(wb: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    try:
        return wb.sheets[sheet_name]
    except Exception:
        return None


def iter_candidate_files(folder: Path) -> Iterable[Path]:
    for file_path in sorted(folder.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"Skipping {file_path.name}: temporary file")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipping {file_path.name}: not an .xlsx file")
            continue
        yield file_path


def main() -> None:
    source_folder = input_dir.expanduser().resolve()
    target_folder = output_dir.expanduser().resolve()
    target_folder.mkdir(parents=True, exist_ok=True)

    if not source_folder.exists() or not source_folder.is_dir():
        raise FileNotFoundError(f"Input folder does not exist: {source_folder}")

    output_path = build_output_path(source_folder, target_folder)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in iter_candidate_files(source_folder):
            print(f"Processing {file_path.name}")
            wb: Optional[xw.Book] = None

            try:
                wb = app.books.open(str(file_path), update_links=False)
                metadata = parse_file_labels(file_path.name)

                empirical_sheet = safe_get_sheet(wb, "Empirical Model")
                if empirical_sheet is None:
                    print(f"Skipping empirical extraction for {file_path.name}: missing 'Empirical Model'")
                else:
                    empirical_rows.extend(
                        extract_empirical_rows(empirical_sheet, metadata, file_path.name, app)
                    )

                regression_sheet = safe_get_sheet(wb, "Regression Model")
                if regression_sheet is None:
                    print(f"Skipping regression extraction for {file_path.name}: missing 'Regression Model'")
                else:
                    regression_rows.extend(
                        extract_regression_rows(regression_sheet, metadata, file_path.name, app)
                    )

                processed_files += 1
            except Exception as exc:
                print(f"Skipping {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    close_workbook_without_save(wb)
    finally:
        app.quit()

    out_wb = Workbook()
    default_sheet = out_wb.active
    out_wb.remove(default_sheet)

    empirical_ws = out_wb.create_sheet("empirical_candidates")
    regression_ws = out_wb.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    out_wb.save(output_path)

    print(f"Output path: {output_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

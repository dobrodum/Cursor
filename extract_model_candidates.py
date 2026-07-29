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

# ===== User inputs =====
input_dir = r"./input"
output_dir = r"./output"


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

MONTH_MAP = {
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

PHASE_TO_DAY = {
    "early": 5,
    "mid": 15,
    "late": 25,
}


@dataclass
class FileLabel:
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
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def to_int_if_whole(value: Any) -> Any:
    num = to_float(value)
    if num is None:
        return value
    if num.is_integer():
        return int(num)
    return num


def get_matrix_value(
    matrix: Sequence[Sequence[Any]],
    start_row: int,
    start_col: int,
    row: int,
    col: int,
) -> Any:
    row_idx = row - start_row
    col_idx = col - start_col
    if row_idx < 0 or col_idx < 0:
        return None
    if row_idx >= len(matrix):
        return None
    row_values = matrix[row_idx]
    if col_idx >= len(row_values):
        return None
    return row_values[col_idx]


def find_anchor_max(
    matrix: Sequence[Sequence[Any]],
    start_row: int,
    start_col: int,
) -> Optional[Tuple[int, int]]:
    for row_offset, row_values in enumerate(matrix):
        for col_offset, value in enumerate(row_values):
            if normalize_text(value) == "max":
                return start_row + row_offset, start_col + col_offset
    return None


def build_header_maps(
    matrix: Sequence[Sequence[Any]],
    start_row: int,
    start_col: int,
    anchor_row: int,
) -> List[Dict[int, str]]:
    header_maps: List[Dict[int, str]] = []
    for header_row in (anchor_row, anchor_row - 1):
        row_idx = header_row - start_row
        if row_idx < 0 or row_idx >= len(matrix):
            continue
        mapping: Dict[int, str] = {}
        for col_offset, value in enumerate(matrix[row_idx]):
            normalized = normalize_text(value)
            if normalized:
                mapping[start_col + col_offset] = normalized
        if mapping:
            header_maps.append(mapping)
    return header_maps


def find_col_by_keywords(
    header_maps: Sequence[Dict[int, str]],
    keyword_options: Sequence[Sequence[str]],
) -> Optional[int]:
    for keywords in keyword_options:
        normalized_keywords = [normalize_text(keyword) for keyword in keywords]
        for header_map in header_maps:
            for col_idx, header_text in header_map.items():
                if all(keyword in header_text for keyword in normalized_keywords):
                    return col_idx
    return None


def find_numeric_rows(
    matrix: Sequence[Sequence[Any]],
    start_row: int,
    start_col: int,
    col_idx: int,
    max_row_exclusive: int,
) -> List[int]:
    rows: List[int] = []
    for row_idx in range(start_row, max_row_exclusive):
        value = get_matrix_value(matrix, start_row, start_col, row_idx, col_idx)
        if to_float(value) is not None:
            rows.append(row_idx)
    return rows


def close_source_workbook(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    workbook_api = getattr(workbook, "api", None)
    if workbook_api is not None:
        try:
            workbook_api.Close(SaveChanges=False)
            return
        except TypeError:
            try:
                workbook_api.Close(False)
                return
            except Exception:
                pass
        except Exception:
            pass

    try:
        workbook.close()
    except Exception:
        pass


def get_sheet(book: xw.Book, target_name: str) -> Optional[xw.Sheet]:
    target_lower = target_name.strip().lower()
    for sheet in book.sheets:
        if sheet.name.strip().lower() == target_lower:
            return sheet
    return None


def parse_file_label(file_name: str) -> FileLabel:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ") if part.strip()]

    ticker = ""
    if len(parts) >= 2:
        ticker = parts[1]
    else:
        ticker_candidates = re.findall(r"\b[A-Z]{2,8}\b", stem)
        ticker = next((t for t in ticker_candidates if t not in {"MODEL", "SEND"}), "")

    if not ticker:
        raise ValueError("could not parse ticker")

    period_match = re.search(
        r"(Early|Mid|Late)[ _-]*([A-Za-z]{3,9})[ _-]*(\d{4})",
        stem,
        re.IGNORECASE,
    )
    if not period_match:
        raise ValueError("could not parse model period")

    phase = period_match.group(1).title()
    month_token = period_match.group(2)[:3].title()
    year = int(period_match.group(3))

    month_num = MONTH_MAP.get(month_token.lower())
    if month_num is None:
        raise ValueError(f"unknown month token '{month_token}'")

    day = PHASE_TO_DAY[phase.lower()]
    model_period = f"{phase}{month_token}_{year}"
    model_date = date(year, month_num, day).isoformat()
    model = f"{ticker}_{model_period}"

    return FileLabel(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def build_output_path(input_path: Path, output_path: Path) -> Path:
    base_name = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = output_path / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def calc_range_width(max_value: Any, min_value: Any) -> Optional[float]:
    max_num = to_float(max_value)
    min_num = to_float(min_value)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


def normalized_signature(values: Iterable[Any]) -> Tuple[Any, ...]:
    signature: List[Any] = []
    for value in values:
        num = to_float(value)
        if num is None:
            signature.append(value)
        else:
            signature.append(round(num, 10))
    return tuple(signature)


def extract_empirical_candidates(
    workbook: xw.Book,
    label: FileLabel,
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = get_sheet(workbook, "Empirical Model")
    if sheet is None:
        print(f"Skipping empirical extraction for {source_file}: missing sheet 'Empirical Model'")
        return []

    used_range = sheet.used_range
    matrix = to_matrix(used_range.value)
    if not matrix:
        print(f"Skipping empirical extraction for {source_file}: empty used range")
        return []

    start_row = used_range.row
    start_col = used_range.column
    anchor = find_anchor_max(matrix, start_row, start_col)
    if anchor is None:
        print(f"Skipping empirical extraction for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    header_maps = build_header_maps(matrix, start_row, start_col, anchor_row)

    min_col = find_col_by_keywords(header_maps, (("min",),)) or (anchor_col + 1)
    num_quarters_col = find_col_by_keywords(
        header_maps,
        (
            ("num", "quarter"),
            ("quarters", "used"),
            ("n", "quarter"),
        ),
    )
    last_quarter_col = find_col_by_keywords(header_maps, (("last", "quarter"),))
    forecast_value_col = find_col_by_keywords(
        header_maps,
        (
            ("estimated", "total", "sold"),
            ("forecast", "value"),
            ("tot", "fcst"),
        ),
    )
    actual_value_col = find_col_by_keywords(
        header_maps,
        (
            ("reported", "sales"),
            ("actual", "sales"),
            ("actual",),
        ),
    )
    quarterly_sales_col = find_col_by_keywords(header_maps, (("quarterly", "sales"),))
    reported_sales_col = find_col_by_keywords(header_maps, (("reported", "sales"),))
    growth_rate_col = find_col_by_keywords(header_maps, (("growth", "rate"),))
    captured_pct_col = find_col_by_keywords(
        header_maps,
        (
            ("captured", "db"),
            ("sales", "captured"),
        ),
    )
    penetration_col = find_col_by_keywords(header_maps, (("penetration",),)) or (anchor_col - 11)
    quarter_label_col = find_col_by_keywords(
        header_maps,
        (
            ("quarter",),
            ("period",),
        ),
    )

    penetration_rows = find_numeric_rows(matrix, start_row, start_col, penetration_col, anchor_row)

    matrix_width = max((len(row_values) for row_values in matrix), default=0)
    calc_col = start_col + matrix_width + 5
    calc_row_base = start_row

    avg_penetration_cells: Dict[int, xw.Range] = {}
    for n_quarters in range(1, N_QUARTERS + 1):
        if len(penetration_rows) < n_quarters:
            continue
        start_data_row = penetration_rows[-n_quarters]
        end_data_row = penetration_rows[-1]
        calc_cell = sheet.range((calc_row_base + n_quarters - 1, calc_col))
        calc_cell.formula2 = (
            f"=AVERAGE(R{start_data_row}C{penetration_col}:R{end_data_row}C{penetration_col})"
        )
        avg_penetration_cells[n_quarters] = calc_cell

    if avg_penetration_cells:
        workbook.app.calculate()

    avg_penetration_values = {
        n_quarters: calc_cell.value for n_quarters, calc_cell in avg_penetration_cells.items()
    }

    extracted_rows: List[Dict[str, Any]] = []
    for n_quarters in range(1, N_QUARTERS + 1):
        candidate_row = anchor_row + n_quarters
        num_quarters_used = (
            get_matrix_value(matrix, start_row, start_col, candidate_row, num_quarters_col)
            if num_quarters_col
            else n_quarters
        )
        if num_quarters_used in (None, ""):
            num_quarters_used = n_quarters
        num_quarters_used = to_int_if_whole(num_quarters_used)

        last_quarter_used = (
            get_matrix_value(matrix, start_row, start_col, candidate_row, last_quarter_col)
            if last_quarter_col
            else None
        )
        if last_quarter_used in (None, "") and quarter_label_col and penetration_rows:
            last_quarter_used = get_matrix_value(
                matrix,
                start_row,
                start_col,
                penetration_rows[-1],
                quarter_label_col,
            )

        forecast_value = (
            get_matrix_value(matrix, start_row, start_col, candidate_row, forecast_value_col)
            if forecast_value_col
            else None
        )
        actual_value = (
            get_matrix_value(matrix, start_row, start_col, candidate_row, actual_value_col)
            if actual_value_col
            else None
        )
        forecast_max = get_matrix_value(matrix, start_row, start_col, candidate_row, anchor_col)
        forecast_min = get_matrix_value(matrix, start_row, start_col, candidate_row, min_col)
        avg_penetration_pct = avg_penetration_values.get(n_quarters)

        quarterly_sales = (
            get_matrix_value(matrix, start_row, start_col, candidate_row, quarterly_sales_col)
            if quarterly_sales_col
            else None
        )
        reported_sales = (
            get_matrix_value(matrix, start_row, start_col, candidate_row, reported_sales_col)
            if reported_sales_col
            else None
        )
        growth_rate_pct = (
            get_matrix_value(matrix, start_row, start_col, candidate_row, growth_rate_col)
            if growth_rate_col
            else None
        )
        sales_captured_pct = (
            get_matrix_value(matrix, start_row, start_col, candidate_row, captured_pct_col)
            if captured_pct_col
            else None
        )

        if (
            forecast_value is None
            and forecast_max is None
            and forecast_min is None
            and avg_penetration_pct is None
        ):
            continue

        extracted_rows.append(
            {
                "model": label.model,
                "ticker": label.ticker,
                "model_period": label.model_period,
                "model_date": label.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": calc_range_width(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_pct,
                "source_file": source_file,
            }
        )

    return extracted_rows


def extract_regression_candidates(
    workbook: xw.Book,
    label: FileLabel,
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = get_sheet(workbook, "Regression Model")
    if sheet is None:
        print(f"Skipping regression extraction for {source_file}: missing sheet 'Regression Model'")
        return []

    used_range = sheet.used_range
    matrix = to_matrix(used_range.value)
    if not matrix:
        print(f"Skipping regression extraction for {source_file}: empty used range")
        return []

    start_row = used_range.row
    start_col = used_range.column
    anchor = find_anchor_max(matrix, start_row, start_col)
    if anchor is None:
        print(f"Skipping regression extraction for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    header_maps = build_header_maps(matrix, start_row, start_col, anchor_row)

    min_col = find_col_by_keywords(header_maps, (("min",),)) or (anchor_col + 1)
    num_quarters_col = find_col_by_keywords(
        header_maps,
        (
            ("num", "quarter"),
            ("quarters", "used"),
            ("n", "quarter"),
        ),
    )
    forecast_total_col = find_col_by_keywords(
        header_maps,
        (
            ("tot", "fcst", "w", "o", "sa"),
            ("forecast", "without", "sa"),
            ("tot", "forecast"),
        ),
    ) or (anchor_col - 1)
    actual_value_col = find_col_by_keywords(
        header_maps,
        (
            ("actual", "sales"),
            ("reported", "sales"),
            ("actual",),
        ),
    )

    # Required by spec: derive x/y from max anchor.
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    xy_rows: List[int] = []
    for row_idx in range(start_row, anchor_row):
        x_value = get_matrix_value(matrix, start_row, start_col, row_idx, x_col)
        y_value = get_matrix_value(matrix, start_row, start_col, row_idx, y_col)
        if to_float(x_value) is not None and to_float(y_value) is not None:
            xy_rows.append(row_idx)

    matrix_width = max((len(row_values) for row_values in matrix), default=0)
    calc_col = start_col + matrix_width + 8
    calc_row_base = start_row

    intercept_cells: Dict[int, xw.Range] = {}
    slope_cells: Dict[int, xw.Range] = {}
    for n_quarters in range(1, N_QUARTERS + 1):
        if len(xy_rows) < n_quarters:
            continue
        start_data_row = xy_rows[-n_quarters]
        end_data_row = xy_rows[-1]

        intercept_cell = sheet.range((calc_row_base + n_quarters - 1, calc_col))
        slope_cell = sheet.range((calc_row_base + n_quarters - 1, calc_col + 1))

        intercept_cell.formula2 = (
            f"=INTERCEPT(R{start_data_row}C{y_col}:R{end_data_row}C{y_col},"
            f"R{start_data_row}C{x_col}:R{end_data_row}C{x_col})"
        )
        slope_cell.formula2 = (
            f"=SLOPE(R{start_data_row}C{y_col}:R{end_data_row}C{y_col},"
            f"R{start_data_row}C{x_col}:R{end_data_row}C{x_col})"
        )

        intercept_cells[n_quarters] = intercept_cell
        slope_cells[n_quarters] = slope_cell

    if intercept_cells or slope_cells:
        workbook.app.calculate()

    extracted_rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for n_quarters in range(1, N_QUARTERS + 1):
        candidate_row = anchor_row + n_quarters
        num_quarters_used = (
            get_matrix_value(matrix, start_row, start_col, candidate_row, num_quarters_col)
            if num_quarters_col
            else n_quarters
        )
        if num_quarters_used in (None, ""):
            num_quarters_used = n_quarters
        num_quarters_used = to_int_if_whole(num_quarters_used)

        forecast_value = get_matrix_value(
            matrix,
            start_row,
            start_col,
            candidate_row,
            forecast_total_col,
        )
        actual_value = (
            get_matrix_value(matrix, start_row, start_col, candidate_row, actual_value_col)
            if actual_value_col
            else ""
        )
        forecast_max = get_matrix_value(matrix, start_row, start_col, candidate_row, anchor_col)
        forecast_min = get_matrix_value(matrix, start_row, start_col, candidate_row, min_col)
        intercept = intercept_cells[n_quarters].value if n_quarters in intercept_cells else None
        slope = slope_cells[n_quarters].value if n_quarters in slope_cells else None

        if (
            forecast_value is None
            and forecast_max is None
            and forecast_min is None
            and intercept is None
            and slope is None
        ):
            continue

        row_signature = normalized_signature(
            (
                num_quarters_used,
                forecast_value,
                forecast_max,
                forecast_min,
                intercept,
                slope,
            )
        )
        if previous_signature is not None and row_signature == previous_signature:
            continue
        previous_signature = row_signature

        extracted_rows.append(
            {
                "model": label.model,
                "ticker": label.ticker,
                "model_period": label.model_period,
                "model_date": label.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value if actual_value is not None else "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": calc_range_width(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return extracted_rows


def write_sheet(
    workbook: Workbook,
    sheet_name: str,
    columns: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    worksheet = workbook.create_sheet(title=sheet_name)
    worksheet.append(list(columns))
    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    for row in rows:
        worksheet.append([row.get(column, "") for column in columns])

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(worksheet.max_row, 1)}"

    for col_idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row_idx in range(2, worksheet.max_row + 1):
            cell_value = worksheet.cell(row=row_idx, column=col_idx).value
            if cell_value is None:
                continue
            cell_len = len(str(cell_value))
            if cell_len > max_len:
                max_len = cell_len
        worksheet.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    write_sheet(workbook, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(workbook, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    workbook.save(output_path)


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists():
        raise FileNotFoundError(f"input_dir does not exist: {input_path}")
    if not input_path.is_dir():
        raise NotADirectoryError(f"input_dir is not a directory: {input_path}")

    output_path.mkdir(parents=True, exist_ok=True)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(input_path.iterdir(), key=lambda path: path.name.lower()):
            if not file_path.is_file():
                print(f"Skipping {file_path.name}: not a file")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipping {file_path.name}: temporary Excel file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipping {file_path.name}: not an .xlsx file")
                continue

            try:
                label = parse_file_label(file_path.name)
            except Exception as exc:
                print(f"Skipping {file_path.name}: filename parse failure ({exc})")
                continue

            print(f"Processing {file_path.name}")
            workbook: Optional[xw.Book] = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(
                    extract_empirical_candidates(workbook, label=label, source_file=file_path.name)
                )
                regression_rows.extend(
                    extract_regression_candidates(workbook, label=label, source_file=file_path.name)
                )
                processed_files += 1
            except Exception as exc:
                print(f"Skipping {file_path.name}: workbook processing failure ({exc})")
            finally:
                if workbook is not None:
                    close_source_workbook(workbook)
    finally:
        app.quit()

    final_output_path = build_output_path(input_path, output_path)
    write_output_workbook(final_output_path, empirical_rows, regression_rows)

    print(f"Output path: {final_output_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

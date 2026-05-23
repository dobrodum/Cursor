from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Update these two paths for your environment.
input_dir = Path("input")
output_dir = Path("output")

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


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def to_list(values: Any) -> List[Any]:
    if values is None:
        return []
    if isinstance(values, list):
        if values and isinstance(values[0], list):
            return [row[0] if row else None for row in values]
        return values
    return [values]


def safe_float(value: Any) -> Optional[float]:
    if value in (None, ""):
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(",", "").strip())
    except (TypeError, ValueError):
        return None


def subtract_or_none(a: Any, b: Any) -> Optional[float]:
    a_float = safe_float(a)
    b_float = safe_float(b)
    if a_float is None or b_float is None:
        return None
    return a_float - b_float


def unique_output_path(input_folder: Path, output_folder: Path) -> Path:
    output_folder.mkdir(parents=True, exist_ok=True)
    stem = f"{input_folder.name}_PARAM"
    candidate = output_folder / f"{stem}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = output_folder / f"{stem}.{suffix}.xlsx"
        suffix += 1
    return candidate


def parse_month_token(token: str) -> int:
    token = token.strip()
    if not token:
        raise ValueError("Empty month token")

    token_clean = token[:3].title()
    return datetime.strptime(token_clean, "%b").month


def parse_file_labels(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = parts[1].upper() if len(parts) > 1 else ""
    period_blob = parts[2] if len(parts) > 2 else ""
    period_token = period_blob.split("_")[0]

    period_match = re.search(
        r"(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})",
        period_token,
        flags=re.IGNORECASE,
    )

    model_period = period_token
    model_date = ""

    if period_match:
        period_bucket = period_match.group(1).title()
        month_token = period_match.group(2)
        year = int(period_match.group(3))

        month_number = parse_month_token(month_token)
        month_abbrev = datetime(year, month_number, 1).strftime("%b")
        day_lookup = {"Early": 5, "Mid": 15, "Late": 25}
        day = day_lookup[period_bucket]

        model_period = f"{period_bucket}{month_abbrev}_{year}"
        model_date = date(year, month_number, day).isoformat()

    model = f"{ticker}_{model_period}" if ticker and model_period else (ticker or stem)
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def safe_close_book(book: xw.Book) -> None:
    try:
        book.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        book.close(False)
        return
    except Exception:
        pass

    try:
        book.api.Close(SaveChanges=False)
    except Exception:
        # Final fallback: best-effort close with no save
        book.api.Close(False)


def get_sheet_or_none(book: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    try:
        return book.sheets[sheet_name]
    except Exception:
        return None


def used_range_bounds(sheet: xw.Sheet) -> Tuple[int, int, int, int]:
    used = sheet.used_range
    first_row = used.row
    first_col = used.column
    last_row = first_row + used.rows.count - 1
    last_col = first_col + used.columns.count - 1
    return first_row, first_col, last_row, last_col


def find_max_anchor(sheet: xw.Sheet) -> Tuple[int, int]:
    first_row, first_col, _, _ = used_range_bounds(sheet)
    grid = sheet.used_range.value

    if grid is None:
        raise ValueError("Sheet is empty")
    if not isinstance(grid, list):
        grid = [[grid]]
    elif grid and not isinstance(grid[0], list):
        grid = [grid]

    for row_idx, row_values in enumerate(grid):
        for col_idx, value in enumerate(row_values):
            if normalize_text(value) == "max":
                return first_row + row_idx, first_col + col_idx

    raise ValueError("Could not find 'max' anchor")


def build_header_map(sheet: xw.Sheet, header_row: int, first_col: int, last_col: int) -> Dict[str, List[int]]:
    row_values = sheet.range((header_row, first_col), (header_row, last_col)).value
    values = to_list(row_values)
    header_map: Dict[str, List[int]] = {}
    for idx, value in enumerate(values):
        key = normalize_text(value)
        if not key:
            continue
        header_map.setdefault(key, []).append(first_col + idx)
    return header_map


def pick_column(
    header_map: Dict[str, List[int]],
    anchor_col: int,
    include_terms: Sequence[str],
    default: Optional[int] = None,
    exclude_terms: Sequence[str] = (),
) -> Optional[int]:
    candidates: List[int] = []
    for header_key, cols in header_map.items():
        if all(term in header_key for term in include_terms) and not any(
            term in header_key for term in exclude_terms
        ):
            candidates.extend(cols)
    if not candidates:
        return default
    return min(candidates, key=lambda col: abs(col - anchor_col))


def read_columns(
    sheet: xw.Sheet,
    start_row: int,
    end_row: int,
    columns: Sequence[Optional[int]],
) -> Dict[int, List[Any]]:
    col_values: Dict[int, List[Any]] = {}
    for col in sorted({col for col in columns if col and col > 0}):
        values = sheet.range((start_row, col), (end_row, col)).value
        col_values[col] = to_list(values)
    return col_values


def value_at(col_map: Dict[int, List[Any]], col: Optional[int], row_idx: int) -> Any:
    if col is None:
        return None
    values = col_map.get(col)
    if not values or row_idx >= len(values):
        return None
    return values[row_idx]


def build_empirical_formulas(
    helper_rows: Sequence[int],
    n_quarters: int,
    penetration_col: Optional[int],
    anchor_row: int,
    first_row: int,
) -> List[Optional[str]]:
    formulas: List[Optional[str]] = []
    if penetration_col is None:
        return [None] * len(helper_rows)

    history_end_row = anchor_row - 1
    for n in range(1, n_quarters + 1):
        start_row = max(first_row, history_end_row - n + 1)
        if start_row > history_end_row:
            formulas.append(None)
            continue
        formula = (
            f'=IFERROR(AVERAGE(R{start_row}C{penetration_col}:'
            f"R{history_end_row}C{penetration_col}),\"\")"
        )
        formulas.append(formula)
    return formulas


def extract_empirical_rows(
    sheet: xw.Sheet,
    labels: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    n_quarters = 10
    anchor_row, anchor_col = find_max_anchor(sheet)
    first_row, first_col, _, last_col = used_range_bounds(sheet)
    header_map = build_header_map(sheet, anchor_row, first_col, last_col)

    data_start = anchor_row + 1
    data_end = data_start + n_quarters - 1

    forecast_max_col = anchor_col
    forecast_min_col = pick_column(header_map, anchor_col, ["min"], default=anchor_col + 1)
    forecast_value_col = pick_column(
        header_map,
        anchor_col,
        ["estimated", "totalsold"],
        default=pick_column(
            header_map,
            anchor_col,
            ["totfcst"],
            default=anchor_col - 1,
            exclude_terms=("max", "min"),
        ),
    )
    actual_value_col = pick_column(
        header_map,
        anchor_col,
        ["reportedsales"],
        default=pick_column(header_map, anchor_col, ["actual"], default=anchor_col - 2),
    )
    num_quarters_col = pick_column(
        header_map, anchor_col, ["numquartersused"], default=pick_column(header_map, anchor_col, ["quartersused"])
    )
    last_quarter_col = pick_column(
        header_map, anchor_col, ["lastquarterused"], default=pick_column(header_map, anchor_col, ["lastquarter"])
    )
    quarterly_sales_col = pick_column(header_map, anchor_col, ["quarterlysales"])
    growth_rate_col = pick_column(header_map, anchor_col, ["growthrate"])
    sales_captured_col = pick_column(
        header_map,
        anchor_col,
        ["salescapturedindb"],
        default=pick_column(header_map, anchor_col, ["captured", "db"]),
    )
    penetration_col = pick_column(
        header_map,
        anchor_col,
        ["penetration"],
        exclude_terms=("avg",),
    )

    helper_col = last_col + 2
    helper_rows = list(range(data_start, data_end + 1))
    empirical_formulas = build_empirical_formulas(
        helper_rows=helper_rows,
        n_quarters=n_quarters,
        penetration_col=penetration_col,
        anchor_row=anchor_row,
        first_row=first_row,
    )

    formulas_to_write = [[formula if formula else ""] for formula in empirical_formulas]
    sheet.range((data_start, helper_col), (data_end, helper_col)).formula2 = formulas_to_write
    sheet.book.app.calculate()
    avg_penetration_values = to_list(
        sheet.range((data_start, helper_col), (data_end, helper_col)).value
    )
    sheet.range((data_start, helper_col), (data_end, helper_col)).clear_contents()

    col_map = read_columns(
        sheet,
        data_start,
        data_end,
        [
            forecast_value_col,
            actual_value_col,
            forecast_max_col,
            forecast_min_col,
            num_quarters_col,
            last_quarter_col,
            quarterly_sales_col,
            growth_rate_col,
            sales_captured_col,
        ],
    )

    rows: List[Dict[str, Any]] = []
    for idx in range(n_quarters):
        forecast_max = value_at(col_map, forecast_max_col, idx)
        forecast_min = value_at(col_map, forecast_min_col, idx)
        forecast_value = value_at(col_map, forecast_value_col, idx)
        actual_value = value_at(col_map, actual_value_col, idx)
        avg_penetration = avg_penetration_values[idx] if idx < len(avg_penetration_values) else None

        if all(value in (None, "") for value in [forecast_value, forecast_max, forecast_min, avg_penetration]):
            continue

        num_quarters_used = value_at(col_map, num_quarters_col, idx)
        if num_quarters_used in (None, ""):
            num_quarters_used = idx + 1

        reported_sales = actual_value

        row = {
            "model": labels["model"],
            "ticker": labels["ticker"],
            "model_period": labels["model_period"],
            "model_date": labels["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration,
            "num_quarters_used": num_quarters_used,
            "last_quarter_used": value_at(col_map, last_quarter_col, idx),
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": subtract_or_none(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration,
            "quarterly_sales": value_at(col_map, quarterly_sales_col, idx),
            "reported_sales": reported_sales,
            "growth_rate_pct": value_at(col_map, growth_rate_col, idx),
            "sales_captured_in_db_pct": value_at(col_map, sales_captured_col, idx),
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def build_regression_formulas(
    n_quarters: int,
    anchor_row: int,
    first_row: int,
    y_col: int,
    x_col: int,
) -> Tuple[List[str], List[str]]:
    intercept_formulas: List[str] = []
    slope_formulas: List[str] = []
    history_end_row = anchor_row - 1

    for n in range(1, n_quarters + 1):
        start_row = max(first_row, history_end_row - n + 1)
        if start_row > history_end_row:
            intercept_formulas.append("")
            slope_formulas.append("")
            continue

        intercept_formulas.append(
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{history_end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{history_end_row}C{x_col}),"")'
        )
        slope_formulas.append(
            f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{history_end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{history_end_row}C{x_col}),"")'
        )

    return intercept_formulas, slope_formulas


def dedupe_key(values: Sequence[Any]) -> Tuple[Any, ...]:
    normalized: List[Any] = []
    for value in values:
        as_float = safe_float(value)
        if as_float is None:
            normalized.append(value)
        else:
            normalized.append(round(as_float, 10))
    return tuple(normalized)


def extract_regression_rows(
    sheet: xw.Sheet,
    labels: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    n_quarters = 10
    anchor_row, anchor_col = find_max_anchor(sheet)
    first_row, first_col, _, last_col = used_range_bounds(sheet)
    header_map = build_header_map(sheet, anchor_row, first_col, last_col)

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    data_start = anchor_row + 1
    data_end = data_start + n_quarters - 1

    forecast_max_col = anchor_col
    forecast_min_col = pick_column(header_map, anchor_col, ["min"], default=anchor_col + 1)
    forecast_total_col = pick_column(
        header_map,
        anchor_col,
        ["totfcstwosa"],
        default=pick_column(
            header_map,
            anchor_col,
            ["forecast"],
            default=anchor_col - 1,
            exclude_terms=("max", "min"),
        ),
    )
    actual_value_col = pick_column(
        header_map,
        anchor_col,
        ["actual"],
        default=pick_column(header_map, anchor_col, ["reportedsales"]),
    )
    num_quarters_col = pick_column(
        header_map, anchor_col, ["numquartersused"], default=pick_column(header_map, anchor_col, ["quartersused"])
    )

    helper_intercept_col = last_col + 2
    helper_slope_col = last_col + 3
    intercept_formulas, slope_formulas = build_regression_formulas(
        n_quarters=n_quarters,
        anchor_row=anchor_row,
        first_row=first_row,
        y_col=y_col,
        x_col=x_col,
    )

    sheet.range((data_start, helper_intercept_col), (data_end, helper_intercept_col)).formula2 = [
        [formula] for formula in intercept_formulas
    ]
    sheet.range((data_start, helper_slope_col), (data_end, helper_slope_col)).formula2 = [
        [formula] for formula in slope_formulas
    ]
    sheet.book.app.calculate()

    intercept_values = to_list(
        sheet.range((data_start, helper_intercept_col), (data_end, helper_intercept_col)).value
    )
    slope_values = to_list(sheet.range((data_start, helper_slope_col), (data_end, helper_slope_col)).value)
    sheet.range((data_start, helper_intercept_col), (data_end, helper_slope_col)).clear_contents()

    col_map = read_columns(
        sheet,
        data_start,
        data_end,
        [
            num_quarters_col,
            forecast_total_col,
            actual_value_col,
            forecast_max_col,
            forecast_min_col,
        ],
    )

    rows: List[Dict[str, Any]] = []
    last_key: Optional[Tuple[Any, ...]] = None
    for idx in range(n_quarters):
        num_quarters_used = value_at(col_map, num_quarters_col, idx)
        if num_quarters_used in (None, ""):
            num_quarters_used = idx + 1

        forecast_total = value_at(col_map, forecast_total_col, idx)
        actual_value = value_at(col_map, actual_value_col, idx)
        forecast_max = value_at(col_map, forecast_max_col, idx)
        forecast_min = value_at(col_map, forecast_min_col, idx)
        intercept = intercept_values[idx] if idx < len(intercept_values) else None
        slope = slope_values[idx] if idx < len(slope_values) else None

        if all(value in (None, "") for value in [forecast_total, forecast_max, forecast_min, intercept, slope]):
            continue

        current_key = dedupe_key([num_quarters_used, intercept, slope, forecast_total, forecast_max, forecast_min])
        if last_key is not None and current_key == last_key:
            continue
        last_key = current_key

        rows.append(
            {
                "model": labels["model"],
                "ticker": labels["ticker"],
                "model_period": labels["model_period"],
                "model_date": labels["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_total,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": subtract_or_none(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def apply_sheet_formatting(ws, headers: Sequence[str]) -> None:
    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"

    for col_idx, header in enumerate(headers, start=1):
        max_width = len(header)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_width = max(max_width, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_width + 2, 42)


def write_output_workbook(
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
    output_path: Path,
) -> None:
    wb = Workbook()
    default_ws = wb.active
    wb.remove(default_ws)

    empirical_ws = wb.create_sheet("empirical_candidates")
    empirical_ws.append(EMPIRICAL_COLUMNS)
    for row in empirical_rows:
        empirical_ws.append([row.get(col) for col in EMPIRICAL_COLUMNS])
    apply_sheet_formatting(empirical_ws, EMPIRICAL_COLUMNS)

    regression_ws = wb.create_sheet("regression_candidates")
    regression_ws.append(REGRESSION_COLUMNS)
    for row in regression_rows:
        regression_ws.append([row.get(col) for col in REGRESSION_COLUMNS])
    apply_sheet_formatting(regression_ws, REGRESSION_COLUMNS)

    wb.save(output_path)


def list_source_files(folder: Path) -> List[Path]:
    if not folder.exists():
        raise FileNotFoundError(f"Input folder not found: {folder}")
    if not folder.is_dir():
        raise NotADirectoryError(f"Input path is not a folder: {folder}")

    return sorted(path for path in folder.iterdir() if path.is_file())


def process_files(input_folder: Path, output_folder: Path) -> Path:
    source_files = list_source_files(input_folder)
    output_path = unique_output_path(input_folder, output_folder)

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    processed_files = 0
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []

    try:
        for file_path in source_files:
            file_name = file_path.name
            if file_name.startswith("~"):
                print(f"Skipping {file_name}: temp file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipping {file_name}: not an .xlsx file")
                continue

            print(f"Processing {file_name}")
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                labels = parse_file_labels(file_name)

                empirical_sheet = get_sheet_or_none(wb, "Empirical Model")
                if empirical_sheet is None:
                    print(f"Skipping empirical extraction for {file_name}: sheet missing")
                else:
                    empirical_rows.extend(extract_empirical_rows(empirical_sheet, labels, file_name))

                regression_sheet = get_sheet_or_none(wb, "Regression Model")
                if regression_sheet is None:
                    print(f"Skipping regression extraction for {file_name}: sheet missing")
                else:
                    regression_rows.extend(extract_regression_rows(regression_sheet, labels, file_name))

                processed_files += 1
            except Exception as exc:
                print(f"Skipping {file_name}: {exc}")
            finally:
                if wb is not None:
                    safe_close_book(wb)
    finally:
        app.quit()

    write_output_workbook(empirical_rows, regression_rows, output_path)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")

    return output_path


def main() -> None:
    process_files(input_folder=input_dir, output_folder=output_dir)


if __name__ == "__main__":
    main()

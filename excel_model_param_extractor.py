#!/usr/bin/env python3
"""
Extract empirical and regression model candidates from .xlsx files.

The script opens each source workbook only once, processes both model sheets while
the workbook is open, and writes a single output workbook with:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

from datetime import date
from pathlib import Path
import re
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

try:
    import xlwings as xw
except ImportError:  # pragma: no cover - depends on host machine setup
    xw = None  # type: ignore[assignment]
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter
except ImportError:  # pragma: no cover - depends on host machine setup
    Workbook = None  # type: ignore[assignment]
    Font = None  # type: ignore[assignment]
    get_column_letter = None  # type: ignore[assignment]

# User inputs
input_dir = "./input"
output_dir = "./output"

EMPIRICAL_COLUMNS: List[str] = [
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

REGRESSION_COLUMNS: List[str] = [
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
    "oct": 10,
    "nov": 11,
    "dec": 12,
}

DAY_BY_PERIOD = {"early": 5, "mid": 15, "late": 25}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def to_float(value: Any) -> Optional[float]:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    is_pct = text.endswith("%")
    if is_pct:
        text = text[:-1]
    try:
        number = float(text)
    except ValueError:
        return None
    return number / 100.0 if is_pct else number


def as_2d_list(value: Any) -> List[List[Any]]:
    if value is None:
        return []
    if isinstance(value, list):
        if not value:
            return []
        if isinstance(value[0], list):
            return value
        return [value]
    return [[value]]


def first_supported_sheet(wb: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    try:
        return wb.sheets[sheet_name]
    except Exception:
        return None


def read_used_matrix(sheet: xw.Sheet) -> Tuple[List[List[Any]], int, int]:
    used = sheet.used_range
    matrix = as_2d_list(used.value)
    return matrix, used.row, used.column


def find_anchor_cell(
    matrix: Sequence[Sequence[Any]],
    start_row: int,
    start_col: int,
    target: str = "max",
) -> Optional[Tuple[int, int]]:
    target_norm = normalize_text(target)
    for r_idx, row_vals in enumerate(matrix):
        for c_idx, cell_val in enumerate(row_vals):
            if normalize_text(cell_val) == target_norm:
                return start_row + r_idx, start_col + c_idx
    return None


def safe_close_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.api.Close(SaveChanges=False)
    except Exception:
        try:
            wb.close()
        except Exception:
            # Last-resort fallback: workbook was not saved and will be discarded.
            pass


def parse_file_labels(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    parts = [p.strip() for p in stem.split(" - ") if p.strip()]

    ticker = parts[1] if len(parts) >= 2 else "UNKNOWN"

    period_token = ""
    if len(parts) >= 3:
        period_token = parts[2]
    elif parts:
        period_token = parts[-1]

    period_token = re.sub(r"_send.*$", "", period_token, flags=re.IGNORECASE)
    period_compact = re.sub(r"[^A-Za-z0-9]", "", period_token)

    match = re.search(
        r"(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})",
        period_compact,
        flags=re.IGNORECASE,
    )

    model_period = "unknown_period"
    model_date = ""

    if match:
        bucket = match.group(1).title()
        month_token_raw = match.group(2)
        year = int(match.group(3))
        month_token = month_token_raw[:3].title()
        month_num = MONTHS.get(month_token.lower())

        if month_num is not None:
            model_period = f"{bucket}{month_token}_{year}"
            day = DAY_BY_PERIOD[bucket.lower()]
            model_date = date(year, month_num, day).isoformat()
        else:
            model_period = f"{bucket}{month_token}_{year}"

    if ticker == "UNKNOWN":
        model = model_period
    else:
        model = f"{ticker}_{model_period}"

    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def resolve_output_path(input_path: Path, output_path: Path) -> Path:
    base_name = f"{input_path.name}_PARAM.xlsx"
    candidate = output_path / base_name
    if not candidate.exists():
        return candidate

    idx = 1
    while True:
        candidate = output_path / f"{input_path.name}_PARAM.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def get_header_candidates(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    radius: int = 30,
) -> List[Tuple[int, str]]:
    left = max(1, anchor_col - radius)
    right = max(left, anchor_col + radius)
    values = sheet.range((anchor_row, left), (anchor_row, right)).value
    if not isinstance(values, list):
        values = [values]
    return [(left + idx, normalize_text(val)) for idx, val in enumerate(values)]


def find_column(
    headers: Sequence[Tuple[int, str]],
    patterns: Sequence[Sequence[str]],
    fallback: Optional[int] = None,
) -> Optional[int]:
    for pattern in patterns:
        for col, header in headers:
            if header and all(part in header for part in pattern):
                return col
    if fallback is None or fallback < 1:
        return None
    return fallback


def sheet_value(sheet: xw.Sheet, row: int, col: Optional[int]) -> Any:
    if row < 1 or col is None or col < 1:
        return None
    try:
        return sheet.range((row, col)).value
    except Exception:
        return None


def value_or_none(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, str) and not value.strip():
        return None
    return value


def ensure_len(values: Any, target_len: int) -> List[Any]:
    if isinstance(values, list):
        if len(values) >= target_len:
            return values[:target_len]
        return values + [None] * (target_len - len(values))
    if values is None:
        return [None] * target_len
    return [values] + [None] * (target_len - 1)


def almost_equal(a: Any, b: Any, tol: float = 1e-9) -> bool:
    fa = to_float(a)
    fb = to_float(b)
    if fa is not None and fb is not None:
        return abs(fa - fb) <= tol
    return value_or_none(a) == value_or_none(b)


def is_duplicate_regression_row(
    previous: Dict[str, Any],
    current: Dict[str, Any],
) -> bool:
    keys = (
        "num_quarters_used",
        "forecast_value",
        "forecast_max",
        "forecast_min",
        "intercept",
        "slope",
    )
    return all(almost_equal(previous.get(key), current.get(key)) for key in keys)


def calculate_empirical_avg_penetration(
    sheet: xw.Sheet,
    anchor_row: int,
    captured_col: Optional[int],
    n_quarters: int = 10,
) -> List[Any]:
    if captured_col is None or captured_col < 1 or anchor_row < 2:
        return [None] * n_quarters

    used = sheet.used_range
    scratch_row = used.last_cell.row + 2
    scratch_col = used.last_cell.column + 2
    end_row = anchor_row - 1

    for n in range(1, n_quarters + 1):
        start_row = max(1, end_row - n + 1)
        formula = f"=AVERAGE(R{start_row}C{captured_col}:R{end_row}C{captured_col})"
        sheet.range((scratch_row + n - 1, scratch_col)).formula2 = formula

    sheet.book.app.calculate()
    calculated = sheet.range(
        (scratch_row, scratch_col),
        (scratch_row + n_quarters - 1, scratch_col),
    ).value
    return ensure_len(calculated, n_quarters)


def calculate_regression_stats(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    n_quarters: int = 10,
) -> Tuple[List[Any], List[Any], List[Any]]:
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if y_col < 1 or x_col < 1 or anchor_row < 2:
        blank = [None] * n_quarters
        return blank, blank, blank

    used = sheet.used_range
    scratch_row = used.last_cell.row + 2
    scratch_col = used.last_cell.column + 2
    data_end = anchor_row - 1

    for n in range(1, n_quarters + 1):
        start_row = max(1, data_end - n + 1)
        i_cell = sheet.range((scratch_row + n - 1, scratch_col))
        s_cell = sheet.range((scratch_row + n - 1, scratch_col + 1))
        f_cell = sheet.range((scratch_row + n - 1, scratch_col + 2))

        i_cell.formula2 = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{data_end}C{y_col},"
            f"R{start_row}C{x_col}:R{data_end}C{x_col})"
        )
        s_cell.formula2 = (
            f"=SLOPE(R{start_row}C{y_col}:R{data_end}C{y_col},"
            f"R{start_row}C{x_col}:R{data_end}C{x_col})"
        )
        f_cell.formula2 = f"=RC[-2]+RC[-1]*R{data_end}C{x_col}"

    sheet.book.app.calculate()

    intercepts = sheet.range(
        (scratch_row, scratch_col),
        (scratch_row + n_quarters - 1, scratch_col),
    ).value
    slopes = sheet.range(
        (scratch_row, scratch_col + 1),
        (scratch_row + n_quarters - 1, scratch_col + 1),
    ).value
    forecasts = sheet.range(
        (scratch_row, scratch_col + 2),
        (scratch_row + n_quarters - 1, scratch_col + 2),
    ).value

    return (
        ensure_len(intercepts, n_quarters),
        ensure_len(slopes, n_quarters),
        ensure_len(forecasts, n_quarters),
    )


def extract_empirical_rows(
    wb: xw.Book,
    labels: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = first_supported_sheet(wb, "Empirical Model")
    if sheet is None:
        print(f"skipped file: {source_file} (missing Empirical Model sheet)")
        return []

    matrix, start_row, start_col = read_used_matrix(sheet)
    anchor = find_anchor_cell(matrix, start_row, start_col, target="max")
    if anchor is None:
        print(f"skipped file: {source_file} (Empirical Model max anchor not found)")
        return []

    anchor_row, anchor_col = anchor
    headers = get_header_candidates(sheet, anchor_row, anchor_col)

    num_q_col = find_column(
        headers,
        patterns=[("num", "quarter"), ("quarters", "used")],
        fallback=anchor_col - 9,
    )
    last_q_col = find_column(
        headers,
        patterns=[("last", "quarter"), ("quarter", "used")],
        fallback=anchor_col - 8,
    )
    forecast_col = find_column(
        headers,
        patterns=[
            ("estimated", "total", "sold"),
            ("estimate", "total", "sold"),
            ("forecast", "value"),
            ("forecast",),
        ],
        fallback=anchor_col - 1,
    )
    min_col = find_column(headers, patterns=[("min",)], fallback=anchor_col + 1)
    avg_pen_col = find_column(
        headers,
        patterns=[("avg", "penetration"), ("average", "penetration"), ("penetration",)],
        fallback=anchor_col - 6,
    )
    quarterly_sales_col = find_column(
        headers,
        patterns=[("quarterly", "sales"), ("qtr", "sales")],
        fallback=anchor_col - 5,
    )
    reported_sales_col = find_column(
        headers,
        patterns=[("reported", "sales"), ("actual", "sales"), ("actual",)],
        fallback=anchor_col - 4,
    )
    growth_col = find_column(
        headers,
        patterns=[("growth", "rate"), ("growth",)],
        fallback=anchor_col - 3,
    )
    captured_col = find_column(
        headers,
        patterns=[
            ("captured", "db"),
            ("sales", "captured"),
            ("penetration", "db"),
        ],
        fallback=anchor_col - 2,
    )

    avg_calculated = calculate_empirical_avg_penetration(
        sheet=sheet,
        anchor_row=anchor_row,
        captured_col=captured_col,
        n_quarters=10,
    )

    rows: List[Dict[str, Any]] = []
    for idx in range(10):
        row_number = anchor_row + idx + 1
        num_quarters_used = value_or_none(sheet_value(sheet, row_number, num_q_col))
        if num_quarters_used is None:
            num_quarters_used = idx + 1

        last_quarter_used = value_or_none(sheet_value(sheet, row_number, last_q_col))
        forecast_value = value_or_none(sheet_value(sheet, row_number, forecast_col))
        forecast_max = value_or_none(sheet_value(sheet, row_number, anchor_col))
        forecast_min = value_or_none(sheet_value(sheet, row_number, min_col))
        avg_penetration_pct = value_or_none(sheet_value(sheet, row_number, avg_pen_col))
        quarterly_sales = value_or_none(sheet_value(sheet, row_number, quarterly_sales_col))
        reported_sales = value_or_none(sheet_value(sheet, row_number, reported_sales_col))
        growth_rate_pct = value_or_none(sheet_value(sheet, row_number, growth_col))
        sales_captured_pct = value_or_none(sheet_value(sheet, row_number, captured_col))

        if avg_penetration_pct is None:
            avg_penetration_pct = value_or_none(avg_calculated[idx])

        if forecast_value is None:
            q_sales = to_float(quarterly_sales)
            avg_pen = to_float(avg_penetration_pct)
            if q_sales is not None and avg_pen not in (None, 0):
                forecast_value = q_sales / avg_pen

        range_width = None
        fmax = to_float(forecast_max)
        fmin = to_float(forecast_min)
        if fmax is not None and fmin is not None:
            range_width = fmax - fmin

        signal_values = (
            forecast_value,
            forecast_max,
            forecast_min,
            avg_penetration_pct,
            quarterly_sales,
            reported_sales,
        )
        if all(value_or_none(v) is None for v in signal_values):
            continue

        rows.append(
            {
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
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_pct,
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_rows(
    wb: xw.Book,
    labels: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = first_supported_sheet(wb, "Regression Model")
    if sheet is None:
        print(f"skipped file: {source_file} (missing Regression Model sheet)")
        return []

    matrix, start_row, start_col = read_used_matrix(sheet)
    anchor = find_anchor_cell(matrix, start_row, start_col, target="max")
    if anchor is None:
        print(f"skipped file: {source_file} (Regression Model max anchor not found)")
        return []

    anchor_row, anchor_col = anchor
    headers = get_header_candidates(sheet, anchor_row, anchor_col)

    num_q_col = find_column(
        headers,
        patterns=[("num", "quarter"), ("quarters", "used")],
        fallback=anchor_col - 4,
    )
    forecast_col = find_column(
        headers,
        patterns=[
            ("tot", "fcst", "w", "o", "sa"),
            ("total", "forecast", "without", "sa"),
            ("forecast", "value"),
            ("forecast",),
        ],
        fallback=anchor_col - 2,
    )
    min_col = find_column(headers, patterns=[("min",)], fallback=anchor_col + 1)
    actual_col = find_column(
        headers,
        patterns=[("actual", "sales"), ("reported", "sales"), ("actual",)],
        fallback=None,
    )

    intercepts, slopes, forecasts_calc = calculate_regression_stats(
        sheet=sheet,
        anchor_row=anchor_row,
        anchor_col=anchor_col,
        n_quarters=10,
    )

    rows: List[Dict[str, Any]] = []
    for idx in range(10):
        row_number = anchor_row + idx + 1
        num_quarters_used = value_or_none(sheet_value(sheet, row_number, num_q_col))
        if num_quarters_used is None:
            num_quarters_used = idx + 1

        forecast_value = value_or_none(sheet_value(sheet, row_number, forecast_col))
        if forecast_value is None:
            forecast_value = value_or_none(forecasts_calc[idx])

        forecast_max = value_or_none(sheet_value(sheet, row_number, anchor_col))
        forecast_min = value_or_none(sheet_value(sheet, row_number, min_col))
        actual_value = value_or_none(sheet_value(sheet, row_number, actual_col))
        intercept = value_or_none(intercepts[idx])
        slope = value_or_none(slopes[idx])

        fmax = to_float(forecast_max)
        fmin = to_float(forecast_min)
        range_width = None
        if fmax is not None and fmin is not None:
            range_width = fmax - fmin

        signal_values = (
            forecast_value,
            forecast_max,
            forecast_min,
            intercept,
            slope,
        )
        if all(value_or_none(v) is None for v in signal_values):
            continue

        row_payload = {
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
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        if rows and is_duplicate_regression_row(rows[-1], row_payload):
            continue

        rows.append(row_payload)

    return rows


def write_sheet(
    workbook: Workbook,
    sheet_name: str,
    columns: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    ws = workbook.create_sheet(title=sheet_name)
    ws.append(list(columns))

    for row in rows:
        ws.append([row.get(col) for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row_idx in range(2, ws.max_row + 1):
            cell_value = ws.cell(row=row_idx, column=col_idx).value
            cell_len = len("" if cell_value is None else str(cell_value))
            if cell_len > max_len:
                max_len = cell_len
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 48)


def write_output_workbook(
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
    output_path: Path,
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    write_sheet(
        workbook=wb,
        sheet_name="empirical_candidates",
        columns=EMPIRICAL_COLUMNS,
        rows=empirical_rows,
    )
    write_sheet(
        workbook=wb,
        sheet_name="regression_candidates",
        columns=REGRESSION_COLUMNS,
        rows=regression_rows,
    )
    wb.save(output_path)


def iter_source_files(input_path: Path) -> Iterable[Path]:
    for file_path in sorted(input_path.iterdir()):
        if not file_path.is_file():
            print(f"skipped file: {file_path.name} (not a file)")
            continue
        if file_path.name.startswith("~"):
            print(f"skipped file: {file_path.name} (temporary file)")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"skipped file: {file_path.name} (not .xlsx)")
            continue
        yield file_path


def main() -> None:
    if xw is None:
        print("xlwings is required. Install it with: pip install xlwings")
        return
    if Workbook is None or Font is None or get_column_letter is None:
        print("openpyxl is required. Install it with: pip install openpyxl")
        return

    input_path = Path(input_dir).expanduser()
    output_path_dir = Path(output_dir).expanduser()

    if not input_path.exists() or not input_path.is_dir():
        print(f"Input directory does not exist: {input_path}")
        return

    output_path_dir.mkdir(parents=True, exist_ok=True)
    output_path = resolve_output_path(input_path=input_path, output_path=output_path_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in iter_source_files(input_path):
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                labels = parse_file_labels(file_path)

                empirical_rows.extend(
                    extract_empirical_rows(
                        wb=wb,
                        labels=labels,
                        source_file=file_path.name,
                    )
                )
                regression_rows.extend(
                    extract_regression_rows(
                        wb=wb,
                        labels=labels,
                        source_file=file_path.name,
                    )
                )

                files_processed += 1
                print(f"processed file: {file_path.name}")
            except Exception as exc:
                print(f"skipped file: {file_path.name} (open/process error: {exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        try:
            app.quit()
        except Exception:
            pass

    write_output_workbook(
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
        output_path=output_path,
    )

    print(f"output path: {output_path}")
    print(f"number of files processed: {files_processed}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

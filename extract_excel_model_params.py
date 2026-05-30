from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any, Iterable, Sequence

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# User-configurable paths
input_dir = Path("input")
output_dir = Path("output")

EMPIRICAL_SHEET = "Empirical Model"
REGRESSION_SHEET = "Regression Model"
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

FILE_LABEL_RE = re.compile(
    r"""
    -\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*
    (?P<period>
        (?P<phase>Early|Mid|Late)
        (?P<month>Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)
        (?P<year>\d{4})
    )
    """,
    re.IGNORECASE | re.VERBOSE,
)

MONTH_TO_NUMBER = {
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


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[\s_/\\-]+", " ", str(value).strip().lower())


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def to_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    try:
        text = str(value).replace(",", "").replace("%", "").strip()
        return float(text)
    except (TypeError, ValueError):
        return None


def to_int(value: Any) -> int | None:
    numeric = to_float(value)
    if numeric is None:
        return None
    return int(round(numeric))


def ensure_2d(values: Any, expected_rows: int) -> list[list[Any]]:
    if expected_rows <= 0:
        return []
    if isinstance(values, list):
        if not values:
            return [[None] for _ in range(expected_rows)]
        if isinstance(values[0], list):
            matrix = values
        else:
            matrix = [values]
    else:
        matrix = [[values]]

    while len(matrix) < expected_rows:
        matrix.append([None] * len(matrix[0]))
    return matrix


def write_formula2(rng: xw.main.Range, formula: str) -> None:
    try:
        rng.formula2 = formula
    except Exception:
        try:
            rng.api.Formula2 = formula
        except Exception:
            rng.formula = formula


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

    try:
        wb.close()
    except Exception:
        pass


def get_sheet(wb: xw.Book, sheet_name: str) -> xw.Sheet | None:
    try:
        return wb.sheets[sheet_name]
    except Exception:
        return None


def find_anchor_max(sheet: xw.Sheet) -> tuple[int, int] | None:
    used = sheet.used_range
    values = ensure_2d(used.value, used.rows.count)
    base_row = used.row
    base_col = used.column

    for r_idx, row_values in enumerate(values):
        for c_idx, cell in enumerate(row_values):
            if normalize_text(cell) == "max":
                return base_row + r_idx, base_col + c_idx
    return None


def scan_header_tokens(
    sheet: xw.Sheet,
    header_row: int,
    anchor_col: int,
    span: int = 40,
) -> list[tuple[str, int]]:
    left = max(1, anchor_col - span)
    right = max(left, anchor_col + span)
    row_values = sheet.range((header_row, left), (header_row, right)).value
    if not isinstance(row_values, list):
        row_values = [row_values]
    tokens: list[tuple[str, int]] = []
    for idx, value in enumerate(row_values):
        norm = normalize_text(value)
        if norm:
            tokens.append((norm, left + idx))
    return tokens


def find_col(
    header_tokens: Sequence[tuple[str, int]],
    patterns: Iterable[Iterable[str]],
    default: int | None = None,
) -> int | None:
    for pattern in patterns:
        needle = [word.lower() for word in pattern]
        for token, col in header_tokens:
            if all(word in token for word in needle):
                return col
    return default


def parse_file_label(file_path: Path) -> dict[str, str]:
    stem = file_path.stem
    match = FILE_LABEL_RE.search(stem)
    ticker = "UNKNOWN"
    model_period = "UNKNOWN"
    model_date = ""

    if match:
        ticker = match.group("ticker").upper()
        phase = match.group("phase").capitalize()
        month = match.group("month").capitalize()
        year = int(match.group("year"))
        model_period = f"{phase}{month}_{year}"
        month_num = MONTH_TO_NUMBER[month.lower()]
        day = PHASE_TO_DAY[phase.lower()]
        model_date = date(year, month_num, day).isoformat()
    else:
        parts = [part.strip() for part in stem.split("-")]
        if len(parts) >= 2 and parts[1]:
            ticker = parts[1].replace(" ", "").upper()

    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": f"{ticker}_{model_period}",
    }


def next_output_path(in_dir: Path, out_dir: Path) -> Path:
    folder_name = in_dir.resolve().name
    base = f"{folder_name}_PARAM"
    candidate = out_dir / f"{base}.xlsx"
    if not candidate.exists():
        return candidate

    index = 1
    while True:
        candidate = out_dir / f"{base}.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def safe_width(value: Any) -> int:
    if value is None:
        return 0
    return len(str(value))


def compute_range_width(max_value: Any, min_value: Any) -> float | None:
    max_num = to_float(max_value)
    min_num = to_float(min_value)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


def values_match(left: Any, right: Any, tolerance: float = 1e-10) -> bool:
    l_num = to_float(left)
    r_num = to_float(right)
    if l_num is not None and r_num is not None:
        return abs(l_num - r_num) <= tolerance
    return left == right


def pull_block(
    sheet: xw.Sheet,
    start_row: int,
    n_rows: int,
    cols: Sequence[int | None],
) -> tuple[list[list[Any]], int, int]:
    valid_cols = sorted({col for col in cols if col is not None and col > 0})
    if not valid_cols:
        return [], 0, 0
    left_col = valid_cols[0]
    right_col = valid_cols[-1]
    values = sheet.range(
        (start_row, left_col),
        (start_row + n_rows - 1, right_col),
    ).value
    return ensure_2d(values, n_rows), left_col, right_col


def extract_empirical_rows(
    wb: xw.Book,
    metadata: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    sheet = get_sheet(wb, EMPIRICAL_SHEET)
    if sheet is None:
        return []

    anchor = find_anchor_max(sheet)
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    header_tokens = scan_header_tokens(sheet, anchor_row, anchor_col)
    data_start = anchor_row + 1

    num_q_col = find_col(
        header_tokens,
        [
            ("num", "quarters", "used"),
            ("quarters", "used"),
            ("n", "quarters"),
        ],
        default=max(1, anchor_col - 8),
    )
    last_q_col = find_col(
        header_tokens,
        [
            ("last", "quarter", "used"),
            ("last", "quarter"),
        ],
        default=max(1, anchor_col - 7),
    )
    forecast_col = find_col(
        header_tokens,
        [
            ("estimated", "total", "sold"),
            ("forecast", "value"),
            ("total", "forecast"),
            ("tot", "fcst"),
        ],
        default=max(1, anchor_col - 1),
    )
    actual_col = find_col(
        header_tokens,
        [
            ("reported", "sales"),
            ("actual", "value"),
            ("actual",),
        ],
        default=max(1, anchor_col - 2),
    )
    max_col = anchor_col
    min_col = find_col(
        header_tokens,
        [
            ("min",),
        ],
        default=anchor_col + 1,
    )
    quarterly_sales_col = find_col(
        header_tokens,
        [
            ("quarterly", "sales"),
            ("quarter", "sales"),
        ],
        default=max(1, anchor_col - 5),
    )
    reported_sales_col = find_col(
        header_tokens,
        [
            ("reported", "sales"),
            ("sales", "reported"),
        ],
        default=actual_col,
    )
    growth_col = find_col(
        header_tokens,
        [
            ("growth", "rate"),
            ("growth",),
        ],
        default=max(1, anchor_col - 4),
    )
    sales_captured_col = find_col(
        header_tokens,
        [
            ("sales", "captured", "db"),
            ("captured", "db"),
        ],
    )
    avg_pen_col = find_col(
        header_tokens,
        [
            ("avg", "penetration"),
            ("average", "penetration"),
        ],
    )
    penetration_series_col = find_col(
        header_tokens,
        [
            ("penetration",),
        ],
    )
    if penetration_series_col == avg_pen_col:
        penetration_series_col = None

    matrix, left_col, _ = pull_block(
        sheet,
        data_start,
        N_QUARTERS,
        [
            num_q_col,
            last_q_col,
            forecast_col,
            actual_col,
            max_col,
            min_col,
            quarterly_sales_col,
            reported_sales_col,
            growth_col,
            sales_captured_col,
            avg_pen_col,
        ],
    )

    if not matrix:
        return []

    def value_at(row_index: int, col: int | None) -> Any:
        if col is None or col < left_col:
            return None
        col_idx = col - left_col
        row = matrix[row_index]
        if col_idx >= len(row):
            return None
        return row[col_idx]

    quarter_counts: list[int] = []
    for idx in range(N_QUARTERS):
        explicit_q = to_int(value_at(idx, num_q_col))
        quarter_counts.append(explicit_q if explicit_q and explicit_q > 0 else idx + 1)

    helper_col = max(c for c in [max_col, min_col or 0, forecast_col or 0] if c) + 2
    wrote_formula = False
    for idx, n_quarters in enumerate(quarter_counts):
        row_num = data_start + idx
        start_row = data_start
        end_row = data_start + n_quarters - 1
        if penetration_series_col:
            formula = (
                f'=IFERROR(AVERAGE('
                f'R{start_row}C{penetration_series_col}:R{end_row}C{penetration_series_col}'
                f'),"")'
            )
            write_formula2(sheet.range((row_num, helper_col)), formula)
            wrote_formula = True
        elif quarterly_sales_col and reported_sales_col:
            formula = (
                f'=IFERROR('
                f'SUM(R{start_row}C{quarterly_sales_col}:R{end_row}C{quarterly_sales_col})/'
                f'SUM(R{start_row}C{reported_sales_col}:R{end_row}C{reported_sales_col})'
                f',"")'
            )
            write_formula2(sheet.range((row_num, helper_col)), formula)
            wrote_formula = True

    avg_pen_values: list[Any] = [None] * N_QUARTERS
    if wrote_formula:
        wb.app.calculate()
        avg_matrix = ensure_2d(
            sheet.range((data_start, helper_col), (data_start + N_QUARTERS - 1, helper_col)).value,
            N_QUARTERS,
        )
        avg_pen_values = [avg_matrix[idx][0] for idx in range(N_QUARTERS)]
    elif avg_pen_col:
        avg_pen_values = [value_at(idx, avg_pen_col) for idx in range(N_QUARTERS)]

    rows: list[dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        forecast_value = value_at(idx, forecast_col)
        actual_value = value_at(idx, actual_col)
        forecast_max = value_at(idx, max_col)
        forecast_min = value_at(idx, min_col)

        if all(is_blank(v) for v in (forecast_value, actual_value, forecast_max, forecast_min)):
            continue

        avg_penetration = avg_pen_values[idx]
        range_width = compute_range_width(forecast_max, forecast_min)

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": quarter_counts[idx],
                "last_quarter_used": value_at(idx, last_q_col),
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": value_at(idx, quarterly_sales_col),
                "reported_sales": value_at(idx, reported_sales_col) or actual_value,
                "growth_rate_pct": value_at(idx, growth_col),
                "sales_captured_in_db_pct": value_at(idx, sales_captured_col),
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_rows(
    wb: xw.Book,
    metadata: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    sheet = get_sheet(wb, REGRESSION_SHEET)
    if sheet is None:
        return []

    anchor = find_anchor_max(sheet)
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    y_col = max(1, anchor_col - 7)
    x_col = max(1, anchor_col - 11)
    header_tokens = scan_header_tokens(sheet, anchor_row, anchor_col)
    data_start = anchor_row + 1

    num_q_col = find_col(
        header_tokens,
        [
            ("num", "quarters", "used"),
            ("quarters", "used"),
            ("n", "quarters"),
        ],
        default=max(1, anchor_col - 8),
    )
    forecast_col = find_col(
        header_tokens,
        [
            ("tot", "fcst", "w", "o", "sa"),
            ("total", "forecast", "without", "sa"),
            ("fcst", "without", "sa"),
        ],
        default=max(1, anchor_col - 1),
    )
    actual_col = find_col(
        header_tokens,
        [
            ("actual",),
            ("reported", "sales"),
        ],
    )
    max_col = anchor_col
    min_col = find_col(
        header_tokens,
        [
            ("min",),
        ],
        default=anchor_col + 1,
    )

    matrix, left_col, right_col = pull_block(
        sheet,
        data_start,
        N_QUARTERS,
        [num_q_col, forecast_col, actual_col, max_col, min_col],
    )
    if not matrix:
        return []

    def value_at(row_index: int, col: int | None) -> Any:
        if col is None or col < left_col:
            return None
        col_idx = col - left_col
        row = matrix[row_index]
        if col_idx >= len(row):
            return None
        return row[col_idx]

    quarter_counts: list[int] = []
    for idx in range(N_QUARTERS):
        explicit_q = to_int(value_at(idx, num_q_col))
        quarter_counts.append(explicit_q if explicit_q and explicit_q > 0 else idx + 1)

    intercept_col = right_col + 2
    slope_col = right_col + 3
    history_end = anchor_row - 1
    wrote_formula = False

    if history_end >= 1:
        for idx, n_quarters in enumerate(quarter_counts):
            calc_row = data_start + idx
            history_start = max(1, history_end - n_quarters + 1)
            intercept_formula = (
                f'=IFERROR(INTERCEPT('
                f'R{history_start}C{y_col}:R{history_end}C{y_col},'
                f'R{history_start}C{x_col}:R{history_end}C{x_col}'
                f'),"")'
            )
            slope_formula = (
                f'=IFERROR(SLOPE('
                f'R{history_start}C{y_col}:R{history_end}C{y_col},'
                f'R{history_start}C{x_col}:R{history_end}C{x_col}'
                f'),"")'
            )
            write_formula2(sheet.range((calc_row, intercept_col)), intercept_formula)
            write_formula2(sheet.range((calc_row, slope_col)), slope_formula)
            wrote_formula = True

    intercept_values: list[Any] = [None] * N_QUARTERS
    slope_values: list[Any] = [None] * N_QUARTERS
    if wrote_formula:
        wb.app.calculate()
        intercept_matrix = ensure_2d(
            sheet.range((data_start, intercept_col), (data_start + N_QUARTERS - 1, intercept_col)).value,
            N_QUARTERS,
        )
        slope_matrix = ensure_2d(
            sheet.range((data_start, slope_col), (data_start + N_QUARTERS - 1, slope_col)).value,
            N_QUARTERS,
        )
        intercept_values = [intercept_matrix[idx][0] for idx in range(N_QUARTERS)]
        slope_values = [slope_matrix[idx][0] for idx in range(N_QUARTERS)]

    rows: list[dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        forecast_value = value_at(idx, forecast_col)
        actual_value = value_at(idx, actual_col)
        forecast_max = value_at(idx, max_col)
        forecast_min = value_at(idx, min_col)

        if all(is_blank(v) for v in (forecast_value, forecast_max, forecast_min, actual_value)):
            continue

        row = {
            "model": metadata["model"],
            "ticker": metadata["ticker"],
            "model_period": metadata["model_period"],
            "model_date": metadata["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": quarter_counts[idx],
            "num_quarters_used": quarter_counts[idx],
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": compute_range_width(forecast_max, forecast_min),
            "intercept": intercept_values[idx],
            "slope": slope_values[idx],
            "source_file": source_file,
        }

        if rows:
            prev = rows[-1]
            duplicate = all(
                values_match(row[key], prev[key])
                for key in (
                    "parameter_value",
                    "forecast_value",
                    "forecast_max",
                    "forecast_min",
                    "intercept",
                    "slope",
                )
            )
            if duplicate:
                continue

        rows.append(row)

    return rows


def write_sheet(
    wb: Workbook,
    sheet_name: str,
    headers: Sequence[str],
    rows: Sequence[dict[str, Any]],
) -> None:
    ws = wb.create_sheet(title=sheet_name)
    ws.append(list(headers))

    for row in rows:
        ws.append([row.get(header) for header in headers])

    for col_idx in range(1, len(headers) + 1):
        ws.cell(row=1, column=col_idx).font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            max_len = max(max_len, safe_width(value))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(60, max(12, max_len + 2))


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[dict[str, Any]],
    regression_rows: Sequence[dict[str, Any]],
) -> None:
    out_wb = Workbook()
    out_wb.remove(out_wb.active)
    write_sheet(out_wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(out_wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    out_wb.save(output_path)


def should_skip_file(file_path: Path) -> tuple[bool, str]:
    if not file_path.is_file():
        return True, "not a file"
    if file_path.name.startswith("~"):
        return True, "temporary workbook"
    if file_path.suffix.lower() != ".xlsx":
        return True, "not .xlsx"
    if re.search(r"_PARAM(\.\d+)?\.xlsx$", file_path.name, flags=re.IGNORECASE):
        return True, "generated output workbook"
    return False, ""


def main() -> None:
    in_dir = Path(input_dir)
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    if not in_dir.exists():
        raise FileNotFoundError(f"Input directory does not exist: {in_dir}")

    output_path = next_output_path(in_dir, out_dir)
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        app.calculation = "manual"
    except Exception:
        pass

    try:
        for file_path in sorted(in_dir.iterdir()):
            should_skip, reason = should_skip_file(file_path)
            if should_skip:
                print(f"SKIPPED: {file_path.name} ({reason})")
                continue

            print(f"PROCESSING: {file_path.name}")
            metadata = parse_file_label(file_path)
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_rows(wb, metadata, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, metadata, file_path.name))
                files_processed += 1
            except Exception as exc:
                print(f"SKIPPED: {file_path.name} (error: {exc})")
            finally:
                if wb is not None:
                    close_workbook_without_save(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"OUTPUT: {output_path}")
    print(f"FILES PROCESSED: {files_processed}")
    print(f"EMPIRICAL ROWS: {len(empirical_rows)}")
    print(f"REGRESSION ROWS: {len(regression_rows)}")


if __name__ == "__main__":
    main()

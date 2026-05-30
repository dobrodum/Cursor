from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

try:
    import xlwings as xw
except ImportError as exc:
    raise SystemExit(
        "xlwings is required for this script (Excel desktop automation). "
        "Install it with: pip install xlwings"
    ) from exc


# User-configurable paths
input_dir = Path("./input")
output_dir = Path("./output")


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

PERIOD_DAY = {
    "early": 5,
    "mid": 15,
    "late": 25,
}

PERIOD_PATTERN = re.compile(
    r"(?i)(early|mid|late)\s*([a-z]{3,9})\s*[_-]?\s*(\d{4})"
)


@dataclass
class FileMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


def month_abbr(month_token: str) -> str:
    return month_token.strip()[:3].title()


def parse_filename_metadata(file_path: Path) -> FileMetadata:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split("-")]

    ticker = "UNKNOWN"
    if len(parts) >= 2 and parts[1]:
        ticker = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper() or "UNKNOWN"

    period_source = stem
    if len(parts) >= 3 and parts[2]:
        period_source = parts[2]

    match = PERIOD_PATTERN.search(period_source) or PERIOD_PATTERN.search(stem)
    if not match:
        model_period = "unknown_period"
        model_date = ""
    else:
        period_word, month_word, year_str = match.groups()
        period_title = period_word.title()
        month_title = month_abbr(month_word)
        month_num = MONTHS.get(month_title.lower())
        day = PERIOD_DAY.get(period_word.lower())
        model_period = f"{period_title}{month_title}_{year_str}"
        if month_num is not None and day is not None:
            model_date = date(int(year_str), month_num, day).isoformat()
        else:
            model_date = ""

    model = f"{ticker}_{model_period}"
    return FileMetadata(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def ensure_list(values: Any, size: int) -> list[Any]:
    if isinstance(values, tuple):
        values = list(values)
    if not isinstance(values, list):
        return [values] * size
    if len(values) >= size:
        return values[:size]
    return values + [None] * (size - len(values))


def to_number(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "")
        if cleaned.endswith("%"):
            cleaned = cleaned[:-1]
        value = cleaned
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def pick_output_path(src_input_dir: Path, dst_output_dir: Path) -> Path:
    dst_output_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{src_input_dir.name}_PARAM"
    first_choice = dst_output_dir / f"{base_name}.xlsx"
    if not first_choice.exists():
        return first_choice

    index = 1
    while True:
        candidate = dst_output_dir / f"{base_name}.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def close_without_saving(wb: xw.Book) -> None:
    attempts = [
        lambda: wb.close(save=False),
        lambda: wb.close(False),
        lambda: wb.api.Close(SaveChanges=False),
        lambda: wb.api.Close(False),
    ]
    for close_call in attempts:
        try:
            close_call()
            return
        except Exception:
            continue


def _to_matrix(values: Any) -> list[list[Any]]:
    if isinstance(values, tuple):
        values = list(values)
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return [[]]
    if isinstance(values[0], tuple):
        values = [list(row) for row in values]
    if not isinstance(values[0], list):
        return [values]
    return values


def find_max_anchor(sheet: xw.Sheet) -> tuple[int, int] | None:
    used = sheet.used_range
    matrix = _to_matrix(used.value)

    candidates: list[tuple[int, int, bool]] = []
    for r_idx, row in enumerate(matrix):
        for c_idx, cell in enumerate(row):
            if normalize_text(cell) != "max":
                continue
            right_cell = row[c_idx + 1] if c_idx + 1 < len(row) else None
            has_min_on_right = normalize_text(right_cell) == "min"
            candidates.append((r_idx, c_idx, has_min_on_right))

    if not candidates:
        return None

    selected = next((item for item in candidates if item[2]), candidates[0])
    return (used.row + selected[0], used.column + selected[1])


def discover_columns(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    synonyms: dict[str, list[str]],
    col_window: int = 22,
) -> dict[str, int]:
    first_col = max(1, anchor_col - col_window)
    last_col = anchor_col + col_window
    top_row = max(1, anchor_row - 2)
    bottom_row = anchor_row + 2

    block = sheet.range((top_row, first_col), (bottom_row, last_col)).value
    matrix = _to_matrix(block)

    found: dict[str, int] = {}
    for r_offset, row in enumerate(matrix):
        _ = r_offset  # Helps readability while preserving row iteration
        for c_offset, cell in enumerate(row):
            text = normalize_text(cell)
            if not text:
                continue
            abs_col = first_col + c_offset
            for key, terms in synonyms.items():
                if key in found:
                    continue
                if any(term in text for term in terms):
                    found[key] = abs_col
    return found


def read_column_block(
    sheet: xw.Sheet, row_start: int, row_count: int, col: int | None
) -> list[Any]:
    if col is None or col < 1:
        return [None] * row_count
    values = sheet.range((row_start, col), (row_start + row_count - 1, col)).value
    return ensure_list(values, row_count)


def read_field_blocks(
    sheet: xw.Sheet, row_start: int, row_count: int, field_cols: dict[str, int | None]
) -> dict[str, list[Any]]:
    valid_cols = [col for col in field_cols.values() if isinstance(col, int) and col >= 1]
    if not valid_cols:
        return {key: [None] * row_count for key in field_cols}

    min_col = min(valid_cols)
    max_col = max(valid_cols)
    raw_block = sheet.range(
        (row_start, min_col), (row_start + row_count - 1, max_col)
    ).value
    matrix = _to_matrix(raw_block)

    if len(matrix) < row_count:
        matrix.extend([[]] * (row_count - len(matrix)))

    output: dict[str, list[Any]] = {}
    for field, col in field_cols.items():
        if col is None or col < 1:
            output[field] = [None] * row_count
            continue
        rel_col = col - min_col
        output[field] = [
            row[rel_col] if rel_col < len(row) else None for row in matrix[:row_count]
        ]
    return output


def extract_empirical_rows(
    wb: xw.Book, metadata: FileMetadata, source_file: str
) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        return []

    anchor = find_max_anchor(sheet)
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    fallback = {
        "num_quarters_used": anchor_col - 12,
        "last_quarter_used": anchor_col - 11,
        "forecast_value": anchor_col - 3,
        "actual_value": anchor_col - 2,
        "forecast_max": anchor_col,
        "forecast_min": anchor_col + 1,
        "quarterly_sales": anchor_col - 10,
        "reported_sales": anchor_col - 9,
        "growth_rate_pct": anchor_col - 8,
        "sales_captured_in_db_pct": anchor_col - 7,
        "avg_penetration_pct": anchor_col - 6,
    }

    synonyms = {
        "num_quarters_used": ["num quarters", "quarters used", "n quarters"],
        "last_quarter_used": ["last quarter", "latest quarter"],
        "forecast_value": ["estimated total sold", "tot fcst", "forecast"],
        "actual_value": ["actual", "reported sales"],
        "forecast_max": ["max"],
        "forecast_min": ["min"],
        "quarterly_sales": ["quarterly sales", "q sales"],
        "reported_sales": ["reported sales", "sales in db"],
        "growth_rate_pct": ["growth rate", "growth pct", "growth"],
        "sales_captured_in_db_pct": [
            "sales captured in db",
            "captured in db",
            "captured",
        ],
        "avg_penetration_pct": ["avg penetration", "average penetration"],
    }

    discovered = discover_columns(sheet, anchor_row, anchor_col, synonyms)
    cols = {key: discovered.get(key, fallback.get(key)) for key in fallback}

    row_start = anchor_row + 1
    row_count = N_QUARTERS

    values = read_field_blocks(sheet, row_start, row_count, cols)

    used_last_col = sheet.used_range.last_cell.column
    scratch_col = max(used_last_col + 2, anchor_col + 25)
    scratch_vals: list[Any] = [None] * row_count
    formulas_written = False

    quarterly_col = cols.get("quarterly_sales")
    reported_col = cols.get("reported_sales")
    if (
        isinstance(quarterly_col, int)
        and quarterly_col > 0
        and isinstance(reported_col, int)
        and reported_col > 0
    ):
        for idx in range(row_count):
            row_num = row_start + idx
            n_used = values["num_quarters_used"][idx]
            n_quarters = int(n_used) if str(n_used).strip().isdigit() else idx + 1
            n_quarters = max(1, min(N_QUARTERS, n_quarters))
            range_start = max(row_start, row_num - n_quarters + 1)
            formula = (
                f'=IFERROR(AVERAGE(IFERROR(R{range_start}C{quarterly_col}:'
                f'R{row_num}C{quarterly_col}/R{range_start}C{reported_col}:'
                f'R{row_num}C{reported_col},"")),"")'
            )
            sheet.range((row_num, scratch_col)).formula2 = formula
            formulas_written = True

        if formulas_written:
            wb.app.calculate()
            scratch_vals = read_column_block(sheet, row_start, row_count, scratch_col)

    rows: list[dict[str, Any]] = []
    for idx in range(row_count):
        num_q = values["num_quarters_used"][idx]
        if num_q in (None, ""):
            num_q = idx + 1

        avg_pen = values["avg_penetration_pct"][idx]
        if avg_pen in (None, "") and formulas_written:
            avg_pen = scratch_vals[idx]

        forecast_max = values["forecast_max"][idx]
        forecast_min = values["forecast_min"][idx]
        max_num = to_number(forecast_max)
        min_num = to_number(forecast_min)
        range_width = (
            (max_num - min_num) if max_num is not None and min_num is not None else None
        )

        row = {
            "model": metadata.model,
            "ticker": metadata.ticker,
            "model_period": metadata.model_period,
            "model_date": metadata.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_pen,
            "num_quarters_used": num_q,
            "last_quarter_used": values["last_quarter_used"][idx],
            "forecast_value": values["forecast_value"][idx],
            "actual_value": values["actual_value"][idx],
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "avg_penetration_pct": avg_pen,
            "quarterly_sales": values["quarterly_sales"][idx],
            "reported_sales": values["reported_sales"][idx],
            "growth_rate_pct": values["growth_rate_pct"][idx],
            "sales_captured_in_db_pct": values["sales_captured_in_db_pct"][idx],
            "source_file": source_file,
        }

        material_fields = [
            row["forecast_value"],
            row["actual_value"],
            row["forecast_max"],
            row["forecast_min"],
            row["avg_penetration_pct"],
        ]
        if all(item in (None, "") for item in material_fields):
            continue
        rows.append(row)

    return rows


def _sig(value: Any) -> Any:
    number = to_number(value)
    if number is not None:
        return round(number, 8)
    if value in ("", None):
        return None
    return str(value).strip()


def extract_regression_rows(
    wb: xw.Book, metadata: FileMetadata, source_file: str
) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        return []

    anchor = find_max_anchor(sheet)
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if y_col < 1 or x_col < 1:
        return []

    fallback = {
        "num_quarters_used": x_col - 1,
        "forecast_value": anchor_col - 3,
        "actual_value": anchor_col - 2,
        "forecast_max": anchor_col,
        "forecast_min": anchor_col + 1,
    }

    synonyms = {
        "num_quarters_used": ["num quarters", "quarters used", "n quarters"],
        "forecast_value": ["tot fcst w o sa", "tot fcst wo sa", "tot fcst", "forecast"],
        "actual_value": ["actual", "reported sales"],
        "forecast_max": ["max"],
        "forecast_min": ["min"],
    }
    discovered = discover_columns(sheet, anchor_row, anchor_col, synonyms)
    cols = {key: discovered.get(key, fallback.get(key)) for key in fallback}

    row_start = anchor_row + 1
    row_count = N_QUARTERS
    values = read_field_blocks(sheet, row_start, row_count, cols)

    used_last_col = sheet.used_range.last_cell.column
    intercept_col = max(used_last_col + 2, anchor_col + 22)
    slope_col = intercept_col + 1

    for idx in range(row_count):
        row_num = row_start + idx
        n_used = values["num_quarters_used"][idx]
        n_quarters = int(n_used) if str(n_used).strip().isdigit() else idx + 1
        n_quarters = max(2, min(N_QUARTERS, n_quarters))
        range_start = max(row_start, row_num - n_quarters + 1)
        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{range_start}C{y_col}:R{row_num}C{y_col},'
            f'R{range_start}C{x_col}:R{row_num}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{range_start}C{y_col}:R{row_num}C{y_col},'
            f'R{range_start}C{x_col}:R{row_num}C{x_col}),"")'
        )
        sheet.range((row_num, intercept_col)).formula2 = intercept_formula
        sheet.range((row_num, slope_col)).formula2 = slope_formula

    wb.app.calculate()
    intercept_values = read_column_block(sheet, row_start, row_count, intercept_col)
    slope_values = read_column_block(sheet, row_start, row_count, slope_col)

    rows: list[dict[str, Any]] = []
    for idx in range(row_count):
        num_q = values["num_quarters_used"][idx]
        if num_q in (None, ""):
            num_q = idx + 1

        forecast_max = values["forecast_max"][idx]
        forecast_min = values["forecast_min"][idx]
        max_num = to_number(forecast_max)
        min_num = to_number(forecast_min)
        range_width = (
            (max_num - min_num) if max_num is not None and min_num is not None else None
        )

        row = {
            "model": metadata.model,
            "ticker": metadata.ticker,
            "model_period": metadata.model_period,
            "model_date": metadata.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_q,
            "num_quarters_used": num_q,
            "forecast_value": values["forecast_value"][idx],
            "actual_value": values["actual_value"][idx],
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept_values[idx],
            "slope": slope_values[idx],
            "source_file": source_file,
        }

        material_fields = [
            row["forecast_value"],
            row["forecast_max"],
            row["forecast_min"],
            row["intercept"],
            row["slope"],
        ]
        if all(item in (None, "") for item in material_fields):
            continue

        if idx == row_count - 1 and rows:
            previous = rows[-1]
            current_sig = (
                _sig(row["num_quarters_used"]),
                _sig(row["intercept"]),
                _sig(row["slope"]),
                _sig(row["forecast_value"]),
                _sig(row["forecast_max"]),
                _sig(row["forecast_min"]),
            )
            previous_sig = (
                _sig(previous["num_quarters_used"]),
                _sig(previous["intercept"]),
                _sig(previous["slope"]),
                _sig(previous["forecast_value"]),
                _sig(previous["forecast_max"]),
                _sig(previous["forecast_min"]),
            )
            if current_sig == previous_sig:
                continue

        rows.append(row)

    return rows


def format_sheet(ws: Any, columns: list[str], rows: list[dict[str, Any]]) -> None:
    ws.append(columns)
    for row in rows:
        ws.append([row.get(column) for column in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for idx, col_name in enumerate(columns, start=1):
        values = [col_name]
        values.extend(
            "" if row.get(col_name) is None else str(row.get(col_name)) for row in rows
        )
        max_len = max((len(item) for item in values), default=len(col_name))
        ws.column_dimensions[get_column_letter(idx)].width = min(54, max(12, max_len + 2))


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    empirical_ws = wb.create_sheet("empirical_candidates")
    format_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)

    regression_ws = wb.create_sheet("regression_candidates")
    format_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def iter_input_files(src_input_dir: Path) -> list[Path]:
    files: list[Path] = []
    for file_path in sorted(src_input_dir.iterdir(), key=lambda item: item.name.lower()):
        if not file_path.is_file():
            continue
        files.append(file_path)
    return files


def main() -> None:
    src_input_dir = input_dir.expanduser().resolve()
    dst_output_dir = output_dir.expanduser().resolve()

    if not src_input_dir.exists():
        raise FileNotFoundError(f"Input directory does not exist: {src_input_dir}")
    if not src_input_dir.is_dir():
        raise NotADirectoryError(f"Input path is not a directory: {src_input_dir}")

    output_path = pick_output_path(src_input_dir, dst_output_dir)

    processed_files = 0
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []

    app: xw.App | None = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in iter_input_files(src_input_dir):
            if file_path.name.startswith("~"):
                print(f"skipped file: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped file: {file_path.name} (not .xlsx)")
                continue

            print(f"processing file: {file_path.name}")

            wb: xw.Book | None = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                metadata = parse_filename_metadata(file_path)
                empirical_rows.extend(
                    extract_empirical_rows(wb, metadata, file_path.name)
                )
                regression_rows.extend(
                    extract_regression_rows(wb, metadata, file_path.name)
                )
                processed_files += 1
                print(f"processed file: {file_path.name}")
            except Exception as exc:
                print(f"skipped file: {file_path.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    close_without_saving(wb)
    finally:
        if app is not None:
            app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"output path: {output_path}")
    print(f"number of files processed: {processed_files}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

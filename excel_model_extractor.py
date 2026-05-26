from __future__ import annotations

import calendar
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -----------------------------------------------------------------------------
# Configure paths here
# -----------------------------------------------------------------------------
input_dir = Path("/workspace/input")
output_dir = Path("/workspace/output")

# -----------------------------------------------------------------------------
# Constants
# -----------------------------------------------------------------------------
EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"
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

DAY_BY_PERIOD_PREFIX = {"early": 5, "mid": 15, "late": 25}


@dataclass(frozen=True)
class ParsedLabel:
    ticker: str
    model_period: str
    model_date: str
    model: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def slug_text(value: Any) -> str:
    text = normalize_text(value)
    text = re.sub(r"[%()/\-]+", " ", text)
    text = re.sub(r"[^a-z0-9 ]+", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def to_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, (list, tuple)):
        return [[values]]
    if values and isinstance(values[0], (list, tuple)):
        return [list(row) for row in values]
    return [list(values)]


def parse_file_label(file_path: Path) -> ParsedLabel:
    stem = file_path.stem

    ticker = ""
    ticker_match = re.search(r"\bmodel\s*-\s*([A-Za-z0-9]+)\s*-\s*", stem, flags=re.IGNORECASE)
    if ticker_match:
        ticker = ticker_match.group(1).upper()
    else:
        parts = [p.strip() for p in stem.split("-")]
        if len(parts) >= 2:
            ticker = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper()

    period_match = re.search(
        r"\b(Early|Mid|Late)\s*[-_ ]?\s*([A-Za-z]{3,9})\s*[-_ ]?\s*(\d{4})\b",
        stem,
        flags=re.IGNORECASE,
    )
    if not period_match:
        raise ValueError(f"Cannot parse period from filename: {file_path.name}")

    period_prefix_raw = period_match.group(1)
    month_raw = period_match.group(2)
    year = int(period_match.group(3))

    period_prefix = period_prefix_raw.capitalize()
    month_abbr = month_raw[:3].title()
    month_num = list(calendar.month_abbr).index(month_abbr)
    if month_num <= 0:
        raise ValueError(f"Invalid month in filename: {file_path.name}")

    day = DAY_BY_PERIOD_PREFIX[period_prefix.lower()]
    model_period = f"{period_prefix}{month_abbr}_{year}"
    model_date = date(year, month_num, day).isoformat()
    model = f"{ticker}_{model_period}"

    return ParsedLabel(
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
        model=model,
    )


def get_output_path(input_folder: Path, destination: Path) -> Path:
    destination.mkdir(parents=True, exist_ok=True)
    base = f"{input_folder.name}_PARAM"
    candidate = destination / f"{base}.xlsx"
    if not candidate.exists():
        return candidate

    index = 1
    while True:
        numbered = destination / f"{base}.{index}.xlsx"
        if not numbered.exists():
            return numbered
        index += 1


def safe_close_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
    except TypeError:
        try:
            wb.api.Close(SaveChanges=False)
        except Exception:
            wb.close()


def iter_source_files(folder: Path) -> Iterable[Path]:
    for file_path in sorted(folder.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped: {file_path.name} (temporary file)")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped: {file_path.name} (not .xlsx)")
            continue
        yield file_path


def get_used_range_matrix(sheet: xw.Sheet) -> Tuple[int, int, List[List[Any]]]:
    used = sheet.used_range
    base_row = used.row
    base_col = used.column
    matrix = to_2d(used.value)
    return base_row, base_col, matrix


def find_anchor_max(
    base_row: int,
    base_col: int,
    matrix: Sequence[Sequence[Any]],
) -> Tuple[int, int]:
    for r_idx, row in enumerate(matrix):
        for c_idx, value in enumerate(row):
            if normalize_text(value) == "max":
                return base_row + r_idx, base_col + c_idx
    raise ValueError("Could not find 'max' anchor.")


def header_row_values(
    matrix: Sequence[Sequence[Any]],
    base_row: int,
    anchor_row: int,
) -> List[Any]:
    idx = anchor_row - base_row
    if idx < 0 or idx >= len(matrix):
        return []
    return list(matrix[idx])


def find_col_in_header(
    row_values: Sequence[Any],
    base_col: int,
    required_tokens: Sequence[str],
    forbidden_tokens: Sequence[str] = (),
) -> Optional[int]:
    for idx, value in enumerate(row_values):
        text = slug_text(value)
        if not text:
            continue
        if any(token not in text for token in required_tokens):
            continue
        if any(token in text for token in forbidden_tokens):
            continue
        return base_col + idx
    return None


def pick_col(
    discovered: Optional[int],
    anchor_col: int,
    fallback_offset: int,
) -> int:
    if discovered is not None:
        return discovered
    return anchor_col + fallback_offset


def nfloat(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace(",", "").strip()
    if text == "":
        return None
    try:
        return float(text)
    except ValueError:
        return None


def nint(value: Any) -> Optional[int]:
    num = nfloat(value)
    if num is None:
        return None
    try:
        return int(round(num))
    except (TypeError, ValueError):
        return None


def as_value(value: Any) -> Any:
    if value is None:
        return None
    return value


def build_empirical_column_map(
    row_values: Sequence[Any],
    base_col: int,
    anchor_col: int,
) -> Dict[str, int]:
    return {
        "num_quarters_used": pick_col(
            find_col_in_header(row_values, base_col, ("quarter",), ("last",)),
            anchor_col,
            -9,
        ),
        "last_quarter_used": pick_col(
            find_col_in_header(row_values, base_col, ("last", "quarter")),
            anchor_col,
            -8,
        ),
        "forecast_value": pick_col(
            find_col_in_header(row_values, base_col, ("estimated", "sold")),
            anchor_col,
            -4,
        ),
        "actual_value": pick_col(
            find_col_in_header(row_values, base_col, ("reported", "sales")),
            anchor_col,
            -3,
        ),
        "forecast_max": pick_col(
            find_col_in_header(row_values, base_col, ("max",)),
            anchor_col,
            0,
        ),
        "forecast_min": pick_col(
            find_col_in_header(row_values, base_col, ("min",)),
            anchor_col,
            1,
        ),
        "avg_penetration_pct": pick_col(
            find_col_in_header(row_values, base_col, ("penetration",)),
            anchor_col,
            -6,
        ),
        "quarterly_sales": pick_col(
            find_col_in_header(row_values, base_col, ("quarterly", "sales")),
            anchor_col,
            -2,
        ),
        "reported_sales": pick_col(
            find_col_in_header(row_values, base_col, ("reported", "sales")),
            anchor_col,
            -3,
        ),
        "growth_rate_pct": pick_col(
            find_col_in_header(row_values, base_col, ("growth",)),
            anchor_col,
            2,
        ),
        "sales_captured_in_db_pct": pick_col(
            find_col_in_header(row_values, base_col, ("captured", "db")),
            anchor_col,
            3,
        ),
    }


def build_regression_column_map(
    row_values: Sequence[Any],
    base_col: int,
    anchor_col: int,
) -> Dict[str, int]:
    return {
        "num_quarters_used": pick_col(
            find_col_in_header(row_values, base_col, ("quarter",), ("last",)),
            anchor_col,
            -9,
        ),
        "forecast_value": pick_col(
            find_col_in_header(row_values, base_col, ("tot", "fcst", "sa"), ("with",)),
            anchor_col,
            -2,
        ),
        "forecast_max": pick_col(
            find_col_in_header(row_values, base_col, ("max",)),
            anchor_col,
            0,
        ),
        "forecast_min": pick_col(
            find_col_in_header(row_values, base_col, ("min",)),
            anchor_col,
            1,
        ),
        "intercept": pick_col(
            find_col_in_header(row_values, base_col, ("intercept",)),
            anchor_col,
            2,
        ),
        "slope": pick_col(
            find_col_in_header(row_values, base_col, ("slope",)),
            anchor_col,
            3,
        ),
    }


def read_cell(sheet: xw.Sheet, row: int, col: int) -> Any:
    return sheet.cells(row, col).value


def process_empirical_sheet(
    wb: xw.Book,
    sheet: xw.Sheet,
    label: ParsedLabel,
    source_file: str,
) -> List[Dict[str, Any]]:
    base_row, base_col, matrix = get_used_range_matrix(sheet)
    anchor_row, anchor_col = find_anchor_max(base_row, base_col, matrix)
    hdr_values = header_row_values(matrix, base_row, anchor_row)
    cols = build_empirical_column_map(hdr_values, base_col, anchor_col)

    start_row = anchor_row + 1

    # Formula updates are batched; calculate once after writing all formulas.
    for i in range(N_QUARTERS):
        row = start_row + i
        n_quarters = nint(read_cell(sheet, row, cols["num_quarters_used"])) or (i + 1)
        avg_cell = sheet.cells(row, cols["avg_penetration_pct"])
        avg_cell.formula2 = f"=IFERROR(AVERAGE(R[-{n_quarters}]C:R[-1]C),0)"

    wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    for i in range(N_QUARTERS):
        row = start_row + i
        num_quarters_used = nint(read_cell(sheet, row, cols["num_quarters_used"])) or (i + 1)
        last_quarter_used = as_value(read_cell(sheet, row, cols["last_quarter_used"]))
        forecast_value = as_value(read_cell(sheet, row, cols["forecast_value"]))
        actual_value = as_value(read_cell(sheet, row, cols["actual_value"]))
        forecast_max = as_value(read_cell(sheet, row, cols["forecast_max"]))
        forecast_min = as_value(read_cell(sheet, row, cols["forecast_min"]))
        avg_penetration_pct = as_value(read_cell(sheet, row, cols["avg_penetration_pct"]))
        quarterly_sales = as_value(read_cell(sheet, row, cols["quarterly_sales"]))
        reported_sales = as_value(read_cell(sheet, row, cols["reported_sales"]))
        growth_rate_pct = as_value(read_cell(sheet, row, cols["growth_rate_pct"]))
        sales_captured_in_db_pct = as_value(read_cell(sheet, row, cols["sales_captured_in_db_pct"]))

        if (
            forecast_value is None
            and forecast_max is None
            and forecast_min is None
            and avg_penetration_pct is None
        ):
            continue

        max_num = nfloat(forecast_max)
        min_num = nfloat(forecast_min)
        range_width = (max_num - min_num) if (max_num is not None and min_num is not None) else None

        rows.append(
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
    label: ParsedLabel,
    source_file: str,
) -> List[Dict[str, Any]]:
    base_row, base_col, matrix = get_used_range_matrix(sheet)
    anchor_row, anchor_col = find_anchor_max(base_row, base_col, matrix)
    hdr_values = header_row_values(matrix, base_row, anchor_row)
    cols = build_regression_column_map(hdr_values, base_col, anchor_col)

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    start_row = anchor_row + 1

    for i in range(N_QUARTERS):
        row = start_row + i
        n_quarters = nint(read_cell(sheet, row, cols["num_quarters_used"])) or (i + 1)
        rng_start = max(start_row, row - n_quarters + 1)

        intercept_cell = sheet.cells(row, cols["intercept"])
        slope_cell = sheet.cells(row, cols["slope"])

        intercept_cell.formula2 = (
            f'=IFERROR(INTERCEPT(R{rng_start}C{y_col}:R{row}C{y_col},'
            f'R{rng_start}C{x_col}:R{row}C{x_col}),"")'
        )
        slope_cell.formula2 = (
            f'=IFERROR(SLOPE(R{rng_start}C{y_col}:R{row}C{y_col},'
            f'R{rng_start}C{x_col}:R{row}C{x_col}),"")'
        )

    wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    prev_signature: Optional[Tuple[Any, ...]] = None

    for i in range(N_QUARTERS):
        row = start_row + i
        num_quarters_used = nint(read_cell(sheet, row, cols["num_quarters_used"])) or (i + 1)
        forecast_value = as_value(read_cell(sheet, row, cols["forecast_value"]))
        forecast_max = as_value(read_cell(sheet, row, cols["forecast_max"]))
        forecast_min = as_value(read_cell(sheet, row, cols["forecast_min"]))
        intercept = as_value(read_cell(sheet, row, cols["intercept"]))
        slope = as_value(read_cell(sheet, row, cols["slope"]))

        if (
            forecast_value is None
            and forecast_max is None
            and forecast_min is None
            and intercept in (None, "")
            and slope in (None, "")
        ):
            continue

        max_num = nfloat(forecast_max)
        min_num = nfloat(forecast_min)
        range_width = (max_num - min_num) if (max_num is not None and min_num is not None) else None

        signature = (
            num_quarters_used,
            nfloat(intercept) if intercept not in ("", None) else intercept,
            nfloat(slope) if slope not in ("", None) else slope,
            nfloat(forecast_value) if forecast_value not in ("", None) else forecast_value,
            max_num,
            min_num,
        )
        if signature == prev_signature:
            continue
        prev_signature = signature

        rows.append(
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
                "actual_value": None,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def write_sheet(
    wb: Workbook,
    sheet_name: str,
    headers: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    ws = wb.create_sheet(title=sheet_name)
    ws.append(list(headers))

    for row in rows:
        ws.append([row.get(col) for col in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for value, in ws.iter_rows(
            min_row=2,
            max_row=ws.max_row,
            min_col=col_idx,
            max_col=col_idx,
            values_only=True,
        ):
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 48)


def run() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"Input directory does not exist: {input_dir}")

    output_path = get_output_path(input_dir, output_dir)
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.calculation = "manual"

    try:
        for file_path in iter_source_files(input_dir):
            print(f"Processing: {file_path.name}")
            try:
                label = parse_file_label(file_path)
            except ValueError as exc:
                print(f"Skipped: {file_path.name} ({exc})")
                continue

            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                sheet_names = {s.name: s for s in wb.sheets}

                if EMPIRICAL_SHEET_NAME in sheet_names:
                    empirical_rows.extend(
                        process_empirical_sheet(
                            wb=wb,
                            sheet=sheet_names[EMPIRICAL_SHEET_NAME],
                            label=label,
                            source_file=file_path.name,
                        )
                    )
                else:
                    print(f"Skipped empirical in {file_path.name} (sheet not found)")

                if REGRESSION_SHEET_NAME in sheet_names:
                    regression_rows.extend(
                        process_regression_sheet(
                            wb=wb,
                            sheet=sheet_names[REGRESSION_SHEET_NAME],
                            label=label,
                            source_file=file_path.name,
                        )
                    )
                else:
                    print(f"Skipped regression in {file_path.name} (sheet not found)")

                processed_files += 1
            except Exception as exc:
                print(f"Skipped: {file_path.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    out_wb = Workbook()
    default_ws = out_wb.active
    out_wb.remove(default_ws)

    write_sheet(out_wb, "empirical_candidates", EMPIRICAL_HEADERS, empirical_rows)
    write_sheet(out_wb, "regression_candidates", REGRESSION_HEADERS, regression_rows)
    out_wb.save(output_path)

    print(f"Output: {output_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run()

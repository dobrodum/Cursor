from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, NamedTuple, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ------------------------------
# User-configurable paths
# ------------------------------
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


PHASE_DAY_MAP = {"early": 5, "mid": 15, "late": 25}
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


class FileMetadata(NamedTuple):
    model: str
    ticker: str
    model_period: str
    model_date: str
    source_file: str


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return re.sub(r"[^a-z0-9]+", " ", value.lower()).strip()
    return re.sub(r"[^a-z0-9]+", " ", str(value).lower()).strip()


def is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def to_float(value: Any) -> Optional[float]:
    if is_number(value):
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "")
        if not cleaned:
            return None
        if cleaned.endswith("%"):
            try:
                return float(cleaned[:-1]) / 100.0
            except ValueError:
                return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def round_for_signature(value: Any) -> Any:
    numeric = to_float(value)
    if numeric is None:
        return value
    return round(numeric, 8)


def safe_value(value: Any) -> Any:
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return value


def as_row_matrix(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def parse_file_metadata(file_path: Path) -> Optional[FileMetadata]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = ""
    if len(parts) >= 2:
        ticker = parts[1].strip().upper()
    else:
        ticker_match = re.search(r"\b([A-Z]{2,6})\b", stem)
        if ticker_match:
            ticker = ticker_match.group(1).upper()

    period_match = re.search(
        r"(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(\d{4})",
        stem,
        flags=re.IGNORECASE,
    )
    if not period_match or not ticker:
        return None

    phase = period_match.group(1).title()
    month_abbrev = period_match.group(2).title()
    year = period_match.group(3)

    day = PHASE_DAY_MAP[phase.lower()]
    month = MONTH_MAP[month_abbrev.lower()]
    model_period = f"{phase}{month_abbrev}_{year}"
    model_date = date(int(year), month, day).isoformat()

    return FileMetadata(
        model=f"{ticker}_{model_period}",
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
        source_file=file_path.name,
    )


def build_output_path(in_dir: Path, out_dir: Path) -> Path:
    input_folder_name = in_dir.name or "input"
    base_name = f"{input_folder_name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    idx = 1
    while candidate.exists():
        candidate = out_dir / f"{base_name}.{idx}.xlsx"
        idx += 1
    return candidate


def get_sheet_if_exists(wb: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    try:
        return wb.sheets[sheet_name]
    except Exception:
        return None


def find_anchor_max(sheet: xw.Sheet) -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    used_values = as_row_matrix(used.value)
    base_row = used.row
    base_col = used.column

    for r_idx, row_values in enumerate(used_values):
        for c_idx, cell_value in enumerate(row_values):
            if isinstance(cell_value, str) and cell_value.strip().lower() == "max":
                return base_row + r_idx, base_col + c_idx
    return None


def build_header_map(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    span: int = 24,
) -> Dict[str, int]:
    left_col = max(1, anchor_col - span)
    right_col = anchor_col + span
    header_map: Dict[str, int] = {}

    for row in (anchor_row - 1, anchor_row, anchor_row + 1):
        if row < 1:
            continue
        row_values = sheet.range((row, left_col), (row, right_col)).value
        row_values = row_values if isinstance(row_values, list) else [row_values]
        for idx, value in enumerate(row_values):
            label = normalize_label(value)
            if label:
                header_map[label] = left_col + idx
    return header_map


def resolve_column(
    header_map: Dict[str, int],
    label_candidates: Sequence[str],
    anchor_col: int,
    default_offset: int,
) -> int:
    for candidate in label_candidates:
        pattern = normalize_label(candidate)
        for label, col in header_map.items():
            if pattern and pattern in label:
                return col
    return anchor_col + default_offset


def get_col_value(row_values: Sequence[Any], start_col: int, target_col: int) -> Any:
    idx = target_col - start_col
    if idx < 0 or idx >= len(row_values):
        return None
    return row_values[idx]


def read_numeric_rows_upward(
    sheet: xw.Sheet,
    col: int,
    start_row: int,
    max_scan: int = 240,
) -> List[int]:
    if start_row < 1:
        return []
    top_row = max(1, start_row - max_scan + 1)
    values = sheet.range((top_row, col), (start_row, col)).value
    values = values if isinstance(values, list) else [values]

    rows: List[int] = []
    for idx in range(len(values) - 1, -1, -1):
        if to_float(values[idx]) is not None:
            rows.append(top_row + idx)
        elif rows:
            # Stop at first gap once we started collecting the trailing block.
            break
    rows.reverse()
    return rows


def read_paired_numeric_rows_upward(
    sheet: xw.Sheet,
    x_col: int,
    y_col: int,
    start_row: int,
    max_scan: int = 240,
) -> List[int]:
    if start_row < 1:
        return []
    top_row = max(1, start_row - max_scan + 1)
    x_values = sheet.range((top_row, x_col), (start_row, x_col)).value
    y_values = sheet.range((top_row, y_col), (start_row, y_col)).value
    x_values = x_values if isinstance(x_values, list) else [x_values]
    y_values = y_values if isinstance(y_values, list) else [y_values]

    rows: List[int] = []
    for idx in range(len(x_values) - 1, -1, -1):
        if to_float(x_values[idx]) is not None and to_float(y_values[idx]) is not None:
            rows.append(top_row + idx)
        elif rows:
            break
    rows.reverse()
    return rows


def set_formula2_r1c1(target: xw.Range, formula_r1c1: str) -> None:
    # Requirement asks for R1C1 formulas via formula2.
    try:
        target.formula2 = formula_r1c1
        return
    except Exception:
        pass
    try:
        target.api.Formula2R1C1 = formula_r1c1
        return
    except Exception:
        pass
    target.formula = formula_r1c1


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
    except Exception:
        try:
            wb.close()
        except Exception:
            pass


def process_empirical_sheet(wb: xw.Book, meta: FileMetadata) -> List[Dict[str, Any]]:
    sheet = get_sheet_if_exists(wb, "Empirical Model")
    if sheet is None:
        return []

    anchor = find_anchor_max(sheet)
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    header_map = build_header_map(sheet, anchor_row, anchor_col)
    cols = {
        "num_quarters_used": resolve_column(
            header_map, ["num quarters used", "quarters used"], anchor_col, -6
        ),
        "last_quarter_used": resolve_column(
            header_map, ["last quarter used", "last quarter"], anchor_col, -5
        ),
        "avg_penetration_pct": resolve_column(
            header_map, ["avg penetration", "average penetration"], anchor_col, -4
        ),
        "forecast_value": resolve_column(
            header_map,
            ["estimated total sold", "total sold estimate", "forecast total"],
            anchor_col,
            -3,
        ),
        "actual_value": resolve_column(
            header_map, ["reported sales", "actual sales"], anchor_col, -2
        ),
        "quarterly_sales": resolve_column(
            header_map, ["quarterly sales"], anchor_col, -1
        ),
        "forecast_max": anchor_col,
        "forecast_min": resolve_column(header_map, ["min"], anchor_col, 1),
        "growth_rate_pct": resolve_column(
            header_map, ["growth rate", "growth rate pct"], anchor_col, 2
        ),
        "sales_captured_in_db_pct": resolve_column(
            header_map, ["sales captured in db", "sales captured"], anchor_col, 3
        ),
        "reported_sales": resolve_column(
            header_map, ["reported sales", "actual sales"], anchor_col, -2
        ),
    }

    penetration_source_col = resolve_column(
        header_map,
        ["penetration", "penetration pct", "quarter penetration"],
        anchor_col,
        -10,
    )
    penetration_history_rows = read_numeric_rows_upward(
        sheet, penetration_source_col, anchor_row - 1, max_scan=240
    )

    formulas_written = False
    n_quarters = 10
    for idx in range(n_quarters):
        row = anchor_row + 1 + idx
        num_q = idx + 1

        num_q_cell = sheet.range((row, cols["num_quarters_used"]))
        if to_float(num_q_cell.value) != float(num_q):
            num_q_cell.value = num_q
            formulas_written = True

        if len(penetration_history_rows) >= num_q:
            start = penetration_history_rows[-num_q]
            end = penetration_history_rows[-1]
            avg_formula = (
                f'=IFERROR(AVERAGE(R{start}C{penetration_source_col}:'
                f'R{end}C{penetration_source_col}),"")'
            )
            set_formula2_r1c1(sheet.range((row, cols["avg_penetration_pct"])), avg_formula)
            formulas_written = True

    if formulas_written:
        wb.app.calculate()

    read_start_col = min(cols.values())
    read_end_col = max(cols.values())
    table_values = as_row_matrix(
        sheet.range(
            (anchor_row + 1, read_start_col),
            (anchor_row + n_quarters, read_end_col),
        ).value
    )

    rows: List[Dict[str, Any]] = []
    for idx, row_values in enumerate(table_values):
        num_quarters_used = get_col_value(row_values, read_start_col, cols["num_quarters_used"])
        if to_float(num_quarters_used) is None:
            num_quarters_used = idx + 1

        forecast_max = get_col_value(row_values, read_start_col, cols["forecast_max"])
        forecast_min = get_col_value(row_values, read_start_col, cols["forecast_min"])
        forecast_max_n = to_float(forecast_max)
        forecast_min_n = to_float(forecast_min)
        range_width = (
            forecast_max_n - forecast_min_n
            if forecast_max_n is not None and forecast_min_n is not None
            else None
        )

        avg_pen = get_col_value(row_values, read_start_col, cols["avg_penetration_pct"])
        reported_sales = get_col_value(row_values, read_start_col, cols["reported_sales"])

        entry = {
            "model": meta.model,
            "ticker": meta.ticker,
            "model_period": meta.model_period,
            "model_date": meta.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": safe_value(avg_pen),
            "num_quarters_used": safe_value(num_quarters_used),
            "last_quarter_used": safe_value(
                get_col_value(row_values, read_start_col, cols["last_quarter_used"])
            ),
            "forecast_value": safe_value(
                get_col_value(row_values, read_start_col, cols["forecast_value"])
            ),
            "actual_value": safe_value(
                get_col_value(row_values, read_start_col, cols["actual_value"])
            ),
            "forecast_max": safe_value(forecast_max),
            "forecast_min": safe_value(forecast_min),
            "range_width": safe_value(range_width),
            "avg_penetration_pct": safe_value(avg_pen),
            "quarterly_sales": safe_value(
                get_col_value(row_values, read_start_col, cols["quarterly_sales"])
            ),
            "reported_sales": safe_value(reported_sales),
            "growth_rate_pct": safe_value(
                get_col_value(row_values, read_start_col, cols["growth_rate_pct"])
            ),
            "sales_captured_in_db_pct": safe_value(
                get_col_value(row_values, read_start_col, cols["sales_captured_in_db_pct"])
            ),
            "source_file": meta.source_file,
        }

        rows.append(entry)

    return rows


def process_regression_sheet(wb: xw.Book, meta: FileMetadata) -> List[Dict[str, Any]]:
    sheet = get_sheet_if_exists(wb, "Regression Model")
    if sheet is None:
        return []

    anchor = find_anchor_max(sheet)
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    # Required anchor-based source columns.
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    header_map = build_header_map(sheet, anchor_row, anchor_col)
    cols = {
        "num_quarters_used": resolve_column(
            header_map, ["num quarters used", "quarters used"], anchor_col, -2
        ),
        "forecast_value": resolve_column(
            header_map,
            ["tot fcst w o sa", "tot fcst w/o sa", "forecast total without sa", "total fcst"],
            anchor_col,
            -1,
        ),
        "forecast_max": anchor_col,
        "forecast_min": resolve_column(header_map, ["min"], anchor_col, 1),
        "actual_value": resolve_column(
            header_map, ["actual", "actual value", "reported sales"], anchor_col, 4
        ),
        "intercept": resolve_column(header_map, ["intercept"], anchor_col, 2),
        "slope": resolve_column(header_map, ["slope"], anchor_col, 3),
    }

    paired_rows = read_paired_numeric_rows_upward(
        sheet, x_col, y_col, anchor_row - 1, max_scan=240
    )
    max_rows = min(10, len(paired_rows)) if paired_rows else 10

    formulas_written = False
    for idx in range(max_rows):
        row = anchor_row + 1 + idx
        n_quarters = idx + 1

        num_q_cell = sheet.range((row, cols["num_quarters_used"]))
        if to_float(num_q_cell.value) != float(n_quarters):
            num_q_cell.value = n_quarters
            formulas_written = True

        if len(paired_rows) >= n_quarters:
            start = paired_rows[-n_quarters]
            end = paired_rows[-1]
            intercept_formula = (
                f'=IFERROR(INTERCEPT(R{start}C{y_col}:R{end}C{y_col},'
                f'R{start}C{x_col}:R{end}C{x_col}),"")'
            )
            slope_formula = (
                f'=IFERROR(SLOPE(R{start}C{y_col}:R{end}C{y_col},'
                f'R{start}C{x_col}:R{end}C{x_col}),"")'
            )
            set_formula2_r1c1(sheet.range((row, cols["intercept"])), intercept_formula)
            set_formula2_r1c1(sheet.range((row, cols["slope"])), slope_formula)
            formulas_written = True

            forecast_cell = sheet.range((row, cols["forecast_value"]))
            if forecast_cell.value in (None, ""):
                forecast_formula = (
                    f'=IFERROR(R{row}C{cols["intercept"]}+'
                    f'R{row}C{cols["slope"]}*R{end}C{x_col},"")'
                )
                set_formula2_r1c1(forecast_cell, forecast_formula)
                formulas_written = True

    if formulas_written:
        wb.app.calculate()

    read_start_col = min(cols.values())
    read_end_col = max(cols.values())
    table_values = as_row_matrix(
        sheet.range(
            (anchor_row + 1, read_start_col),
            (anchor_row + max_rows, read_end_col),
        ).value
    )

    rows: List[Dict[str, Any]] = []
    prev_signature: Optional[Tuple[Any, ...]] = None

    for idx, row_values in enumerate(table_values):
        num_quarters_used = get_col_value(row_values, read_start_col, cols["num_quarters_used"])
        if to_float(num_quarters_used) is None:
            num_quarters_used = idx + 1

        forecast_max = get_col_value(row_values, read_start_col, cols["forecast_max"])
        forecast_min = get_col_value(row_values, read_start_col, cols["forecast_min"])
        forecast_max_n = to_float(forecast_max)
        forecast_min_n = to_float(forecast_min)
        range_width = (
            forecast_max_n - forecast_min_n
            if forecast_max_n is not None and forecast_min_n is not None
            else None
        )

        intercept_value = get_col_value(row_values, read_start_col, cols["intercept"])
        slope_value = get_col_value(row_values, read_start_col, cols["slope"])
        forecast_value = get_col_value(row_values, read_start_col, cols["forecast_value"])

        current_signature = (
            round_for_signature(num_quarters_used),
            round_for_signature(forecast_value),
            round_for_signature(forecast_max),
            round_for_signature(forecast_min),
            round_for_signature(intercept_value),
            round_for_signature(slope_value),
        )

        if prev_signature is not None and current_signature == prev_signature:
            continue
        prev_signature = current_signature

        entry = {
            "model": meta.model,
            "ticker": meta.ticker,
            "model_period": meta.model_period,
            "model_date": meta.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": safe_value(num_quarters_used),
            "num_quarters_used": safe_value(num_quarters_used),
            "forecast_value": safe_value(forecast_value),
            "actual_value": safe_value(
                get_col_value(row_values, read_start_col, cols["actual_value"])
            ),
            "forecast_max": safe_value(forecast_max),
            "forecast_min": safe_value(forecast_min),
            "range_width": safe_value(range_width),
            "intercept": safe_value(intercept_value),
            "slope": safe_value(slope_value),
            "source_file": meta.source_file,
        }
        rows.append(entry)

    return rows


def write_rows_to_sheet(
    ws,
    columns: Sequence[str],
    rows: Iterable[Dict[str, Any]],
) -> None:
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(col, "") for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 42)


def create_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    wb = Workbook()

    ws_empirical = wb.active
    ws_empirical.title = "empirical_candidates"
    write_rows_to_sheet(ws_empirical, EMPIRICAL_COLUMNS, empirical_rows)

    ws_regression = wb.create_sheet("regression_candidates")
    write_rows_to_sheet(ws_regression, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def iter_input_files(in_dir: Path) -> Iterable[Path]:
    for path in sorted(in_dir.iterdir()):
        if path.is_file():
            yield path


def main() -> None:
    in_dir = Path(input_dir).expanduser().resolve()
    out_dir = Path(output_dir).expanduser().resolve()

    if not in_dir.exists() or not in_dir.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a folder: {in_dir}")
    out_dir.mkdir(parents=True, exist_ok=True)

    output_path = build_output_path(in_dir, out_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    original_calc_mode = app.calculation
    app.calculation = "manual"

    try:
        for file_path in iter_input_files(in_dir):
            if file_path.name.startswith("~"):
                print(f"skipped: {file_path.name} (temporary workbook)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped: {file_path.name} (not an .xlsx file)")
                continue

            metadata = parse_file_metadata(file_path)
            if metadata is None:
                print(f"skipped: {file_path.name} (unable to parse ticker/model period)")
                continue

            print(f"processing: {file_path.name}")
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(process_empirical_sheet(wb, metadata))
                regression_rows.extend(process_regression_sheet(wb, metadata))
                processed_files += 1
            except Exception as exc:
                print(f"skipped: {file_path.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        try:
            app.calculation = original_calc_mode
        except Exception:
            pass
        app.quit()

    create_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"output: {output_path}")
    print(f"files_processed: {processed_files}")
    print(f"empirical_rows: {len(empirical_rows)}")
    print(f"regression_rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""Extract empirical and regression candidates from Excel model workbooks."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import re
import sys
from typing import Any, Dict, List, Optional, Sequence, Tuple

try:
    import xlwings as xw
except ImportError as exc:  # pragma: no cover
    raise SystemExit("xlwings is required. Install with: pip install xlwings") from exc

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter
except ImportError as exc:  # pragma: no cover
    raise SystemExit("openpyxl is required. Install with: pip install openpyxl") from exc


# ----------------------------
# User-configurable directories
# ----------------------------
input_dir = Path("input")
output_dir = Path("output")


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

DAY_MAP = {"early": 5, "mid": 15, "late": 25}

FILE_PATTERN = re.compile(
    r"""
    .*?Model\s*-\s*
    (?P<ticker>[A-Za-z0-9._-]+)\s*-\s*
    (?P<window>Early|Mid|Late)
    (?P<month>[A-Za-z]+)
    (?P<year>\d{4})
    _Send
    $
    """,
    re.IGNORECASE | re.VERBOSE,
)


@dataclass(frozen=True)
class ModelMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetSnapshot:
    text_positions: Dict[str, List[Tuple[int, int]]]
    last_col: int


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    if not text:
        return ""
    text = text.replace("%", " pct ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def ensure_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, tuple):
        values = list(values)
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    first = values[0]
    if isinstance(first, tuple):
        return [list(row) if isinstance(row, (list, tuple)) else [row] for row in values]
    if isinstance(first, list):
        return [list(row) for row in values]
    return [list(values)]


def to_float(value: Any) -> Optional[float]:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    cleaned = text.replace(",", "").replace("$", "").replace("%", "")
    try:
        return float(cleaned)
    except ValueError:
        return None


def to_int(value: Any) -> Optional[int]:
    number = to_float(value)
    if number is None:
        return None
    return int(round(number))


def month_to_number(token: str) -> Optional[int]:
    token = token.strip()
    if not token:
        return None
    for fmt in ("%b", "%B"):
        try:
            return datetime.strptime(token.title(), fmt).month
        except ValueError:
            continue
    short = token[:3].title()
    try:
        return datetime.strptime(short, "%b").month
    except ValueError:
        return None


def parse_metadata(file_name: str) -> Optional[ModelMetadata]:
    stem = Path(file_name).stem
    match = FILE_PATTERN.match(stem)
    if not match:
        return None

    ticker = match.group("ticker").strip().upper()
    window = match.group("window").strip().title()
    month_token = match.group("month").strip()
    year = int(match.group("year"))

    month_num = month_to_number(month_token)
    if month_num is None:
        return None

    month_abbr = datetime(year, month_num, 1).strftime("%b")
    model_period = f"{window}{month_abbr}_{year}"
    model_date = datetime(year, month_num, DAY_MAP[window.lower()]).date().isoformat()
    model = f"{ticker}_{model_period}"
    return ModelMetadata(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def next_output_path(input_path: Path, output_path: Path) -> Path:
    base_name = f"{input_path.name}_PARAM.xlsx"
    candidate = output_path / base_name
    if not candidate.exists():
        return candidate

    index = 1
    while True:
        candidate = output_path / f"{input_path.name}_PARAM.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def snapshot_sheet(sheet: xw.Sheet) -> SheetSnapshot:
    used = sheet.used_range
    matrix = ensure_2d(used.value)
    text_positions: Dict[str, List[Tuple[int, int]]] = {}

    for r_idx, row in enumerate(matrix):
        for c_idx, cell_value in enumerate(row):
            normalized = normalize_text(cell_value)
            if normalized:
                text_positions.setdefault(normalized, []).append(
                    (used.row + r_idx, used.column + c_idx)
                )

    last_col = used.column + (len(matrix[0]) - 1 if matrix and matrix[0] else 0)
    return SheetSnapshot(text_positions=text_positions, last_col=last_col)


def choose_max_anchor(snapshot: SheetSnapshot) -> Optional[Tuple[int, int]]:
    max_positions = snapshot.text_positions.get("max", [])
    if not max_positions:
        return None

    min_positions = snapshot.text_positions.get("min", [])
    for max_row, max_col in max_positions:
        for min_row, min_col in min_positions:
            if abs(max_row - min_row) <= 2 and 0 < (min_col - max_col) <= 8:
                return (max_row, max_col)
    return max_positions[0]


def find_offset(
    snapshot: SheetSnapshot,
    anchor_row: int,
    anchor_col: int,
    fragments: Sequence[str],
    default: int,
) -> int:
    fragment_tokens = [normalize_text(fragment) for fragment in fragments if fragment]
    best: Optional[Tuple[int, int]] = None

    for normalized_text, positions in snapshot.text_positions.items():
        if not any(fragment in normalized_text for fragment in fragment_tokens):
            continue
        for row, col in positions:
            if row < anchor_row - 40 or row > anchor_row + 5:
                continue
            offset = col - anchor_col
            distance = abs(row - anchor_row) + abs(offset)
            if best is None or distance < best[1]:
                best = (offset, distance)

    if best is not None:
        return best[0]
    return default


def read_block(
    sheet: xw.Sheet,
    top_row: int,
    bottom_row: int,
    left_col: int,
    right_col: int,
) -> List[List[Any]]:
    if top_row > bottom_row or left_col > right_col:
        return []
    values = sheet.range((top_row, left_col), (bottom_row, right_col)).value
    return ensure_2d(values)


def value_from_block(
    block: Sequence[Sequence[Any]],
    row: int,
    col: int,
    top_row: int,
    left_col: int,
) -> Any:
    row_idx = row - top_row
    col_idx = col - left_col
    if row_idx < 0 or col_idx < 0:
        return None
    if row_idx >= len(block):
        return None
    if col_idx >= len(block[row_idx]):
        return None
    return block[row_idx][col_idx]


def rounded_signature(values: Sequence[Any]) -> Tuple[Any, ...]:
    signature: List[Any] = []
    for value in values:
        if isinstance(value, float):
            signature.append(round(value, 8))
        else:
            signature.append(value)
    return tuple(signature)


def safe_close_workbook(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.api.Close(False)
        return
    except Exception:
        pass

    try:
        workbook.close(False)
        return
    except Exception:
        pass

    workbook.close()


def prepare_empirical_formulas(
    sheet: xw.Sheet,
    anchor_row: int,
    penetration_col: int,
    helper_col: int,
) -> List[Optional[float]]:
    helper_start_row = anchor_row + 1
    for idx in range(1, N_QUARTERS + 1):
        start_row = anchor_row + 1
        end_row = anchor_row + idx
        formula = (
            f'=IFERROR(AVERAGE(R{start_row}C{penetration_col}:R{end_row}C{penetration_col}),"")'
        )
        sheet.range((helper_start_row + idx - 1, helper_col)).formula2 = formula

    sheet.book.app.calculate()
    helper_values = ensure_2d(
        sheet.range(
            (helper_start_row, helper_col),
            (helper_start_row + N_QUARTERS - 1, helper_col),
        ).value
    )

    results = [to_float(row[0] if row else None) for row in helper_values]
    while len(results) < N_QUARTERS:
        results.append(None)
    return results


def extract_empirical_rows(
    workbook: xw.Book,
    metadata: ModelMetadata,
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = workbook.sheets["Empirical Model"]
    except Exception:
        return []

    snapshot = snapshot_sheet(sheet)
    anchor = choose_max_anchor(snapshot)
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    offsets = {
        "num_quarters_used": find_offset(
            snapshot,
            anchor_row,
            anchor_col,
            ["num quarters used", "quarters used", "num quarters", "n quarters"],
            -8,
        ),
        "last_quarter_used": find_offset(
            snapshot,
            anchor_row,
            anchor_col,
            ["last quarter used", "last quarter"],
            -7,
        ),
        "forecast_value": find_offset(
            snapshot,
            anchor_row,
            anchor_col,
            ["estimated total sold", "est total sold", "tot fcst", "forecast"],
            -2,
        ),
        "actual_value": find_offset(
            snapshot,
            anchor_row,
            anchor_col,
            ["reported sales", "actual sales", "actual value"],
            -1,
        ),
        "forecast_max": 0,
        "forecast_min": find_offset(snapshot, anchor_row, anchor_col, ["min"], 1),
        "avg_penetration_pct": find_offset(
            snapshot,
            anchor_row,
            anchor_col,
            ["avg penetration", "average penetration", "penetration pct"],
            -4,
        ),
        "quarterly_sales": find_offset(
            snapshot,
            anchor_row,
            anchor_col,
            ["quarterly sales", "quarter sales"],
            -6,
        ),
        "reported_sales": find_offset(
            snapshot,
            anchor_row,
            anchor_col,
            ["reported sales"],
            -1,
        ),
        "growth_rate_pct": find_offset(
            snapshot,
            anchor_row,
            anchor_col,
            ["growth rate", "growth pct"],
            -3,
        ),
        "sales_captured_in_db_pct": find_offset(
            snapshot,
            anchor_row,
            anchor_col,
            ["sales captured in db", "captured in db", "sales captured"],
            -5,
        ),
    }

    all_cols = [anchor_col + offset for offset in offsets.values()]
    block_top = anchor_row + 1
    block_bottom = anchor_row + N_QUARTERS
    block_left = min(all_cols)
    block_right = max(all_cols)
    block = read_block(sheet, block_top, block_bottom, block_left, block_right)

    helper_col = max(snapshot.last_col + 2, anchor_col + 20)
    avg_formula_values = prepare_empirical_formulas(
        sheet=sheet,
        anchor_row=anchor_row,
        penetration_col=anchor_col + offsets["avg_penetration_pct"],
        helper_col=helper_col,
    )

    rows: List[Dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        row_num = anchor_row + 1 + idx

        num_quarters_used = to_int(
            value_from_block(
                block,
                row=row_num,
                col=anchor_col + offsets["num_quarters_used"],
                top_row=block_top,
                left_col=block_left,
            )
        )
        if num_quarters_used is None:
            num_quarters_used = idx + 1

        last_quarter_used = value_from_block(
            block,
            row=row_num,
            col=anchor_col + offsets["last_quarter_used"],
            top_row=block_top,
            left_col=block_left,
        )

        forecast_value = to_float(
            value_from_block(
                block,
                row=row_num,
                col=anchor_col + offsets["forecast_value"],
                top_row=block_top,
                left_col=block_left,
            )
        )
        reported_sales = to_float(
            value_from_block(
                block,
                row=row_num,
                col=anchor_col + offsets["reported_sales"],
                top_row=block_top,
                left_col=block_left,
            )
        )
        actual_value = to_float(
            value_from_block(
                block,
                row=row_num,
                col=anchor_col + offsets["actual_value"],
                top_row=block_top,
                left_col=block_left,
            )
        )
        if actual_value is None:
            actual_value = reported_sales
        if reported_sales is None:
            reported_sales = actual_value

        forecast_max = to_float(
            value_from_block(
                block,
                row=row_num,
                col=anchor_col + offsets["forecast_max"],
                top_row=block_top,
                left_col=block_left,
            )
        )
        forecast_min = to_float(
            value_from_block(
                block,
                row=row_num,
                col=anchor_col + offsets["forecast_min"],
                top_row=block_top,
                left_col=block_left,
            )
        )
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        avg_penetration_pct = avg_formula_values[idx]
        if avg_penetration_pct is None:
            avg_penetration_pct = to_float(
                value_from_block(
                    block,
                    row=row_num,
                    col=anchor_col + offsets["avg_penetration_pct"],
                    top_row=block_top,
                    left_col=block_left,
                )
            )

        quarterly_sales = to_float(
            value_from_block(
                block,
                row=row_num,
                col=anchor_col + offsets["quarterly_sales"],
                top_row=block_top,
                left_col=block_left,
            )
        )
        growth_rate_pct = to_float(
            value_from_block(
                block,
                row=row_num,
                col=anchor_col + offsets["growth_rate_pct"],
                top_row=block_top,
                left_col=block_left,
            )
        )
        sales_captured_in_db_pct = to_float(
            value_from_block(
                block,
                row=row_num,
                col=anchor_col + offsets["sales_captured_in_db_pct"],
                top_row=block_top,
                left_col=block_left,
            )
        )

        if all(
            value is None or value == ""
            for value in (
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
                quarterly_sales,
            )
        ):
            continue

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
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


def prepare_regression_formulas(
    sheet: xw.Sheet,
    anchor_row: int,
    x_col: int,
    y_col: int,
    helper_col: int,
) -> Tuple[List[Optional[float]], List[Optional[float]]]:
    data_scan_top = anchor_row + 1
    data_scan_bottom = anchor_row + 200
    xy_left = min(x_col, y_col)
    xy_right = max(x_col, y_col)
    xy_block = read_block(
        sheet,
        top_row=data_scan_top,
        bottom_row=data_scan_bottom,
        left_col=xy_left,
        right_col=xy_right,
    )

    valid_rows: List[int] = []
    for idx, row in enumerate(xy_block):
        x_value = to_float(row[x_col - xy_left] if (x_col - xy_left) < len(row) else None)
        y_value = to_float(row[y_col - xy_left] if (y_col - xy_left) < len(row) else None)
        if x_value is not None and y_value is not None:
            valid_rows.append(data_scan_top + idx)

    helper_start_row = anchor_row + 1
    intercept_col = helper_col
    slope_col = helper_col + 1

    for idx in range(1, N_QUARTERS + 1):
        intercept_cell = sheet.range((helper_start_row + idx - 1, intercept_col))
        slope_cell = sheet.range((helper_start_row + idx - 1, slope_col))
        if len(valid_rows) < idx:
            intercept_cell.value = None
            slope_cell.value = None
            continue

        start_row = valid_rows[-idx]
        end_row = valid_rows[-1]
        intercept_cell.formula2 = (
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{end_row}C{x_col}),"")'
        )
        slope_cell.formula2 = (
            f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{end_row}C{x_col}),"")'
        )

    sheet.book.app.calculate()
    intercept_matrix = ensure_2d(
        sheet.range(
            (helper_start_row, intercept_col),
            (helper_start_row + N_QUARTERS - 1, intercept_col),
        ).value
    )
    slope_matrix = ensure_2d(
        sheet.range(
            (helper_start_row, slope_col),
            (helper_start_row + N_QUARTERS - 1, slope_col),
        ).value
    )

    intercept_values = [to_float(row[0] if row else None) for row in intercept_matrix]
    slope_values = [to_float(row[0] if row else None) for row in slope_matrix]

    while len(intercept_values) < N_QUARTERS:
        intercept_values.append(None)
    while len(slope_values) < N_QUARTERS:
        slope_values.append(None)
    return intercept_values, slope_values


def extract_regression_rows(
    workbook: xw.Book,
    metadata: ModelMetadata,
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = workbook.sheets["Regression Model"]
    except Exception:
        return []

    snapshot = snapshot_sheet(sheet)
    anchor = choose_max_anchor(snapshot)
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    offsets = {
        "num_quarters_used": find_offset(
            snapshot,
            anchor_row,
            anchor_col,
            ["num quarters used", "quarters used", "num quarters", "n quarters"],
            -8,
        ),
        "forecast_value": find_offset(
            snapshot,
            anchor_row,
            anchor_col,
            ["tot fcst w o sa", "tot fcst without sa", "forecast", "fcst"],
            -2,
        ),
        "actual_value": find_offset(
            snapshot,
            anchor_row,
            anchor_col,
            ["actual sales", "reported sales", "actual value"],
            -1,
        ),
        "forecast_max": 0,
        "forecast_min": find_offset(snapshot, anchor_row, anchor_col, ["min"], 1),
    }

    all_cols = [anchor_col + offset for offset in offsets.values()]
    block_top = anchor_row + 1
    block_bottom = anchor_row + N_QUARTERS
    block_left = min(all_cols)
    block_right = max(all_cols)
    block = read_block(sheet, block_top, block_bottom, block_left, block_right)

    helper_col = max(snapshot.last_col + 2, anchor_col + 20)
    intercept_values, slope_values = prepare_regression_formulas(
        sheet=sheet,
        anchor_row=anchor_row,
        x_col=x_col,
        y_col=y_col,
        helper_col=helper_col,
    )

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None
    for idx in range(N_QUARTERS):
        row_num = anchor_row + 1 + idx

        num_quarters_used = to_int(
            value_from_block(
                block,
                row=row_num,
                col=anchor_col + offsets["num_quarters_used"],
                top_row=block_top,
                left_col=block_left,
            )
        )
        if num_quarters_used is None:
            num_quarters_used = idx + 1

        forecast_value = to_float(
            value_from_block(
                block,
                row=row_num,
                col=anchor_col + offsets["forecast_value"],
                top_row=block_top,
                left_col=block_left,
            )
        )
        actual_value = to_float(
            value_from_block(
                block,
                row=row_num,
                col=anchor_col + offsets["actual_value"],
                top_row=block_top,
                left_col=block_left,
            )
        )
        forecast_max = to_float(
            value_from_block(
                block,
                row=row_num,
                col=anchor_col + offsets["forecast_max"],
                top_row=block_top,
                left_col=block_left,
            )
        )
        forecast_min = to_float(
            value_from_block(
                block,
                row=row_num,
                col=anchor_col + offsets["forecast_min"],
                top_row=block_top,
                left_col=block_left,
            )
        )
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        intercept = intercept_values[idx]
        slope = slope_values[idx]

        if all(
            value is None
            for value in (forecast_value, forecast_max, forecast_min, intercept, slope)
        ):
            continue

        signature = rounded_signature(
            (
                num_quarters_used,
                forecast_value,
                forecast_max,
                forecast_min,
                intercept,
                slope,
            )
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
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


def write_sheet(worksheet: Any, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    worksheet.append(list(columns))
    for col_idx in range(1, len(columns) + 1):
        worksheet.cell(row=1, column=col_idx).font = Font(bold=True)

    for row in rows:
        worksheet.append([row.get(column, "") for column in columns])

    worksheet.freeze_panes = "A2"
    if worksheet.max_row >= 1 and worksheet.max_column >= 1:
        worksheet.auto_filter.ref = worksheet.dimensions

    for col_idx, header in enumerate(columns, start=1):
        max_len = len(header)
        for row_idx in range(2, worksheet.max_row + 1):
            value = worksheet.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        worksheet.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    empirical_sheet = workbook.active
    empirical_sheet.title = "empirical_candidates"
    write_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)

    regression_sheet = workbook.create_sheet("regression_candidates")
    write_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)

    workbook.save(output_path)


def should_skip_file(file_path: Path) -> Optional[str]:
    if not file_path.is_file():
        return "not a file"
    if file_path.name.startswith("~"):
        return "temporary workbook"
    if file_path.suffix.lower() != ".xlsx":
        return "not an .xlsx file"
    return None


def process_files(input_path: Path) -> Tuple[int, List[Dict[str, Any]], List[Dict[str, Any]]]:
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_count = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in sorted(input_path.iterdir()):
            skip_reason = should_skip_file(file_path)
            if skip_reason:
                print(f"skipped: {file_path.name} ({skip_reason})")
                continue

            metadata = parse_metadata(file_path.name)
            if metadata is None:
                print(f"skipped: {file_path.name} (unable to parse model metadata)")
                continue

            workbook: Optional[xw.Book] = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(
                    extract_empirical_rows(
                        workbook=workbook,
                        metadata=metadata,
                        source_file=file_path.name,
                    )
                )
                regression_rows.extend(
                    extract_regression_rows(
                        workbook=workbook,
                        metadata=metadata,
                        source_file=file_path.name,
                    )
                )
                processed_count += 1
                print(f"processed: {file_path.name}")
            except Exception as exc:
                print(f"skipped: {file_path.name} (processing error: {exc})")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        try:
            app.quit()
        except Exception:
            app.kill()

    return processed_count, empirical_rows, regression_rows


def main() -> int:
    input_path = Path(input_dir)
    output_path = Path(output_dir)

    if not input_path.exists():
        print(f"input directory does not exist: {input_path}")
        return 1

    output_path.mkdir(parents=True, exist_ok=True)
    final_output = next_output_path(input_path, output_path)

    processed_count, empirical_rows, regression_rows = process_files(input_path)
    write_output_workbook(final_output, empirical_rows, regression_rows)

    print(f"output: {final_output}")
    print(f"files_processed: {processed_count}")
    print(f"empirical_rows: {len(empirical_rows)}")
    print(f"regression_rows: {len(regression_rows)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

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

PERIOD_DAY_MAP = {"early": 5, "mid": 15, "late": 25}

# Header fallback offsets from the detected "max" anchor.
EMPIRICAL_DEFAULT_OFFSETS = {
    "num_quarters_used": -10,
    "last_quarter_used": -9,
    "forecast_value": -2,
    "actual_value": -1,
    "forecast_max": 0,
    "forecast_min": 1,
    "quarterly_sales": -7,
    "reported_sales": -6,
    "growth_rate_pct": -5,
    "sales_captured_in_db_pct": -4,
}

REGRESSION_DEFAULT_OFFSETS = {
    "num_quarters_used": -10,
    "forecast_value": -2,
    "actual_value": -1,
    "forecast_max": 0,
    "forecast_min": 1,
}


@dataclass
class SheetSnapshot:
    sheet: xw.main.Sheet
    start_row: int
    start_col: int
    matrix: List[List[Any]]

    @property
    def end_row(self) -> int:
        return self.start_row + len(self.matrix) - 1

    @property
    def end_col(self) -> int:
        if not self.matrix:
            return self.start_col - 1
        return self.start_col + len(self.matrix[0]) - 1

    def get_value(self, row: int, col: int) -> Any:
        if row < self.start_row or col < self.start_col:
            return None
        row_idx = row - self.start_row
        col_idx = col - self.start_col
        if row_idx < 0 or row_idx >= len(self.matrix):
            return None
        if col_idx < 0 or col_idx >= len(self.matrix[row_idx]):
            return None
        return self.matrix[row_idx][col_idx]


def to_matrix(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        width = max(len(row) if isinstance(row, list) else 1 for row in values)
        matrix: List[List[Any]] = []
        for row in values:
            if isinstance(row, list):
                matrix.append(row + [None] * (width - len(row)))
            else:
                matrix.append([row] + [None] * (width - 1))
        return matrix
    return [values]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    return re.sub(r"\s+", " ", text)


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip() == ""
    return False


def as_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    if text.endswith("%"):
        number = text[:-1].strip()
        try:
            return float(number) / 100.0
        except ValueError:
            return None
    try:
        return float(text)
    except ValueError:
        return None


def subtract_values(left: Any, right: Any) -> Optional[float]:
    left_float = as_float(left)
    right_float = as_float(right)
    if left_float is None or right_float is None:
        return None
    return left_float - right_float


def safe_int(value: Any) -> Optional[int]:
    number = as_float(value)
    if number is None:
        return None
    return int(round(number))


def parse_metadata(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = parts[1] if len(parts) >= 2 else stem
    ticker = re.sub(r"\s+", "", ticker).upper()

    period_raw = parts[2] if len(parts) >= 3 else ""
    period_raw = re.sub(r"[_\-\s]*send.*$", "", period_raw, flags=re.IGNORECASE)
    period_raw = re.sub(r"[^A-Za-z0-9]", "", period_raw)

    period_match = re.search(
        r"(Early|Mid|Late)([A-Za-z]{3})(\d{4})", period_raw, flags=re.IGNORECASE
    )

    model_period = ""
    model_date = ""
    if period_match:
        period_part = period_match.group(1).title()
        month_part = period_match.group(2).title()
        year_part = int(period_match.group(3))

        month_num = MONTH_MAP.get(month_part.lower())
        day_num = PERIOD_DAY_MAP.get(period_part.lower())

        model_period = f"{period_part}{month_part}_{year_part}"
        if month_num and day_num:
            model_date = date(year_part, month_num, day_num).isoformat()

    model = f"{ticker}_{model_period}" if ticker and model_period else ticker
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def unique_output_path(in_dir: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{in_dir.resolve().name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    index = 1
    while True:
        candidate = out_dir / f"{base_name}.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def build_snapshot(sheet: xw.main.Sheet) -> SheetSnapshot:
    used_range = sheet.used_range
    matrix = to_matrix(used_range.value)
    return SheetSnapshot(
        sheet=sheet,
        start_row=used_range.row,
        start_col=used_range.column,
        matrix=matrix,
    )


def find_max_anchor(snapshot: SheetSnapshot) -> Optional[Tuple[int, int]]:
    candidates: List[Tuple[float, int, int]] = []
    for row in range(snapshot.start_row, snapshot.end_row + 1):
        for col in range(snapshot.start_col, snapshot.end_col + 1):
            if normalize_text(snapshot.get_value(row, col)) != "max":
                continue

            score = 0.0
            if normalize_text(snapshot.get_value(row, col + 1)) == "min":
                score += 5.0
            left_text = normalize_text(snapshot.get_value(row, col - 1))
            if "fcst" in left_text or "forecast" in left_text:
                score += 2.0
            if "tot" in left_text:
                score += 1.0
            # Favor upper anchors when scores tie.
            score -= (row - snapshot.start_row) * 0.01
            candidates.append((score, row, col))

    if not candidates:
        return None

    candidates.sort(key=lambda item: item[0], reverse=True)
    _, anchor_row, anchor_col = candidates[0]
    return anchor_row, anchor_col


def _header_match(text: str, token_groups: Sequence[Sequence[str]]) -> bool:
    for tokens in token_groups:
        if all(token in text for token in tokens):
            return True
    return False


def find_column(
    snapshot: SheetSnapshot,
    anchor_row: int,
    anchor_col: int,
    token_groups: Sequence[Sequence[str]],
    default_offset: int,
    *,
    side: Optional[str] = None,
    row_offsets: Sequence[int] = (0, -1, 1),
    max_distance: int = 30,
) -> int:
    best_match: Optional[Tuple[float, int]] = None
    min_col = max(snapshot.start_col, anchor_col - max_distance)
    max_col = min(snapshot.end_col, anchor_col + max_distance)

    for row_offset in row_offsets:
        row = anchor_row + row_offset
        if row < snapshot.start_row or row > snapshot.end_row:
            continue

        for col in range(min_col, max_col + 1):
            if side == "left" and col > anchor_col:
                continue
            if side == "right" and col < anchor_col:
                continue

            header = normalize_text(snapshot.get_value(row, col))
            if not header:
                continue
            if not _header_match(header, token_groups):
                continue

            score = -abs(col - anchor_col) - abs(row_offset) * 0.25
            if best_match is None or score > best_match[0]:
                best_match = (score, col)

    if best_match is not None:
        return best_match[1]
    return anchor_col + default_offset


def set_formula2(cell: xw.main.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def close_workbook_safe(workbook: xw.main.Book) -> None:
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

    workbook.close()


def map_empirical_columns(
    snapshot: SheetSnapshot, anchor_row: int, anchor_col: int
) -> Dict[str, int]:
    return {
        "num_quarters_used": find_column(
            snapshot,
            anchor_row,
            anchor_col,
            [
                ("num", "quarter"),
                ("quarters", "used"),
                ("num", "qtr"),
                ("n", "qtr"),
            ],
            EMPIRICAL_DEFAULT_OFFSETS["num_quarters_used"],
            side="left",
        ),
        "last_quarter_used": find_column(
            snapshot,
            anchor_row,
            anchor_col,
            [("last", "quarter"), ("last", "qtr"), ("quarter", "used")],
            EMPIRICAL_DEFAULT_OFFSETS["last_quarter_used"],
            side="left",
        ),
        "forecast_value": find_column(
            snapshot,
            anchor_row,
            anchor_col,
            [
                ("estimated", "total", "sold"),
                ("est", "total", "sold"),
                ("forecast",),
                ("tot", "fcst"),
            ],
            EMPIRICAL_DEFAULT_OFFSETS["forecast_value"],
            side="left",
        ),
        "actual_value": find_column(
            snapshot,
            anchor_row,
            anchor_col,
            [("reported", "sales"), ("actual", "sales"), ("actual",)],
            EMPIRICAL_DEFAULT_OFFSETS["actual_value"],
            side="left",
        ),
        "forecast_max": anchor_col,
        "forecast_min": find_column(
            snapshot,
            anchor_row,
            anchor_col,
            [("min",)],
            EMPIRICAL_DEFAULT_OFFSETS["forecast_min"],
            side="right",
        ),
        "quarterly_sales": find_column(
            snapshot,
            anchor_row,
            anchor_col,
            [("quarterly", "sales"), ("qtr", "sales"), ("sales", "quarterly")],
            EMPIRICAL_DEFAULT_OFFSETS["quarterly_sales"],
            side="left",
        ),
        "reported_sales": find_column(
            snapshot,
            anchor_row,
            anchor_col,
            [("reported", "sales"), ("reported",)],
            EMPIRICAL_DEFAULT_OFFSETS["reported_sales"],
            side="left",
        ),
        "growth_rate_pct": find_column(
            snapshot,
            anchor_row,
            anchor_col,
            [("growth", "rate"), ("growth",)],
            EMPIRICAL_DEFAULT_OFFSETS["growth_rate_pct"],
            side="left",
        ),
        "sales_captured_in_db_pct": find_column(
            snapshot,
            anchor_row,
            anchor_col,
            [
                ("captured", "db"),
                ("captured", "in", "db"),
                ("sales", "captured"),
                ("penetration",),
            ],
            EMPIRICAL_DEFAULT_OFFSETS["sales_captured_in_db_pct"],
            side="left",
        ),
    }


def map_regression_columns(
    snapshot: SheetSnapshot, anchor_row: int, anchor_col: int
) -> Dict[str, int]:
    return {
        "num_quarters_used": find_column(
            snapshot,
            anchor_row,
            anchor_col,
            [
                ("num", "quarter"),
                ("quarters", "used"),
                ("num", "qtr"),
                ("n", "qtr"),
            ],
            REGRESSION_DEFAULT_OFFSETS["num_quarters_used"],
            side="left",
        ),
        "forecast_value": find_column(
            snapshot,
            anchor_row,
            anchor_col,
            [
                ("tot", "fcst", "w/o", "sa"),
                ("total", "forecast", "without", "sa"),
                ("tot", "fcst"),
                ("forecast",),
            ],
            REGRESSION_DEFAULT_OFFSETS["forecast_value"],
            side="left",
        ),
        "actual_value": find_column(
            snapshot,
            anchor_row,
            anchor_col,
            [("actual", "sales"), ("reported", "sales"), ("actual",)],
            REGRESSION_DEFAULT_OFFSETS["actual_value"],
            side="left",
        ),
        "forecast_max": anchor_col,
        "forecast_min": find_column(
            snapshot,
            anchor_row,
            anchor_col,
            [("min",)],
            REGRESSION_DEFAULT_OFFSETS["forecast_min"],
            side="right",
        ),
    }


def candidate_rows(
    snapshot: SheetSnapshot, anchor_row: int, anchor_col: int, cols: Dict[str, int]
) -> List[int]:
    rows: List[int] = []
    blank_streak = 0

    for index in range(1, N_QUARTERS + 1):
        row = anchor_row + index
        values = [
            snapshot.get_value(row, cols.get("num_quarters_used", anchor_col)),
            snapshot.get_value(row, cols.get("forecast_value", anchor_col)),
            snapshot.get_value(row, cols.get("forecast_max", anchor_col)),
            snapshot.get_value(row, cols.get("forecast_min", anchor_col)),
        ]
        if all(is_blank(value) for value in values):
            blank_streak += 1
            if blank_streak >= 2:
                break
            continue

        blank_streak = 0
        rows.append(row)

    return rows


def find_numeric_series(
    snapshot: SheetSnapshot, col: int, anchor_row: int
) -> Optional[Tuple[int, int]]:
    end_row = anchor_row - 1
    while end_row >= snapshot.start_row:
        if as_float(snapshot.get_value(end_row, col)) is not None:
            break
        end_row -= 1

    if end_row < snapshot.start_row:
        return None

    start_row = end_row
    while start_row >= snapshot.start_row:
        if as_float(snapshot.get_value(start_row, col)) is None:
            start_row += 1
            break
        start_row -= 1
    else:
        start_row = snapshot.start_row

    return start_row, end_row


def find_numeric_block(
    snapshot: SheetSnapshot, x_col: int, y_col: int, anchor_row: int
) -> Optional[Tuple[int, int]]:
    end_row = anchor_row - 1
    while end_row >= snapshot.start_row:
        x_val = as_float(snapshot.get_value(end_row, x_col))
        y_val = as_float(snapshot.get_value(end_row, y_col))
        if x_val is not None and y_val is not None:
            break
        end_row -= 1

    if end_row < snapshot.start_row:
        return None

    start_row = end_row
    while start_row >= snapshot.start_row:
        x_val = as_float(snapshot.get_value(start_row, x_col))
        y_val = as_float(snapshot.get_value(start_row, y_col))
        if x_val is None or y_val is None:
            start_row += 1
            break
        start_row -= 1
    else:
        start_row = snapshot.start_row

    return start_row, end_row


def values_equal(left: Any, right: Any, tolerance: float = 1e-9) -> bool:
    left_float = as_float(left)
    right_float = as_float(right)
    if left_float is not None and right_float is not None:
        return abs(left_float - right_float) <= tolerance
    return (left in (None, "")) and (right in (None, "")) or str(left) == str(right)


def is_duplicate_regression_row(
    previous_row: Dict[str, Any], current_row: Dict[str, Any]
) -> bool:
    keys = (
        "num_quarters_used",
        "forecast_value",
        "forecast_max",
        "forecast_min",
        "intercept",
        "slope",
    )
    return all(values_equal(previous_row.get(key), current_row.get(key)) for key in keys)


def extract_empirical_candidates(
    workbook: xw.main.Book, metadata: Dict[str, str], file_name: str
) -> List[Dict[str, Any]]:
    output_rows: List[Dict[str, Any]] = []

    try:
        sheet = workbook.sheets["Empirical Model"]
    except Exception:
        print(f"Skipping empirical extraction for {file_name}: sheet not found")
        return output_rows

    snapshot = build_snapshot(sheet)
    anchor = find_max_anchor(snapshot)
    if anchor is None:
        print(f"Skipping empirical extraction for {file_name}: no 'max' anchor found")
        return output_rows

    anchor_row, anchor_col = anchor
    cols = map_empirical_columns(snapshot, anchor_row, anchor_col)
    rows = candidate_rows(snapshot, anchor_row, anchor_col, cols)
    if not rows:
        return output_rows

    temp_col = max(snapshot.end_col, anchor_col) + 3
    quarterly_col = cols["quarterly_sales"]
    reported_col = cols["reported_sales"]
    captured_col = cols["sales_captured_in_db_pct"]
    captured_series = find_numeric_series(snapshot, captured_col, anchor_row)

    for fallback_index, row in enumerate(rows, start=1):
        num_quarters = safe_int(snapshot.get_value(row, cols["num_quarters_used"]))
        if num_quarters is None:
            num_quarters = fallback_index

        formula: str
        if captured_series is not None and num_quarters > 0:
            series_start, series_end = captured_series
            avg_start = max(series_start, series_end - num_quarters + 1)
            formula = f'=IFERROR(AVERAGE(R{avg_start}C{captured_col}:R{series_end}C{captured_col}),"")'
        else:
            formula = f'=IFERROR(R{row}C{quarterly_col}/R{row}C{reported_col},"")'

        set_formula2(sheet.range((row, temp_col)), formula)

    workbook.app.calculate()

    for fallback_index, row in enumerate(rows, start=1):
        num_quarters = snapshot.get_value(row, cols["num_quarters_used"])
        if is_blank(num_quarters):
            num_quarters = fallback_index

        forecast_max = snapshot.get_value(row, cols["forecast_max"])
        forecast_min = snapshot.get_value(row, cols["forecast_min"])
        forecast_value = snapshot.get_value(row, cols["forecast_value"])
        actual_value = snapshot.get_value(row, cols["actual_value"])
        quarterly_sales = snapshot.get_value(row, cols["quarterly_sales"])
        reported_sales = snapshot.get_value(row, cols["reported_sales"])
        growth_rate_pct = snapshot.get_value(row, cols["growth_rate_pct"])
        captured_pct = snapshot.get_value(row, cols["sales_captured_in_db_pct"])
        avg_penetration = sheet.range((row, temp_col)).value

        if all(
            is_blank(value)
            for value in (
                forecast_value,
                forecast_max,
                forecast_min,
                quarterly_sales,
                reported_sales,
            )
        ):
            continue

        range_width = subtract_values(forecast_max, forecast_min)
        output_rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters,
                "last_quarter_used": snapshot.get_value(
                    row, cols["last_quarter_used"]
                ),
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": captured_pct,
                "source_file": file_name,
            }
        )

    sheet.range((rows[0], temp_col), (rows[-1], temp_col)).clear_contents()
    return output_rows


def extract_regression_candidates(
    workbook: xw.main.Book, metadata: Dict[str, str], file_name: str
) -> List[Dict[str, Any]]:
    output_rows: List[Dict[str, Any]] = []

    try:
        sheet = workbook.sheets["Regression Model"]
    except Exception:
        print(f"Skipping regression extraction for {file_name}: sheet not found")
        return output_rows

    snapshot = build_snapshot(sheet)
    anchor = find_max_anchor(snapshot)
    if anchor is None:
        print(f"Skipping regression extraction for {file_name}: no 'max' anchor found")
        return output_rows

    anchor_row, anchor_col = anchor
    cols = map_regression_columns(snapshot, anchor_row, anchor_col)
    rows = candidate_rows(snapshot, anchor_row, anchor_col, cols)
    if not rows:
        return output_rows

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    numeric_block = find_numeric_block(snapshot, x_col, y_col, anchor_row)

    intercept_col = max(snapshot.end_col, anchor_col) + 3
    slope_col = intercept_col + 1
    formula_rows: List[int] = []

    if numeric_block is not None:
        block_start, block_end = numeric_block
        for fallback_index, row in enumerate(rows, start=1):
            num_quarters = safe_int(snapshot.get_value(row, cols["num_quarters_used"]))
            if num_quarters is None:
                num_quarters = fallback_index
            if num_quarters < 2:
                continue

            start_row = max(block_start, block_end - num_quarters + 1)
            if block_end - start_row + 1 < 2:
                continue

            y_range = f"R{start_row}C{y_col}:R{block_end}C{y_col}"
            x_range = f"R{start_row}C{x_col}:R{block_end}C{x_col}"
            intercept_formula = f'=IFERROR(INTERCEPT({y_range},{x_range}),"")'
            slope_formula = f'=IFERROR(SLOPE({y_range},{x_range}),"")'
            set_formula2(sheet.range((row, intercept_col)), intercept_formula)
            set_formula2(sheet.range((row, slope_col)), slope_formula)
            formula_rows.append(row)

        if formula_rows:
            workbook.app.calculate()

    for fallback_index, row in enumerate(rows, start=1):
        num_quarters = snapshot.get_value(row, cols["num_quarters_used"])
        if is_blank(num_quarters):
            num_quarters = fallback_index

        forecast_value = snapshot.get_value(row, cols["forecast_value"])
        actual_value = snapshot.get_value(row, cols["actual_value"])
        forecast_max = snapshot.get_value(row, cols["forecast_max"])
        forecast_min = snapshot.get_value(row, cols["forecast_min"])

        intercept = sheet.range((row, intercept_col)).value if formula_rows else None
        slope = sheet.range((row, slope_col)).value if formula_rows else None

        if all(
            is_blank(value)
            for value in (forecast_value, forecast_max, forecast_min, intercept, slope)
        ):
            continue

        candidate = {
            "model": metadata["model"],
            "ticker": metadata["ticker"],
            "model_period": metadata["model_period"],
            "model_date": metadata["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters,
            "num_quarters_used": num_quarters,
            "forecast_value": forecast_value,
            "actual_value": actual_value if not is_blank(actual_value) else "",
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": subtract_values(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": file_name,
        }

        if output_rows and is_duplicate_regression_row(output_rows[-1], candidate):
            continue
        output_rows.append(candidate)

    if formula_rows:
        sheet.range((formula_rows[0], intercept_col), (formula_rows[-1], slope_col)).clear_contents()
    return output_rows


def normalize_output_value(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    return value


def write_output_sheet(
    workbook: Workbook, sheet_name: str, columns: Sequence[str], rows: Sequence[Dict[str, Any]]
) -> None:
    worksheet = workbook.create_sheet(title=sheet_name)
    worksheet.append(list(columns))

    for row in rows:
        worksheet.append([normalize_output_value(row.get(column)) for column in columns])

    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    worksheet.freeze_panes = "A2"
    if worksheet.max_row >= 1 and worksheet.max_column >= 1:
        worksheet.auto_filter.ref = worksheet.dimensions

    for col_idx, column in enumerate(columns, start=1):
        max_length = len(column)
        for row_idx in range(2, worksheet.max_row + 1):
            value = worksheet.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_length = max(max_length, len(str(value)))
        worksheet.column_dimensions[get_column_letter(col_idx)].width = min(max_length + 2, 44)


def iter_source_files(source_dir: Path) -> Iterable[Path]:
    for file_path in sorted(source_dir.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"Skipping {file_path.name}: temporary workbook")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipping {file_path.name}: not an .xlsx file")
            continue
        yield file_path


def main() -> None:
    source_dir = Path(input_dir)
    target_dir = Path(output_dir)

    if not source_dir.exists():
        raise FileNotFoundError(f"input_dir does not exist: {source_dir}")
    if not source_dir.is_dir():
        raise NotADirectoryError(f"input_dir is not a directory: {source_dir}")

    output_path = unique_output_path(source_dir, target_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app: Optional[xw.main.App] = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in iter_source_files(source_dir):
            print(f"Processing {file_path.name}")
            workbook: Optional[xw.main.Book] = None
            metadata = parse_metadata(file_path.name)

            try:
                workbook = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(
                    extract_empirical_candidates(workbook, metadata, file_path.name)
                )
                regression_rows.extend(
                    extract_regression_candidates(workbook, metadata, file_path.name)
                )
                processed_files += 1
            except Exception as exc:
                print(f"Skipping {file_path.name}: {exc}")
            finally:
                if workbook is not None:
                    close_workbook_safe(workbook)
    finally:
        if app is not None:
            app.quit()

    output_wb = Workbook()
    default_sheet = output_wb.active
    output_wb.remove(default_sheet)

    write_output_sheet(output_wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_output_sheet(output_wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    output_wb.save(output_path)

    print(f"Output path: {output_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# Configure these two paths before running.
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


@dataclass
class SheetSnapshot:
    start_row: int
    start_col: int
    end_row: int
    end_col: int
    values: List[List[Any]]

    @property
    def n_rows(self) -> int:
        return len(self.values)

    @property
    def n_cols(self) -> int:
        if not self.values:
            return 0
        return len(self.values[0])

    def get_value(self, row: int, col: int) -> Any:
        row_idx = row - self.start_row
        col_idx = col - self.start_col
        if row_idx < 0 or col_idx < 0:
            return None
        if row_idx >= self.n_rows or col_idx >= self.n_cols:
            return None
        return self.values[row_idx][col_idx]

    def row_values(self, row: int) -> List[Any]:
        row_idx = row - self.start_row
        if row_idx < 0 or row_idx >= self.n_rows:
            return []
        return self.values[row_idx]


def normalize_matrix(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, tuple):
        values = list(values)
    if not isinstance(values, list):
        return [[values]]

    if not values:
        return []

    if isinstance(values[0], tuple):
        values = [list(row) for row in values]
    elif not isinstance(values[0], list):
        values = [list(values)]

    max_len = max(len(row) for row in values) if values else 0
    normalized: List[List[Any]] = []
    for row in values:
        if isinstance(row, tuple):
            row = list(row)
        if len(row) < max_len:
            row = row + [None] * (max_len - len(row))
        normalized.append(row)
    return normalized


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = text.replace("_", " ").replace("-", " ")
    text = re.sub(r"\s+", " ", text)
    return text


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and not value.strip():
        return True
    return False


def to_float(value: Any) -> Optional[float]:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return None

    pct = text.endswith("%")
    if pct:
        text = text[:-1].strip()

    text = text.replace(",", "")
    if text.startswith("(") and text.endswith(")"):
        text = "-" + text[1:-1]

    try:
        number = float(text)
    except ValueError:
        return None

    if pct:
        number = number / 100.0
    return number


def to_int(value: Any) -> Optional[int]:
    number = to_float(value)
    if number is None:
        return None
    return int(round(number))


def create_snapshot(sheet: xw.Sheet) -> SheetSnapshot:
    used = sheet.used_range
    values = normalize_matrix(used.value)
    start_row = int(used.row)
    start_col = int(used.column)

    if not values:
        return SheetSnapshot(
            start_row=start_row,
            start_col=start_col,
            end_row=start_row,
            end_col=start_col,
            values=[],
        )

    end_row = start_row + len(values) - 1
    end_col = start_col + len(values[0]) - 1
    return SheetSnapshot(
        start_row=start_row,
        start_col=start_col,
        end_row=end_row,
        end_col=end_col,
        values=values,
    )


def find_anchor(snapshot: SheetSnapshot, anchor_text: str = "max") -> Optional[Tuple[int, int]]:
    target = normalize_label(anchor_text)
    candidates: List[Tuple[int, int, int]] = []

    for row_idx, row_values in enumerate(snapshot.values):
        abs_row = snapshot.start_row + row_idx
        for col_idx, value in enumerate(row_values):
            abs_col = snapshot.start_col + col_idx
            if normalize_label(value) == target:
                score = 0
                for delta in (-2, -1, 1, 2):
                    if normalize_label(snapshot.get_value(abs_row, abs_col + delta)) == "min":
                        score += 1
                candidates.append((score, abs_row, abs_col))

    if not candidates:
        return None

    candidates.sort(key=lambda item: (-item[0], item[1], item[2]))
    _, row, col = candidates[0]
    return row, col


def find_column_from_header(
    snapshot: SheetSnapshot,
    anchor_row: int,
    anchor_col: int,
    keywords: Sequence[str],
    fallback_offset: int,
    window: int = 24,
) -> int:
    key_list = [normalize_label(word) for word in keywords]
    candidate_rows = [anchor_row, anchor_row - 1, anchor_row + 1]

    for row in candidate_rows:
        if row < snapshot.start_row or row > snapshot.end_row:
            continue
        col_start = max(snapshot.start_col, anchor_col - window)
        col_end = min(snapshot.end_col, anchor_col + window)
        for col in range(col_start, col_end + 1):
            header = normalize_label(snapshot.get_value(row, col))
            if not header:
                continue
            if any(key in header for key in key_list):
                return col

    fallback_col = anchor_col + fallback_offset
    if fallback_col < snapshot.start_col:
        fallback_col = snapshot.start_col
    if fallback_col > snapshot.end_col:
        fallback_col = snapshot.end_col
    return fallback_col


def find_column_anywhere(
    snapshot: SheetSnapshot,
    keywords: Sequence[str],
    anchor_col: int,
    max_row: int,
) -> Optional[int]:
    key_list = [normalize_label(word) for word in keywords]
    best_match: Optional[Tuple[int, int]] = None

    row_end = min(snapshot.end_row, max_row)
    for row in range(snapshot.start_row, row_end + 1):
        for col in range(snapshot.start_col, snapshot.end_col + 1):
            header = normalize_label(snapshot.get_value(row, col))
            if not header:
                continue
            if any(key in header for key in key_list):
                distance = abs(anchor_col - col)
                if best_match is None or distance < best_match[0]:
                    best_match = (distance, col)

    return best_match[1] if best_match else None


def gather_numeric_series(
    snapshot: SheetSnapshot,
    col: int,
    max_row: int,
) -> List[Tuple[int, float]]:
    values: List[Tuple[int, float]] = []
    for row in range(snapshot.start_row, min(snapshot.end_row, max_row) + 1):
        number = to_float(snapshot.get_value(row, col))
        if number is not None:
            values.append((row, number))
    return values


def pick_scratch_cells(
    sheet: xw.Sheet,
    snapshot: SheetSnapshot,
    count: int = 1,
) -> List[xw.Range]:
    scratch_col = min(16384, snapshot.end_col + 5)
    scratch_row = max(2, snapshot.start_row)
    return [sheet.range((scratch_row + idx, scratch_col)) for idx in range(count)]


def parse_file_metadata(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = "UNKNOWN"
    if len(parts) >= 2 and parts[1]:
        ticker = parts[1].split()[0].upper()
    else:
        fallback_ticker = re.search(r"\b[A-Z]{2,6}\b", stem)
        if fallback_ticker:
            ticker = fallback_ticker.group(0)

    period_match = re.search(r"(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})", stem, re.IGNORECASE)
    model_period = "UnknownPeriod"
    model_date = ""

    if period_match:
        phase = period_match.group(1).title()
        month_token = period_match.group(2)
        year = period_match.group(3)
        month_abbrev = month_token[:3].title()
        model_period = f"{phase}{month_abbrev}_{year}"

        month_number = MONTH_MAP.get(month_token[:3].lower())
        day_map = {"Early": 5, "Mid": 15, "Late": 25}
        if month_number is not None:
            model_date = date(int(year), month_number, day_map[phase]).isoformat()

    model = f"{ticker}_{model_period}" if model_period else ticker
    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def safe_close_workbook(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        workbook.saved = True
        workbook.close()
    except Exception:
        pass


def almost_equal(left: Optional[float], right: Optional[float], tolerance: float = 1e-10) -> bool:
    if left is None and right is None:
        return True
    if left is None or right is None:
        return False
    return abs(left - right) <= tolerance


def row_has_any_value(snapshot: SheetSnapshot, row: int, cols: Sequence[int]) -> bool:
    for col in cols:
        value = snapshot.get_value(row, col)
        if not is_blank(value):
            return True
    return False


def extract_empirical_candidates(
    workbook: xw.Book,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = workbook.sheets["Empirical Model"]
    except Exception:
        print(f"Skipped empirical extraction for {source_file}: missing 'Empirical Model' sheet")
        return []

    snapshot = create_snapshot(sheet)
    anchor = find_anchor(snapshot, "max")
    if anchor is None:
        print(f"Skipped empirical extraction for {source_file}: could not find 'max' anchor")
        return []

    anchor_row, anchor_col = anchor
    min_col = find_column_from_header(snapshot, anchor_row, anchor_col, ["min"], fallback_offset=1, window=6)
    if min_col == anchor_col:
        min_col = min(anchor_col + 1, snapshot.end_col)

    cols = {
        "num_quarters_used": find_column_from_header(
            snapshot, anchor_row, anchor_col, ["num quarters", "quarters used", "n quarters"], fallback_offset=-9
        ),
        "last_quarter_used": find_column_from_header(
            snapshot, anchor_row, anchor_col, ["last quarter", "quarter used"], fallback_offset=-8
        ),
        "avg_penetration_pct": find_column_from_header(
            snapshot, anchor_row, anchor_col, ["avg penetration", "average penetration", "penetration"], fallback_offset=-6
        ),
        "quarterly_sales": find_column_from_header(
            snapshot, anchor_row, anchor_col, ["quarterly sales", "quarter sales", "sold"], fallback_offset=-7
        ),
        "reported_sales": find_column_from_header(
            snapshot, anchor_row, anchor_col, ["reported sales", "reported", "actual"], fallback_offset=-5
        ),
        "growth_rate_pct": find_column_from_header(
            snapshot, anchor_row, anchor_col, ["growth rate", "growth"], fallback_offset=-4
        ),
        "sales_captured_in_db_pct": find_column_from_header(
            snapshot, anchor_row, anchor_col, ["captured in db", "sales captured", "db pct"], fallback_offset=-3
        ),
        "forecast_value": find_column_from_header(
            snapshot, anchor_row, anchor_col, ["estimated total sold", "forecast", "tot fcst"], fallback_offset=-2
        ),
        "actual_value": find_column_from_header(
            snapshot, anchor_row, anchor_col, ["actual value", "actual", "reported sales"], fallback_offset=-1
        ),
        "forecast_max": anchor_col,
        "forecast_min": min_col,
    }

    penetration_col = find_column_anywhere(
        snapshot=snapshot,
        keywords=["penetration"],
        anchor_col=anchor_col,
        max_row=anchor_row - 1,
    )
    if penetration_col is None:
        penetration_col = cols["avg_penetration_pct"]

    penetration_series = gather_numeric_series(snapshot, penetration_col, anchor_row - 1)
    scratch_cell = pick_scratch_cells(sheet, snapshot, count=1)[0]

    rows: List[Dict[str, Any]] = []
    for idx in range(1, 11):
        row_num = anchor_row + idx
        has_raw_row = row_num <= snapshot.end_row and row_has_any_value(
            snapshot,
            row_num,
            list(cols.values()),
        )
        if not has_raw_row and not penetration_series:
            break

        num_quarters_used = to_int(snapshot.get_value(row_num, cols["num_quarters_used"]))
        if num_quarters_used is None or num_quarters_used <= 0:
            num_quarters_used = idx

        avg_penetration = to_float(snapshot.get_value(row_num, cols["avg_penetration_pct"]))
        if penetration_series:
            sample_size = min(num_quarters_used, len(penetration_series))
            if sample_size > 0:
                start_row = penetration_series[-sample_size][0]
                end_row = penetration_series[-1][0]
                scratch_cell.formula2 = f"=AVERAGE(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})"
                workbook.app.calculate()
                calc_avg_penetration = to_float(scratch_cell.value)
                if calc_avg_penetration is not None:
                    avg_penetration = calc_avg_penetration

        forecast_value = to_float(snapshot.get_value(row_num, cols["forecast_value"]))
        actual_value = to_float(snapshot.get_value(row_num, cols["actual_value"]))
        forecast_max = to_float(snapshot.get_value(row_num, cols["forecast_max"]))
        forecast_min = to_float(snapshot.get_value(row_num, cols["forecast_min"]))
        quarterly_sales = to_float(snapshot.get_value(row_num, cols["quarterly_sales"]))
        reported_sales = to_float(snapshot.get_value(row_num, cols["reported_sales"]))
        growth_rate_pct = to_float(snapshot.get_value(row_num, cols["growth_rate_pct"]))
        sales_captured_in_db_pct = to_float(snapshot.get_value(row_num, cols["sales_captured_in_db_pct"]))

        if actual_value is None and reported_sales is not None:
            actual_value = reported_sales
        if reported_sales is None and actual_value is not None:
            reported_sales = actual_value
        if forecast_value is None and reported_sales is not None and avg_penetration not in (None, 0.0):
            forecast_value = reported_sales / avg_penetration
        if sales_captured_in_db_pct is None and forecast_value not in (None, 0.0) and reported_sales is not None:
            sales_captured_in_db_pct = reported_sales / forecast_value

        range_width: Optional[float] = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        row_payload = {
            "model": metadata["model"],
            "ticker": metadata["ticker"],
            "model_period": metadata["model_period"],
            "model_date": metadata["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration,
            "num_quarters_used": num_quarters_used,
            "last_quarter_used": snapshot.get_value(row_num, cols["last_quarter_used"]),
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "avg_penetration_pct": avg_penetration,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate_pct,
            "sales_captured_in_db_pct": sales_captured_in_db_pct,
            "source_file": source_file,
        }

        meaningful_values = [
            row_payload["parameter_value"],
            row_payload["forecast_value"],
            row_payload["actual_value"],
            row_payload["forecast_max"],
            row_payload["forecast_min"],
        ]
        if any(value is not None for value in meaningful_values):
            rows.append(row_payload)

    return rows


def extract_regression_candidates(
    workbook: xw.Book,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = workbook.sheets["Regression Model"]
    except Exception:
        print(f"Skipped regression extraction for {source_file}: missing 'Regression Model' sheet")
        return []

    snapshot = create_snapshot(sheet)
    anchor = find_anchor(snapshot, "max")
    if anchor is None:
        print(f"Skipped regression extraction for {source_file}: could not find 'max' anchor")
        return []

    anchor_row, anchor_col = anchor
    min_col = find_column_from_header(snapshot, anchor_row, anchor_col, ["min"], fallback_offset=1, window=6)
    if min_col == anchor_col:
        min_col = min(anchor_col + 1, snapshot.end_col)

    cols = {
        "num_quarters_used": find_column_from_header(
            snapshot, anchor_row, anchor_col, ["num quarters", "quarters used", "n quarters"], fallback_offset=-9
        ),
        "forecast_value": find_column_from_header(
            snapshot, anchor_row, anchor_col, ["tot fcst w/o sa", "tot fcst", "without sa", "forecast"], fallback_offset=-1
        ),
        "forecast_max": anchor_col,
        "forecast_min": min_col,
        "intercept_existing": find_column_from_header(
            snapshot, anchor_row, anchor_col, ["intercept"], fallback_offset=-4
        ),
        "slope_existing": find_column_from_header(snapshot, anchor_row, anchor_col, ["slope"], fallback_offset=-3),
    }

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    xy_pairs: List[Tuple[int, float, float]] = []
    for row in range(snapshot.start_row, anchor_row):
        x_value = to_float(snapshot.get_value(row, x_col))
        y_value = to_float(snapshot.get_value(row, y_col))
        if x_value is not None and y_value is not None:
            xy_pairs.append((row, x_value, y_value))

    scratch_intercept, scratch_slope = pick_scratch_cells(sheet, snapshot, count=2)

    rows: List[Dict[str, Any]] = []
    previous_key: Optional[Tuple[Any, ...]] = None
    for idx in range(1, 11):
        row_num = anchor_row + idx
        has_raw_row = row_num <= snapshot.end_row and row_has_any_value(
            snapshot,
            row_num,
            [cols["num_quarters_used"], cols["forecast_value"], cols["forecast_max"], cols["forecast_min"]],
        )
        if not has_raw_row and not xy_pairs:
            break

        num_quarters_used = to_int(snapshot.get_value(row_num, cols["num_quarters_used"]))
        if num_quarters_used is None or num_quarters_used <= 0:
            num_quarters_used = idx

        intercept_value = to_float(snapshot.get_value(row_num, cols["intercept_existing"]))
        slope_value = to_float(snapshot.get_value(row_num, cols["slope_existing"]))

        if len(xy_pairs) >= 2:
            sample_size = min(num_quarters_used, len(xy_pairs))
            if sample_size >= 2:
                start_row = xy_pairs[-sample_size][0]
                end_row = xy_pairs[-1][0]
                scratch_intercept.formula2 = (
                    f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
                )
                scratch_slope.formula2 = (
                    f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
                )
                workbook.app.calculate()
                calc_intercept = to_float(scratch_intercept.value)
                calc_slope = to_float(scratch_slope.value)
                if calc_intercept is not None:
                    intercept_value = calc_intercept
                if calc_slope is not None:
                    slope_value = calc_slope

        forecast_value = to_float(snapshot.get_value(row_num, cols["forecast_value"]))
        forecast_max = to_float(snapshot.get_value(row_num, cols["forecast_max"]))
        forecast_min = to_float(snapshot.get_value(row_num, cols["forecast_min"]))

        if forecast_value is None and intercept_value is not None and slope_value is not None and xy_pairs:
            latest_x = xy_pairs[-1][1]
            forecast_value = intercept_value + slope_value * latest_x

        range_width: Optional[float] = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        dedupe_key = (
            num_quarters_used,
            round(forecast_value, 10) if forecast_value is not None else None,
            round(forecast_max, 10) if forecast_max is not None else None,
            round(forecast_min, 10) if forecast_min is not None else None,
            round(intercept_value, 10) if intercept_value is not None else None,
            round(slope_value, 10) if slope_value is not None else None,
        )
        if previous_key is not None and dedupe_key == previous_key:
            continue
        previous_key = dedupe_key

        row_payload = {
            "model": metadata["model"],
            "ticker": metadata["ticker"],
            "model_period": metadata["model_period"],
            "model_date": metadata["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": None,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept_value,
            "slope": slope_value,
            "source_file": source_file,
        }

        meaningful_values = [
            row_payload["forecast_value"],
            row_payload["forecast_max"],
            row_payload["forecast_min"],
            row_payload["intercept"],
            row_payload["slope"],
        ]
        if any(value is not None for value in meaningful_values):
            rows.append(row_payload)

    return rows


def next_output_file(input_path: Path, output_path: Path) -> Path:
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


def should_skip_input_file(file_path: Path, input_path: Path) -> Optional[str]:
    if not file_path.is_file():
        return "not a file"
    if file_path.name.startswith("~"):
        return "temporary file"
    if file_path.suffix.lower() != ".xlsx":
        return "not an .xlsx file"

    output_name_pattern = rf"^{re.escape(input_path.name)}_PARAM(\.\d+)?\.xlsx$"
    if re.match(output_name_pattern, file_path.name, re.IGNORECASE):
        return "appears to be a previously generated output"
    return None


def write_sheet(workbook: Workbook, sheet_name: str, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws = workbook.create_sheet(title=sheet_name)

    for col_idx, column_name in enumerate(columns, start=1):
        ws.cell(row=1, column=col_idx, value=column_name)

    for row_idx, row_payload in enumerate(rows, start=2):
        for col_idx, column_name in enumerate(columns, start=1):
            ws.cell(row=row_idx, column=col_idx, value=row_payload.get(column_name))

    for header_cell in ws[1]:
        header_cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    last_column_letter = get_column_letter(len(columns))
    last_row = max(1, len(rows) + 1)
    ws.auto_filter.ref = f"A1:{last_column_letter}{last_row}"

    for col_idx, column_name in enumerate(columns, start=1):
        values = [column_name]
        for row_payload in rows:
            value = row_payload.get(column_name)
            values.append("" if value is None else str(value))
        max_length = max(len(value) for value in values)
        width = min(48, max(12, max_length + 2))
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    output_book = Workbook()
    default_sheet = output_book.active
    output_book.remove(default_sheet)

    write_sheet(
        workbook=output_book,
        sheet_name="empirical_candidates",
        columns=EMPIRICAL_COLUMNS,
        rows=empirical_rows,
    )
    write_sheet(
        workbook=output_book,
        sheet_name="regression_candidates",
        columns=REGRESSION_COLUMNS,
        rows=regression_rows,
    )
    output_book.save(output_path)


def run_extraction() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise RuntimeError(f"Input folder does not exist or is not a directory: {input_path}")

    output_path.mkdir(parents=True, exist_ok=True)
    generated_output_path = next_output_file(input_path=input_path, output_path=output_path)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_count = 0

    excel_app: Optional[xw.App] = None
    try:
        excel_app = xw.App(visible=False, add_book=False)
        excel_app.display_alerts = False
        excel_app.screen_updating = False

        for file_path in sorted(input_path.iterdir()):
            skip_reason = should_skip_input_file(file_path, input_path)
            if skip_reason is not None:
                print(f"Skipped file: {file_path.name} ({skip_reason})")
                continue

            workbook: Optional[xw.Book] = None
            try:
                workbook = excel_app.books.open(str(file_path), update_links=False)
                metadata = parse_file_metadata(file_path)
                empirical_rows.extend(
                    extract_empirical_candidates(
                        workbook=workbook,
                        metadata=metadata,
                        source_file=file_path.name,
                    )
                )
                regression_rows.extend(
                    extract_regression_candidates(
                        workbook=workbook,
                        metadata=metadata,
                        source_file=file_path.name,
                    )
                )
                processed_count += 1
                print(f"Processed file: {file_path.name}")
            except Exception as error:
                print(f"Skipped file: {file_path.name} (processing error: {error})")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        if excel_app is not None:
            excel_app.quit()

    write_output_workbook(
        output_path=generated_output_path,
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
    )

    print(f"Output path: {generated_output_path}")
    print(f"Number of files processed: {processed_count}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run_extraction()

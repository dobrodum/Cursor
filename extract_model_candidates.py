#!/usr/bin/env python3
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple


# ---------------------------------------------------------------------------
# User-configurable paths
# ---------------------------------------------------------------------------
input_dir = Path("./input")
output_dir = Path("./output")


N_QUARTERS = 10
PHASE_TO_DAY = {"early": 5, "mid": 15, "late": 25}

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


@dataclass
class SheetSnapshot:
    sheet: Any
    start_row: int
    start_col: int
    values: List[List[Any]]

    @property
    def end_row(self) -> int:
        return self.start_row + len(self.values) - 1

    @property
    def end_col(self) -> int:
        if not self.values or not self.values[0]:
            return self.start_col
        return self.start_col + len(self.values[0]) - 1

    def in_bounds(self, row: int, col: int) -> bool:
        return self.start_row <= row <= self.end_row and self.start_col <= col <= self.end_col

    def get_value(self, row: int, col: int) -> Any:
        if not self.in_bounds(row, col):
            return None
        return self.values[row - self.start_row][col - self.start_col]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = text.replace("\n", " ").replace("\r", " ")
    text = re.sub(r"[_\-/()%]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip() == ""
    return False


def as_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if text == "":
            return None
        multiplier = 1.0
        if text.endswith("%"):
            text = text[:-1]
            multiplier = 0.01
        try:
            return float(text) * multiplier
        except ValueError:
            return None
    return None


def as_int(value: Any) -> Optional[int]:
    parsed = as_float(value)
    if parsed is None:
        return None
    return int(round(parsed))


def build_output_path(source_input_dir: Path, source_output_dir: Path) -> Path:
    source_output_dir.mkdir(parents=True, exist_ok=True)
    folder_name = source_input_dir.resolve().name
    base_name = f"{folder_name}_PARAM"
    candidate = source_output_dir / f"{base_name}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = source_output_dir / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def parse_file_labels(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = ""
    if len(parts) >= 2 and parts[1]:
        ticker = parts[1].upper()
    if not ticker:
        ticker_match = re.search(r"\b([A-Z]{1,6})\b", stem)
        ticker = ticker_match.group(1) if ticker_match else "UNKNOWN"

    period_match = re.search(
        r"(Early|Mid|Late)\s*[-_ ]*([A-Za-z]{3,9})\s*[-_ ]*(20\d{2})",
        stem,
        flags=re.IGNORECASE,
    )

    if period_match:
        phase = period_match.group(1).title()
        month_text = re.sub(r"[^A-Za-z]", "", period_match.group(2))
        year = int(period_match.group(3))
    else:
        phase = "Mid"
        month_text = "Jan"
        year = date.today().year

    try:
        month_abbr = datetime.strptime(month_text[:3].title(), "%b").strftime("%b")
    except ValueError:
        month_abbr = "Jan"

    month_num = datetime.strptime(month_abbr, "%b").month
    day = PHASE_TO_DAY.get(phase.lower(), 15)
    model_period = f"{phase}{month_abbr}_{year}"
    model_date = date(year, month_num, day).isoformat()
    model = f"{ticker}_{model_period}"

    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def snapshot_sheet(sheet: Any) -> Optional[SheetSnapshot]:
    used = sheet.used_range
    raw_values = used.options(ndim=2).value
    if raw_values is None:
        return None

    if not isinstance(raw_values, list):
        raw_values = [[raw_values]]
    elif raw_values and not isinstance(raw_values[0], list):
        raw_values = [raw_values]

    normalized_rows: List[List[Any]] = []
    max_cols = 0
    for row in raw_values:
        if not isinstance(row, list):
            row = [row]
        max_cols = max(max_cols, len(row))
        normalized_rows.append(row)

    for idx, row in enumerate(normalized_rows):
        if len(row) < max_cols:
            normalized_rows[idx] = row + [None] * (max_cols - len(row))

    return SheetSnapshot(
        sheet=sheet,
        start_row=used.row,
        start_col=used.column,
        values=normalized_rows,
    )


def find_max_anchor(snapshot: SheetSnapshot) -> Optional[Tuple[int, int]]:
    candidates: List[Tuple[int, int, int]] = []
    for row in range(snapshot.start_row, snapshot.end_row + 1):
        for col in range(snapshot.start_col, snapshot.end_col + 1):
            if normalize_text(snapshot.get_value(row, col)) != "max":
                continue
            score = 0
            for probe_col in range(col - 2, col + 3):
                if as_float(snapshot.get_value(row + 1, probe_col)) is not None:
                    score += 1
            for probe_row in range(row + 1, row + 4):
                if as_float(snapshot.get_value(probe_row, col)) is not None:
                    score += 2
            if "min" in normalize_text(snapshot.get_value(row, col + 1)):
                score += 2
            candidates.append((score, row, col))

    if not candidates:
        return None

    candidates.sort(key=lambda item: item[0], reverse=True)
    _, anchor_row, anchor_col = candidates[0]
    return anchor_row, anchor_col


def choose_header_row(snapshot: SheetSnapshot, anchor_row: int) -> int:
    candidate_rows = range(max(snapshot.start_row, anchor_row - 2), min(snapshot.end_row, anchor_row + 2) + 1)
    best_row = anchor_row
    best_score = -1
    keywords = ("quarter", "forecast", "max", "min", "slope", "intercept", "sales", "penetration", "tot")

    for row in candidate_rows:
        score = 0
        for col in range(snapshot.start_col, snapshot.end_col + 1):
            text = normalize_text(snapshot.get_value(row, col))
            if not text:
                continue
            score += 1
            if any(keyword in text for keyword in keywords):
                score += 3
        if score > best_score:
            best_score = score
            best_row = row
    return best_row


def header_map(snapshot: SheetSnapshot, header_row: int) -> Dict[int, str]:
    headers: Dict[int, str] = {}
    for col in range(snapshot.start_col, snapshot.end_col + 1):
        text = normalize_text(snapshot.get_value(header_row, col))
        if text:
            headers[col] = text
    return headers


def find_column(headers: Dict[int, str], keyword_groups: Sequence[Sequence[str]], anchor_col: Optional[int] = None) -> Optional[int]:
    best_col: Optional[int] = None
    best_rank: Optional[Tuple[int, int]] = None

    for col, text in headers.items():
        for keywords in keyword_groups:
            if all(keyword in text for keyword in keywords):
                distance = abs(col - anchor_col) if anchor_col is not None else 0
                rank = (len(keywords), -distance)
                if best_rank is None or rank > best_rank:
                    best_rank = rank
                    best_col = col
    return best_col


def select_data_rows(
    snapshot: SheetSnapshot,
    data_start_row: int,
    preferred_cols: Iterable[Optional[int]],
    max_rows: int,
) -> List[int]:
    preferred = [col for col in preferred_cols if col is not None]
    rows: List[int] = []
    blank_streak = 0
    current = data_start_row

    while current <= snapshot.end_row and len(rows) < max_rows:
        has_data = any(not is_blank(snapshot.get_value(current, col)) for col in preferred) if preferred else False
        if has_data:
            rows.append(current)
            blank_streak = 0
        else:
            blank_streak += 1
            if rows and blank_streak >= 2:
                break
            if not rows and blank_streak >= 5:
                break
        current += 1

    if not rows:
        last_row = min(snapshot.end_row, data_start_row + max_rows - 1)
        rows = list(range(data_start_row, last_row + 1))

    return rows[:max_rows]


def safe_close_workbook(workbook: Any) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    close_attempts = (
        lambda: workbook.close(SaveChanges=False),
        lambda: workbook.api.Close(False),
        lambda: workbook.close(),
    )
    for close_call in close_attempts:
        try:
            close_call()
            return
        except Exception:
            continue


def subtract_if_numeric(left_value: Any, right_value: Any) -> Any:
    left = as_float(left_value)
    right = as_float(right_value)
    if left is None or right is None:
        return ""
    return left - right


def extract_empirical_rows(workbook: Any, meta: Dict[str, str], source_file: str) -> List[Dict[str, Any]]:
    if "Empirical Model" not in [sheet.name for sheet in workbook.sheets]:
        print(f"Skipped {source_file}: missing sheet 'Empirical Model'")
        return []

    sheet = workbook.sheets["Empirical Model"]
    snapshot = snapshot_sheet(sheet)
    if snapshot is None:
        print(f"Skipped {source_file}: empty sheet 'Empirical Model'")
        return []

    anchor = find_max_anchor(snapshot)
    if anchor is None:
        print(f"Skipped {source_file}: could not find 'max' anchor in 'Empirical Model'")
        return []

    anchor_row, anchor_col = anchor
    header_row = choose_header_row(snapshot, anchor_row)
    headers = header_map(snapshot, header_row)

    cols = {
        "num_quarters_used": find_column(
            headers,
            [
                ("num", "quarter"),
                ("n", "quarter"),
                ("quarters", "used"),
                ("#", "quarter"),
            ],
            anchor_col,
        ),
        "last_quarter_used": find_column(
            headers,
            [("last", "quarter"), ("quarter", "used")],
            anchor_col,
        ),
        "forecast_value": find_column(
            headers,
            [
                ("estimated", "total", "sold"),
                ("estimate", "total", "sold"),
                ("forecast", "value"),
                ("tot", "fcst"),
            ],
            anchor_col,
        ),
        "actual_value": find_column(
            headers,
            [("reported", "sales"), ("actual", "value"), ("actual", "sales")],
            anchor_col,
        ),
        "forecast_max": anchor_col,
        "forecast_min": find_column(headers, [("min",), ("minimum",)], anchor_col),
        "avg_penetration_pct": find_column(
            headers,
            [
                ("avg", "penetration"),
                ("average", "penetration"),
                ("penetration", "pct"),
            ],
            anchor_col,
        ),
        "quarterly_sales": find_column(
            headers,
            [("quarterly", "sales"), ("quarter", "sales")],
            anchor_col,
        ),
        "reported_sales": find_column(
            headers,
            [("reported", "sales")],
            anchor_col,
        ),
        "growth_rate_pct": find_column(
            headers,
            [("growth", "rate"), ("growth", "%"), ("growth", "pct")],
            anchor_col,
        ),
        "sales_captured_in_db_pct": find_column(
            headers,
            [
                ("sales", "captured", "db"),
                ("captured", "in", "db"),
                ("captured", "db"),
            ],
            anchor_col,
        ),
    }

    data_start_row = header_row + 1
    data_rows = select_data_rows(
        snapshot=snapshot,
        data_start_row=data_start_row,
        preferred_cols=[
            cols["num_quarters_used"],
            cols["forecast_value"],
            cols["forecast_max"],
            cols["forecast_min"],
        ],
        max_rows=N_QUARTERS,
    )

    scratch_col = snapshot.end_col + 2
    avg_formulas_written: Dict[int, bool] = {}
    for index, row in enumerate(data_rows, start=1):
        num_quarters = as_int(snapshot.get_value(row, cols["num_quarters_used"])) if cols["num_quarters_used"] else index
        lookback = max(1, num_quarters or index)

        formula = ""
        avg_source_col = cols["avg_penetration_pct"] or cols["sales_captured_in_db_pct"]
        if avg_source_col is not None:
            window_start = max(data_start_row, row - lookback + 1)
            formula = f'=IFERROR(AVERAGE(R{window_start}C{avg_source_col}:R{row}C{avg_source_col}),"")'
        elif cols["quarterly_sales"] and cols["reported_sales"]:
            formula = (
                f'=IFERROR(R{row}C{cols["quarterly_sales"]}/R{row}C{cols["reported_sales"]},"")'
            )

        if formula:
            sheet.range((row, scratch_col)).formula2 = formula
            avg_formulas_written[row] = True

    avg_calculated: Dict[int, Any] = {}
    if avg_formulas_written:
        workbook.app.calculate()
        first_formula_row = min(avg_formulas_written)
        last_formula_row = max(avg_formulas_written)
        range_values = sheet.range((first_formula_row, scratch_col), (last_formula_row, scratch_col)).options(ndim=2).value
        if range_values and not isinstance(range_values[0], list):
            range_values = [range_values]
        for offset, row in enumerate(range(first_formula_row, last_formula_row + 1)):
            avg_calculated[row] = range_values[offset][0]

    rows: List[Dict[str, Any]] = []
    for index, row in enumerate(data_rows, start=1):
        num_quarters = snapshot.get_value(row, cols["num_quarters_used"]) if cols["num_quarters_used"] else index
        last_quarter = snapshot.get_value(row, cols["last_quarter_used"]) if cols["last_quarter_used"] else ""
        forecast_value = snapshot.get_value(row, cols["forecast_value"]) if cols["forecast_value"] else ""
        actual_value = snapshot.get_value(row, cols["actual_value"]) if cols["actual_value"] else ""
        forecast_max = snapshot.get_value(row, cols["forecast_max"]) if cols["forecast_max"] else ""
        forecast_min = snapshot.get_value(row, cols["forecast_min"]) if cols["forecast_min"] else ""
        quarterly_sales = snapshot.get_value(row, cols["quarterly_sales"]) if cols["quarterly_sales"] else ""
        reported_sales = snapshot.get_value(row, cols["reported_sales"]) if cols["reported_sales"] else ""
        growth_rate = snapshot.get_value(row, cols["growth_rate_pct"]) if cols["growth_rate_pct"] else ""
        sales_captured = (
            snapshot.get_value(row, cols["sales_captured_in_db_pct"]) if cols["sales_captured_in_db_pct"] else ""
        )

        avg_penetration = avg_calculated.get(row, "")
        if is_blank(avg_penetration) and cols["avg_penetration_pct"] is not None:
            avg_penetration = snapshot.get_value(row, cols["avg_penetration_pct"])

        if all(is_blank(v) for v in [num_quarters, forecast_value, forecast_max, forecast_min, avg_penetration]):
            continue

        rows.append(
            {
                "model": meta["model"],
                "ticker": meta["ticker"],
                "model_period": meta["model_period"],
                "model_date": meta["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters,
                "last_quarter_used": last_quarter,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": subtract_if_numeric(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate,
                "sales_captured_in_db_pct": sales_captured,
                "source_file": source_file,
            }
        )

    return rows


def value_signature(*values: Any) -> Tuple[Any, ...]:
    signature: List[Any] = []
    for value in values:
        parsed = as_float(value)
        if parsed is not None:
            signature.append(round(parsed, 10))
        elif is_blank(value):
            signature.append("")
        else:
            signature.append(str(value).strip())
    return tuple(signature)


def extract_regression_rows(workbook: Any, meta: Dict[str, str], source_file: str) -> List[Dict[str, Any]]:
    if "Regression Model" not in [sheet.name for sheet in workbook.sheets]:
        print(f"Skipped {source_file}: missing sheet 'Regression Model'")
        return []

    sheet = workbook.sheets["Regression Model"]
    snapshot = snapshot_sheet(sheet)
    if snapshot is None:
        print(f"Skipped {source_file}: empty sheet 'Regression Model'")
        return []

    anchor = find_max_anchor(snapshot)
    if anchor is None:
        print(f"Skipped {source_file}: could not find 'max' anchor in 'Regression Model'")
        return []

    anchor_row, anchor_col = anchor
    header_row = choose_header_row(snapshot, anchor_row)
    headers = header_map(snapshot, header_row)

    cols = {
        "num_quarters_used": find_column(
            headers,
            [
                ("num", "quarter"),
                ("n", "quarter"),
                ("quarters", "used"),
                ("#", "quarter"),
            ],
            anchor_col,
        ),
        "forecast_value": find_column(
            headers,
            [
                ("tot", "fcst", "w", "sa"),
                ("tot", "fcst", "without", "sa"),
                ("forecast", "without", "sa"),
                ("tot", "fcst"),
            ],
            anchor_col,
        ),
        "actual_value": find_column(
            headers,
            [("actual", "value"), ("actual", "sales"), ("reported", "sales")],
            anchor_col,
        ),
        "forecast_max": anchor_col,
        "forecast_min": find_column(headers, [("min",), ("minimum",)], anchor_col),
        "intercept": find_column(headers, [("intercept",)], anchor_col),
        "slope": find_column(headers, [("slope",)], anchor_col),
    }

    # Required anchor-based offsets for source regression series.
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    data_start_row = header_row + 1
    data_rows = select_data_rows(
        snapshot=snapshot,
        data_start_row=data_start_row,
        preferred_cols=[
            cols["num_quarters_used"],
            cols["forecast_value"],
            cols["forecast_max"],
            cols["forecast_min"],
        ],
        max_rows=N_QUARTERS,
    )

    # Build source data rows once for intercept/slope formulas.
    series_rows = [
        row
        for row in range(snapshot.start_row, header_row)
        if as_float(snapshot.get_value(row, x_col)) is not None and as_float(snapshot.get_value(row, y_col)) is not None
    ]
    if len(series_rows) < 2:
        series_rows = [
            row
            for row in range(snapshot.start_row, snapshot.end_row + 1)
            if as_float(snapshot.get_value(row, x_col)) is not None and as_float(snapshot.get_value(row, y_col)) is not None
        ]

    intercept_scratch_col = snapshot.end_col + 3
    slope_scratch_col = snapshot.end_col + 4
    formulas_written = False

    for index, row in enumerate(data_rows, start=1):
        num_q_raw = snapshot.get_value(row, cols["num_quarters_used"]) if cols["num_quarters_used"] else index
        num_q = max(2, as_int(num_q_raw) or index)
        if len(series_rows) < 2:
            continue
        num_q = min(num_q, len(series_rows))
        start_row = series_rows[-num_q]
        end_row = series_rows[-1]
        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col}),"")'
        )
        sheet.range((row, intercept_scratch_col)).formula2 = intercept_formula
        sheet.range((row, slope_scratch_col)).formula2 = slope_formula
        formulas_written = True

    calculated_intercepts: Dict[int, Any] = {}
    calculated_slopes: Dict[int, Any] = {}
    if formulas_written:
        workbook.app.calculate()
        first_row = min(data_rows)
        last_row = max(data_rows)
        intercept_values = sheet.range((first_row, intercept_scratch_col), (last_row, intercept_scratch_col)).options(ndim=2).value
        slope_values = sheet.range((first_row, slope_scratch_col), (last_row, slope_scratch_col)).options(ndim=2).value
        if intercept_values and not isinstance(intercept_values[0], list):
            intercept_values = [intercept_values]
        if slope_values and not isinstance(slope_values[0], list):
            slope_values = [slope_values]
        for offset, row in enumerate(range(first_row, last_row + 1)):
            calculated_intercepts[row] = intercept_values[offset][0]
            calculated_slopes[row] = slope_values[offset][0]

    rows: List[Dict[str, Any]] = []
    prev_signature: Optional[Tuple[Any, ...]] = None
    latest_x = as_float(snapshot.get_value(series_rows[-1], x_col)) if series_rows else None

    for index, row in enumerate(data_rows, start=1):
        num_quarters = snapshot.get_value(row, cols["num_quarters_used"]) if cols["num_quarters_used"] else index
        forecast_value = snapshot.get_value(row, cols["forecast_value"]) if cols["forecast_value"] else ""
        actual_value = snapshot.get_value(row, cols["actual_value"]) if cols["actual_value"] else ""
        forecast_max = snapshot.get_value(row, cols["forecast_max"]) if cols["forecast_max"] else ""
        forecast_min = snapshot.get_value(row, cols["forecast_min"]) if cols["forecast_min"] else ""

        intercept = snapshot.get_value(row, cols["intercept"]) if cols["intercept"] else ""
        slope = snapshot.get_value(row, cols["slope"]) if cols["slope"] else ""
        if is_blank(intercept):
            intercept = calculated_intercepts.get(row, "")
        if is_blank(slope):
            slope = calculated_slopes.get(row, "")

        # Fallback if workbook does not expose TOT FCST w/o SA directly.
        if is_blank(forecast_value) and latest_x is not None and as_float(intercept) is not None and as_float(slope) is not None:
            forecast_value = as_float(intercept) + as_float(slope) * (latest_x + 1.0)

        if is_blank(actual_value):
            actual_value = ""

        if all(is_blank(v) for v in [num_quarters, forecast_value, forecast_max, forecast_min, intercept, slope]):
            continue

        current_signature = value_signature(num_quarters, forecast_value, forecast_max, forecast_min, intercept, slope)
        if prev_signature is not None and current_signature == prev_signature:
            continue
        prev_signature = current_signature

        rows.append(
            {
                "model": meta["model"],
                "ticker": meta["ticker"],
                "model_period": meta["model_period"],
                "model_date": meta["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters,
                "num_quarters_used": num_quarters,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": subtract_if_numeric(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def apply_sheet_formatting(worksheet: Any, columns: Sequence[str]) -> None:
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter

    for cell in worksheet[1]:
        cell.font = Font(bold=True)
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    max_scan_rows = min(worksheet.max_row, 5000)
    for index, column_name in enumerate(columns, start=1):
        width = max(12, len(column_name) + 2)
        for row in range(2, max_scan_rows + 1):
            value = worksheet.cell(row=row, column=index).value
            if value is None:
                continue
            width = max(width, len(str(value)) + 2)
        worksheet.column_dimensions[get_column_letter(index)].width = min(width, 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    from openpyxl import Workbook

    workbook = Workbook()
    empirical_sheet = workbook.active
    empirical_sheet.title = "empirical_candidates"
    regression_sheet = workbook.create_sheet("regression_candidates")

    empirical_sheet.append(EMPIRICAL_COLUMNS)
    for row in empirical_rows:
        empirical_sheet.append([row.get(column, "") for column in EMPIRICAL_COLUMNS])

    regression_sheet.append(REGRESSION_COLUMNS)
    for row in regression_rows:
        regression_sheet.append([row.get(column, "") for column in REGRESSION_COLUMNS])

    apply_sheet_formatting(empirical_sheet, EMPIRICAL_COLUMNS)
    apply_sheet_formatting(regression_sheet, REGRESSION_COLUMNS)

    workbook.save(output_path)


def main() -> None:
    if not input_dir.exists():
        print(f"Skipped run: input_dir does not exist -> {input_dir.resolve()}")
        return

    try:
        import openpyxl  # noqa: F401
    except ModuleNotFoundError:
        print("Skipped run: missing dependency 'openpyxl'. Install it and rerun.")
        return

    try:
        import xlwings as xw  # type: ignore
    except ModuleNotFoundError:
        print("Skipped run: missing dependency 'xlwings'. Install it and rerun.")
        return

    output_path = build_output_path(input_dir, output_dir)
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(input_dir.iterdir(), key=lambda path: path.name.lower()):
            if not file_path.is_file():
                print(f"Skipped {file_path.name}: not a file")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary Excel file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file")
                continue

            print(f"Processing {file_path.name}")
            workbook = None
            metadata = parse_file_labels(file_path)

            try:
                workbook = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_rows(workbook, metadata, file_path.name))
                regression_rows.extend(extract_regression_rows(workbook, metadata, file_path.name))
                processed_files += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        try:
            app.quit()
        except Exception:
            pass

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path.resolve()}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

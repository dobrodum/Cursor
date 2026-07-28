from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ---------- user-configurable paths ----------
input_dir = Path("input")
output_dir = Path("output")

# ---------- constants ----------
N_QUARTERS = 10

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

MODEL_DAY = {
    "early": 5,
    "mid": 15,
    "late": 25,
}

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

# Anchor-based fallback offsets from the "max" label cell.
EMPIRICAL_OFFSETS = {
    "forecast_max": (0, 1),
    "forecast_min": (1, 1),
    "num_quarters_used": (-1, -1),
    "last_quarter_used": (-2, 1),
    "forecast_value": (-3, 1),
    "quarterly_sales": (-4, 1),
    "reported_sales": (-5, 1),
    "growth_rate_pct": (-6, 1),
    "sales_captured_in_db_pct": (-7, 1),
}

REGRESSION_OFFSETS = {
    "forecast_max": (0, 1),
    "forecast_min": (1, 1),
    "num_quarters_used": (-1, -1),
    "forecast_value": (-3, 1),
}


@dataclass
class ModelMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetSnapshot:
    sheet: xw.main.Sheet
    start_row: int
    start_col: int
    values: List[List[Any]]
    text_cells: List[Tuple[int, int, str]]

    @classmethod
    def capture(cls, sheet: xw.main.Sheet) -> "SheetSnapshot":
        used = sheet.used_range
        raw = used.value
        values = _as_2d(raw)
        start_row = used.row
        start_col = used.column
        text_cells: List[Tuple[int, int, str]] = []
        for row_idx, row_vals in enumerate(values, start=start_row):
            for col_idx, cell_val in enumerate(row_vals, start=start_col):
                if isinstance(cell_val, str) and cell_val.strip():
                    text_cells.append((row_idx, col_idx, cell_val.strip()))
        return cls(
            sheet=sheet,
            start_row=start_row,
            start_col=start_col,
            values=values,
            text_cells=text_cells,
        )

    @property
    def end_row(self) -> int:
        if not self.values:
            return self.start_row
        return self.start_row + len(self.values) - 1

    @property
    def end_col(self) -> int:
        width = max((len(row_vals) for row_vals in self.values), default=1)
        return self.start_col + width - 1

    def value_at(self, row: int, col: int) -> Any:
        row_idx = row - self.start_row
        col_idx = col - self.start_col
        if row_idx < 0 or col_idx < 0 or row_idx >= len(self.values):
            return None
        row_vals = self.values[row_idx]
        if col_idx >= len(row_vals):
            return None
        return row_vals[col_idx]


def _as_2d(raw: Any) -> List[List[Any]]:
    if raw is None:
        return []
    if isinstance(raw, list):
        if raw and isinstance(raw[0], list):
            return raw
        return [raw]
    return [[raw]]


def normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", value.strip().lower())


def as_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.replace(",", "").replace("%", "").strip()
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def subtract_or_none(left: Any, right: Any) -> Optional[float]:
    left_num = as_float(left)
    right_num = as_float(right)
    if left_num is None or right_num is None:
        return None
    return left_num - right_num


def round_for_signature(value: Any) -> Any:
    value_num = as_float(value)
    if value_num is None:
        return value
    return round(value_num, 10)


def parse_model_meta(filename: str) -> Optional[ModelMeta]:
    stem = Path(filename).stem
    stem = re.sub(r"_send$", "", stem, flags=re.IGNORECASE)

    pattern = re.compile(
        r"\s-\s(?P<ticker>[A-Za-z0-9]+)\s-\s"
        r"(?P<period>(?P<phase>Early|Mid|Late)(?P<month>Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(?P<year>\d{4}))",
        flags=re.IGNORECASE,
    )
    match = pattern.search(stem)
    if not match:
        return None

    ticker = match.group("ticker").upper()
    phase_key = match.group("phase").lower()
    month_key = match.group("month").lower()
    year = int(match.group("year"))

    if phase_key not in MODEL_DAY or month_key not in MONTH_TO_NUMBER:
        return None

    phase_formatted = phase_key.capitalize()
    month_formatted = month_key.capitalize()
    model_period = f"{phase_formatted}{month_formatted}_{year}"
    model_date = date(year, MONTH_TO_NUMBER[month_key], MODEL_DAY[phase_key]).isoformat()
    model = f"{ticker}_{model_period}"
    return ModelMeta(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def next_output_path(src_dir: Path, dst_dir: Path) -> Path:
    folder_name = src_dir.resolve().name
    base_name = f"{folder_name}_PARAM.xlsx"
    first_candidate = dst_dir / base_name
    if not first_candidate.exists():
        return first_candidate

    index = 1
    while True:
        candidate = dst_dir / f"{folder_name}_PARAM.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def safe_close_workbook(wb: xw.main.Book) -> None:
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
    except Exception as exc:
        print(f"warning: could not close workbook cleanly ({exc})")


def find_anchor_max(snapshot: SheetSnapshot) -> Optional[Tuple[int, int]]:
    exact_candidates: List[Tuple[int, int]] = []
    with_min_below: List[Tuple[int, int]] = []

    for row, col, text in snapshot.text_cells:
        norm = normalize_text(text)
        if norm == "max":
            exact_candidates.append((row, col))
            below = snapshot.value_at(row + 1, col)
            if isinstance(below, str) and normalize_text(below) == "min":
                with_min_below.append((row, col))

    if with_min_below:
        return with_min_below[0]
    if exact_candidates:
        return exact_candidates[0]

    for row, col, text in snapshot.text_cells:
        if normalize_text(text) == "maximum":
            return (row, col)
    return None


def find_label_cell(
    snapshot: SheetSnapshot,
    required_terms: Sequence[str],
    anchor: Optional[Tuple[int, int]] = None,
) -> Optional[Tuple[int, int]]:
    target_terms = [normalize_text(term) for term in required_terms]
    matches: List[Tuple[int, int, int]] = []  # distance, row, col

    for row, col, text in snapshot.text_cells:
        norm = normalize_text(text)
        if all(term in norm for term in target_terms):
            if anchor is None:
                distance = 0
            else:
                distance = abs(row - anchor[0]) + abs(col - anchor[1])
            matches.append((distance, row, col))

    if not matches:
        return None
    matches.sort(key=lambda item: item[0])
    _, row, col = matches[0]
    return row, col


def resolve_value_cell(
    snapshot: SheetSnapshot,
    label_terms: Sequence[str],
    anchor: Tuple[int, int],
    fallback_offset: Tuple[int, int],
) -> Tuple[int, int]:
    label_cell = find_label_cell(snapshot, label_terms, anchor=anchor)
    if label_cell:
        row, col = label_cell
        for row_delta, col_delta in ((0, 1), (0, 2), (1, 0), (1, 1), (0, -1)):
            candidate_row = row + row_delta
            candidate_col = col + col_delta
            candidate_val = snapshot.value_at(candidate_row, candidate_col)
            if candidate_val not in (None, ""):
                return candidate_row, candidate_col
        return row, col + 1

    return anchor[0] + fallback_offset[0], anchor[1] + fallback_offset[1]


def read_cell(sheet: xw.main.Sheet, cell: Optional[Tuple[int, int]]) -> Any:
    if cell is None:
        return None
    return sheet.range(cell).value


def r1c1_ref(
    base_row: int,
    base_col: int,
    target_row: int,
    target_col: int,
) -> str:
    row_delta = target_row - base_row
    col_delta = target_col - base_col
    row_ref = "R" if row_delta == 0 else f"R[{row_delta}]"
    col_ref = "C" if col_delta == 0 else f"C[{col_delta}]"
    return f"{row_ref}{col_ref}"


def find_penetration_series(
    snapshot: SheetSnapshot,
    anchor: Tuple[int, int],
) -> Optional[Tuple[int, List[int]]]:
    candidates: List[Tuple[int, int, List[int]]] = []
    for row, col, text in snapshot.text_cells:
        norm = normalize_text(text)
        if "penetration" not in norm:
            continue
        numeric_cols = []
        for test_col in range(snapshot.start_col, snapshot.end_col + 1):
            value = snapshot.value_at(row, test_col)
            if as_float(value) is not None:
                numeric_cols.append(test_col)
        if len(numeric_cols) >= 2:
            distance = abs(anchor[0] - row) + abs(anchor[1] - col)
            candidates.append((distance, row, numeric_cols))

    if not candidates:
        return None

    # Prefer richer series first, then nearest to anchor.
    candidates.sort(key=lambda item: (-len(item[2]), item[0]))
    _, row, numeric_cols = candidates[0]
    numeric_cols.sort()
    return row, numeric_cols


def build_empirical_rows(
    wb: xw.main.Book,
    meta: ModelMeta,
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet_name = "Empirical Model"
    if sheet_name not in [sht.name for sht in wb.sheets]:
        print(f"skipped empirical sheet for {source_file}: missing '{sheet_name}'")
        return []

    sheet = wb.sheets[sheet_name]
    snapshot = SheetSnapshot.capture(sheet)
    anchor = find_anchor_max(snapshot)
    if anchor is None:
        print(f"skipped empirical sheet for {source_file}: 'max' anchor not found")
        return []

    forecast_max_cell = (anchor[0] + EMPIRICAL_OFFSETS["forecast_max"][0], anchor[1] + EMPIRICAL_OFFSETS["forecast_max"][1])
    forecast_min_cell = (anchor[0] + EMPIRICAL_OFFSETS["forecast_min"][0], anchor[1] + EMPIRICAL_OFFSETS["forecast_min"][1])
    num_quarters_cell = resolve_value_cell(
        snapshot,
        label_terms=("num", "quarter"),
        anchor=anchor,
        fallback_offset=EMPIRICAL_OFFSETS["num_quarters_used"],
    )
    last_quarter_cell = resolve_value_cell(
        snapshot,
        label_terms=("last", "quarter"),
        anchor=anchor,
        fallback_offset=EMPIRICAL_OFFSETS["last_quarter_used"],
    )
    forecast_value_cell = resolve_value_cell(
        snapshot,
        label_terms=("estimated", "total", "sold"),
        anchor=anchor,
        fallback_offset=EMPIRICAL_OFFSETS["forecast_value"],
    )
    quarterly_sales_cell = resolve_value_cell(
        snapshot,
        label_terms=("quarterly", "sales"),
        anchor=anchor,
        fallback_offset=EMPIRICAL_OFFSETS["quarterly_sales"],
    )
    reported_sales_cell = resolve_value_cell(
        snapshot,
        label_terms=("reported", "sales"),
        anchor=anchor,
        fallback_offset=EMPIRICAL_OFFSETS["reported_sales"],
    )
    growth_rate_cell = resolve_value_cell(
        snapshot,
        label_terms=("growth", "rate"),
        anchor=anchor,
        fallback_offset=EMPIRICAL_OFFSETS["growth_rate_pct"],
    )
    sales_captured_cell = resolve_value_cell(
        snapshot,
        label_terms=("sales", "captured", "db"),
        anchor=anchor,
        fallback_offset=EMPIRICAL_OFFSETS["sales_captured_in_db_pct"],
    )

    penetration_series = find_penetration_series(snapshot, anchor)
    helper_row = anchor[0]
    helper_col = snapshot.end_col + 6
    avg_pen_cell = sheet.range((helper_row, helper_col))

    output_rows: List[Dict[str, Any]] = []
    for quarter_count in range(1, N_QUARTERS + 1):
        formulas_updated = False

        # Keep workbook model logic in control by setting quarter count input.
        sheet.range(num_quarters_cell).value = quarter_count
        formulas_updated = True

        if penetration_series is not None:
            pen_row, numeric_cols = penetration_series
            subset_size = min(quarter_count, len(numeric_cols))
            selected_cols = numeric_cols[-subset_size:]
            start_col = selected_cols[0]
            end_col = selected_cols[-1]
            start_ref = r1c1_ref(helper_row, helper_col, pen_row, start_col)
            end_ref = r1c1_ref(helper_row, helper_col, pen_row, end_col)
            avg_pen_cell.formula2 = f"=AVERAGE({start_ref}:{end_ref})"
            formulas_updated = True

        if formulas_updated:
            wb.app.calculate()

        avg_penetration_pct = avg_pen_cell.value if penetration_series is not None else None
        forecast_value = read_cell(sheet, forecast_value_cell)
        quarterly_sales = read_cell(sheet, quarterly_sales_cell)
        reported_sales = read_cell(sheet, reported_sales_cell)
        growth_rate_pct = read_cell(sheet, growth_rate_cell)
        sales_captured_pct = read_cell(sheet, sales_captured_cell)
        forecast_max = read_cell(sheet, forecast_max_cell)
        forecast_min = read_cell(sheet, forecast_min_cell)
        range_width = subtract_or_none(forecast_max, forecast_min)
        last_quarter_used = read_cell(sheet, last_quarter_cell)

        if (forecast_value is None or forecast_value == "") and as_float(quarterly_sales) is not None:
            avg_num = as_float(avg_penetration_pct)
            sales_num = as_float(quarterly_sales)
            if avg_num not in (None, 0.0) and sales_num is not None:
                forecast_value = sales_num / avg_num

        output_rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": quarter_count,
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

    return output_rows


def find_numeric_rows(
    snapshot: SheetSnapshot,
    x_col: int,
    y_col: int,
    max_row_exclusive: int,
) -> List[int]:
    numeric_rows: List[int] = []
    for row in range(snapshot.start_row, max_row_exclusive):
        x_val = snapshot.value_at(row, x_col)
        y_val = snapshot.value_at(row, y_col)
        if as_float(x_val) is not None and as_float(y_val) is not None:
            numeric_rows.append(row)
    return numeric_rows


def build_regression_rows(
    wb: xw.main.Book,
    meta: ModelMeta,
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet_name = "Regression Model"
    if sheet_name not in [sht.name for sht in wb.sheets]:
        print(f"skipped regression sheet for {source_file}: missing '{sheet_name}'")
        return []

    sheet = wb.sheets[sheet_name]
    snapshot = SheetSnapshot.capture(sheet)
    anchor = find_anchor_max(snapshot)
    if anchor is None:
        print(f"skipped regression sheet for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    numeric_rows = find_numeric_rows(snapshot, x_col=x_col, y_col=y_col, max_row_exclusive=anchor_row)
    if not numeric_rows:
        print(f"skipped regression rows for {source_file}: no x/y numeric rows found")
        return []

    num_quarters_cell = resolve_value_cell(
        snapshot,
        label_terms=("num", "quarter"),
        anchor=anchor,
        fallback_offset=REGRESSION_OFFSETS["num_quarters_used"],
    )
    forecast_total_cell = resolve_value_cell(
        snapshot,
        label_terms=("tot", "fcst", "w/o", "sa"),
        anchor=anchor,
        fallback_offset=REGRESSION_OFFSETS["forecast_value"],
    )
    forecast_max_cell = (anchor_row + REGRESSION_OFFSETS["forecast_max"][0], anchor_col + REGRESSION_OFFSETS["forecast_max"][1])
    forecast_min_cell = (anchor_row + REGRESSION_OFFSETS["forecast_min"][0], anchor_col + REGRESSION_OFFSETS["forecast_min"][1])

    helper_col = snapshot.end_col + 8
    intercept_cell = sheet.range((anchor_row, helper_col))
    slope_cell = sheet.range((anchor_row + 1, helper_col))

    output_rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for quarter_count in range(1, N_QUARTERS + 1):
        effective_quarters = min(quarter_count, len(numeric_rows))
        rows_used = numeric_rows[-effective_quarters:]
        if not rows_used:
            continue
        start_row = rows_used[0]
        end_row = rows_used[-1]

        sheet.range(num_quarters_cell).value = effective_quarters

        y_start = r1c1_ref(anchor_row, helper_col, start_row, y_col)
        y_end = r1c1_ref(anchor_row, helper_col, end_row, y_col)
        x_start = r1c1_ref(anchor_row, helper_col, start_row, x_col)
        x_end = r1c1_ref(anchor_row, helper_col, end_row, x_col)

        intercept_cell.formula2 = f"=INTERCEPT({y_start}:{y_end},{x_start}:{x_end})"
        slope_cell.formula2 = f"=SLOPE({y_start}:{y_end},{x_start}:{x_end})"

        wb.app.calculate()

        intercept = intercept_cell.value
        slope = slope_cell.value
        forecast_value = read_cell(sheet, forecast_total_cell)
        forecast_max = read_cell(sheet, forecast_max_cell)
        forecast_min = read_cell(sheet, forecast_min_cell)
        range_width = subtract_or_none(forecast_max, forecast_min)

        signature = (
            round_for_signature(forecast_value),
            round_for_signature(forecast_max),
            round_for_signature(forecast_min),
            round_for_signature(intercept),
            round_for_signature(slope),
        )
        if signature == previous_signature:
            continue
        previous_signature = signature

        output_rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": effective_quarters,
                "num_quarters_used": effective_quarters,
                "forecast_value": forecast_value,
                "actual_value": "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return output_rows


def write_sheet(
    workbook: Workbook,
    sheet_name: str,
    headers: Sequence[str],
    rows: Iterable[Dict[str, Any]],
) -> None:
    ws = workbook.create_sheet(title=sheet_name)
    ws.append(list(headers))

    for row_data in rows:
        ws.append([row_data.get(col) for col in headers])

    header_font = Font(bold=True)
    for cell in ws[1]:
        cell.font = header_font

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"

    for col_idx, header in enumerate(headers, start=1):
        max_length = len(header)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_length = max(max_length, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = max(12, min(50, max_length + 2))


def write_output_workbook(
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
    output_path: Path,
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    write_sheet(workbook, "empirical_candidates", EMPIRICAL_HEADERS, empirical_rows)
    write_sheet(workbook, "regression_candidates", REGRESSION_HEADERS, regression_rows)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


def process_workbooks(src_dir: Path, dst_dir: Path) -> None:
    output_path = next_output_path(src_dir, dst_dir)
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(src_dir.iterdir()):
            if not file_path.is_file():
                print(f"skipped: {file_path.name} (not a file)")
                continue
            if file_path.name.startswith("~"):
                print(f"skipped: {file_path.name} (temporary workbook)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped: {file_path.name} (not .xlsx)")
                continue

            meta = parse_model_meta(file_path.name)
            if meta is None:
                print(f"skipped: {file_path.name} (filename format not recognized)")
                continue

            wb: Optional[xw.main.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(build_empirical_rows(wb, meta, file_path.name))
                regression_rows.extend(build_regression_rows(wb, meta, file_path.name))
                processed_files += 1
                print(f"processed: {file_path.name}")
            except Exception as exc:
                print(f"skipped: {file_path.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(empirical_rows, regression_rows, output_path)

    print(f"output_path: {output_path}")
    print(f"files_processed: {processed_files}")
    print(f"empirical_rows: {len(empirical_rows)}")
    print(f"regression_rows: {len(regression_rows)}")


def main() -> None:
    src_dir = Path(input_dir).expanduser().resolve()
    dst_dir = Path(output_dir).expanduser().resolve()

    if not src_dir.exists():
        raise FileNotFoundError(f"input_dir does not exist: {src_dir}")
    if not src_dir.is_dir():
        raise NotADirectoryError(f"input_dir is not a directory: {src_dir}")

    dst_dir.mkdir(parents=True, exist_ok=True)
    process_workbooks(src_dir=src_dir, dst_dir=dst_dir)


if __name__ == "__main__":
    main()

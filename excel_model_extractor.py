#!/usr/bin/env python3
from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ----------------------------
# User-configurable locations
# ----------------------------
input_dir = "/workspace/input"
output_dir = "/workspace/output"

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

DAY_MAP = {"Early": 5, "Mid": 15, "Late": 25}

_NON_ALNUM = re.compile(r"[^a-z0-9]+")


@dataclass
class FileMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetSnapshot:
    start_row: int
    start_col: int
    values: List[List[Any]]
    labels: Dict[str, List[Tuple[int, int]]]

    @property
    def row_count(self) -> int:
        return len(self.values)

    @property
    def col_count(self) -> int:
        return len(self.values[0]) if self.values else 0

    @property
    def last_row(self) -> int:
        return self.start_row + self.row_count - 1

    @property
    def last_col(self) -> int:
        return self.start_col + self.col_count - 1

    def get(self, row: int, col: int) -> Any:
        if row < self.start_row or col < self.start_col:
            return None
        row_idx = row - self.start_row
        col_idx = col - self.start_col
        if row_idx < 0 or col_idx < 0:
            return None
        if row_idx >= self.row_count or col_idx >= self.col_count:
            return None
        return self.values[row_idx][col_idx]


def normalize_label(value: Any) -> str:
    return _NON_ALNUM.sub(" ", str(value).strip().lower()).strip()


def to_float(value: Any) -> Optional[float]:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "")
        if not cleaned:
            return None
        if cleaned.endswith("%"):
            cleaned = cleaned[:-1]
            try:
                return float(cleaned) / 100.0
            except ValueError:
                return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def normalize_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            rows = values
        else:
            rows = [values]
    else:
        rows = [[values]]

    col_count = max((len(row) for row in rows), default=0)
    if col_count == 0:
        return []
    normalized_rows: List[List[Any]] = []
    for row in rows:
        normalized_rows.append(row + [None] * (col_count - len(row)))
    return normalized_rows


def build_snapshot(sheet: xw.Sheet) -> Optional[SheetSnapshot]:
    used = sheet.used_range
    values = normalize_2d(used.options(ndim=2).value)
    if not values:
        return None

    labels: Dict[str, List[Tuple[int, int]]] = {}
    for r_idx, row in enumerate(values):
        for c_idx, value in enumerate(row):
            if isinstance(value, str) and value.strip():
                key = normalize_label(value)
                labels.setdefault(key, []).append((used.row + r_idx, used.column + c_idx))

    return SheetSnapshot(
        start_row=used.row,
        start_col=used.column,
        values=values,
        labels=labels,
    )


def choose_bottom_right(coords: Sequence[Tuple[int, int]]) -> Optional[Tuple[int, int]]:
    if not coords:
        return None
    return max(coords, key=lambda x: (x[0], x[1]))


def find_label_exact(snapshot: SheetSnapshot, label: str) -> Optional[Tuple[int, int]]:
    coords = snapshot.labels.get(normalize_label(label), [])
    return choose_bottom_right(coords)


def find_label_contains(snapshot: SheetSnapshot, *terms: str) -> Optional[Tuple[int, int]]:
    normalized_terms = [normalize_label(term) for term in terms]
    candidates: List[Tuple[int, int]] = []
    for key, coords in snapshot.labels.items():
        if all(term in key for term in normalized_terms):
            candidates.extend(coords)
    return choose_bottom_right(candidates)


def find_max_anchor(snapshot: SheetSnapshot) -> Optional[Tuple[int, int]]:
    anchor = find_label_exact(snapshot, "max")
    if anchor is not None:
        return anchor
    return find_label_contains(snapshot, "max")


def nearby_numeric(
    snapshot: SheetSnapshot,
    row: int,
    col: int,
    offsets: Sequence[Tuple[int, int]],
) -> Optional[float]:
    for d_row, d_col in offsets:
        value = to_float(snapshot.get(row + d_row, col + d_col))
        if value is not None:
            return value
    return None


def value_next_to_label(snapshot: SheetSnapshot, *terms: str) -> Optional[float]:
    coord = find_label_contains(snapshot, *terms)
    if coord is None:
        return None
    return nearby_numeric(
        snapshot,
        coord[0],
        coord[1],
        offsets=((0, 1), (1, 0), (0, -1), (-1, 0), (1, 1), (-1, 1), (0, 2), (2, 0)),
    )


def collect_numeric_column(
    snapshot: SheetSnapshot,
    start_row: int,
    col: int,
    max_rows: int = 600,
) -> List[Tuple[int, float]]:
    series: List[Tuple[int, float]] = []
    blank_streak = 0
    stop_row = min(snapshot.last_row, start_row + max_rows)
    for row in range(start_row, stop_row + 1):
        raw_value = snapshot.get(row, col)
        value = to_float(raw_value)
        if value is None:
            if raw_value in (None, ""):
                blank_streak += 1
                if blank_streak >= 2 and series:
                    break
            elif series:
                break
            continue
        blank_streak = 0
        series.append((row, value))
    return series


def collect_numeric_series_near_header(
    snapshot: SheetSnapshot,
    header_coord: Optional[Tuple[int, int]],
) -> Tuple[List[Tuple[int, float]], Optional[int]]:
    if header_coord is None:
        return [], None
    header_row, header_col = header_coord
    for shift in (0, 1, -1, 2, -2):
        target_col = header_col + shift
        if target_col < 1:
            continue
        series = collect_numeric_column(snapshot, header_row + 1, target_col)
        if len(series) >= 2:
            return series, target_col
    return [], header_col


def collect_values_for_rows(
    snapshot: SheetSnapshot,
    rows: Sequence[int],
    col: Optional[int],
) -> Dict[int, Any]:
    if col is None:
        return {}
    return {row: snapshot.get(row, col) for row in rows}


def lookup_row_or_latest(value_map: Dict[int, Any], row: int) -> Any:
    if not value_map:
        return None
    if row in value_map and value_map[row] not in (None, ""):
        return value_map[row]
    valid_rows = sorted(r for r, value in value_map.items() if value not in (None, ""))
    if not valid_rows:
        return None
    prior_rows = [r for r in valid_rows if r <= row]
    chosen_row = prior_rows[-1] if prior_rows else valid_rows[-1]
    return value_map[chosen_row]


def parse_file_metadata(file_name: str) -> FileMetadata:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = parts[1].strip() if len(parts) > 1 and parts[1].strip() else "UNKNOWN"
    period_token = ""
    if len(parts) > 2:
        period_token = parts[2].split("_")[0].strip()

    period_match = re.search(r"(?i)(early|mid|late)\s*([a-z]{3,9})\s*(\d{4})", period_token)
    if period_match:
        phase = period_match.group(1).capitalize()
        month_raw = period_match.group(2)[:3].lower()
        year = period_match.group(3)
        month_num = MONTH_MAP.get(month_raw, 1)
        day = DAY_MAP[phase]
        month_title = month_raw.capitalize()
        model_period = f"{phase}{month_title}_{year}"
        model_date = f"{year}-{month_num:02d}-{day:02d}"
    else:
        cleaned_period = re.sub(r"[^A-Za-z0-9]", "", period_token) or "Unknown"
        model_period = f"{cleaned_period}_0000"
        model_date = ""

    model = f"{ticker}_{model_period}"
    return FileMetadata(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def set_formula2(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        # Fallback for older Excel APIs while still keeping R1C1 formula text.
        cell.formula = formula_r1c1


def close_workbook_safe(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception as close_error:
        print(f"  workbook close(save=False) fallback: {close_error}")

    try:
        wb.api.Close(SaveChanges=False)
    except Exception:
        try:
            wb.close()
        except Exception as final_error:
            print(f"  workbook close fallback failed: {final_error}")


def safe_signature(values: Sequence[Any]) -> Tuple[Any, ...]:
    normalized: List[Any] = []
    for value in values:
        value_float = to_float(value)
        if value_float is None:
            normalized.append(value)
        else:
            normalized.append(round(value_float, 10))
    return tuple(normalized)


def extract_empirical_rows(wb: xw.Book, metadata: FileMetadata, source_file: str) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"  skipped empirical: {source_file} (missing sheet 'Empirical Model')")
        return []

    snapshot = build_snapshot(sheet)
    if snapshot is None:
        print(f"  skipped empirical: {source_file} (empty or unreadable sheet)")
        return []

    max_anchor = find_max_anchor(snapshot)
    if max_anchor is None:
        print(f"  skipped empirical: {source_file} (missing 'max' anchor)")
        return []

    min_anchor = find_label_exact(snapshot, "min") or find_label_contains(snapshot, "min")

    # Anchor-based offsets from the max cell for max/min extraction.
    forecast_max = nearby_numeric(
        snapshot,
        max_anchor[0],
        max_anchor[1],
        offsets=((0, 1), (1, 0), (1, 1), (0, 2), (2, 0)),
    )
    forecast_min = None
    if min_anchor is not None:
        forecast_min = nearby_numeric(
            snapshot,
            min_anchor[0],
            min_anchor[1],
            offsets=((0, 1), (1, 0), (1, 1), (0, 2), (2, 0), (-1, 0)),
        )
    if forecast_min is None:
        forecast_min = nearby_numeric(
            snapshot,
            max_anchor[0],
            max_anchor[1],
            offsets=((1, 1), (2, 0), (0, 2), (2, 1)),
        )

    penetration_header = find_label_contains(snapshot, "penetration")
    penetration_series, penetration_col = collect_numeric_series_near_header(snapshot, penetration_header)
    if not penetration_series or penetration_col is None:
        print(f"  skipped empirical: {source_file} (no penetration history found)")
        return []

    quarterly_sales_header = find_label_contains(snapshot, "quarterly", "sales")
    quarterly_sales_series, quarterly_sales_col = collect_numeric_series_near_header(snapshot, quarterly_sales_header)
    quarterly_sales_map = {row: value for row, value in quarterly_sales_series}

    reported_sales_header = find_label_contains(snapshot, "reported", "sales")
    reported_sales_series, reported_sales_col = collect_numeric_series_near_header(snapshot, reported_sales_header)
    reported_sales_map = {row: value for row, value in reported_sales_series}

    growth_header = find_label_contains(snapshot, "growth", "rate")
    growth_series, growth_col = collect_numeric_series_near_header(snapshot, growth_header)
    growth_map = {row: value for row, value in growth_series}

    captured_header = find_label_contains(snapshot, "captured", "db") or find_label_contains(
        snapshot, "sales", "captured"
    )
    captured_series, captured_col = collect_numeric_series_near_header(snapshot, captured_header)
    captured_map = {row: value for row, value in captured_series}

    quarter_header = find_label_contains(snapshot, "quarter")
    quarter_col = quarter_header[1] if quarter_header is not None else penetration_col - 1

    reported_sales_scalar = value_next_to_label(snapshot, "reported", "sales")
    estimated_total_sold_scalar = value_next_to_label(snapshot, "estimated", "total", "sold")
    growth_scalar = value_next_to_label(snapshot, "growth", "rate")
    captured_scalar = value_next_to_label(snapshot, "captured", "db")

    rows: List[Dict[str, Any]] = []
    end_row = penetration_series[-1][0]
    n_limit = min(N_QUARTERS, len(penetration_series))
    if n_limit == 0:
        return rows

    scratch_start_row = snapshot.last_row + 5
    scratch_col = snapshot.last_col + 5

    start_rows: List[int] = []
    for idx in range(n_limit):
        n_quarters = idx + 1
        start_row = penetration_series[-n_quarters][0]
        start_rows.append(start_row)
        formula = f"=AVERAGE(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})"
        set_formula2(sheet.cells(scratch_start_row + idx, scratch_col), formula)

    wb.app.calculate()

    all_rows_for_lookup = [row for row, _ in penetration_series]
    quarter_values = collect_values_for_rows(snapshot, all_rows_for_lookup, quarter_col)

    for idx in range(n_limit):
        n_quarters = idx + 1
        start_row = start_rows[idx]
        avg_penetration = to_float(sheet.cells(scratch_start_row + idx, scratch_col).value)
        if avg_penetration is None:
            continue

        latest_reported_sales = to_float(lookup_row_or_latest(reported_sales_map, end_row))
        if latest_reported_sales is None:
            latest_reported_sales = reported_sales_scalar

        quarterly_sales = to_float(lookup_row_or_latest(quarterly_sales_map, start_row))
        reported_sales = latest_reported_sales
        actual_value = reported_sales

        if reported_sales is not None and avg_penetration not in (None, 0):
            forecast_value = reported_sales / avg_penetration
        else:
            forecast_value = estimated_total_sold_scalar

        growth_rate = to_float(lookup_row_or_latest(growth_map, end_row))
        if growth_rate is None:
            growth_rate = growth_scalar

        sales_captured = to_float(lookup_row_or_latest(captured_map, end_row))
        if sales_captured is None:
            sales_captured = captured_scalar

        last_quarter_used = quarter_values.get(start_row)
        if last_quarter_used in (None, ""):
            last_quarter_used = quarter_values.get(end_row)

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": n_quarters,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate,
                "sales_captured_in_db_pct": sales_captured,
                "source_file": source_file,
            }
        )

    return rows


def collect_xy_rows(snapshot: SheetSnapshot, x_col: int, y_col: int, stop_before_row: int) -> List[int]:
    numeric_rows: List[int] = []
    for row in range(snapshot.start_row, stop_before_row):
        x_val = to_float(snapshot.get(row, x_col))
        y_val = to_float(snapshot.get(row, y_col))
        if x_val is not None and y_val is not None:
            numeric_rows.append(row)
    if not numeric_rows:
        return []

    # Keep the latest contiguous block (allowing a one-row spacing gap).
    block = [numeric_rows[-1]]
    for row in reversed(numeric_rows[:-1]):
        if block[0] - row <= 2:
            block.insert(0, row)
        else:
            break
    return block


def extract_regression_rows(wb: xw.Book, metadata: FileMetadata, source_file: str) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"  skipped regression: {source_file} (missing sheet 'Regression Model')")
        return []

    snapshot = build_snapshot(sheet)
    if snapshot is None:
        print(f"  skipped regression: {source_file} (empty or unreadable sheet)")
        return []

    max_anchor = find_max_anchor(snapshot)
    if max_anchor is None:
        print(f"  skipped regression: {source_file} (missing 'max' anchor)")
        return []

    y_col = max_anchor[1] - 7
    x_col = max_anchor[1] - 11
    if x_col < 1 or y_col < 1:
        print(f"  skipped regression: {source_file} (invalid anchor offsets for x/y columns)")
        return []

    xy_rows = collect_xy_rows(snapshot, x_col=x_col, y_col=y_col, stop_before_row=max_anchor[0])
    if not xy_rows:
        print(f"  skipped regression: {source_file} (no x/y history in anchor-offset columns)")
        return []

    min_anchor = find_label_exact(snapshot, "min") or find_label_contains(snapshot, "min")
    forecast_max = nearby_numeric(
        snapshot,
        max_anchor[0],
        max_anchor[1],
        offsets=((0, 1), (1, 0), (1, 1), (0, 2), (2, 0)),
    )
    forecast_min = None
    if min_anchor is not None:
        forecast_min = nearby_numeric(
            snapshot,
            min_anchor[0],
            min_anchor[1],
            offsets=((0, 1), (1, 0), (1, 1), (0, 2), (2, 0), (-1, 0)),
        )
    if forecast_min is None:
        forecast_min = nearby_numeric(
            snapshot,
            max_anchor[0],
            max_anchor[1],
            offsets=((1, 1), (2, 0), (0, 2), (2, 1)),
        )

    actual_value = value_next_to_label(snapshot, "actual") or value_next_to_label(snapshot, "reported", "sales")

    rows: List[Dict[str, Any]] = []
    n_limit = min(N_QUARTERS, len(xy_rows))
    end_row = xy_rows[-1]
    scratch_start_row = snapshot.last_row + 5
    scratch_col = snapshot.last_col + 5

    start_rows: List[int] = []
    for idx in range(n_limit):
        n_quarters = idx + 1
        start_row = xy_rows[-n_quarters]
        start_rows.append(start_row)

        intercept_formula = f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        slope_formula = f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        forecast_formula = (
            f"=R{scratch_start_row + idx}C{scratch_col}"
            f"+R{scratch_start_row + idx}C{scratch_col + 1}*R{end_row}C{x_col}"
        )

        set_formula2(sheet.cells(scratch_start_row + idx, scratch_col), intercept_formula)
        set_formula2(sheet.cells(scratch_start_row + idx, scratch_col + 1), slope_formula)
        set_formula2(sheet.cells(scratch_start_row + idx, scratch_col + 2), forecast_formula)

    wb.app.calculate()

    previous_signature: Optional[Tuple[Any, ...]] = None
    for idx in range(n_limit):
        n_quarters = idx + 1
        intercept_value = to_float(sheet.cells(scratch_start_row + idx, scratch_col).value)
        slope_value = to_float(sheet.cells(scratch_start_row + idx, scratch_col + 1).value)
        forecast_total_without_sa = to_float(sheet.cells(scratch_start_row + idx, scratch_col + 2).value)

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        signature = safe_signature(
            (
                intercept_value,
                slope_value,
                forecast_total_without_sa,
                forecast_max,
                forecast_min,
            )
        )
        if signature == previous_signature:
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
                "parameter_value": n_quarters,
                "num_quarters_used": n_quarters,
                "forecast_value": forecast_total_without_sa,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept_value,
                "slope": slope_value,
                "source_file": source_file,
            }
        )

    return rows


def next_output_path(input_path: Path, output_path: Path) -> Path:
    base_name = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    version = 1
    while candidate.exists():
        candidate = output_path / f"{base_name}.{version}.xlsx"
        version += 1
    return candidate


def write_sheet(ws, headers: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(headers))
    for row in rows:
        ws.append([row.get(header) for header in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(idx)].width = min(max_len + 2, 50)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    ws_empirical = wb.create_sheet("empirical_candidates")
    ws_regression = wb.create_sheet("regression_candidates")

    write_sheet(ws_empirical, EMPIRICAL_HEADERS, empirical_rows)
    write_sheet(ws_regression, REGRESSION_HEADERS, regression_rows)
    wb.save(output_path)


def skip_reason(path: Path) -> Optional[str]:
    if not path.is_file():
        return "not a file"
    if path.name.startswith("~"):
        return "temporary lock file"
    if path.suffix.lower() != ".xlsx":
        return "not an .xlsx file"
    return None


def main() -> None:
    in_path = Path(input_dir).expanduser().resolve()
    out_path = Path(output_dir).expanduser().resolve()

    if not in_path.exists():
        raise SystemExit(f"Input folder not found: {in_path}")

    out_path.mkdir(parents=True, exist_ok=True)
    output_path = next_output_path(in_path, out_path)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_count = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(in_path.iterdir(), key=lambda p: p.name.lower()):
            reason = skip_reason(file_path)
            if reason is not None:
                print(f"skipped: {file_path.name} ({reason})")
                continue

            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                metadata = parse_file_metadata(file_path.name)

                empirical_rows.extend(extract_empirical_rows(wb, metadata, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, metadata, file_path.name))
                processed_count += 1
                print(f"processed: {file_path.name}")
            except Exception as process_error:
                print(f"skipped: {file_path.name} (processing error: {process_error})")
            finally:
                if wb is not None:
                    close_workbook_safe(wb)
    finally:
        try:
            app.quit()
        except Exception:
            pass

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"output_path: {output_path}")
    print(f"files_processed: {processed_count}")
    print(f"empirical_rows: {len(empirical_rows)}")
    print(f"regression_rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
from __future__ import annotations

import math
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure these two paths before running.
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")

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

MONTH_TO_NUM = {
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

DAY_BY_PERIOD = {"early": 5, "mid": 15, "late": 25}


@dataclass
class SheetCache:
    sheet: Any
    start_row: int
    start_col: int
    values: List[List[Any]]

    @property
    def end_row(self) -> int:
        return self.start_row + len(self.values) - 1

    @property
    def end_col(self) -> int:
        if not self.values:
            return self.start_col
        max_cols = max((len(row) for row in self.values), default=1)
        return self.start_col + max_cols - 1

    def get(self, row: int, col: int) -> Any:
        row_idx = row - self.start_row
        col_idx = col - self.start_col
        if row_idx < 0 or row_idx >= len(self.values):
            return None
        row_values = self.values[row_idx]
        if col_idx < 0 or col_idx >= len(row_values):
            return None
        return row_values[col_idx]


def normalize_matrix(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def build_sheet_cache(sheet: Any) -> SheetCache:
    used = sheet.used_range
    values = normalize_matrix(used.value)
    return SheetCache(sheet=sheet, start_row=used.row, start_col=used.column, values=values)


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
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "")
        if cleaned == "":
            return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def as_int(value: Any) -> Optional[int]:
    numeric = as_float(value)
    if numeric is None:
        return None
    return int(round(numeric))


def is_number(value: Any) -> bool:
    return as_float(value) is not None


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def matches_alias(value: Any, aliases: Sequence[str]) -> bool:
    normalized_value = normalize_label(value)
    if not normalized_value:
        return False
    for alias in aliases:
        alias_key = normalize_label(alias)
        if alias_key and alias_key in normalized_value:
            return True
    return False


def find_nearest_cell(
    cache: SheetCache,
    aliases: Sequence[str],
    row_min: int,
    row_max: int,
    col_min: int,
    col_max: int,
    anchor: Optional[Tuple[int, int]] = None,
) -> Optional[Tuple[int, int]]:
    best_position: Optional[Tuple[int, int]] = None
    best_distance: Optional[int] = None
    row_min = max(row_min, cache.start_row)
    row_max = min(row_max, cache.end_row)
    col_min = max(col_min, cache.start_col)
    col_max = min(col_max, cache.end_col)
    if row_min > row_max or col_min > col_max:
        return None

    for row in range(row_min, row_max + 1):
        for col in range(col_min, col_max + 1):
            if not matches_alias(cache.get(row, col), aliases):
                continue
            if anchor is None:
                return row, col
            distance = abs(row - anchor[0]) + abs(col - anchor[1])
            if best_distance is None or distance < best_distance:
                best_distance = distance
                best_position = (row, col)
    return best_position


def find_anchor_max(cache: SheetCache) -> Optional[Tuple[int, int]]:
    return find_nearest_cell(
        cache=cache,
        aliases=["max"],
        row_min=cache.start_row,
        row_max=cache.end_row,
        col_min=cache.start_col,
        col_max=cache.end_col,
        anchor=(cache.start_row, cache.start_col),
    )


def resolve_column_from_anchor(
    cache: SheetCache,
    anchor_row: int,
    anchor_col: int,
    aliases: Sequence[str],
    default_offset: int,
) -> int:
    found = find_nearest_cell(
        cache=cache,
        aliases=aliases,
        row_min=anchor_row - 4,
        row_max=anchor_row + 4,
        col_min=anchor_col - 45,
        col_max=anchor_col + 20,
        anchor=(anchor_row, anchor_col),
    )
    if found:
        return found[1]
    return anchor_col + default_offset


def find_first_data_row(cache: SheetCache, start_row: int, columns: Sequence[int]) -> int:
    probe_start = max(start_row, cache.start_row)
    for row in range(probe_start, cache.end_row + 1):
        for col in columns:
            if is_number(cache.get(row, col)):
                return row
    return probe_start


def collect_candidate_rows(
    cache: SheetCache,
    start_row: int,
    key_columns: Sequence[int],
    max_rows: int = 10,
) -> List[Tuple[int, int]]:
    rows: List[Tuple[int, int]] = []
    blank_streak = 0
    for n_quarters in range(1, max_rows + 1):
        row = start_row + n_quarters - 1
        if row > cache.end_row:
            break
        has_signal = False
        for col in key_columns:
            if not is_blank(cache.get(row, col)):
                has_signal = True
                break
        if has_signal:
            blank_streak = 0
            rows.append((n_quarters, row))
        else:
            blank_streak += 1
            if rows and blank_streak >= 2:
                break
    return rows


def set_formula2(cell: Any, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def parse_filename_metadata(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    ticker_match = re.search(r"-\s*([A-Za-z0-9]+)\s*-\s*", stem)
    ticker = ticker_match.group(1).upper() if ticker_match else "UNKNOWN"

    period_match = re.search(r"\b(Early|Mid|Late)([A-Za-z]{3})(\d{4})\b", stem, flags=re.IGNORECASE)
    model_period = ""
    model_date = ""

    if period_match:
        period_label = period_match.group(1).title()
        month_abbr = period_match.group(2).title()
        month_key = month_abbr.lower()
        year = period_match.group(3)
        month_num = MONTH_TO_NUM.get(month_key)
        if month_num is not None:
            model_period = f"{period_label}{month_abbr}_{year}"
            day = DAY_BY_PERIOD[period_label.lower()]
            model_date = f"{year}-{month_num:02d}-{day:02d}"

    model = f"{ticker}_{model_period}" if model_period else ticker
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def close_workbook_without_saving(wb: Any) -> None:
    last_error: Optional[Exception] = None
    close_attempts = [
        lambda: wb.close(save=False),
        lambda: wb.close(False),
        lambda: wb.api.Close(SaveChanges=False),
        lambda: wb.close(),
    ]
    for close_attempt in close_attempts:
        try:
            close_attempt()
            return
        except Exception as exc:  # pragma: no cover - defensive COM fallback
            last_error = exc
    if last_error is not None:
        raise last_error


def safe_division(left: Any, right: Any) -> Optional[float]:
    left_value = as_float(left)
    right_value = as_float(right)
    if left_value is None or right_value is None or right_value == 0:
        return None
    return left_value / right_value


def round_key(value: Any) -> Any:
    numeric = as_float(value)
    if numeric is None:
        return value
    return round(numeric, 10)


def extract_empirical_rows(
    wb: Any,
    metadata: Dict[str, str],
    source_file: Path,
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"  Skipped empirical extraction ({source_file.name}): missing 'Empirical Model' sheet")
        return []

    cache = build_sheet_cache(sheet)
    anchor = find_anchor_max(cache)
    if anchor is None:
        print(f"  Skipped empirical extraction ({source_file.name}): could not find 'max' anchor")
        return []

    anchor_row, anchor_col = anchor

    col_map = {
        "num_quarters_used": resolve_column_from_anchor(
            cache, anchor_row, anchor_col, ["num quarters", "quarters used", "# qtrs"], -10
        ),
        "last_quarter_used": resolve_column_from_anchor(
            cache, anchor_row, anchor_col, ["last quarter used", "last qtr used", "last quarter"], -9
        ),
        "forecast_value": resolve_column_from_anchor(
            cache, anchor_row, anchor_col, ["estimated total sold", "tot fcst", "forecast"], -2
        ),
        "actual_value": resolve_column_from_anchor(
            cache, anchor_row, anchor_col, ["reported sales", "actual value", "actual"], -4
        ),
        "forecast_max": anchor_col,
        "forecast_min": resolve_column_from_anchor(cache, anchor_row, anchor_col, ["min"], 1),
        "avg_penetration_pct": resolve_column_from_anchor(
            cache, anchor_row, anchor_col, ["avg penetration", "average penetration"], -6
        ),
        "quarterly_sales": resolve_column_from_anchor(
            cache, anchor_row, anchor_col, ["quarterly sales", "qtr sales"], -5
        ),
        "reported_sales": resolve_column_from_anchor(
            cache, anchor_row, anchor_col, ["reported sales", "reported"], -4
        ),
        "growth_rate_pct": resolve_column_from_anchor(
            cache, anchor_row, anchor_col, ["growth rate", "growth %"], -3
        ),
        "sales_captured_in_db_pct": resolve_column_from_anchor(
            cache, anchor_row, anchor_col, ["captured in db", "sales captured in db", "db %"], -7
        ),
        "penetration_source": resolve_column_from_anchor(
            cache, anchor_row, anchor_col, ["penetration"], -6
        ),
    }

    data_start_row = find_first_data_row(
        cache,
        anchor_row + 1,
        [
            col_map["num_quarters_used"],
            col_map["forecast_value"],
            col_map["forecast_max"],
            col_map["forecast_min"],
        ],
    )

    candidate_rows = collect_candidate_rows(
        cache,
        start_row=data_start_row,
        key_columns=[
            col_map["num_quarters_used"],
            col_map["forecast_value"],
            col_map["forecast_max"],
            col_map["forecast_min"],
        ],
        max_rows=10,
    )
    if not candidate_rows:
        return []

    helper_col = cache.end_col + 2
    for n_quarters, row in candidate_rows:
        start = max(data_start_row, row - n_quarters + 1)
        formula = (
            f"=AVERAGE(R{start}C{col_map['penetration_source']}:"
            f"R{row}C{col_map['penetration_source']})"
        )
        set_formula2(sheet.range((row, helper_col)), formula)

    try:
        wb.app.calculate()
    except Exception:
        pass

    avg_penetration_by_row: Dict[int, Optional[float]] = {}
    for _, row in candidate_rows:
        avg_penetration_by_row[row] = as_float(sheet.range((row, helper_col)).value)

    rows: List[Dict[str, Any]] = []
    for n_quarters, row in candidate_rows:
        forecast_value = cache.get(row, col_map["forecast_value"])
        forecast_max = cache.get(row, col_map["forecast_max"])
        forecast_min = cache.get(row, col_map["forecast_min"])
        signal_values = [forecast_value, forecast_max, forecast_min]
        if all(is_blank(value) for value in signal_values):
            continue

        num_quarters_used = as_int(cache.get(row, col_map["num_quarters_used"])) or n_quarters
        actual_value = cache.get(row, col_map["actual_value"])
        reported_sales = cache.get(row, col_map["reported_sales"])

        avg_penetration = avg_penetration_by_row.get(row)
        if avg_penetration is None:
            avg_penetration = as_float(cache.get(row, col_map["avg_penetration_pct"]))
        if avg_penetration is None:
            avg_penetration = safe_division(
                cache.get(row, col_map["quarterly_sales"]),
                reported_sales,
            )

        forecast_max_f = as_float(forecast_max)
        forecast_min_f = as_float(forecast_min)
        range_width = (
            forecast_max_f - forecast_min_f
            if forecast_max_f is not None and forecast_min_f is not None
            else None
        )

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": cache.get(row, col_map["last_quarter_used"]),
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": cache.get(row, col_map["quarterly_sales"]),
                "reported_sales": reported_sales,
                "growth_rate_pct": cache.get(row, col_map["growth_rate_pct"]),
                "sales_captured_in_db_pct": cache.get(row, col_map["sales_captured_in_db_pct"]),
                "source_file": source_file.name,
            }
        )
    return rows


def extract_regression_rows(
    wb: Any,
    metadata: Dict[str, str],
    source_file: Path,
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"  Skipped regression extraction ({source_file.name}): missing 'Regression Model' sheet")
        return []

    cache = build_sheet_cache(sheet)
    anchor = find_anchor_max(cache)
    if anchor is None:
        print(f"  Skipped regression extraction ({source_file.name}): could not find 'max' anchor")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    col_map = {
        "num_quarters_used": resolve_column_from_anchor(
            cache, anchor_row, anchor_col, ["num quarters", "quarters used", "# qtrs"], -8
        ),
        "forecast_value": resolve_column_from_anchor(
            cache, anchor_row, anchor_col, ["tot fcst w/o sa", "tot fcst wo sa", "forecast"], -2
        ),
        "actual_value": resolve_column_from_anchor(
            cache, anchor_row, anchor_col, ["actual", "reported sales"], -4
        ),
        "forecast_max": anchor_col,
        "forecast_min": resolve_column_from_anchor(cache, anchor_row, anchor_col, ["min"], 1),
    }

    data_start_row = find_first_data_row(
        cache,
        anchor_row + 1,
        [
            col_map["num_quarters_used"],
            col_map["forecast_value"],
            col_map["forecast_max"],
            col_map["forecast_min"],
        ],
    )
    candidate_rows = collect_candidate_rows(
        cache,
        start_row=data_start_row,
        key_columns=[
            col_map["num_quarters_used"],
            col_map["forecast_value"],
            col_map["forecast_max"],
            col_map["forecast_min"],
        ],
        max_rows=10,
    )
    if not candidate_rows:
        return []

    xy_rows = [
        row
        for row in range(cache.start_row, anchor_row + 1)
        if is_number(cache.get(row, x_col)) and is_number(cache.get(row, y_col))
    ]

    helper_intercept_col = cache.end_col + 2
    helper_slope_col = cache.end_col + 3
    formula_rows: List[int] = []

    for n_quarters, row in candidate_rows:
        declared_quarters = as_int(cache.get(row, col_map["num_quarters_used"])) or n_quarters
        if len(xy_rows) < 2:
            continue
        span = max(2, min(declared_quarters, len(xy_rows)))
        selected_rows = xy_rows[-span:]
        start = selected_rows[0]
        end = selected_rows[-1]
        intercept_formula = f"=INTERCEPT(R{start}C{y_col}:R{end}C{y_col},R{start}C{x_col}:R{end}C{x_col})"
        slope_formula = f"=SLOPE(R{start}C{y_col}:R{end}C{y_col},R{start}C{x_col}:R{end}C{x_col})"
        set_formula2(sheet.range((row, helper_intercept_col)), intercept_formula)
        set_formula2(sheet.range((row, helper_slope_col)), slope_formula)
        formula_rows.append(row)

    if formula_rows:
        try:
            wb.app.calculate()
        except Exception:
            pass

    intercept_by_row: Dict[int, Optional[float]] = {}
    slope_by_row: Dict[int, Optional[float]] = {}
    for row in formula_rows:
        intercept_by_row[row] = as_float(sheet.range((row, helper_intercept_col)).value)
        slope_by_row[row] = as_float(sheet.range((row, helper_slope_col)).value)

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for n_quarters, row in candidate_rows:
        forecast_value = cache.get(row, col_map["forecast_value"])
        forecast_max = cache.get(row, col_map["forecast_max"])
        forecast_min = cache.get(row, col_map["forecast_min"])
        signal_values = [forecast_value, forecast_max, forecast_min]
        if all(is_blank(value) for value in signal_values):
            continue

        num_quarters_used = as_int(cache.get(row, col_map["num_quarters_used"])) or n_quarters
        intercept = intercept_by_row.get(row)
        slope = slope_by_row.get(row)
        forecast_max_f = as_float(forecast_max)
        forecast_min_f = as_float(forecast_min)
        range_width = (
            forecast_max_f - forecast_min_f
            if forecast_max_f is not None and forecast_min_f is not None
            else None
        )

        signature = (
            round_key(num_quarters_used),
            round_key(intercept),
            round_key(slope),
            round_key(forecast_value),
            round_key(forecast_max),
            round_key(forecast_min),
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": cache.get(row, col_map["actual_value"]) or "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file.name,
            }
        )
    return rows


def ensure_unique_output_path(input_folder: Path, destination_folder: Path) -> Path:
    base_name = f"{input_folder.name}_PARAM.xlsx"
    candidate = destination_folder / base_name
    if not candidate.exists():
        return candidate
    counter = 1
    while True:
        candidate = destination_folder / f"{input_folder.name}_PARAM.{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


def write_sheet(
    workbook: Workbook,
    sheet_name: str,
    headers: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    ws = workbook.create_sheet(title=sheet_name)
    ws.append(list(headers))
    for row in rows:
        ws.append([row.get(header) for header in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, header in enumerate(headers, start=1):
        max_length = len(header)
        preview_limit = min(ws.max_row, 400)
        for row_idx in range(2, preview_limit + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_length = max(max_length, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_length + 2, 50)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)
    write_sheet(workbook, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(workbook, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    workbook.save(output_path)


def iter_input_files(folder: Path) -> Iterable[Path]:
    for item in sorted(folder.iterdir()):
        if item.is_file():
            yield item


def process_workbooks() -> None:
    source_folder = Path(input_dir)
    destination_folder = Path(output_dir)

    if not source_folder.exists():
        raise FileNotFoundError(f"input_dir does not exist: {source_folder}")
    destination_folder.mkdir(parents=True, exist_ok=True)

    output_path = ensure_unique_output_path(source_folder, destination_folder)

    processed_files = 0
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in iter_input_files(source_folder):
            if file_path.name.startswith("~"):
                print(f"Skipped file: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped file: {file_path.name} (not .xlsx)")
                continue

            wb = None
            try:
                print(f"Processing file: {file_path.name}")
                wb = app.books.open(str(file_path), update_links=False)
                metadata = parse_filename_metadata(file_path)

                empirical_rows.extend(extract_empirical_rows(wb, metadata, file_path))
                regression_rows.extend(extract_regression_rows(wb, metadata, file_path))

                processed_files += 1
                print(f"Processed file: {file_path.name}")
            except Exception as exc:
                print(f"Skipped file: {file_path.name} (error: {exc})")
            finally:
                if wb is not None:
                    try:
                        close_workbook_without_saving(wb)
                    except Exception as close_exc:
                        print(f"Warning: failed to close {file_path.name} cleanly ({close_exc})")
    finally:
        try:
            app.quit()
        except Exception:
            pass

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    process_workbooks()

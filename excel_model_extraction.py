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

# ==========================
# User-configurable settings
# ==========================
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")

N_QUARTERS = 10
EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"

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

MONTH_ABBR_TO_NUM = {
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

DAY_BY_PERIOD_PREFIX = {
    "early": 5,
    "mid": 15,
    "late": 25,
}


@dataclass(frozen=True)
class FileModelMeta:
    ticker: str
    model_period: str
    model_date: str
    model: str


@dataclass
class SheetCache:
    values: List[List[Any]]
    last_row: int
    last_col: int
    labels: Dict[str, List[Tuple[int, int]]]


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"\s+", " ", text)
    return text


def to_2d_list(value: Any) -> List[List[Any]]:
    if value is None:
        return []
    if isinstance(value, list):
        if not value:
            return []
        if isinstance(value[0], list):
            return value
        return [value]
    return [[value]]


def read_sheet_cache(sheet: xw.Sheet) -> SheetCache:
    used = sheet.used_range
    values = to_2d_list(used.value)
    if not values:
        return SheetCache(values=[], last_row=0, last_col=0, labels={})

    last_row = len(values)
    last_col = max((len(row) for row in values), default=0)
    labels: Dict[str, List[Tuple[int, int]]] = {}

    for r_idx, row in enumerate(values, start=1):
        for c_idx, cell_value in enumerate(row, start=1):
            key = normalize_label(cell_value)
            if key:
                labels.setdefault(key, []).append((r_idx, c_idx))

    return SheetCache(values=values, last_row=last_row, last_col=last_col, labels=labels)


def cache_value(cache: SheetCache, row: int, col: int) -> Any:
    if row <= 0 or col <= 0:
        return None
    if row > cache.last_row:
        return None
    row_values = cache.values[row - 1]
    if col > len(row_values):
        return None
    return row_values[col - 1]


def is_number(value: Any) -> bool:
    if value is None:
        return False
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return not math.isnan(float(value))
    return False


def as_float(value: Any) -> Optional[float]:
    if is_number(value):
        return float(value)
    if isinstance(value, str):
        stripped = value.strip().replace(",", "")
        if stripped.endswith("%"):
            stripped = stripped[:-1]
            try:
                return float(stripped) / 100.0
            except ValueError:
                return None
        try:
            return float(stripped)
        except ValueError:
            return None
    return None


def first_existing_sheet(wb: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    for sheet in wb.sheets:
        if normalize_label(sheet.name) == normalize_label(sheet_name):
            return sheet
    return None


def find_anchor_max(cache: SheetCache) -> Optional[Tuple[int, int]]:
    max_positions = cache.labels.get("max", [])
    if not max_positions:
        return None
    min_positions = cache.labels.get("min", [])
    if not min_positions:
        return max_positions[0]

    def score(pos: Tuple[int, int]) -> int:
        row, col = pos
        best = min(abs(row - mr) + abs(col - mc) for mr, mc in min_positions)
        return best

    return min(max_positions, key=score)


def find_label_positions(cache: SheetCache, label_variants: Sequence[str]) -> List[Tuple[int, int]]:
    candidates: List[Tuple[int, int]] = []
    for key, positions in cache.labels.items():
        if any(variant in key for variant in label_variants):
            candidates.extend(positions)
    return candidates


def nearest_position(
    positions: Sequence[Tuple[int, int]],
    anchor_row: int,
    anchor_col: int,
    max_row_distance: int = 40,
    max_col_distance: int = 30,
) -> Optional[Tuple[int, int]]:
    best: Optional[Tuple[int, int]] = None
    best_distance: Optional[int] = None
    for row, col in positions:
        if abs(row - anchor_row) > max_row_distance:
            continue
        if abs(col - anchor_col) > max_col_distance:
            continue
        distance = abs(row - anchor_row) + abs(col - anchor_col)
        if best is None or distance < (best_distance or 10**9):
            best = (row, col)
            best_distance = distance
    return best


def value_next_to_label(cache: SheetCache, label_row: int, label_col: int) -> Any:
    offsets = [(0, 1), (1, 0), (0, -1), (-1, 0), (1, 1), (-1, 1)]
    for dr, dc in offsets:
        value = cache_value(cache, label_row + dr, label_col + dc)
        if value is not None and str(value).strip() != "":
            return value
    return None


def find_value_by_label_near_anchor(
    cache: SheetCache,
    anchor_row: int,
    anchor_col: int,
    label_variants: Sequence[str],
) -> Any:
    positions = find_label_positions(cache, label_variants)
    nearest = nearest_position(positions, anchor_row, anchor_col)
    if not nearest:
        return None
    return value_next_to_label(cache, nearest[0], nearest[1])


def find_column_with_density(
    cache: SheetCache,
    start_col: int,
    end_col: int,
    max_row: int,
    value_predicate,
) -> Optional[int]:
    best_col: Optional[int] = None
    best_count = 0
    start_col = max(start_col, 1)
    end_col = min(end_col, cache.last_col)
    for col in range(start_col, end_col + 1):
        count = 0
        for row in range(1, max_row + 1):
            value = cache_value(cache, row, col)
            if value_predicate(value):
                count += 1
        if count > best_count:
            best_count = count
            best_col = col
    if best_count == 0:
        return None
    return best_col


def collect_numeric_pairs(
    cache: SheetCache, x_col: int, y_col: int, max_row: int
) -> List[Tuple[int, float, float]]:
    pairs: List[Tuple[int, float, float]] = []
    for row in range(1, max_row + 1):
        x_value = as_float(cache_value(cache, row, x_col))
        y_value = as_float(cache_value(cache, row, y_col))
        if x_value is None or y_value is None:
            continue
        pairs.append((row, x_value, y_value))
    return pairs


def set_formula2(target_range: xw.Range, formula: str) -> None:
    try:
        target_range.formula2 = formula
    except Exception:
        target_range.formula = formula


def safe_close_source_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception as exc:
        message = str(exc).lower()
        if "save" not in message and "keyword" not in message:
            raise
    try:
        wb.api.Close(SaveChanges=False)
    except Exception:
        wb.close()


def parse_file_model_meta(file_name: str) -> Optional[FileModelMeta]:
    stem = Path(file_name).stem
    # Handles names like:
    # "MedMiner_Model - AORT - MidJan2026_Send"
    # while allowing case-insensitive "_send" suffix.
    match = re.search(
        r"-\s*([A-Za-z0-9]+)\s*-\s*(Early|Mid|Late)([A-Za-z]{3})(\d{4})(?:_send)?$",
        stem,
        flags=re.IGNORECASE,
    )
    if not match:
        return None

    ticker, prefix_raw, month_abbr_raw, year_raw = match.groups()
    prefix = prefix_raw.capitalize()
    month_abbr = month_abbr_raw.capitalize()
    year_int = int(year_raw)

    month_num = MONTH_ABBR_TO_NUM.get(month_abbr.lower())
    day_num = DAY_BY_PERIOD_PREFIX.get(prefix.lower())
    if not month_num or not day_num:
        return None

    model_period = f"{prefix}{month_abbr}_{year_int}"
    model_date = date(year_int, month_num, day_num).isoformat()
    model = f"{ticker}_{model_period}"

    return FileModelMeta(
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
        model=model,
    )


def output_path_with_suffix(input_dir_path: Path, output_dir_path: Path) -> Path:
    base_name = f"{input_dir_path.name}_PARAM"
    candidate = output_dir_path / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate
    suffix = 1
    while True:
        candidate = output_dir_path / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def iter_source_files(input_dir_path: Path) -> Iterable[Path]:
    for file_path in sorted(input_dir_path.iterdir(), key=lambda p: p.name.lower()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped {file_path.name}: temporary file")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped {file_path.name}: not an .xlsx file")
            continue
        yield file_path


def compute_column_widths(rows: List[Dict[str, Any]], columns: Sequence[str]) -> Dict[str, int]:
    widths: Dict[str, int] = {}
    for col in columns:
        max_len = len(col)
        for row in rows:
            value = row.get(col)
            if value is None:
                continue
            text = str(value)
            if len(text) > max_len:
                max_len = len(text)
        widths[col] = min(max(max_len + 2, 12), 48)
    return widths


def write_output_workbook(
    output_file_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    def write_sheet(sheet_name: str, columns: Sequence[str], rows: List[Dict[str, Any]]) -> None:
        ws = wb.create_sheet(title=sheet_name)
        ws.append(list(columns))
        for row in rows:
            ws.append([row.get(col) for col in columns])

        for cell in ws[1]:
            cell.font = Font(bold=True)

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        widths = compute_column_widths(rows, columns)
        for idx, col_name in enumerate(columns, start=1):
            ws.column_dimensions[get_column_letter(idx)].width = widths[col_name]

    write_sheet("empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet("regression_candidates", REGRESSION_COLUMNS, regression_rows)
    output_file_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_file_path)


def approx_equal(a: Optional[float], b: Optional[float], tolerance: float = 1e-9) -> bool:
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    return abs(a - b) <= tolerance


def parse_penetration(value: Any) -> Optional[float]:
    number = as_float(value)
    if number is None:
        return None
    if number > 1.5:
        return None
    if number < 0:
        return None
    return number


def infer_empirical_support_columns(
    cache: SheetCache, anchor_row: int, anchor_col: int
) -> Tuple[Optional[int], Optional[int], Optional[int]]:
    penetration_col = find_column_with_density(
        cache=cache,
        start_col=anchor_col - 20,
        end_col=anchor_col + 2,
        max_row=max(anchor_row - 1, 1),
        value_predicate=lambda v: parse_penetration(v) is not None,
    )

    quarterly_sales_col = find_column_with_density(
        cache=cache,
        start_col=anchor_col - 20,
        end_col=anchor_col + 2,
        max_row=max(anchor_row - 1, 1),
        value_predicate=lambda v: (as_float(v) is not None) and (as_float(v) or 0.0) > 1.5,
    )

    quarter_label_col = find_column_with_density(
        cache=cache,
        start_col=anchor_col - 24,
        end_col=anchor_col,
        max_row=max(anchor_row - 1, 1),
        value_predicate=lambda v: isinstance(v, str) and bool(v.strip()),
    )

    return penetration_col, quarterly_sales_col, quarter_label_col


def extract_empirical_rows(
    wb: xw.Book,
    model_meta: FileModelMeta,
    source_file_name: str,
) -> List[Dict[str, Any]]:
    sheet = first_existing_sheet(wb, EMPIRICAL_SHEET_NAME)
    if not sheet:
        return []

    cache = read_sheet_cache(sheet)
    anchor = find_anchor_max(cache)
    if not anchor:
        return []
    anchor_row, anchor_col = anchor

    penetration_col, quarterly_sales_col, quarter_label_col = infer_empirical_support_columns(
        cache=cache,
        anchor_row=anchor_row,
        anchor_col=anchor_col,
    )

    if penetration_col is None:
        return []

    penetration_series: List[Tuple[int, float]] = []
    for row in range(1, anchor_row):
        value = parse_penetration(cache_value(cache, row, penetration_col))
        if value is None:
            continue
        penetration_series.append((row, value))

    if not penetration_series:
        return []

    helper_row = cache.last_row + 3
    helper_col = cache.last_col + 3
    avg_helper = sheet.range((helper_row, helper_col))

    reported_sales = as_float(
        find_value_by_label_near_anchor(
            cache,
            anchor_row,
            anchor_col,
            ["reported sales", "actual", "actual sales"],
        )
    )
    quarterly_sales_labeled = as_float(
        find_value_by_label_near_anchor(
            cache,
            anchor_row,
            anchor_col,
            ["quarterly sales", "sales in db", "captured sales"],
        )
    )
    growth_rate_labeled = as_float(
        find_value_by_label_near_anchor(
            cache,
            anchor_row,
            anchor_col,
            ["growth", "growth rate"],
        )
    )
    sales_captured_labeled = as_float(
        find_value_by_label_near_anchor(
            cache,
            anchor_row,
            anchor_col,
            ["captured in db", "captured %", "sales captured"],
        )
    )

    all_quarterly_sales: List[Tuple[int, float]] = []
    if quarterly_sales_col is not None:
        for row in range(1, anchor_row):
            value = as_float(cache_value(cache, row, quarterly_sales_col))
            if value is None:
                continue
            all_quarterly_sales.append((row, value))

    rows: List[Dict[str, Any]] = []

    for requested_n in range(1, N_QUARTERS + 1):
        effective_n = min(requested_n, len(penetration_series))
        selected_pen = penetration_series[-effective_n:]
        start_row = selected_pen[0][0]
        end_row = selected_pen[-1][0]

        # Keep formula evaluation in Excel with R1C1 references for parity with model logic.
        set_formula2(avg_helper, f"=AVERAGE(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})")
        wb.app.calculate()
        avg_penetration = as_float(avg_helper.value)

        if avg_penetration is None or avg_penetration == 0:
            continue

        if quarterly_sales_labeled is not None:
            quarterly_sales = quarterly_sales_labeled
        else:
            quarterly_sales = None
            if all_quarterly_sales:
                lookup = {row_idx: val for row_idx, val in all_quarterly_sales}
                quarterly_sales = lookup.get(end_row)
                if quarterly_sales is None:
                    quarterly_sales = all_quarterly_sales[-1][1]

        if quarterly_sales is None:
            continue

        forecast_value = quarterly_sales / avg_penetration if avg_penetration else None
        max_pen = max(val for _, val in selected_pen)
        min_pen = min(val for _, val in selected_pen)
        forecast_max = quarterly_sales / min_pen if min_pen else None
        forecast_min = quarterly_sales / max_pen if max_pen else None
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        last_quarter_used: Any = end_row
        if quarter_label_col is not None:
            q_label = cache_value(cache, end_row, quarter_label_col)
            if q_label is not None and str(q_label).strip():
                last_quarter_used = q_label

        growth_rate_pct = growth_rate_labeled
        if growth_rate_pct is None and len(all_quarterly_sales) >= 2:
            prev_sales = all_quarterly_sales[-2][1]
            curr_sales = all_quarterly_sales[-1][1]
            if prev_sales:
                growth_rate_pct = (curr_sales - prev_sales) / prev_sales

        sales_captured_pct = sales_captured_labeled
        if sales_captured_pct is None and reported_sales not in (None, 0):
            sales_captured_pct = quarterly_sales / reported_sales

        rows.append(
            {
                "model": model_meta.model,
                "ticker": model_meta.ticker,
                "model_period": model_meta.model_period,
                "model_date": model_meta.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": effective_n,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_pct,
                "source_file": source_file_name,
            }
        )

    avg_helper.clear_contents()
    return rows


def extract_regression_rows(
    wb: xw.Book,
    model_meta: FileModelMeta,
    source_file_name: str,
) -> List[Dict[str, Any]]:
    sheet = first_existing_sheet(wb, REGRESSION_SHEET_NAME)
    if not sheet:
        return []

    cache = read_sheet_cache(sheet)
    anchor = find_anchor_max(cache)
    if not anchor:
        return []
    anchor_row, anchor_col = anchor

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if x_col <= 0 or y_col <= 0:
        return []

    pairs = collect_numeric_pairs(cache, x_col=x_col, y_col=y_col, max_row=max(anchor_row - 1, 1))
    if len(pairs) < 2:
        return []

    helper_row = cache.last_row + 3
    helper_col = cache.last_col + 3
    intercept_cell = sheet.range((helper_row, helper_col))
    slope_cell = sheet.range((helper_row, helper_col + 1))

    actual_value = as_float(
        find_value_by_label_near_anchor(
            cache,
            anchor_row,
            anchor_col,
            ["actual", "reported sales"],
        )
    )

    tot_fcst_wo_sa = as_float(
        find_value_by_label_near_anchor(
            cache,
            anchor_row,
            anchor_col,
            ["tot fcst w/o sa", "tot fcst wo sa", "total forecast without sa"],
        )
    )

    rows: List[Dict[str, Any]] = []
    previous_row: Optional[Dict[str, Any]] = None

    for requested_n in range(1, N_QUARTERS + 1):
        effective_n = min(requested_n, len(pairs))
        selected = pairs[-effective_n:]
        start_row = selected[0][0]
        end_row = selected[-1][0]

        set_formula2(
            intercept_cell,
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})",
        )
        set_formula2(
            slope_cell,
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})",
        )
        wb.app.calculate()

        intercept = as_float(intercept_cell.value)
        slope = as_float(slope_cell.value)
        if intercept is None or slope is None:
            continue

        x_target = selected[-1][1]
        forecast_value = intercept + slope * x_target
        if tot_fcst_wo_sa is not None and requested_n == 1:
            # Prefer workbook value for the default case when present.
            forecast_value = tot_fcst_wo_sa

        residuals = [y_val - (intercept + slope * x_val) for _, x_val, y_val in selected]
        mean_sq = sum(r * r for r in residuals) / len(residuals) if residuals else 0.0
        spread = math.sqrt(mean_sq)
        forecast_max = forecast_value + spread
        forecast_min = forecast_value - spread
        range_width = forecast_max - forecast_min

        current_row = {
            "model": model_meta.model,
            "ticker": model_meta.ticker,
            "model_period": model_meta.model_period,
            "model_date": model_meta.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": effective_n,
            "num_quarters_used": effective_n,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file_name,
        }

        is_duplicate = False
        if previous_row is not None:
            is_duplicate = (
                approx_equal(as_float(current_row["forecast_value"]), as_float(previous_row["forecast_value"]))
                and approx_equal(as_float(current_row["forecast_max"]), as_float(previous_row["forecast_max"]))
                and approx_equal(as_float(current_row["forecast_min"]), as_float(previous_row["forecast_min"]))
                and approx_equal(as_float(current_row["intercept"]), as_float(previous_row["intercept"]))
                and approx_equal(as_float(current_row["slope"]), as_float(previous_row["slope"]))
            )

        if not is_duplicate:
            rows.append(current_row)
            previous_row = current_row

    intercept_cell.clear_contents()
    slope_cell.clear_contents()
    return rows


def main() -> None:
    if not input_dir.exists() or not input_dir.is_dir():
        raise SystemExit(f"Input directory does not exist: {input_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)
    output_file = output_path_with_suffix(input_dir, output_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    previous_calculation_mode: Optional[Any] = None
    try:
        previous_calculation_mode = app.calculation
        app.calculation = "manual"
    except Exception:
        previous_calculation_mode = None

    try:
        for file_path in iter_source_files(input_dir):
            model_meta = parse_file_model_meta(file_path.name)
            if model_meta is None:
                print(f"Skipped {file_path.name}: filename pattern not recognized")
                continue

            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                print(f"Processed {file_path.name}")

                empirical_rows.extend(
                    extract_empirical_rows(
                        wb=wb,
                        model_meta=model_meta,
                        source_file_name=file_path.name,
                    )
                )
                regression_rows.extend(
                    extract_regression_rows(
                        wb=wb,
                        model_meta=model_meta,
                        source_file_name=file_path.name,
                    )
                )
                processed_files += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error ({exc})")
            finally:
                if wb is not None:
                    safe_close_source_workbook(wb)
    finally:
        if previous_calculation_mode is not None:
            try:
                app.calculation = previous_calculation_mode
            except Exception:
                pass
        app.quit()

    write_output_workbook(
        output_file_path=output_file,
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
    )

    print(f"Output path: {output_file}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

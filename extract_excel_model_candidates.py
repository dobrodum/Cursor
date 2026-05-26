from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

try:
    import xlwings as xw
except ImportError as exc:  # pragma: no cover - import guard for runtime environments
    raise SystemExit(
        "xlwings is required to run this script. Install it with `pip install xlwings`."
    ) from exc

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter
except ImportError as exc:  # pragma: no cover - import guard for runtime environments
    raise SystemExit(
        "openpyxl is required to run this script. Install it with `pip install openpyxl`."
    ) from exc


# ---------------------------------------------------------------------------
# User-configurable paths
# ---------------------------------------------------------------------------
input_dir = Path("/workspace/input")
output_dir = Path("/workspace/output")


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

PERIOD_DAY_MAP = {"early": 5, "mid": 15, "late": 25}
MONTH_NUM_MAP = {
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
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetSnapshot:
    values: List[List[Any]]
    top_row: int
    left_col: int
    row_count: int
    col_count: int

    @property
    def right_col(self) -> int:
        return self.left_col + self.col_count - 1

    @property
    def bottom_row(self) -> int:
        return self.top_row + self.row_count - 1


def normalize_header(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[%()/\-]+", " ", text)
    text = re.sub(r"\s+", "_", text)
    return text.strip("_")


def is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def has_value(value: Any) -> bool:
    return value not in (None, "")


def numeric_or_none(value: Any) -> Optional[float]:
    if is_number(value):
        return float(value)
    if isinstance(value, str):
        stripped = value.strip().replace(",", "")
        if stripped == "":
            return None
        try:
            return float(stripped)
        except ValueError:
            return None
    return None


def safe_subtract(left: Any, right: Any) -> Optional[float]:
    left_num = numeric_or_none(left)
    right_num = numeric_or_none(right)
    if left_num is None or right_num is None:
        return None
    return left_num - right_num


def to_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            max_cols = max(len(row) for row in values) if values else 0
            return [list(row) + [None] * (max_cols - len(row)) for row in values]
        return [list(values)]
    return [[values]]


def flatten_col(values: Any, expected_len: int) -> List[Any]:
    if isinstance(values, list):
        if values and isinstance(values[0], list):
            flattened = [row[0] if row else None for row in values]
        else:
            flattened = list(values)
    else:
        flattened = [values]
    if len(flattened) < expected_len:
        flattened.extend([None] * (expected_len - len(flattened)))
    return flattened[:expected_len]


def parse_file_label(file_name: str) -> FileLabel:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1] if len(parts) >= 2 else ""
    period_token = parts[2].split("_")[0] if len(parts) >= 3 else ""

    model_period = ""
    model_date = ""
    period_match = re.match(
        r"^(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(\d{4})$",
        period_token,
        flags=re.IGNORECASE,
    )
    if period_match:
        bucket = period_match.group(1).title()
        month_abbrev = period_match.group(2).title()
        year = int(period_match.group(3))
        month_num = MONTH_NUM_MAP[month_abbrev.lower()]
        day_num = PERIOD_DAY_MAP[bucket.lower()]
        model_period = f"{bucket}{month_abbrev}_{year}"
        model_date = f"{year:04d}-{month_num:02d}-{day_num:02d}"

    if not ticker:
        ticker = parts[0].split("_")[0].strip() if parts else stem

    model = f"{ticker}_{model_period}" if model_period else ticker
    return FileLabel(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def get_next_output_path(input_folder: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_folder.name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    version = 1
    while candidate.exists():
        candidate = out_dir / f"{base_name}.{version}.xlsx"
        version += 1
    return candidate


def get_sheet_by_name(workbook: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    target = sheet_name.strip().lower()
    for sheet in workbook.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def snapshot_sheet(sheet: xw.Sheet) -> Optional[SheetSnapshot]:
    used = sheet.used_range
    values = to_2d(used.value)
    if not values:
        return None
    row_count = len(values)
    col_count = max((len(row) for row in values), default=0)
    if col_count == 0:
        return None
    normalized_values = [row + [None] * (col_count - len(row)) for row in values]
    return SheetSnapshot(
        values=normalized_values,
        top_row=used.row,
        left_col=used.column,
        row_count=row_count,
        col_count=col_count,
    )


def get_cell(snapshot: SheetSnapshot, row: int, col: int) -> Any:
    r_idx = row - snapshot.top_row
    c_idx = col - snapshot.left_col
    if r_idx < 0 or c_idx < 0:
        return None
    if r_idx >= snapshot.row_count or c_idx >= snapshot.col_count:
        return None
    return snapshot.values[r_idx][c_idx]


def find_anchor(snapshot: SheetSnapshot, label: str = "max") -> Optional[Tuple[int, int]]:
    target = label.strip().lower()
    for r_offset, row_values in enumerate(snapshot.values):
        for c_offset, value in enumerate(row_values):
            if isinstance(value, str) and value.strip().lower() == target:
                return snapshot.top_row + r_offset, snapshot.left_col + c_offset
    return None


def row_label_count(snapshot: SheetSnapshot, row: int) -> int:
    labels = 0
    for col in range(snapshot.left_col, snapshot.right_col + 1):
        value = get_cell(snapshot, row, col)
        if isinstance(value, str) and value.strip():
            labels += 1
    return labels


def pick_header_row(snapshot: SheetSnapshot, anchor_row: int) -> int:
    anchor_score = row_label_count(snapshot, anchor_row)
    prev_row = anchor_row - 1
    if prev_row < snapshot.top_row:
        return anchor_row
    prev_score = row_label_count(snapshot, prev_row)
    return prev_row if prev_score > anchor_score else anchor_row


def build_header_map(snapshot: SheetSnapshot, header_row: int) -> Dict[str, int]:
    headers: Dict[str, int] = {}
    for col in range(snapshot.left_col, snapshot.right_col + 1):
        key = normalize_header(get_cell(snapshot, header_row, col))
        if key and key not in headers:
            headers[key] = col
    return headers


def resolve_column(
    header_map: Dict[str, int],
    aliases: Sequence[str],
    fallback: Optional[int] = None,
) -> Optional[int]:
    normalized_aliases = [normalize_header(alias) for alias in aliases if alias]

    for alias in normalized_aliases:
        if alias in header_map:
            return header_map[alias]

    for key, col in header_map.items():
        for alias in normalized_aliases:
            if alias and alias in key:
                return col
    return fallback


def find_penetration_series_columns(header_map: Dict[str, int]) -> List[int]:
    cols: List[int] = []
    for key, col in header_map.items():
        if "penetration" in key and "avg" not in key and "average" not in key:
            cols.append(col)
    return sorted(set(cols))


def set_formula2(target_range: xw.Range, formulas: List[List[str]]) -> None:
    try:
        target_range.formula2 = formulas
    except Exception:
        target_range.formula = formulas


def clear_contents(target_range: xw.Range) -> None:
    try:
        target_range.clear_contents()
    except Exception:
        target_range.value = None


def close_source_workbook(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    # Fallback for xlwings variants that do not accept save=False.
    try:
        workbook.api.Saved = True
    except Exception:
        pass
    workbook.close()


def infer_num_quarters(raw_value: Any, default_value: int) -> int:
    num = numeric_or_none(raw_value)
    if num is None:
        return default_value
    if num <= 0:
        return default_value
    return int(round(num))


def is_blank_empirical_candidate(
    forecast_value: Any,
    actual_value: Any,
    forecast_max: Any,
    forecast_min: Any,
    avg_penetration_pct: Any,
    quarterly_sales: Any,
    reported_sales: Any,
) -> bool:
    values = [
        forecast_value,
        actual_value,
        forecast_max,
        forecast_min,
        avg_penetration_pct,
        quarterly_sales,
        reported_sales,
    ]
    return not any(has_value(value) for value in values)


def is_blank_regression_candidate(
    forecast_value: Any,
    forecast_max: Any,
    forecast_min: Any,
    intercept: Any,
    slope: Any,
) -> bool:
    values = [forecast_value, forecast_max, forecast_min, intercept, slope]
    return not any(has_value(value) for value in values)


def normalized_compare_key(value: Any) -> Any:
    num = numeric_or_none(value)
    if num is None:
        if value in ("", None):
            return None
        return str(value)
    return round(num, 10)


def extract_empirical_candidates(
    workbook: xw.Book,
    file_meta: FileLabel,
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = get_sheet_by_name(workbook, "Empirical Model")
    if sheet is None:
        return []

    snapshot = snapshot_sheet(sheet)
    if snapshot is None:
        return []

    anchor = find_anchor(snapshot, "max")
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    header_row = pick_header_row(snapshot, anchor_row)
    header_map = build_header_map(snapshot, header_row)

    # Anchor-driven fallbacks keep extraction robust even if headers vary slightly.
    max_col = anchor_col
    min_col = resolve_column(header_map, ["min", "forecast_min"], fallback=anchor_col + 1)
    num_quarters_col = resolve_column(
        header_map,
        ["num_quarters_used", "num quarters used", "n_quarters", "quarters_used"],
        fallback=anchor_col - 8,
    )
    last_quarter_col = resolve_column(
        header_map,
        ["last_quarter_used", "last quarter used", "last_quarter"],
        fallback=anchor_col - 7,
    )
    forecast_col = resolve_column(
        header_map,
        ["estimated_total_sold", "estimated total sold", "forecast_value", "forecast"],
        fallback=anchor_col - 3,
    )
    actual_col = resolve_column(
        header_map,
        ["reported_sales", "actual_value", "actual"],
        fallback=anchor_col - 2,
    )
    avg_penetration_col = resolve_column(
        header_map,
        ["avg_penetration_pct", "avg penetration pct", "average penetration", "avg penetration"],
        fallback=None,
    )
    quarterly_sales_col = resolve_column(
        header_map,
        ["quarterly_sales", "quarterly sales"],
        fallback=anchor_col - 6,
    )
    reported_sales_col = resolve_column(
        header_map,
        ["reported_sales", "reported sales"],
        fallback=actual_col,
    )
    growth_rate_col = resolve_column(
        header_map,
        ["growth_rate_pct", "growth rate pct", "growth rate"],
        fallback=anchor_col - 5,
    )
    captured_col = resolve_column(
        header_map,
        ["sales_captured_in_db_pct", "sales captured in db pct", "sales captured"],
        fallback=anchor_col - 4,
    )
    penetration_series_cols = find_penetration_series_columns(header_map)

    start_row = anchor_row + 1
    end_row = start_row + N_QUARTERS - 1
    helper_col = snapshot.right_col + 2

    avg_formulas: List[List[str]] = []
    needs_calc = False
    for idx in range(N_QUARTERS):
        n_used = idx + 1
        if penetration_series_cols:
            end_index = min(n_used, len(penetration_series_cols)) - 1
            start_col = penetration_series_cols[0]
            end_col = penetration_series_cols[end_index]
            formula = (
                f'=IFERROR(AVERAGE(RC[{start_col - helper_col}]:RC[{end_col - helper_col}]),"")'
            )
            needs_calc = True
        elif quarterly_sales_col is not None and reported_sales_col is not None:
            formula = (
                f'=IFERROR(RC[{reported_sales_col - helper_col}]'
                f'/RC[{quarterly_sales_col - helper_col}], "")'
            )
            needs_calc = True
        elif avg_penetration_col is not None:
            formula = f'=IFERROR(RC[{avg_penetration_col - helper_col}], "")'
            needs_calc = True
        else:
            formula = '=""'
        avg_formulas.append([formula])

    avg_range = sheet.range((start_row, helper_col), (end_row, helper_col))
    set_formula2(avg_range, avg_formulas)
    if needs_calc:
        workbook.app.calculate()
    avg_values = flatten_col(avg_range.value, N_QUARTERS)
    clear_contents(avg_range)

    rows: List[Dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        row_num = start_row + idx
        default_n = idx + 1
        num_quarters_used = infer_num_quarters(get_cell(snapshot, row_num, num_quarters_col), default_n)
        last_quarter_used = get_cell(snapshot, row_num, last_quarter_col) if last_quarter_col else None
        if not has_value(last_quarter_used):
            last_quarter_used = num_quarters_used

        forecast_value = get_cell(snapshot, row_num, forecast_col) if forecast_col else None
        actual_value = get_cell(snapshot, row_num, actual_col) if actual_col else None
        forecast_max = get_cell(snapshot, row_num, max_col)
        forecast_min = get_cell(snapshot, row_num, min_col) if min_col else None
        avg_penetration_pct = avg_values[idx] if idx < len(avg_values) else None
        quarterly_sales = get_cell(snapshot, row_num, quarterly_sales_col) if quarterly_sales_col else None
        reported_sales = get_cell(snapshot, row_num, reported_sales_col) if reported_sales_col else None
        growth_rate_pct = get_cell(snapshot, row_num, growth_rate_col) if growth_rate_col else None
        sales_captured_in_db_pct = get_cell(snapshot, row_num, captured_col) if captured_col else None

        if is_blank_empirical_candidate(
            forecast_value,
            actual_value,
            forecast_max,
            forecast_min,
            avg_penetration_pct,
            quarterly_sales,
            reported_sales,
        ):
            continue

        rows.append(
            {
                "model": file_meta.model,
                "ticker": file_meta.ticker,
                "model_period": file_meta.model_period,
                "model_date": file_meta.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_subtract(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    return rows


def collect_numeric_pair_rows(
    snapshot: SheetSnapshot,
    x_col: int,
    y_col: int,
    max_row: int,
) -> List[int]:
    rows: List[int] = []
    for row in range(snapshot.top_row, max_row + 1):
        if numeric_or_none(get_cell(snapshot, row, x_col)) is not None and numeric_or_none(
            get_cell(snapshot, row, y_col)
        ) is not None:
            rows.append(row)
    return rows


def extract_regression_candidates(
    workbook: xw.Book,
    file_meta: FileLabel,
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = get_sheet_by_name(workbook, "Regression Model")
    if sheet is None:
        return []

    snapshot = snapshot_sheet(sheet)
    if snapshot is None:
        return []

    anchor = find_anchor(snapshot, "max")
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    header_row = pick_header_row(snapshot, anchor_row)
    header_map = build_header_map(snapshot, header_row)

    x_col = anchor_col - 11
    y_col = anchor_col - 7
    max_col = anchor_col
    min_col = resolve_column(header_map, ["min", "forecast_min"], fallback=anchor_col + 1)
    num_quarters_col = resolve_column(
        header_map,
        ["num_quarters_used", "num quarters used", "n_quarters", "quarters_used"],
        fallback=anchor_col - 8,
    )
    forecast_col = resolve_column(
        header_map,
        [
            "tot_fcst_w_o_sa",
            "tot_fcst_wo_sa",
            "tot fcst w/o sa",
            "forecast_total_without_sa",
            "forecast_value",
            "forecast",
        ],
        fallback=anchor_col - 3,
    )
    actual_col = resolve_column(
        header_map,
        ["actual_value", "actual", "reported_sales", "reported sales"],
        fallback=None,
    )

    pair_rows = collect_numeric_pair_rows(snapshot, x_col=x_col, y_col=y_col, max_row=anchor_row - 1)
    start_row = anchor_row + 1
    end_row = start_row + N_QUARTERS - 1
    helper_intercept_col = snapshot.right_col + 2
    helper_slope_col = helper_intercept_col + 1

    intercept_formulas: List[List[str]] = []
    slope_formulas: List[List[str]] = []
    needs_calc = False

    for idx in range(N_QUARTERS):
        row_num = start_row + idx
        default_n = idx + 1
        raw_num_quarters = get_cell(snapshot, row_num, num_quarters_col) if num_quarters_col else None
        num_quarters_used = infer_num_quarters(raw_num_quarters, default_n)

        if len(pair_rows) >= 2:
            window_size = max(2, min(num_quarters_used, len(pair_rows)))
            selected_rows = pair_rows[-window_size:]
            data_start = selected_rows[0]
            data_end = selected_rows[-1]
            intercept_formula = (
                f'=IFERROR(INTERCEPT(R{data_start}C{y_col}:R{data_end}C{y_col},'
                f'R{data_start}C{x_col}:R{data_end}C{x_col}),"")'
            )
            slope_formula = (
                f'=IFERROR(SLOPE(R{data_start}C{y_col}:R{data_end}C{y_col},'
                f'R{data_start}C{x_col}:R{data_end}C{x_col}),"")'
            )
            needs_calc = True
        else:
            intercept_formula = '=""'
            slope_formula = '=""'

        intercept_formulas.append([intercept_formula])
        slope_formulas.append([slope_formula])

    intercept_range = sheet.range((start_row, helper_intercept_col), (end_row, helper_intercept_col))
    slope_range = sheet.range((start_row, helper_slope_col), (end_row, helper_slope_col))
    set_formula2(intercept_range, intercept_formulas)
    set_formula2(slope_range, slope_formulas)
    if needs_calc:
        workbook.app.calculate()
    intercept_values = flatten_col(intercept_range.value, N_QUARTERS)
    slope_values = flatten_col(slope_range.value, N_QUARTERS)
    clear_contents(intercept_range)
    clear_contents(slope_range)

    rows: List[Dict[str, Any]] = []
    previous_key: Optional[Tuple[Any, ...]] = None
    for idx in range(N_QUARTERS):
        row_num = start_row + idx
        default_n = idx + 1
        num_quarters_used = infer_num_quarters(get_cell(snapshot, row_num, num_quarters_col), default_n)
        forecast_value = get_cell(snapshot, row_num, forecast_col) if forecast_col else None
        actual_value = get_cell(snapshot, row_num, actual_col) if actual_col else None
        forecast_max = get_cell(snapshot, row_num, max_col)
        forecast_min = get_cell(snapshot, row_num, min_col) if min_col else None
        intercept = intercept_values[idx] if idx < len(intercept_values) else None
        slope = slope_values[idx] if idx < len(slope_values) else None

        if is_blank_regression_candidate(forecast_value, forecast_max, forecast_min, intercept, slope):
            continue

        current_key = (
            normalized_compare_key(num_quarters_used),
            normalized_compare_key(forecast_value),
            normalized_compare_key(forecast_max),
            normalized_compare_key(forecast_min),
            normalized_compare_key(intercept),
            normalized_compare_key(slope),
        )
        if previous_key is not None and current_key == previous_key:
            # Skip repeated terminal rows when Excel fills trailing rows with duplicates.
            continue
        previous_key = current_key

        rows.append(
            {
                "model": file_meta.model,
                "ticker": file_meta.ticker,
                "model_period": file_meta.model_period,
                "model_date": file_meta.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value if has_value(actual_value) else None,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_subtract(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def write_table_sheet(
    workbook: Workbook,
    sheet_name: str,
    columns: Sequence[str],
    rows: Iterable[Dict[str, Any]],
) -> None:
    if sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        sheet.delete_rows(1, sheet.max_row)
    else:
        sheet = workbook.create_sheet(sheet_name)

    sheet.append(list(columns))
    for row in rows:
        sheet.append([row.get(col) for col in columns])

    for header_cell in sheet[1]:
        header_cell.font = Font(bold=True)

    sheet.freeze_panes = "A2"
    if sheet.max_row >= 1 and sheet.max_column >= 1:
        sheet.auto_filter.ref = sheet.dimensions

    for col_idx, column_name in enumerate(columns, start=1):
        values = [column_name]
        for row_idx in range(2, sheet.max_row + 1):
            cell_value = sheet.cell(row=row_idx, column=col_idx).value
            if cell_value is not None:
                values.append(str(cell_value))
        max_len = max((len(v) for v in values), default=len(column_name))
        width = max(12, min(max_len + 2, 48))
        sheet.column_dimensions[get_column_letter(col_idx)].width = width


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)
    write_table_sheet(workbook, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_table_sheet(workbook, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    workbook.save(output_path)


def iter_input_files(folder: Path) -> Iterable[Path]:
    if not folder.exists():
        print(f"Input directory does not exist: {folder}")
        return []
    return sorted(path for path in folder.iterdir() if path.is_file())


def main() -> None:
    all_empirical_rows: List[Dict[str, Any]] = []
    all_regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    output_path = get_next_output_path(input_dir, output_dir)

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in iter_input_files(input_dir):
            if file_path.name.startswith("~"):
                print(f"skipped: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped: {file_path.name} (not .xlsx)")
                continue

            print(f"processing: {file_path.name}")
            workbook: Optional[xw.Book] = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                file_meta = parse_file_label(file_path.name)

                empirical_rows = extract_empirical_candidates(
                    workbook=workbook,
                    file_meta=file_meta,
                    source_file=file_path.name,
                )
                regression_rows = extract_regression_candidates(
                    workbook=workbook,
                    file_meta=file_meta,
                    source_file=file_path.name,
                )

                all_empirical_rows.extend(empirical_rows)
                all_regression_rows.extend(regression_rows)
                files_processed += 1
                print(
                    f"processed: {file_path.name} "
                    f"(empirical_rows={len(empirical_rows)}, regression_rows={len(regression_rows)})"
                )
            except Exception as exc:
                print(f"skipped: {file_path.name} (error: {exc})")
            finally:
                if workbook is not None:
                    close_source_workbook(workbook)
    finally:
        app.quit()

    write_output_workbook(
        output_path=output_path,
        empirical_rows=all_empirical_rows,
        regression_rows=all_regression_rows,
    )

    print(f"output_path: {output_path}")
    print(f"files_processed: {files_processed}")
    print(f"empirical_rows: {len(all_empirical_rows)}")
    print(f"regression_rows: {len(all_regression_rows)}")


if __name__ == "__main__":
    main()

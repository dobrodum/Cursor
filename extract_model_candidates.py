from __future__ import annotations

import calendar
import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# ----------------------------- User-configurable paths -----------------------------
input_dir = "/path/to/input"
output_dir = "/path/to/output"
# -----------------------------------------------------------------------------------

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

EARLY_MID_LATE_DAY = {"Early": 5, "Mid": 15, "Late": 25}
PERIOD_PATTERN = re.compile(r"(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})", flags=re.IGNORECASE)
NON_ALNUM_PATTERN = re.compile(r"[^a-z0-9]+")

MONTH_LOOKUP: dict[str, int] = {}
for idx, month_name in enumerate(calendar.month_name):
    if month_name:
        MONTH_LOOKUP[month_name.lower()] = idx
for idx, month_abbrev in enumerate(calendar.month_abbr):
    if month_abbrev:
        MONTH_LOOKUP[month_abbrev.lower()] = idx
MONTH_LOOKUP["sept"] = 9


@dataclass(frozen=True)
class FileMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetSnapshot:
    values: list[list[Any]]
    top_row: int
    left_col: int
    last_row: int
    last_col: int

    def get(self, row: int, col: int) -> Any:
        row_idx = row - self.top_row
        col_idx = col - self.left_col
        if row_idx < 0 or col_idx < 0:
            return None
        if row_idx >= len(self.values):
            return None
        row_vals = self.values[row_idx]
        if col_idx >= len(row_vals):
            return None
        return row_vals[col_idx]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    return NON_ALNUM_PATTERN.sub(" ", text).strip()


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return not value.strip()
    return False


def to_2d_list(values: Any) -> list[list[Any]]:
    if isinstance(values, tuple):
        values = [list(item) if isinstance(item, tuple) else item for item in values]
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], tuple):
        values = [list(item) if isinstance(item, tuple) else [item] for item in values]
    if isinstance(values[0], list):
        return values
    return [values]


def to_column_list(values: Any, expected_len: int) -> list[Any]:
    if isinstance(values, list):
        if values and isinstance(values[0], list):
            flat = [row[0] if row else None for row in values]
        else:
            flat = values
    else:
        flat = [values]

    if len(flat) < expected_len:
        flat.extend([None] * (expected_len - len(flat)))
    return flat[:expected_len]


def as_number(value: Any) -> float | None:
    if is_blank(value):
        return None
    if isinstance(value, bool):
        return float(int(value))
    if isinstance(value, (int, float)):
        return float(value)
    try:
        cleaned = str(value).replace(",", "").strip()
        return float(cleaned)
    except (TypeError, ValueError):
        return None


def diff_values(lhs: Any, rhs: Any) -> float | None:
    lhs_num = as_number(lhs)
    rhs_num = as_number(rhs)
    if lhs_num is None or rhs_num is None:
        return None
    return lhs_num - rhs_num


def signature_value(value: Any) -> Any:
    number = as_number(value)
    if number is not None:
        return round(number, 10)
    if isinstance(value, datetime):
        return value.isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if value is None:
        return None
    return str(value).strip()


def parse_month(month_token: str) -> int | None:
    cleaned = re.sub(r"[^A-Za-z]", "", month_token).lower()
    if not cleaned:
        return None
    if cleaned in MONTH_LOOKUP:
        return MONTH_LOOKUP[cleaned]
    short = cleaned[:3]
    return MONTH_LOOKUP.get(short)


def guess_ticker(file_stem: str) -> str:
    token_pattern = re.compile(r"\b[A-Z]{2,8}\b")
    matches = token_pattern.findall(file_stem)
    if matches:
        return matches[0]
    return "UNKNOWN"


def parse_file_metadata(file_name: str) -> FileMetadata:
    file_stem = Path(file_name).stem
    parts = [part.strip() for part in file_stem.split(" - ") if part.strip()]

    ticker = parts[1].upper() if len(parts) >= 2 else guess_ticker(file_stem)

    period_source = parts[2] if len(parts) >= 3 else file_stem
    period_source = period_source.split("_")[0].strip()

    match = PERIOD_PATTERN.search(period_source.replace(" ", ""))
    if not match:
        match = PERIOD_PATTERN.search(file_stem.replace(" ", ""))

    if match:
        phase = match.group(1).title()
        month_token = match.group(2)
        year_text = match.group(3)
        month_num = parse_month(month_token)
        if month_num:
            month_abbrev = calendar.month_abbr[month_num]
            model_period = f"{phase}{month_abbrev}_{year_text}"
            model_day = EARLY_MID_LATE_DAY.get(phase, 15)
            model_date = date(int(year_text), month_num, model_day).isoformat()
        else:
            model_period = f"{phase}{month_token}_{year_text}"
            model_date = ""
    else:
        model_period = "unknown_period"
        model_date = ""

    model = f"{ticker}_{model_period}"
    return FileMetadata(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def next_output_path(input_path: Path, output_path: Path) -> Path:
    base_name = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = output_path / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def set_formula2(target_range: xw.main.Range, formula_r1c1: str) -> None:
    try:
        target_range.formula2 = formula_r1c1
    except Exception:
        target_range.formula = formula_r1c1


def safe_close_workbook(workbook: xw.main.Book) -> None:
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
        workbook.close()
    except Exception:
        pass


def snapshot_sheet(sheet: xw.main.Sheet) -> SheetSnapshot:
    used = sheet.used_range
    values = to_2d_list(used.value)
    top_row = used.row
    left_col = used.column
    if not values:
        return SheetSnapshot(values=[[]], top_row=top_row, left_col=left_col, last_row=top_row, last_col=left_col)
    last_row = top_row + len(values) - 1
    row_width = max((len(row) for row in values if isinstance(row, list)), default=1)
    last_col = left_col + row_width - 1
    return SheetSnapshot(values=values, top_row=top_row, left_col=left_col, last_row=last_row, last_col=last_col)


def find_max_anchor(snapshot: SheetSnapshot) -> tuple[int, int] | None:
    candidates: list[tuple[int, int, int]] = []
    for row_idx, row_vals in enumerate(snapshot.values):
        for col_idx, value in enumerate(row_vals):
            if normalize_text(value) != "max":
                continue
            abs_row = snapshot.top_row + row_idx
            abs_col = snapshot.left_col + col_idx
            score = 0
            if normalize_text(snapshot.get(abs_row, abs_col + 1)) == "min":
                score += 3
            if as_number(snapshot.get(abs_row + 1, abs_col)) is not None:
                score += 1
            candidates.append((score, abs_row, abs_col))
    if not candidates:
        return None
    candidates.sort(reverse=True)
    _, anchor_row, anchor_col = candidates[0]
    return anchor_row, anchor_col


def find_column_by_keywords(
    snapshot: SheetSnapshot,
    header_row: int,
    scan_left: int,
    scan_right: int,
    keyword_sets: list[tuple[str, ...]],
) -> int | None:
    for col in range(scan_left, scan_right + 1):
        header_text = normalize_text(snapshot.get(header_row, col))
        if not header_text:
            continue
        for keyword_set in keyword_sets:
            if all(keyword in header_text for keyword in keyword_set):
                return col
    return None


def resolve_column(
    snapshot: SheetSnapshot,
    header_row: int,
    scan_left: int,
    scan_right: int,
    keyword_sets: list[tuple[str, ...]],
    anchor_col: int,
    fallback_offset: int | None,
) -> int | None:
    detected = find_column_by_keywords(snapshot, header_row, scan_left, scan_right, keyword_sets)
    if detected is not None:
        return detected
    if fallback_offset is None:
        return None
    return anchor_col + fallback_offset


def get_sheet_by_name(workbook: xw.main.Book, target_name: str) -> xw.main.Sheet | None:
    for sheet in workbook.sheets:
        if sheet.name.strip().lower() == target_name.strip().lower():
            return sheet
    return None


def process_empirical_sheet(
    workbook: xw.main.Book,
    sheet: xw.main.Sheet,
    metadata: FileMetadata,
    source_file: str,
) -> list[dict[str, Any]]:
    snapshot = snapshot_sheet(sheet)
    anchor = find_max_anchor(snapshot)
    if anchor is None:
        print(f"SKIPPED {source_file} [Empirical Model]: could not find 'max' anchor")
        return []

    anchor_row, anchor_col = anchor
    header_row = anchor_row
    scan_left = max(snapshot.left_col, anchor_col - 25)
    scan_right = max(scan_left, min(snapshot.last_col, anchor_col + 12))

    empirical_offsets = {
        "num_quarters_used": -8,
        "last_quarter_used": -7,
        "forecast_value": -3,
        "actual_value": -2,
        "forecast_min": 1,
        "avg_penetration_pct": -5,
        "quarterly_sales": -11,
        "reported_sales": -2,
        "growth_rate_pct": -6,
        "sales_captured_in_db_pct": -4,
    }

    num_quarters_col = resolve_column(
        snapshot,
        header_row,
        scan_left,
        scan_right,
        [("num", "quarter"), ("quarters", "used"), ("n", "quarters")],
        anchor_col,
        empirical_offsets["num_quarters_used"],
    )
    last_quarter_col = resolve_column(
        snapshot,
        header_row,
        scan_left,
        scan_right,
        [("last", "quarter"), ("quarter", "used")],
        anchor_col,
        empirical_offsets["last_quarter_used"],
    )
    forecast_value_col = resolve_column(
        snapshot,
        header_row,
        scan_left,
        scan_right,
        [("estimated", "total", "sold"), ("forecast", "value"), ("tot", "fcst")],
        anchor_col,
        empirical_offsets["forecast_value"],
    )
    actual_value_col = resolve_column(
        snapshot,
        header_row,
        scan_left,
        scan_right,
        [("reported", "sales"), ("actual", "sales"), ("actual", "value")],
        anchor_col,
        empirical_offsets["actual_value"],
    )
    min_col = resolve_column(
        snapshot,
        header_row,
        scan_left,
        scan_right,
        [("min",)],
        anchor_col,
        empirical_offsets["forecast_min"],
    )
    avg_penetration_col = resolve_column(
        snapshot,
        header_row,
        scan_left,
        scan_right,
        [("avg", "penetration"), ("average", "penetration"), ("penetration", "pct")],
        anchor_col,
        empirical_offsets["avg_penetration_pct"],
    )
    quarterly_sales_col = resolve_column(
        snapshot,
        header_row,
        scan_left,
        scan_right,
        [("quarterly", "sales"), ("qtr", "sales"), ("quarter", "sales")],
        anchor_col,
        empirical_offsets["quarterly_sales"],
    )
    reported_sales_col = resolve_column(
        snapshot,
        header_row,
        scan_left,
        scan_right,
        [("reported", "sales"), ("actual", "sales")],
        anchor_col,
        empirical_offsets["reported_sales"],
    )
    growth_rate_col = resolve_column(
        snapshot,
        header_row,
        scan_left,
        scan_right,
        [("growth", "rate"), ("growth", "pct")],
        anchor_col,
        empirical_offsets["growth_rate_pct"],
    )
    sales_captured_col = resolve_column(
        snapshot,
        header_row,
        scan_left,
        scan_right,
        [("captured", "db"), ("sales", "captured")],
        anchor_col,
        empirical_offsets["sales_captured_in_db_pct"],
    )

    max_col = anchor_col
    data_start_row = anchor_row + 1
    avg_pen_formula_values = [None] * N_QUARTERS

    if quarterly_sales_col is not None and reported_sales_col is not None:
        temp_col = snapshot.last_col + 2
        for idx in range(N_QUARTERS):
            row = data_start_row + idx
            lookback = idx + 1
            start_row = max(data_start_row, row - lookback + 1)
            formula = (
                f'=IFERROR(SUM(R{start_row}C{quarterly_sales_col}:R{row}C{quarterly_sales_col})/'
                f'SUM(R{start_row}C{reported_sales_col}:R{row}C{reported_sales_col}),"")'
            )
            set_formula2(sheet.range((row, temp_col)), formula)
        workbook.app.calculate()
        avg_values = sheet.range(
            (data_start_row, temp_col), (data_start_row + N_QUARTERS - 1, temp_col)
        ).value
        avg_pen_formula_values = to_column_list(avg_values, N_QUARTERS)

    rows: list[dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        row_num = data_start_row + idx
        num_quarters_used = snapshot.get(row_num, num_quarters_col) if num_quarters_col else idx + 1
        if is_blank(num_quarters_used):
            num_quarters_used = idx + 1

        last_quarter_used = snapshot.get(row_num, last_quarter_col) if last_quarter_col else None
        forecast_value = snapshot.get(row_num, forecast_value_col) if forecast_value_col else None
        actual_value = snapshot.get(row_num, actual_value_col) if actual_value_col else None
        forecast_max = snapshot.get(row_num, max_col)
        forecast_min = snapshot.get(row_num, min_col) if min_col else None
        quarterly_sales = snapshot.get(row_num, quarterly_sales_col) if quarterly_sales_col else None
        reported_sales = snapshot.get(row_num, reported_sales_col) if reported_sales_col else actual_value
        growth_rate_pct = snapshot.get(row_num, growth_rate_col) if growth_rate_col else None
        sales_captured_pct = snapshot.get(row_num, sales_captured_col) if sales_captured_col else None

        avg_penetration_pct = avg_pen_formula_values[idx]
        if is_blank(avg_penetration_pct) and avg_penetration_col is not None:
            avg_penetration_pct = snapshot.get(row_num, avg_penetration_col)

        if all(
            is_blank(value)
            for value in (forecast_value, actual_value, forecast_max, forecast_min, avg_penetration_pct)
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
                "range_width": diff_values(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_pct,
                "source_file": source_file,
            }
        )

    return rows


def process_regression_sheet(
    workbook: xw.main.Book,
    sheet: xw.main.Sheet,
    metadata: FileMetadata,
    source_file: str,
) -> list[dict[str, Any]]:
    snapshot = snapshot_sheet(sheet)
    anchor = find_max_anchor(snapshot)
    if anchor is None:
        print(f"SKIPPED {source_file} [Regression Model]: could not find 'max' anchor")
        return []

    anchor_row, anchor_col = anchor
    header_row = anchor_row
    scan_left = max(snapshot.left_col, anchor_col - 25)
    scan_right = max(scan_left, min(snapshot.last_col, anchor_col + 12))

    # Required anchor-based offsets for regression model.
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    regression_offsets = {
        "num_quarters_used": -8,
        "forecast_value": -3,
        "actual_value": -2,
        "forecast_min": 1,
    }

    num_quarters_col = resolve_column(
        snapshot,
        header_row,
        scan_left,
        scan_right,
        [("num", "quarter"), ("quarters", "used"), ("n", "quarters")],
        anchor_col,
        regression_offsets["num_quarters_used"],
    )
    forecast_value_col = resolve_column(
        snapshot,
        header_row,
        scan_left,
        scan_right,
        [
            ("tot", "fcst", "w", "o", "sa"),
            ("forecast", "without", "sa"),
            ("tot", "fcst"),
            ("forecast", "value"),
        ],
        anchor_col,
        regression_offsets["forecast_value"],
    )
    actual_value_col = resolve_column(
        snapshot,
        header_row,
        scan_left,
        scan_right,
        [("actual", "sales"), ("reported", "sales"), ("actual", "value")],
        anchor_col,
        regression_offsets["actual_value"],
    )
    min_col = resolve_column(
        snapshot,
        header_row,
        scan_left,
        scan_right,
        [("min",)],
        anchor_col,
        regression_offsets["forecast_min"],
    )

    max_col = anchor_col
    data_start_row = anchor_row + 1
    temp_intercept_col = snapshot.last_col + 2
    temp_slope_col = snapshot.last_col + 3

    for idx in range(N_QUARTERS):
        row = data_start_row + idx
        lookback = idx + 1
        start_row = max(data_start_row, row - lookback + 1)

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{row}C{y_col},'
            f'R{start_row}C{x_col}:R{row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{row}C{y_col},'
            f'R{start_row}C{x_col}:R{row}C{x_col}),"")'
        )
        set_formula2(sheet.range((row, temp_intercept_col)), intercept_formula)
        set_formula2(sheet.range((row, temp_slope_col)), slope_formula)

    workbook.app.calculate()

    intercept_values = sheet.range(
        (data_start_row, temp_intercept_col), (data_start_row + N_QUARTERS - 1, temp_intercept_col)
    ).value
    slope_values = sheet.range(
        (data_start_row, temp_slope_col), (data_start_row + N_QUARTERS - 1, temp_slope_col)
    ).value
    intercepts = to_column_list(intercept_values, N_QUARTERS)
    slopes = to_column_list(slope_values, N_QUARTERS)

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None

    for idx in range(N_QUARTERS):
        row_num = data_start_row + idx
        num_quarters_used = snapshot.get(row_num, num_quarters_col) if num_quarters_col else idx + 1
        if is_blank(num_quarters_used):
            num_quarters_used = idx + 1

        forecast_value = snapshot.get(row_num, forecast_value_col) if forecast_value_col else None
        actual_value = snapshot.get(row_num, actual_value_col) if actual_value_col else None
        forecast_max = snapshot.get(row_num, max_col)
        forecast_min = snapshot.get(row_num, min_col) if min_col else None
        intercept = intercepts[idx]
        slope = slopes[idx]

        if all(is_blank(value) for value in (forecast_value, forecast_max, forecast_min, intercept, slope)):
            continue

        current_signature = (
            signature_value(num_quarters_used),
            signature_value(forecast_value),
            signature_value(forecast_max),
            signature_value(forecast_min),
            signature_value(intercept),
            signature_value(slope),
        )
        if previous_signature is not None and current_signature == previous_signature:
            continue
        previous_signature = current_signature

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
                "range_width": diff_values(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def write_output_sheet(
    workbook: Workbook,
    sheet_name: str,
    columns: list[str],
    rows: list[dict[str, Any]],
) -> None:
    worksheet = workbook.create_sheet(title=sheet_name)
    worksheet.append(columns)

    for row in rows:
        worksheet.append([row.get(column) for column in columns])

    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    worksheet.freeze_panes = "A2"
    last_col_letter = get_column_letter(len(columns))
    last_row = max(worksheet.max_row, 1)
    worksheet.auto_filter.ref = f"A1:{last_col_letter}{last_row}"

    for col_idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row_idx in range(2, worksheet.max_row + 1):
            value = worksheet.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            if isinstance(value, (datetime, date)):
                rendered = value.isoformat()
            else:
                rendered = str(value)
            max_len = max(max_len, len(rendered))
        worksheet.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 42)


def write_output_workbook(
    output_file: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    workbook = Workbook()
    if workbook.active:
        workbook.remove(workbook.active)

    write_output_sheet(workbook, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_output_sheet(workbook, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    workbook.save(output_file)


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"Input folder does not exist or is not a directory: {input_path}")

    output_path.mkdir(parents=True, exist_ok=True)
    output_file = next_output_path(input_path, output_path)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app: xw.main.App | None = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in sorted(input_path.iterdir(), key=lambda p: p.name.lower()):
            if not file_path.is_file():
                continue
            if file_path.name.startswith("~"):
                print(f"SKIPPED {file_path.name}: temporary file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"SKIPPED {file_path.name}: not an .xlsx file")
                continue
            if file_path.name.lower().startswith(f"{input_path.name.lower()}_param"):
                print(f"SKIPPED {file_path.name}: generated PARAM workbook")
                continue

            print(f"PROCESSING {file_path.name}")
            workbook: xw.main.Book | None = None
            try:
                metadata = parse_file_metadata(file_path.name)
                workbook = app.books.open(str(file_path), update_links=False)

                empirical_sheet = get_sheet_by_name(workbook, "Empirical Model")
                if empirical_sheet is None:
                    print(f"SKIPPED {file_path.name} [Empirical Model]: sheet missing")
                else:
                    empirical_rows.extend(
                        process_empirical_sheet(
                            workbook=workbook,
                            sheet=empirical_sheet,
                            metadata=metadata,
                            source_file=file_path.name,
                        )
                    )

                regression_sheet = get_sheet_by_name(workbook, "Regression Model")
                if regression_sheet is None:
                    print(f"SKIPPED {file_path.name} [Regression Model]: sheet missing")
                else:
                    regression_rows.extend(
                        process_regression_sheet(
                            workbook=workbook,
                            sheet=regression_sheet,
                            metadata=metadata,
                            source_file=file_path.name,
                        )
                    )

                processed_files += 1
            except Exception as exc:
                print(f"SKIPPED {file_path.name}: {exc}")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass

    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"OUTPUT {output_file}")
    print(f"FILES PROCESSED {processed_files}")
    print(f"EMPIRICAL ROWS {len(empirical_rows)}")
    print(f"REGRESSION ROWS {len(regression_rows)}")


if __name__ == "__main__":
    main()

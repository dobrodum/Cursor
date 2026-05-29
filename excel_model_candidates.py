#!/usr/bin/env python3
"""Extract empirical and regression model candidates from Excel workbooks."""

from __future__ import annotations

import calendar
import math
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Iterable, Sequence

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

try:
    import xlwings as xw
except ImportError as exc:  # pragma: no cover - runtime dependency
    raise SystemExit("xlwings is required: pip install xlwings") from exc


# Update these paths for your environment.
input_dir = "./input"
output_dir = "./output"


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


@dataclass
class ModelMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetSnapshot:
    values: list[list[Any]]
    top_row: int
    left_col: int
    n_rows: int
    n_cols: int


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip() == ""
    if isinstance(value, float) and math.isnan(value):
        return True
    return False


def to_float(value: Any) -> float | None:
    if is_blank(value):
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.replace(",", "").strip().rstrip("%")
        if cleaned == "":
            return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def ensure_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if isinstance(values, Sequence) and not isinstance(values, (str, bytes)):
        values_list = list(values)
        if not values_list:
            return []
        if isinstance(values_list[0], Sequence) and not isinstance(values_list[0], (str, bytes)):
            return [list(row) for row in values_list]
        return [values_list]
    return [[values]]


def get_snapshot(sheet: Any) -> SheetSnapshot:
    used = sheet.used_range
    values = ensure_2d(used.value)
    return SheetSnapshot(
        values=values,
        top_row=used.row,
        left_col=used.column,
        n_rows=len(values),
        n_cols=len(values[0]) if values else 0,
    )


def snapshot_value(snapshot: SheetSnapshot, row: int, col: int) -> Any:
    r_idx = row - snapshot.top_row
    c_idx = col - snapshot.left_col
    if r_idx < 0 or c_idx < 0:
        return None
    if r_idx >= snapshot.n_rows or c_idx >= snapshot.n_cols:
        return None
    return snapshot.values[r_idx][c_idx]


def snapshot_row_values(snapshot: SheetSnapshot, row: int, start_col: int, count: int) -> list[Any]:
    return [snapshot_value(snapshot, row, start_col + idx) for idx in range(count)]


def find_anchor(snapshot: SheetSnapshot, target: str = "max") -> tuple[int, int]:
    norm_target = target.strip().lower()
    for r_idx, row_vals in enumerate(snapshot.values):
        for c_idx, cell_val in enumerate(row_vals):
            text = normalize_text(cell_val)
            if text == norm_target:
                return snapshot.top_row + r_idx, snapshot.left_col + c_idx

    for r_idx, row_vals in enumerate(snapshot.values):
        for c_idx, cell_val in enumerate(row_vals):
            text = normalize_text(cell_val)
            if norm_target in text:
                return snapshot.top_row + r_idx, snapshot.left_col + c_idx

    raise ValueError(f"Could not locate '{target}' anchor")


def find_label_row(
    snapshot: SheetSnapshot,
    anchor_row: int,
    anchor_col: int,
    keywords: Iterable[str],
    row_window: int = 80,
    col_lookback: int = 12,
) -> int | None:
    key_tokens = [k.lower() for k in keywords]
    start_row = max(snapshot.top_row, anchor_row - row_window)
    end_row = min(snapshot.top_row + snapshot.n_rows - 1, anchor_row + row_window)
    start_col = max(snapshot.left_col, anchor_col - col_lookback)
    end_col = min(snapshot.left_col + snapshot.n_cols - 1, anchor_col)

    best_row: int | None = None
    best_distance = float("inf")

    for abs_row in range(start_row, end_row + 1):
        for abs_col in range(start_col, end_col + 1):
            text = normalize_text(snapshot_value(snapshot, abs_row, abs_col))
            if not text:
                continue
            if any(token in text for token in key_tokens):
                distance = abs(abs_row - anchor_row)
                if distance < best_distance:
                    best_distance = distance
                    best_row = abs_row
                break
    return best_row


def safe_close_workbook(workbook: Any) -> None:
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


def parse_model_meta(file_name: str) -> ModelMeta:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]
    if len(parts) < 3:
        raise ValueError("filename does not include expected ' - ticker - period' format")

    ticker = parts[1]
    raw_period = parts[2].split("_")[0].strip()
    match = re.fullmatch(r"(Early|Mid|Late)([A-Za-z]+)(\d{4})", raw_period)
    if not match:
        raise ValueError(f"model period token '{raw_period}' is not in Early/Mid/LateMonthYYYY format")

    period_prefix = match.group(1)
    month_token = match.group(2)
    year = int(match.group(3))

    day_lookup = {"Early": 5, "Mid": 15, "Late": 25}
    month_num = month_from_token(month_token)
    model_day = day_lookup[period_prefix]
    model_date = date(year, month_num, model_day).isoformat()
    month_abbr = calendar.month_abbr[month_num]
    model_period = f"{period_prefix}{month_abbr}_{year}"
    model = f"{ticker}_{model_period}"

    return ModelMeta(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def month_from_token(month_token: str) -> int:
    token = month_token.strip().lower()
    for month_number in range(1, 13):
        if token == calendar.month_abbr[month_number].lower():
            return month_number
        if token == calendar.month_name[month_number].lower():
            return month_number
        if token.startswith(calendar.month_abbr[month_number].lower()):
            return month_number
    raise ValueError(f"unrecognized month token '{month_token}'")


def next_output_path(input_path: Path, out_dir: Path) -> Path:
    prefix = f"{input_path.name}_PARAM"
    candidate = out_dir / f"{prefix}.xlsx"
    if not candidate.exists():
        return candidate

    counter = 1
    while True:
        candidate = out_dir / f"{prefix}.{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


def set_formula2(cell_or_range: Any, formula: str) -> None:
    try:
        cell_or_range.formula2 = formula
    except Exception:
        cell_or_range.formula = formula


def make_signature(values: Iterable[Any]) -> tuple[Any, ...]:
    normalized: list[Any] = []
    for item in values:
        if isinstance(item, bool):
            normalized.append(item)
            continue
        if isinstance(item, (int, float)) and not (isinstance(item, float) and math.isnan(item)):
            normalized.append(round(float(item), 8))
            continue
        if isinstance(item, str):
            normalized.append(item.strip().lower())
            continue
        normalized.append(item)
    return tuple(normalized)


def process_empirical_sheet(workbook: Any, meta: ModelMeta, source_file: str) -> list[dict[str, Any]]:
    try:
        sheet = workbook.sheets["Empirical Model"]
    except Exception:
        return []

    snapshot = get_snapshot(sheet)
    if snapshot.n_rows == 0 or snapshot.n_cols == 0:
        return []

    anchor_row, anchor_col = find_anchor(snapshot, "max")
    first_data_col = anchor_col + 1

    max_row = anchor_row
    min_row = find_label_row(snapshot, anchor_row, anchor_col, ["min"]) or (anchor_row + 1)
    forecast_row = find_label_row(
        snapshot,
        anchor_row,
        anchor_col,
        ["estimated total sold", "est total sold", "forecast total", "tot fcst"],
    ) or (anchor_row - 2)
    actual_row = find_label_row(
        snapshot,
        anchor_row,
        anchor_col,
        ["reported sales", "actual sales", "actual"],
    ) or (anchor_row - 1)
    num_quarters_row = find_label_row(
        snapshot,
        anchor_row,
        anchor_col,
        ["num quarters", "quarters used", "# quarters", "n quarters"],
    ) or (anchor_row - 6)
    last_quarter_row = find_label_row(
        snapshot,
        anchor_row,
        anchor_col,
        ["last quarter used", "last quarter", "last qtr"],
    ) or (anchor_row - 5)
    quarterly_sales_row = find_label_row(
        snapshot,
        anchor_row,
        anchor_col,
        ["quarterly sales", "quarter sales", "qtr sales", "estimated total sold"],
    ) or forecast_row
    reported_sales_row = find_label_row(
        snapshot,
        anchor_row,
        anchor_col,
        ["reported sales", "actual sales", "actual"],
    ) or actual_row
    growth_rate_row = find_label_row(
        snapshot,
        anchor_row,
        anchor_col,
        ["growth rate", "growth %", "growth_rate"],
    )
    sales_captured_row = find_label_row(
        snapshot,
        anchor_row,
        anchor_col,
        ["sales captured in db", "captured in db", "sales captured"],
    )

    max_values = snapshot_row_values(snapshot, max_row, first_data_col, N_QUARTERS)
    min_values = snapshot_row_values(snapshot, min_row, first_data_col, N_QUARTERS)
    forecast_values = snapshot_row_values(snapshot, forecast_row, first_data_col, N_QUARTERS)
    actual_values = snapshot_row_values(snapshot, actual_row, first_data_col, N_QUARTERS)
    num_quarters_values = snapshot_row_values(snapshot, num_quarters_row, first_data_col, N_QUARTERS)
    last_quarter_values = snapshot_row_values(snapshot, last_quarter_row, first_data_col, N_QUARTERS)
    quarterly_sales_values = snapshot_row_values(snapshot, quarterly_sales_row, first_data_col, N_QUARTERS)
    reported_sales_values = snapshot_row_values(snapshot, reported_sales_row, first_data_col, N_QUARTERS)
    growth_rate_values = (
        snapshot_row_values(snapshot, growth_rate_row, first_data_col, N_QUARTERS)
        if growth_rate_row is not None
        else [None] * N_QUARTERS
    )
    sales_captured_values = (
        snapshot_row_values(snapshot, sales_captured_row, first_data_col, N_QUARTERS)
        if sales_captured_row is not None
        else [None] * N_QUARTERS
    )

    # Temporary R1C1 formulas for avg penetration %, then one calculate call.
    avg_penetration_values = [None] * N_QUARTERS
    if reported_sales_row is not None and quarterly_sales_row is not None:
        scratch_row = snapshot.top_row + snapshot.n_rows + 2
        end_col = first_data_col + N_QUARTERS - 1
        scratch_range = sheet.range((scratch_row, first_data_col), (scratch_row, end_col))
        dr_reported = reported_sales_row - scratch_row
        dr_quarterly = quarterly_sales_row - scratch_row
        avg_formula = f'=IFERROR(R[{dr_reported}]C/R[{dr_quarterly}]C,"")'
        set_formula2(scratch_range, avg_formula)
        workbook.app.calculate()
        avg_penetration_values = ensure_2d([scratch_range.value])[0]
        scratch_range.clear_contents()

    rows: list[dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        forecast_max = max_values[idx]
        forecast_min = min_values[idx]
        forecast_value = forecast_values[idx]
        actual_value = actual_values[idx]
        quarterly_sales = quarterly_sales_values[idx]
        reported_sales = reported_sales_values[idx]

        if all(
            is_blank(v)
            for v in [forecast_max, forecast_min, forecast_value, actual_value, quarterly_sales, reported_sales]
        ):
            continue

        max_num = to_float(forecast_max)
        min_num = to_float(forecast_min)
        range_width = None
        if max_num is not None and min_num is not None:
            range_width = max_num - min_num

        num_quarters_used = num_quarters_values[idx]
        if is_blank(num_quarters_used):
            num_quarters_used = idx + 1

        avg_penetration = avg_penetration_values[idx] if idx < len(avg_penetration_values) else None

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_values[idx],
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_values[idx],
                "sales_captured_in_db_pct": sales_captured_values[idx],
                "source_file": source_file,
            }
        )

    return rows


def process_regression_sheet(workbook: Any, meta: ModelMeta, source_file: str) -> list[dict[str, Any]]:
    try:
        sheet = workbook.sheets["Regression Model"]
    except Exception:
        return []

    snapshot = get_snapshot(sheet)
    if snapshot.n_rows == 0 or snapshot.n_cols == 0:
        return []

    anchor_row, anchor_col = find_anchor(snapshot, "max")
    first_data_col = anchor_col + 1
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    max_row = anchor_row
    min_row = find_label_row(snapshot, anchor_row, anchor_col, ["min"]) or (anchor_row + 1)
    num_quarters_row = find_label_row(
        snapshot,
        anchor_row,
        anchor_col,
        ["num quarters", "quarters used", "# quarters", "n quarters"],
    ) or (anchor_row - 6)
    forecast_row = find_label_row(
        snapshot,
        anchor_row,
        anchor_col,
        ["tot fcst w/o sa", "tot fcst wo sa", "forecast total", "tot fcst without sa"],
    ) or (anchor_row - 2)
    actual_row = find_label_row(snapshot, anchor_row, anchor_col, ["actual", "reported sales"])

    max_values = snapshot_row_values(snapshot, max_row, first_data_col, N_QUARTERS)
    min_values = snapshot_row_values(snapshot, min_row, first_data_col, N_QUARTERS)
    num_quarters_values = snapshot_row_values(snapshot, num_quarters_row, first_data_col, N_QUARTERS)
    forecast_values = snapshot_row_values(snapshot, forecast_row, first_data_col, N_QUARTERS)
    actual_values = (
        snapshot_row_values(snapshot, actual_row, first_data_col, N_QUARTERS)
        if actual_row is not None
        else [None] * N_QUARTERS
    )

    numeric_rows: list[int] = []
    for r_idx in range(snapshot.n_rows):
        abs_row = snapshot.top_row + r_idx
        x_val = snapshot_value(snapshot, abs_row, x_col)
        y_val = snapshot_value(snapshot, abs_row, y_col)
        if to_float(x_val) is not None and to_float(y_val) is not None:
            numeric_rows.append(abs_row)
    numeric_rows = sorted(set(numeric_rows))

    intercept_values = [None] * N_QUARTERS
    slope_values = [None] * N_QUARTERS
    if len(numeric_rows) >= 2:
        end_row = numeric_rows[-1]
        scratch_intercept_row = snapshot.top_row + snapshot.n_rows + 4
        scratch_slope_row = scratch_intercept_row + 1

        for idx in range(N_QUARTERS):
            raw_num_quarters = num_quarters_values[idx]
            default_quarters = idx + 1
            quarters = int(to_float(raw_num_quarters) or default_quarters)
            quarters = max(2, min(quarters, len(numeric_rows)))
            start_row = end_row - quarters + 1

            intercept_formula = (
                f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},'
                f'R{start_row}C{x_col}:R{end_row}C{x_col}),"")'
            )
            slope_formula = (
                f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},'
                f'R{start_row}C{x_col}:R{end_row}C{x_col}),"")'
            )
            set_formula2(sheet.range((scratch_intercept_row, first_data_col + idx)), intercept_formula)
            set_formula2(sheet.range((scratch_slope_row, first_data_col + idx)), slope_formula)

        workbook.app.calculate()
        intercept_values = sheet.range(
            (scratch_intercept_row, first_data_col),
            (scratch_intercept_row, first_data_col + N_QUARTERS - 1),
        ).value
        slope_values = sheet.range(
            (scratch_slope_row, first_data_col),
            (scratch_slope_row, first_data_col + N_QUARTERS - 1),
        ).value

        intercept_values = (
            list(intercept_values)
            if isinstance(intercept_values, Sequence) and not isinstance(intercept_values, (str, bytes))
            else [intercept_values]
        )
        slope_values = (
            list(slope_values)
            if isinstance(slope_values, Sequence) and not isinstance(slope_values, (str, bytes))
            else [slope_values]
        )

        sheet.range(
            (scratch_intercept_row, first_data_col),
            (scratch_slope_row, first_data_col + N_QUARTERS - 1),
        ).clear_contents()

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None
    for idx in range(N_QUARTERS):
        forecast_max = max_values[idx]
        forecast_min = min_values[idx]
        forecast_value = forecast_values[idx]
        actual_value = actual_values[idx] if idx < len(actual_values) else None
        intercept = intercept_values[idx] if idx < len(intercept_values) else None
        slope = slope_values[idx] if idx < len(slope_values) else None

        if all(is_blank(v) for v in [forecast_max, forecast_min, forecast_value, intercept, slope]):
            continue

        max_num = to_float(forecast_max)
        min_num = to_float(forecast_min)
        range_width = None
        if max_num is not None and min_num is not None:
            range_width = max_num - min_num

        num_quarters_used = num_quarters_values[idx]
        if is_blank(num_quarters_used):
            num_quarters_used = idx + 1

        signature = make_signature(
            [num_quarters_used, forecast_value, forecast_max, forecast_min, intercept, slope]
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": "" if is_blank(actual_value) else actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def format_sheet(ws: Any, headers: list[str], rows: list[dict[str, Any]]) -> None:
    ws.append(headers)
    for row in rows:
        ws.append([row.get(col, "") for col in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            value_len = len(str(value))
            if value_len > max_len:
                max_len = value_len
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 48)


def write_output_workbook(
    target_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    empirical_ws = workbook.create_sheet("empirical_candidates")
    regression_ws = workbook.create_sheet("regression_candidates")

    format_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    format_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    workbook.save(target_path)


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    out_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise SystemExit(f"Input directory does not exist: {input_path}")

    out_path.mkdir(parents=True, exist_ok=True)
    output_file = next_output_path(input_path, out_path)

    source_files = sorted(input_path.iterdir(), key=lambda p: p.name.lower())

    processed_files = 0
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        app.calculation = "manual"
    except Exception:
        pass

    try:
        for file_path in source_files:
            if not file_path.is_file():
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped file: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped file: {file_path.name} (not .xlsx)")
                continue

            try:
                meta = parse_model_meta(file_path.name)
            except Exception as exc:
                print(f"Skipped file: {file_path.name} (name parse error: {exc})")
                continue

            workbook = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(process_empirical_sheet(workbook, meta, file_path.name))
                regression_rows.extend(process_regression_sheet(workbook, meta, file_path.name))
                processed_files += 1
                print(f"Processed file: {file_path.name}")
            except Exception as exc:
                print(f"Skipped file: {file_path.name} (processing error: {exc})")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        app.quit()

    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"Output path: {output_file}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

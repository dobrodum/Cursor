from __future__ import annotations

import calendar
import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -----------------------------
# User configuration
# -----------------------------
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")


N_QUARTERS = 10
EMPIRICAL_MODEL_SHEET = "Empirical Model"
REGRESSION_MODEL_SHEET = "Regression Model"

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

PERIOD_TO_DAY = {"Early": 5, "Mid": 15, "Late": 25}
MONTH_NAME_TO_NUM = {
    calendar.month_abbr[i].lower(): i for i in range(1, 13) if calendar.month_abbr[i]
}
MONTH_NAME_TO_NUM.update(
    {calendar.month_name[i].lower(): i for i in range(1, 13) if calendar.month_name[i]}
)


def ensure_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = re.sub(r"\s+", " ", str(value).strip().lower())
    return text


def clean_cell(value: Any) -> Any:
    if isinstance(value, str):
        stripped = value.strip()
        if not stripped or stripped.startswith("#"):
            return None
        return stripped
    return value


def to_float(value: Any) -> Optional[float]:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
            return None
        is_negative = text.startswith("(") and text.endswith(")")
        if is_negative:
            text = text[1:-1]
        pct = text.endswith("%")
        if pct:
            text = text[:-1]
        try:
            parsed = float(text)
        except ValueError:
            return None
        if is_negative:
            parsed *= -1
        if pct:
            parsed /= 100
        return parsed
    return None


def range_width(max_value: Any, min_value: Any) -> Optional[float]:
    max_num = to_float(max_value)
    min_num = to_float(min_value)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


def row_is_empty(values: Sequence[Any]) -> bool:
    for value in values:
        if value is None:
            continue
        if isinstance(value, str) and not value.strip():
            continue
        return False
    return True


def as_column(values: Any) -> List[Any]:
    matrix = ensure_2d(values)
    result: List[Any] = []
    for row in matrix:
        result.append(clean_cell(row[0] if row else None))
    return result


def iter_source_files(path: Path) -> Iterable[Path]:
    for item in sorted(path.iterdir()):
        if not item.is_file():
            print(f"Skipped {item.name}: not a file")
            continue
        if item.name.startswith("~"):
            print(f"Skipped {item.name}: temporary file")
            continue
        if item.suffix.lower() != ".xlsx":
            print(f"Skipped {item.name}: not an .xlsx file")
            continue
        yield item


def parse_file_label(file_path: Path) -> Optional[Dict[str, str]]:
    stem = file_path.stem
    parts = [segment.strip() for segment in stem.split(" - ") if segment.strip()]
    if len(parts) < 3:
        return None

    ticker = parts[-2].strip().upper()
    period_token = parts[-1].split("_")[0].strip()
    match = re.match(r"^(Early|Mid|Late)([A-Za-z]+)(\d{4})$", period_token, flags=re.IGNORECASE)
    if not match:
        return None

    period_prefix = match.group(1).title()
    month_token = match.group(2).lower()
    year = int(match.group(3))

    month_num = MONTH_NAME_TO_NUM.get(month_token)
    if month_num is None:
        month_num = MONTH_NAME_TO_NUM.get(month_token[:3])
    if month_num is None:
        return None

    month_abbrev = calendar.month_abbr[month_num]
    model_period = f"{period_prefix}{month_abbrev}_{year}"
    model_date = date(year, month_num, PERIOD_TO_DAY[period_prefix]).isoformat()
    model = f"{ticker}_{model_period}"
    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def choose_output_path(in_dir: Path, out_dir: Path) -> Path:
    base_name = f"{in_dir.name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = out_dir / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def find_sheet(workbook: xw.main.Book, target_name: str) -> Optional[xw.main.Sheet]:
    target = target_name.strip().lower()
    for sheet in workbook.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def scan_sheet(sheet: xw.main.Sheet, anchor_text: str = "max") -> Optional[Dict[str, Any]]:
    used = sheet.used_range
    values = ensure_2d(used.value)
    if not values:
        return None

    start_row = used.row
    start_col = used.column
    col_count = max((len(row) for row in values), default=0)
    anchor_row: Optional[int] = None
    anchor_col: Optional[int] = None
    target = anchor_text.strip().lower()

    for r_idx, row in enumerate(values):
        for c_idx, value in enumerate(row):
            if isinstance(value, str) and value.strip().lower() == target:
                anchor_row = start_row + r_idx
                anchor_col = start_col + c_idx
                break
        if anchor_row is not None and anchor_col is not None:
            break

    if anchor_row is None or anchor_col is None:
        return None

    header_entries: List[Tuple[str, int]] = []
    header_row_idx = anchor_row - start_row
    if 0 <= header_row_idx < len(values):
        header_row = values[header_row_idx]
        for c_idx, value in enumerate(header_row):
            normalized = normalize_text(value)
            if normalized:
                header_entries.append((normalized, start_col + c_idx))

    return {
        "anchor_row": anchor_row,
        "anchor_col": anchor_col,
        "header_entries": header_entries,
        "last_col": start_col + max(col_count - 1, 0),
    }


def resolve_column(
    anchor_col: int,
    header_entries: Sequence[Tuple[str, int]],
    default_offset: int,
    keywords: Sequence[str],
) -> int:
    lowered = [keyword.lower() for keyword in keywords]
    matches: List[Tuple[int, int]] = []
    for header_text, column in header_entries:
        if any(keyword in header_text for keyword in lowered):
            matches.append((abs(column - anchor_col), column))

    if matches:
        matches.sort()
        return matches[0][1]
    return max(1, anchor_col + default_offset)


def read_block(
    sheet: xw.main.Sheet,
    start_row: int,
    end_row: int,
    start_col: int,
    end_col: int,
) -> Tuple[List[List[Any]], int]:
    start_col = max(1, start_col)
    end_col = max(start_col, end_col)
    if end_row < start_row:
        return [], start_col
    values = ensure_2d(sheet.range((start_row, start_col), (end_row, end_col)).value)
    return values, start_col


def block_value(
    block: Sequence[Sequence[Any]],
    block_start_row: int,
    block_start_col: int,
    row: int,
    col: int,
) -> Any:
    r_idx = row - block_start_row
    c_idx = col - block_start_col
    if r_idx < 0 or c_idx < 0 or r_idx >= len(block):
        return None
    row_values = block[r_idx]
    if c_idx >= len(row_values):
        return None
    return clean_cell(row_values[c_idx])


def safe_close_workbook(workbook: xw.main.Book) -> None:
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

    try:
        workbook.close()
    except Exception:
        pass


def normalize_for_compare(value: Any) -> Any:
    as_num = to_float(value)
    if as_num is not None:
        return round(as_num, 10)
    if isinstance(value, str):
        return value.strip().lower()
    return value


def process_empirical_sheet(
    workbook: xw.main.Book,
    sheet: xw.main.Sheet,
    labels: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    scan = scan_sheet(sheet, anchor_text="max")
    if not scan:
        print(f"Skipped {source_file} empirical: 'max' anchor not found")
        return []

    anchor_row = scan["anchor_row"]
    anchor_col = scan["anchor_col"]
    headers = scan["header_entries"]
    data_start_row = anchor_row + 1
    data_end_row = data_start_row + N_QUARTERS - 1

    col_forecast_max = resolve_column(anchor_col, headers, 0, ["max"])
    col_forecast_min = resolve_column(anchor_col, headers, 1, ["min"])
    col_forecast_value = resolve_column(
        anchor_col,
        headers,
        -1,
        ["estimated total sold", "tot fcst", "forecast", "fcst"],
    )
    col_actual_value = resolve_column(
        anchor_col,
        headers,
        2,
        ["reported sales", "actual sales", "actual", "reported"],
    )
    col_penetration = resolve_column(anchor_col, headers, -2, ["penetration"])
    col_quarterly_sales = resolve_column(anchor_col, headers, -5, ["quarterly sales", "quarter sales"])
    col_reported_sales = resolve_column(anchor_col, headers, 2, ["reported sales"])
    col_growth_rate = resolve_column(anchor_col, headers, -4, ["growth rate", "growth"])
    col_sales_captured = resolve_column(
        anchor_col,
        headers,
        -3,
        ["sales captured in db", "captured in db", "captured"],
    )
    col_last_quarter = resolve_column(anchor_col, headers, -6, ["last quarter", "last qtr"])

    needed_cols = [
        col_forecast_max,
        col_forecast_min,
        col_forecast_value,
        col_actual_value,
        col_penetration,
        col_quarterly_sales,
        col_reported_sales,
        col_growth_rate,
        col_sales_captured,
        col_last_quarter,
    ]
    block_start_col = max(1, min(needed_cols))
    block_end_col = max(needed_cols)
    block, block_start_col = read_block(
        sheet=sheet,
        start_row=data_start_row,
        end_row=data_end_row,
        start_col=block_start_col,
        end_col=block_end_col,
    )

    # Use a far-right temporary column for R1C1 formula2 calculations.
    scratch_col = max(scan["last_col"] + 3, 200)
    for idx in range(N_QUARTERS):
        n_quarters = idx + 1
        formula = (
            f"=AVERAGE(R{data_start_row}C{col_penetration}:"
            f"R{data_start_row + n_quarters - 1}C{col_penetration})"
        )
        sheet.range((data_start_row + idx, scratch_col)).formula2 = formula

    workbook.app.calculate()
    avg_penetration_values = as_column(
        sheet.range((data_start_row, scratch_col), (data_end_row, scratch_col)).value
    )
    sheet.range((data_start_row, scratch_col), (data_end_row, scratch_col)).value = None

    rows: List[Dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        row_num = data_start_row + idx
        n_quarters = idx + 1

        forecast_max = block_value(block, data_start_row, block_start_col, row_num, col_forecast_max)
        forecast_min = block_value(block, data_start_row, block_start_col, row_num, col_forecast_min)
        forecast_value = block_value(block, data_start_row, block_start_col, row_num, col_forecast_value)
        actual_value = block_value(block, data_start_row, block_start_col, row_num, col_actual_value)
        quarterly_sales = block_value(block, data_start_row, block_start_col, row_num, col_quarterly_sales)
        reported_sales = block_value(block, data_start_row, block_start_col, row_num, col_reported_sales)
        growth_rate_pct = block_value(block, data_start_row, block_start_col, row_num, col_growth_rate)
        sales_captured_pct = block_value(block, data_start_row, block_start_col, row_num, col_sales_captured)
        last_quarter_used = block_value(block, data_start_row, block_start_col, row_num, col_last_quarter)

        avg_penetration_pct = (
            avg_penetration_values[idx] if idx < len(avg_penetration_values) else None
        )
        if avg_penetration_pct is None:
            avg_penetration_pct = block_value(
                block, data_start_row, block_start_col, row_num, col_penetration
            )

        if row_is_empty(
            [
                forecast_max,
                forecast_min,
                forecast_value,
                actual_value,
                avg_penetration_pct,
                quarterly_sales,
                reported_sales,
                growth_rate_pct,
                sales_captured_pct,
            ]
        ):
            continue

        rows.append(
            {
                "model": labels["model"],
                "ticker": labels["ticker"],
                "model_period": labels["model_period"],
                "model_date": labels["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": n_quarters,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width(forecast_max, forecast_min),
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
    labels: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    scan = scan_sheet(sheet, anchor_text="max")
    if not scan:
        print(f"Skipped {source_file} regression: 'max' anchor not found")
        return []

    anchor_row = scan["anchor_row"]
    anchor_col = scan["anchor_col"]
    headers = scan["header_entries"]
    data_start_row = anchor_row + 1
    data_end_row = data_start_row + N_QUARTERS - 1

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    col_num_quarters = resolve_column(
        anchor_col,
        headers,
        -2,
        ["num quarters used", "num quarters", "# quarters", "quarters used"],
    )
    col_forecast_wo_sa = resolve_column(
        anchor_col,
        headers,
        -1,
        ["tot fcst w/o sa", "tot fcst wo sa", "fcst w/o sa", "forecast w/o sa", "tot fcst"],
    )
    col_actual_value = resolve_column(anchor_col, headers, 2, ["actual", "reported sales"])
    col_forecast_max = resolve_column(anchor_col, headers, 0, ["max"])
    col_forecast_min = resolve_column(anchor_col, headers, 1, ["min"])

    needed_cols = [
        col_num_quarters,
        col_forecast_wo_sa,
        col_actual_value,
        col_forecast_max,
        col_forecast_min,
    ]
    block_start_col = max(1, min(needed_cols))
    block_end_col = max(needed_cols)
    block, block_start_col = read_block(
        sheet=sheet,
        start_row=data_start_row,
        end_row=data_end_row,
        start_col=block_start_col,
        end_col=block_end_col,
    )

    scratch_col = max(scan["last_col"] + 3, 220)
    for idx in range(N_QUARTERS):
        n_quarters = idx + 1
        intercept_formula = (
            "=IFERROR(INTERCEPT("
            f"R{data_start_row}C{y_col}:R{data_start_row + n_quarters - 1}C{y_col},"
            f"R{data_start_row}C{x_col}:R{data_start_row + n_quarters - 1}C{x_col}"
            '),"")'
        )
        slope_formula = (
            "=IFERROR(SLOPE("
            f"R{data_start_row}C{y_col}:R{data_start_row + n_quarters - 1}C{y_col},"
            f"R{data_start_row}C{x_col}:R{data_start_row + n_quarters - 1}C{x_col}"
            '),"")'
        )
        target_row = data_start_row + idx
        sheet.range((target_row, scratch_col)).formula2 = intercept_formula
        sheet.range((target_row, scratch_col + 1)).formula2 = slope_formula

    workbook.app.calculate()
    intercept_values = as_column(
        sheet.range((data_start_row, scratch_col), (data_end_row, scratch_col)).value
    )
    slope_values = as_column(
        sheet.range((data_start_row, scratch_col + 1), (data_end_row, scratch_col + 1)).value
    )
    sheet.range((data_start_row, scratch_col), (data_end_row, scratch_col + 1)).value = None

    rows: List[Dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        row_num = data_start_row + idx
        fallback_quarters = idx + 1

        num_quarters_used = block_value(block, data_start_row, block_start_col, row_num, col_num_quarters)
        if num_quarters_used is None:
            num_quarters_used = fallback_quarters

        forecast_value = block_value(block, data_start_row, block_start_col, row_num, col_forecast_wo_sa)
        actual_value = block_value(block, data_start_row, block_start_col, row_num, col_actual_value)
        forecast_max = block_value(block, data_start_row, block_start_col, row_num, col_forecast_max)
        forecast_min = block_value(block, data_start_row, block_start_col, row_num, col_forecast_min)
        intercept = intercept_values[idx] if idx < len(intercept_values) else None
        slope = slope_values[idx] if idx < len(slope_values) else None

        if row_is_empty([forecast_value, forecast_max, forecast_min, intercept, slope]):
            continue

        row = {
            "model": labels["model"],
            "ticker": labels["ticker"],
            "model_period": labels["model_period"],
            "model_date": labels["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        if idx == N_QUARTERS - 1 and rows:
            previous = rows[-1]
            current_signature = (
                normalize_for_compare(row["num_quarters_used"]),
                normalize_for_compare(row["forecast_value"]),
                normalize_for_compare(row["forecast_max"]),
                normalize_for_compare(row["forecast_min"]),
                normalize_for_compare(row["intercept"]),
                normalize_for_compare(row["slope"]),
            )
            previous_signature = (
                normalize_for_compare(previous["num_quarters_used"]),
                normalize_for_compare(previous["forecast_value"]),
                normalize_for_compare(previous["forecast_max"]),
                normalize_for_compare(previous["forecast_min"]),
                normalize_for_compare(previous["intercept"]),
                normalize_for_compare(previous["slope"]),
            )
            if current_signature == previous_signature:
                continue

        rows.append(row)

    return rows


def set_column_widths(sheet, columns: Sequence[str]) -> None:
    for idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row in range(2, sheet.max_row + 1):
            value = sheet.cell(row=row, column=idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        sheet.column_dimensions[get_column_letter(idx)].width = min(max_len + 2, 40)


def write_table(sheet, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    sheet.append(list(columns))
    for col_idx in range(1, len(columns) + 1):
        sheet.cell(row=1, column=col_idx).font = Font(bold=True)

    for row in rows:
        sheet.append([row.get(column) for column in columns])

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(1, sheet.max_row)}"
    set_column_widths(sheet, columns)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    empirical_sheet = workbook.active
    empirical_sheet.title = "empirical_candidates"
    regression_sheet = workbook.create_sheet("regression_candidates")

    write_table(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)
    write_table(regression_sheet, REGRESSION_COLUMNS, regression_rows)
    workbook.save(output_path)


def main() -> None:
    in_dir = Path(input_dir).expanduser().resolve()
    out_dir = Path(output_dir).expanduser().resolve()

    if not in_dir.exists() or not in_dir.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {in_dir}")
    out_dir.mkdir(parents=True, exist_ok=True)

    output_path = choose_output_path(in_dir, out_dir)
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

        for file_path in iter_source_files(in_dir):
            labels = parse_file_label(file_path)
            if not labels:
                print(f"Skipped {file_path.name}: filename pattern not recognized")
                continue

            workbook: Optional[xw.main.Book] = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                file_empirical_rows: List[Dict[str, Any]] = []
                file_regression_rows: List[Dict[str, Any]] = []

                empirical_sheet = find_sheet(workbook, EMPIRICAL_MODEL_SHEET)
                if empirical_sheet is None:
                    print(f"Skipped {file_path.name} empirical: sheet '{EMPIRICAL_MODEL_SHEET}' not found")
                else:
                    file_empirical_rows = process_empirical_sheet(
                        workbook=workbook,
                        sheet=empirical_sheet,
                        labels=labels,
                        source_file=file_path.name,
                    )

                regression_sheet = find_sheet(workbook, REGRESSION_MODEL_SHEET)
                if regression_sheet is None:
                    print(f"Skipped {file_path.name} regression: sheet '{REGRESSION_MODEL_SHEET}' not found")
                else:
                    file_regression_rows = process_regression_sheet(
                        workbook=workbook,
                        sheet=regression_sheet,
                        labels=labels,
                        source_file=file_path.name,
                    )

                empirical_rows.extend(file_empirical_rows)
                regression_rows.extend(file_regression_rows)
                processed_files += 1
                print(
                    f"Processed {file_path.name}: "
                    f"empirical_rows={len(file_empirical_rows)}, "
                    f"regression_rows={len(file_regression_rows)}"
                )
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing failed ({exc})")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        if app is not None:
            try:
                app.display_alerts = True
                app.screen_updating = True
            except Exception:
                pass
            app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)
    print(f"Output workbook: {output_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

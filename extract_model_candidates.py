#!/usr/bin/env python3
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure these two paths before running.
input_dir = Path("./input")
output_dir = Path("./output")

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

EMPIRICAL_FALLBACK_OFFSETS = {
    "num_quarters_used": -8,
    "last_quarter_used": -7,
    "forecast_value": -4,
    "actual_value": -3,
    "forecast_max": 0,
    "forecast_min": 1,
    "avg_penetration_pct": -5,
    "quarterly_sales": -11,
    "reported_sales": -10,
    "growth_rate_pct": -9,
    "sales_captured_in_db_pct": -6,
}

REGRESSION_FALLBACK_OFFSETS = {
    "num_quarters_used": -8,
    "forecast_value": -1,
    "actual_value": -2,
    "forecast_max": 0,
    "forecast_min": 1,
}

N_QUARTERS = 10


@dataclass
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class AnchorContext:
    row: int
    col: int
    used_first_row: int
    used_last_row: int
    used_first_col: int
    used_last_col: int


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(text.split())


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def to_number(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        stripped = value.strip()
        if not stripped:
            return None
        stripped = stripped.replace(",", "")
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


def safe_subtract(a: Any, b: Any) -> Optional[float]:
    a_num = to_number(a)
    b_num = to_number(b)
    if a_num is None or b_num is None:
        return None
    return a_num - b_num


def safe_int(value: Any, default: int) -> int:
    as_num = to_number(value)
    if as_num is None:
        return default
    return max(1, int(round(as_num)))


def parse_file_labels(file_path: Path) -> FileLabels:
    pattern = re.compile(
        r".*?-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*(?P<period>Early|Mid|Late)"
        r"(?P<month>[A-Za-z]{3,9})(?P<year>\d{4})",
        re.IGNORECASE,
    )
    stem = file_path.stem
    match = pattern.search(stem)
    if not match:
        ticker_guess = stem.split("-")
        ticker = ticker_guess[1].strip() if len(ticker_guess) > 1 else "UNKNOWN"
        model_period = "UNKNOWN_PERIOD"
        return FileLabels(
            model=f"{ticker}_{model_period}",
            ticker=ticker,
            model_period=model_period,
            model_date="",
        )

    ticker = match.group("ticker").upper()
    period_prefix = match.group("period").title()
    month_token = match.group("month").title()
    month_abbrev = month_token[:3]
    year = int(match.group("year"))
    day_by_period = {"early": 5, "mid": 15, "late": 25}
    day = day_by_period[period_prefix.lower()]
    month_number = datetime.strptime(month_abbrev, "%b").month
    model_period = f"{period_prefix}{month_abbrev}_{year}"
    model_date = date(year, month_number, day).isoformat()
    return FileLabels(
        model=f"{ticker}_{model_period}",
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def choose_output_path(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_path.name}_PARAM.xlsx"
    candidate = output_path / base_name
    if not candidate.exists():
        return candidate
    index = 1
    while True:
        candidate = output_path / f"{input_path.name}_PARAM.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def get_sheet_case_insensitive(wb: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    expected = sheet_name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == expected:
            return sheet
    return None


def find_anchor_context(sheet: xw.Sheet, anchor_text: str = "max") -> Optional[AnchorContext]:
    used = sheet.used_range
    values = used.options(ndim=2).value
    if not values:
        return None
    first_row = used.row
    first_col = used.column
    last_row = first_row + len(values) - 1
    last_col = first_col + len(values[0]) - 1

    candidates: List[Tuple[int, int, int]] = []
    for row_idx, row in enumerate(values):
        for col_idx, raw_value in enumerate(row):
            if normalize_text(raw_value) != anchor_text:
                continue
            score = 0
            right = normalize_text(row[col_idx + 1]) if col_idx + 1 < len(row) else ""
            left = normalize_text(row[col_idx - 1]) if col_idx > 0 else ""
            if right == "min" or left == "min":
                score = 1
            candidates.append((score, first_row + row_idx, first_col + col_idx))

    if not candidates:
        return None

    candidates.sort(key=lambda item: (-item[0], item[1], item[2]))
    _, anchor_row, anchor_col = candidates[0]
    return AnchorContext(
        row=anchor_row,
        col=anchor_col,
        used_first_row=first_row,
        used_last_row=last_row,
        used_first_col=first_col,
        used_last_col=last_col,
    )


def get_header_row_values(
    sheet: xw.Sheet, row: int, first_col: int, last_col: int
) -> Dict[int, str]:
    row_values = (
        sheet.range((row, first_col), (row, last_col)).options(ndim=2).value[0]
    )
    return {first_col + idx: normalize_text(value) for idx, value in enumerate(row_values)}


def find_col_by_keywords(
    header_map: Dict[int, str], keyword_groups: Sequence[Sequence[str]]
) -> Optional[int]:
    for col in sorted(header_map.keys()):
        header_text = header_map[col]
        for group in keyword_groups:
            if all(token in header_text for token in group):
                return col
    return None


def resolve_col(
    anchor_col: int,
    header_map: Dict[int, str],
    fallback_offset: int,
    keyword_groups: Sequence[Sequence[str]],
) -> Optional[int]:
    mapped_col = find_col_by_keywords(header_map, keyword_groups)
    if mapped_col is not None:
        return mapped_col
    fallback_col = anchor_col + fallback_offset
    return fallback_col if fallback_col > 0 else None


def load_block(
    sheet: xw.Sheet, start_row: int, end_row: int, start_col: int, end_col: int
) -> List[List[Any]]:
    if start_col > end_col or start_row > end_row:
        return []
    return sheet.range((start_row, start_col), (end_row, end_col)).options(ndim=2).value


def block_value(
    block: List[List[Any]], row_idx: int, target_col: Optional[int], block_start_col: int
) -> Any:
    if not block or target_col is None:
        return None
    offset = target_col - block_start_col
    if offset < 0:
        return None
    if row_idx < 0 or row_idx >= len(block):
        return None
    row = block[row_idx]
    if offset >= len(row):
        return None
    return row[offset]


def write_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.api.Formula2R1C1 = formula_r1c1
    except Exception:
        try:
            cell.formula2 = formula_r1c1
        except Exception:
            cell.formula = formula_r1c1


def close_workbook_safely(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    for closer in (
        lambda: wb.close(False),
        lambda: wb.api.Close(SaveChanges=False),
        lambda: wb.api.Close(False),
    ):
        try:
            closer()
            return
        except Exception:
            continue


def extract_empirical_rows(
    wb: xw.Book, labels: FileLabels, source_file: str
) -> List[Dict[str, Any]]:
    sheet = get_sheet_case_insensitive(wb, "Empirical Model")
    if sheet is None:
        return []

    anchor = find_anchor_context(sheet, "max")
    if anchor is None:
        return []

    header_map = get_header_row_values(
        sheet, anchor.row, anchor.used_first_col, anchor.used_last_col
    )
    cols = {
        "num_quarters_used": resolve_col(
            anchor.col,
            header_map,
            EMPIRICAL_FALLBACK_OFFSETS["num_quarters_used"],
            (("num", "quarter"), ("quarters", "used"), ("n", "quarters")),
        ),
        "last_quarter_used": resolve_col(
            anchor.col,
            header_map,
            EMPIRICAL_FALLBACK_OFFSETS["last_quarter_used"],
            (("last", "quarter"), ("latest", "quarter")),
        ),
        "forecast_value": resolve_col(
            anchor.col,
            header_map,
            EMPIRICAL_FALLBACK_OFFSETS["forecast_value"],
            (
                ("estimated", "total", "sold"),
                ("est", "total", "sold"),
                ("forecast", "value"),
                ("tot", "fcst"),
            ),
        ),
        "actual_value": resolve_col(
            anchor.col,
            header_map,
            EMPIRICAL_FALLBACK_OFFSETS["actual_value"],
            (("reported", "sales"), ("actual", "sales"), ("reported",)),
        ),
        "forecast_max": anchor.col,
        "forecast_min": resolve_col(
            anchor.col,
            header_map,
            EMPIRICAL_FALLBACK_OFFSETS["forecast_min"],
            (("min",),),
        ),
        "avg_penetration_pct": resolve_col(
            anchor.col,
            header_map,
            EMPIRICAL_FALLBACK_OFFSETS["avg_penetration_pct"],
            (("avg", "penetration"), ("average", "penetration"), ("penetration",)),
        ),
        "quarterly_sales": resolve_col(
            anchor.col,
            header_map,
            EMPIRICAL_FALLBACK_OFFSETS["quarterly_sales"],
            (("quarterly", "sales"), ("qtr", "sales"), ("quarter", "sales")),
        ),
        "reported_sales": resolve_col(
            anchor.col,
            header_map,
            EMPIRICAL_FALLBACK_OFFSETS["reported_sales"],
            (("reported", "sales"), ("actual", "sales")),
        ),
        "growth_rate_pct": resolve_col(
            anchor.col,
            header_map,
            EMPIRICAL_FALLBACK_OFFSETS["growth_rate_pct"],
            (("growth", "rate"), ("growth",)),
        ),
        "sales_captured_in_db_pct": resolve_col(
            anchor.col,
            header_map,
            EMPIRICAL_FALLBACK_OFFSETS["sales_captured_in_db_pct"],
            (
                ("sales", "captured", "db"),
                ("captured", "db"),
                ("in", "db"),
                ("captured",),
            ),
        ),
    }

    columns_to_read = [col for col in cols.values() if col is not None]
    if not columns_to_read:
        return []
    start_col = min(columns_to_read)
    end_col = max(columns_to_read)

    start_row = anchor.row + 1
    end_row = start_row + N_QUARTERS - 1
    block = load_block(sheet, start_row, end_row, start_col, end_col)
    if not block:
        return []

    temp_col = max(anchor.used_last_col, end_col) + 2
    formula_rows: List[int] = []
    candidate_rows: List[Tuple[int, int, int]] = []
    blank_streak = 0

    for idx in range(min(N_QUARTERS, len(block))):
        row_num = start_row + idx
        row_num_quarters = block_value(block, idx, cols["num_quarters_used"], start_col)
        row_forecast = block_value(block, idx, cols["forecast_value"], start_col)
        row_max = block_value(block, idx, cols["forecast_max"], start_col)
        row_min = block_value(block, idx, cols["forecast_min"], start_col)
        if all(is_blank(v) for v in (row_num_quarters, row_forecast, row_max, row_min)):
            blank_streak += 1
            if blank_streak >= 2:
                break
            continue
        blank_streak = 0
        n_quarters = safe_int(row_num_quarters, idx + 1)
        candidate_rows.append((idx, row_num, n_quarters))
        sales_captured_col = cols["sales_captured_in_db_pct"]
        if sales_captured_col is not None:
            rel_col = sales_captured_col - temp_col
            start_rel = -n_quarters + 1
            formula = f"=AVERAGE(R[{start_rel}]C[{rel_col}]:RC[{rel_col}])"
            write_formula2_r1c1(sheet.range((row_num, temp_col)), formula)
            formula_rows.append(row_num)

    if formula_rows:
        wb.app.calculate()

    avg_by_row: Dict[int, Any] = {}
    if formula_rows:
        avg_values = (
            sheet.range((formula_rows[0], temp_col), (formula_rows[-1], temp_col))
            .options(ndim=2)
            .value
        )
        for row_num, value in zip(range(formula_rows[0], formula_rows[-1] + 1), avg_values):
            avg_by_row[row_num] = value[0]

    empirical_rows: List[Dict[str, Any]] = []
    for block_idx, row_num, n_quarters in candidate_rows:
        raw_avg = block_value(block, block_idx, cols["avg_penetration_pct"], start_col)
        if row_num in avg_by_row:
            raw_avg = avg_by_row[row_num]

        forecast_max = block_value(block, block_idx, cols["forecast_max"], start_col)
        forecast_min = block_value(block, block_idx, cols["forecast_min"], start_col)
        row = {
            "model": labels.model,
            "ticker": labels.ticker,
            "model_period": labels.model_period,
            "model_date": labels.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": raw_avg,
            "num_quarters_used": n_quarters,
            "last_quarter_used": block_value(
                block, block_idx, cols["last_quarter_used"], start_col
            ),
            "forecast_value": block_value(
                block, block_idx, cols["forecast_value"], start_col
            ),
            "actual_value": block_value(block, block_idx, cols["actual_value"], start_col),
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": safe_subtract(forecast_max, forecast_min),
            "avg_penetration_pct": raw_avg,
            "quarterly_sales": block_value(
                block, block_idx, cols["quarterly_sales"], start_col
            ),
            "reported_sales": block_value(
                block, block_idx, cols["reported_sales"], start_col
            ),
            "growth_rate_pct": block_value(
                block, block_idx, cols["growth_rate_pct"], start_col
            ),
            "sales_captured_in_db_pct": block_value(
                block, block_idx, cols["sales_captured_in_db_pct"], start_col
            ),
            "source_file": source_file,
        }
        empirical_rows.append(row)

    return empirical_rows


def extract_regression_rows(
    wb: xw.Book, labels: FileLabels, source_file: str
) -> List[Dict[str, Any]]:
    sheet = get_sheet_case_insensitive(wb, "Regression Model")
    if sheet is None:
        return []

    anchor = find_anchor_context(sheet, "max")
    if anchor is None:
        return []

    header_map = get_header_row_values(
        sheet, anchor.row, anchor.used_first_col, anchor.used_last_col
    )
    cols = {
        "num_quarters_used": resolve_col(
            anchor.col,
            header_map,
            REGRESSION_FALLBACK_OFFSETS["num_quarters_used"],
            (("num", "quarter"), ("quarters", "used"), ("n", "quarters")),
        ),
        "forecast_value": resolve_col(
            anchor.col,
            header_map,
            REGRESSION_FALLBACK_OFFSETS["forecast_value"],
            (
                ("tot", "fcst", "w", "o", "sa"),
                ("tot", "fcst", "wo", "sa"),
                ("forecast", "without", "sa"),
                ("tot", "fcst"),
            ),
        ),
        "actual_value": resolve_col(
            anchor.col,
            header_map,
            REGRESSION_FALLBACK_OFFSETS["actual_value"],
            (("actual", "sales"), ("reported", "sales"), ("actual",)),
        ),
        "forecast_max": anchor.col,
        "forecast_min": resolve_col(
            anchor.col,
            header_map,
            REGRESSION_FALLBACK_OFFSETS["forecast_min"],
            (("min",),),
        ),
    }
    y_col = anchor.col - 7
    x_col = anchor.col - 11

    columns_to_read = [col for col in cols.values() if col is not None]
    columns_to_read.extend([x_col, y_col])
    start_col = min(columns_to_read)
    end_col = max(columns_to_read)

    start_row = anchor.row + 1
    end_row = start_row + N_QUARTERS - 1
    block = load_block(sheet, start_row, end_row, start_col, end_col)
    if not block:
        return []

    temp_intercept_col = max(anchor.used_last_col, end_col) + 2
    temp_slope_col = temp_intercept_col + 1
    candidate_rows: List[Tuple[int, int, int]] = []
    blank_streak = 0
    formula_rows: List[int] = []

    for idx in range(min(N_QUARTERS, len(block))):
        row_num = start_row + idx
        num_q_val = block_value(block, idx, cols["num_quarters_used"], start_col)
        forecast_val = block_value(block, idx, cols["forecast_value"], start_col)
        max_val = block_value(block, idx, cols["forecast_max"], start_col)
        min_val = block_value(block, idx, cols["forecast_min"], start_col)
        if all(is_blank(v) for v in (num_q_val, forecast_val, max_val, min_val)):
            blank_streak += 1
            if blank_streak >= 2:
                break
            continue
        blank_streak = 0
        n_quarters = max(2, safe_int(num_q_val, idx + 2))
        candidate_rows.append((idx, row_num, n_quarters))

        y_rel_intercept = y_col - temp_intercept_col
        x_rel_intercept = x_col - temp_intercept_col
        y_rel_slope = y_col - temp_slope_col
        x_rel_slope = x_col - temp_slope_col
        start_rel = -n_quarters + 1

        intercept_formula = (
            f"=INTERCEPT(R[{start_rel}]C[{y_rel_intercept}]:RC[{y_rel_intercept}],"
            f"R[{start_rel}]C[{x_rel_intercept}]:RC[{x_rel_intercept}])"
        )
        slope_formula = (
            f"=SLOPE(R[{start_rel}]C[{y_rel_slope}]:RC[{y_rel_slope}],"
            f"R[{start_rel}]C[{x_rel_slope}]:RC[{x_rel_slope}])"
        )
        write_formula2_r1c1(sheet.range((row_num, temp_intercept_col)), intercept_formula)
        write_formula2_r1c1(sheet.range((row_num, temp_slope_col)), slope_formula)
        formula_rows.append(row_num)

    if formula_rows:
        wb.app.calculate()

    intercept_by_row: Dict[int, Any] = {}
    slope_by_row: Dict[int, Any] = {}
    if formula_rows:
        intercept_values = (
            sheet.range(
                (formula_rows[0], temp_intercept_col), (formula_rows[-1], temp_intercept_col)
            )
            .options(ndim=2)
            .value
        )
        slope_values = (
            sheet.range((formula_rows[0], temp_slope_col), (formula_rows[-1], temp_slope_col))
            .options(ndim=2)
            .value
        )
        for row_num, value in zip(range(formula_rows[0], formula_rows[-1] + 1), intercept_values):
            intercept_by_row[row_num] = value[0]
        for row_num, value in zip(range(formula_rows[0], formula_rows[-1] + 1), slope_values):
            slope_by_row[row_num] = value[0]

    rows: List[Dict[str, Any]] = []
    prior_signature: Optional[Tuple[Any, ...]] = None
    for idx, row_num, n_quarters in candidate_rows:
        intercept_val = intercept_by_row.get(row_num)
        slope_val = slope_by_row.get(row_num)
        forecast_max = block_value(block, idx, cols["forecast_max"], start_col)
        forecast_min = block_value(block, idx, cols["forecast_min"], start_col)
        actual_value = block_value(block, idx, cols["actual_value"], start_col)
        if is_blank(actual_value):
            actual_value = ""
        row = {
            "model": labels.model,
            "ticker": labels.ticker,
            "model_period": labels.model_period,
            "model_date": labels.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n_quarters,
            "num_quarters_used": n_quarters,
            "forecast_value": block_value(block, idx, cols["forecast_value"], start_col),
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": safe_subtract(forecast_max, forecast_min),
            "intercept": intercept_val,
            "slope": slope_val,
            "source_file": source_file,
        }

        signature = (
            to_number(row["num_quarters_used"]),
            to_number(row["forecast_value"]),
            to_number(row["forecast_max"]),
            to_number(row["forecast_min"]),
            to_number(row["intercept"]),
            to_number(row["slope"]),
        )
        if prior_signature is not None and signature == prior_signature:
            continue
        prior_signature = signature
        rows.append(row)

    return rows


def write_sheet(
    wb: Workbook, sheet_name: str, headers: Sequence[str], rows: Sequence[Dict[str, Any]]
) -> None:
    ws = wb.create_sheet(title=sheet_name)
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


def main() -> None:
    input_path = input_dir.expanduser().resolve()
    output_path = output_dir.expanduser().resolve()
    if not input_path.exists() or not input_path.is_dir():
        raise SystemExit(f"Input folder does not exist: {input_path}")

    output_file = choose_output_path(input_path, output_path)
    output_prefix = f"{input_path.name}_param"
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app: Optional[xw.App] = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        for file_path in sorted(input_path.iterdir()):
            if not file_path.is_file():
                continue
            file_name = file_path.name
            if file_name.startswith("~"):
                print(f"Skipped {file_name}: temporary file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_name}: not an .xlsx file")
                continue
            if file_path.stem.lower().startswith(output_prefix):
                print(f"Skipped {file_name}: appears to be an output workbook")
                continue

            print(f"Processing {file_name}")
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                labels = parse_file_labels(file_path)
                empirical_rows.extend(extract_empirical_rows(wb, labels, file_name))
                regression_rows.extend(extract_regression_rows(wb, labels, file_name))
                processed_files += 1
                print(f"Processed {file_name}")
            except Exception as exc:
                print(f"Skipped {file_name}: {exc}")
            finally:
                if wb is not None:
                    close_workbook_safely(wb)
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass

    output_wb = Workbook()
    default_sheet = output_wb.active
    output_wb.remove(default_sheet)
    write_sheet(output_wb, "empirical_candidates", EMPIRICAL_HEADERS, empirical_rows)
    write_sheet(output_wb, "regression_candidates", REGRESSION_HEADERS, regression_rows)
    output_wb.save(output_file)

    print(f"Output path: {output_file}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

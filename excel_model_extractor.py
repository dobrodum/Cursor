#!/usr/bin/env python3
"""Extract empirical/regression candidates from Excel model workbooks.

This script:
1. Opens each source workbook once with a single hidden Excel application.
2. Extracts rows from both "Empirical Model" and "Regression Model".
3. Writes one output workbook with two sheets:
   - empirical_candidates
   - regression_candidates
"""

from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ----------------------------
# User-configurable directories
# ----------------------------
input_dir = Path("./input")
output_dir = Path("./output")

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

PHASE_TO_DAY = {"Early": 5, "Mid": 15, "Late": 25}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = re.sub(r"[^a-z0-9]+", " ", str(value).strip().lower())
    return re.sub(r"\s+", " ", text).strip()


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def to_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, (list, tuple)):
        return [[values]]

    if len(values) == 0:
        return []

    first = values[0]
    if not isinstance(first, (list, tuple)):
        return [list(values)]

    return [list(row) if isinstance(row, (list, tuple)) else [row] for row in values]


def get_output_path(src_input_dir: Path, dst_output_dir: Path) -> Path:
    dst_output_dir.mkdir(parents=True, exist_ok=True)
    input_folder_name = src_input_dir.resolve().name
    base_stem = f"{input_folder_name}_PARAM"
    candidate = dst_output_dir / f"{base_stem}.xlsx"
    if not candidate.exists():
        return candidate

    idx = 1
    while True:
        candidate = dst_output_dir / f"{base_stem}.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def parse_filename_label(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    # Example: MedMiner_Model - AORT - MidJan2026_Send
    match = re.search(
        r"-\s*(?P<ticker>[A-Za-z0-9._-]+)\s*-\s*(?P<period>(Early|Mid|Late)([A-Za-z]{3})(\d{4}))",
        stem,
        flags=re.IGNORECASE,
    )
    if not match:
        fallback_ticker = "UNKNOWN"
        split_parts = [p.strip() for p in stem.split("-") if p.strip()]
        if len(split_parts) >= 2:
            fallback_ticker = split_parts[1].split("_")[0].strip() or "UNKNOWN"
        return {
            "ticker": fallback_ticker,
            "model_period": "Unknown",
            "model_date": "",
            "model": f"{fallback_ticker}_Unknown",
        }

    ticker = match.group("ticker").upper()
    period_token = match.group("period")
    parsed_period = re.match(
        r"(?P<phase>Early|Mid|Late)(?P<month>[A-Za-z]{3})(?P<year>\d{4})",
        period_token,
        flags=re.IGNORECASE,
    )
    if not parsed_period:
        return {
            "ticker": ticker,
            "model_period": period_token,
            "model_date": "",
            "model": f"{ticker}_{period_token}",
        }

    phase = parsed_period.group("phase").title()
    month_abbrev = parsed_period.group("month").title()
    year = int(parsed_period.group("year"))
    month_num = datetime.strptime(month_abbrev, "%b").month
    day = PHASE_TO_DAY[phase]

    model_period = f"{phase}{month_abbrev}_{year}"
    model_date = date(year, month_num, day).isoformat()
    model = f"{ticker}_{model_period}"

    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def iter_source_files(src_input_dir: Path) -> Iterable[Path]:
    input_folder_name = src_input_dir.resolve().name
    output_pattern = re.compile(
        rf"^{re.escape(input_folder_name)}_PARAM(\.\d+)?\.xlsx$",
        flags=re.IGNORECASE,
    )

    for path in sorted(src_input_dir.iterdir()):
        if not path.is_file():
            continue
        if path.name.startswith("~"):
            print(f"Skipped file: {path.name} (temporary workbook)")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"Skipped file: {path.name} (not .xlsx)")
            continue
        if output_pattern.match(path.name):
            print(f"Skipped file: {path.name} (output workbook pattern)")
            continue
        yield path


def get_sheet_by_name(workbook: xw.Book, name: str) -> Optional[xw.Sheet]:
    target = name.strip().lower()
    for sheet in workbook.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def find_anchor_cell(sheet: xw.Sheet, anchor_text: str = "max") -> Optional[xw.Range]:
    anchor_norm = normalize_text(anchor_text)
    # Fast path: native Excel Find.
    try:
        # LookIn=-4163 -> xlValues, LookAt=1 -> xlWhole
        found = sheet.api.Cells.Find(
            What=anchor_text,
            LookIn=-4163,
            LookAt=1,
            SearchOrder=1,
            SearchDirection=1,
            MatchCase=False,
        )
        if found is not None:
            return sheet.range((int(found.Row), int(found.Column)))
    except Exception:
        pass

    # Fallback: scan used range once.
    used = sheet.used_range
    if used is None:
        return None

    values_2d = to_2d(used.value)
    if not values_2d:
        return None

    base_row = int(used.row)
    base_col = int(used.column)
    for r_idx, row_values in enumerate(values_2d):
        for c_idx, value in enumerate(row_values):
            if normalize_text(value) == anchor_norm:
                return sheet.range((base_row + r_idx, base_col + c_idx))
    return None


def get_last_used_col(sheet: xw.Sheet) -> int:
    try:
        return int(sheet.used_range.last_cell.column)
    except Exception:
        return 1


def collect_header_cells(
    sheet: xw.Sheet, anchor_row: int, anchor_col: int, col_window: int = 26
) -> List[Tuple[int, int, str]]:
    start_row = max(1, anchor_row - 1)
    end_row = anchor_row + 1
    start_col = max(1, anchor_col - col_window)
    end_col = anchor_col + col_window

    raw_block = sheet.range((start_row, start_col), (end_row, end_col)).value
    block = to_2d(raw_block)

    cells: List[Tuple[int, int, str]] = []
    for r_off, row_values in enumerate(block):
        for c_off, value in enumerate(row_values):
            norm = normalize_text(value)
            if norm:
                cells.append((start_row + r_off, start_col + c_off, norm))
    return cells


def find_column_by_aliases(
    header_cells: Sequence[Tuple[int, int, str]], aliases: Sequence[str]
) -> Optional[int]:
    alias_norm = [normalize_text(alias) for alias in aliases if normalize_text(alias)]
    for _row, col, text in header_cells:
        if any(alias in text for alias in alias_norm):
            return col
    return None


def value_at(sheet: xw.Sheet, row: int, col: Optional[int]) -> Any:
    if col is None or col < 1:
        return None
    try:
        return sheet.range((row, col)).value
    except Exception:
        return None


def offset_to_col(anchor_col: int, offset: Optional[int]) -> Optional[int]:
    if offset is None:
        return None
    col = anchor_col + offset
    return col if col > 0 else None


def calc_range_width(max_value: Any, min_value: Any) -> Optional[float]:
    max_num = to_float(max_value)
    min_num = to_float(min_value)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
    # Primary path required by prompt.
    try:
        cell.formula2 = formula_r1c1
        return
    except Exception:
        pass

    # Safe fallbacks for environments where formula2 behaves differently.
    for attr in ("Formula2R1C1", "FormulaR1C1", "Formula2", "Formula"):
        try:
            setattr(cell.api, attr, formula_r1c1)
            return
        except Exception:
            continue

    # Last fallback: this may still work on some hosts.
    cell.formula = formula_r1c1


def close_workbook_no_save(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.close(False)
        return
    except Exception:
        pass

    for kwargs in ({"SaveChanges": False}, {"SaveChanges": 0}):
        try:
            workbook.api.Close(**kwargs)
            return
        except Exception:
            continue

    try:
        workbook.api.Close(False)
    except Exception:
        pass


def extract_empirical_rows(
    sheet: xw.Sheet,
    meta: Dict[str, str],
    source_file: str,
    n_quarters: int = N_QUARTERS,
) -> List[Dict[str, Any]]:
    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"Skipped empirical extraction in {source_file}: 'max' anchor not found")
        return []

    anchor_row = int(anchor.row)
    anchor_col = int(anchor.column)
    header_cells = collect_header_cells(sheet, anchor_row, anchor_col)

    # Use anchor-based offsets; defaults handle common layouts if labels vary.
    empirical_aliases = {
        "num_quarters_used": ["num quarters used", "quarters used", "n quarters"],
        "last_quarter_used": ["last quarter used", "last quarter"],
        "forecast_value": ["estimated total sold", "forecast value", "tot fcst", "total sold"],
        "actual_value": ["reported sales", "actual sales", "actual value"],
        "forecast_min": ["min"],
        "quarterly_sales": ["quarterly sales", "qtr sales"],
        "growth_rate_pct": ["growth rate", "growth %"],
        "sales_captured_in_db_pct": ["sales captured in db", "captured in db", "sales captured"],
        "avg_penetration_pct": ["avg penetration", "average penetration"],
        "penetration_source": ["penetration"],
    }

    defaults = {
        "num_quarters_used": -9,
        "last_quarter_used": -8,
        "forecast_value": -3,
        "actual_value": -2,
        "quarterly_sales": -7,
        "growth_rate_pct": -6,
        "sales_captured_in_db_pct": -5,
        "avg_penetration_pct": -4,
        "penetration_source": -4,
        "forecast_min": 1,
    }

    offsets: Dict[str, Optional[int]] = {"forecast_max": 0}
    for key, aliases in empirical_aliases.items():
        found_col = find_column_by_aliases(header_cells, aliases)
        offsets[key] = (found_col - anchor_col) if found_col is not None else defaults.get(key)

    data_start_row = anchor_row + 1
    data_rows = [data_start_row + i for i in range(n_quarters)]

    temp_col = max(get_last_used_col(sheet) + 2, anchor_col + 20)
    penetration_source_col = offset_to_col(anchor_col, offsets.get("penetration_source"))
    quarterly_sales_col = offset_to_col(anchor_col, offsets.get("quarterly_sales"))
    reported_sales_col = offset_to_col(anchor_col, offsets.get("actual_value"))

    # Build R1C1 formula2 values first, then calculate once.
    for row in data_rows:
        if penetration_source_col is not None:
            # Running average over the last n rows through current row.
            formula = (
                f'=IFERROR(AVERAGE(R{data_start_row}C{penetration_source_col}:'
                f'R{row}C{penetration_source_col}),"")'
            )
        elif quarterly_sales_col is not None and reported_sales_col is not None:
            num_off = quarterly_sales_col - temp_col
            den_off = reported_sales_col - temp_col
            formula = f'=IFERROR(RC[{num_off}]/RC[{den_off}],"")'
        else:
            formula = '=""'
        set_formula2_r1c1(sheet.range((row, temp_col)), formula)

    sheet.book.app.calculate()

    rows: List[Dict[str, Any]] = []
    for idx, row in enumerate(data_rows, start=1):
        num_quarters_used = value_at(sheet, row, offset_to_col(anchor_col, offsets.get("num_quarters_used")))
        if is_blank(num_quarters_used):
            num_quarters_used = idx

        last_quarter_used = value_at(sheet, row, offset_to_col(anchor_col, offsets.get("last_quarter_used")))
        forecast_value = value_at(sheet, row, offset_to_col(anchor_col, offsets.get("forecast_value")))
        actual_value = value_at(sheet, row, offset_to_col(anchor_col, offsets.get("actual_value")))
        forecast_max = value_at(sheet, row, offset_to_col(anchor_col, offsets.get("forecast_max")))
        forecast_min = value_at(sheet, row, offset_to_col(anchor_col, offsets.get("forecast_min")))
        avg_penetration_pct = value_at(sheet, row, temp_col)

        # If formula result is empty, try existing avg penetration column.
        if is_blank(avg_penetration_pct):
            avg_penetration_pct = value_at(
                sheet,
                row,
                offset_to_col(anchor_col, offsets.get("avg_penetration_pct")),
            )

        quarterly_sales = value_at(sheet, row, offset_to_col(anchor_col, offsets.get("quarterly_sales")))
        reported_sales = value_at(sheet, row, offset_to_col(anchor_col, offsets.get("actual_value")))
        growth_rate_pct = value_at(sheet, row, offset_to_col(anchor_col, offsets.get("growth_rate_pct")))
        sales_captured_in_db_pct = value_at(
            sheet,
            row,
            offset_to_col(anchor_col, offsets.get("sales_captured_in_db_pct")),
        )

        if all(
            is_blank(v)
            for v in (
                forecast_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
                quarterly_sales,
                reported_sales,
            )
        ):
            continue

        row_dict = {
            "model": meta["model"],
            "ticker": meta["ticker"],
            "model_period": meta["model_period"],
            "model_date": meta["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration_pct,
            "num_quarters_used": num_quarters_used,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_value,  # estimated total sold
            "actual_value": actual_value,  # reported sales
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": calc_range_width(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration_pct,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate_pct,
            "sales_captured_in_db_pct": sales_captured_in_db_pct,
            "source_file": source_file,
        }
        rows.append(row_dict)
    return rows


def normalized_key_for_dedupe(*values: Any) -> Tuple[Any, ...]:
    normalized: List[Any] = []
    for value in values:
        as_float = to_float(value)
        if as_float is not None:
            normalized.append(round(as_float, 10))
        else:
            normalized.append(value if not is_blank(value) else None)
    return tuple(normalized)


def extract_regression_rows(
    sheet: xw.Sheet,
    meta: Dict[str, str],
    source_file: str,
    n_quarters: int = N_QUARTERS,
) -> List[Dict[str, Any]]:
    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"Skipped regression extraction in {source_file}: 'max' anchor not found")
        return []

    anchor_row = int(anchor.row)
    anchor_col = int(anchor.column)
    header_cells = collect_header_cells(sheet, anchor_row, anchor_col)

    # Required offsets from prompt:
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    regression_aliases = {
        "num_quarters_used": ["num quarters used", "quarters used", "n quarters"],
        "forecast_value": ["tot fcst w/o sa", "tot fcst wo sa", "forecast w/o sa", "without sa"],
        "actual_value": ["actual value", "actual sales", "reported sales"],
        "forecast_min": ["min"],
    }
    defaults = {
        "num_quarters_used": -9,
        "forecast_value": -1,
        "actual_value": None,
        "forecast_min": 1,
    }

    offsets: Dict[str, Optional[int]] = {"forecast_max": 0}
    for key, aliases in regression_aliases.items():
        found_col = find_column_by_aliases(header_cells, aliases)
        offsets[key] = (found_col - anchor_col) if found_col is not None else defaults.get(key)

    data_start_row = anchor_row + 1
    data_rows = [data_start_row + i for i in range(n_quarters)]

    first_temp_col = max(get_last_used_col(sheet) + 2, anchor_col + 20)
    intercept_col = first_temp_col
    slope_col = first_temp_col + 1

    # Build all formulas, then calculate once.
    for row in data_rows:
        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{data_start_row}C{y_col}:R{row}C{y_col},'
            f'R{data_start_row}C{x_col}:R{row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{data_start_row}C{y_col}:R{row}C{y_col},'
            f'R{data_start_row}C{x_col}:R{row}C{x_col}),"")'
        )
        set_formula2_r1c1(sheet.range((row, intercept_col)), intercept_formula)
        set_formula2_r1c1(sheet.range((row, slope_col)), slope_formula)

    sheet.book.app.calculate()

    rows: List[Dict[str, Any]] = []
    last_key: Optional[Tuple[Any, ...]] = None

    for idx, row in enumerate(data_rows, start=1):
        num_quarters_used = value_at(sheet, row, offset_to_col(anchor_col, offsets.get("num_quarters_used")))
        if is_blank(num_quarters_used):
            num_quarters_used = idx

        forecast_value = value_at(sheet, row, offset_to_col(anchor_col, offsets.get("forecast_value")))
        actual_value = value_at(sheet, row, offset_to_col(anchor_col, offsets.get("actual_value")))
        forecast_max = value_at(sheet, row, offset_to_col(anchor_col, offsets.get("forecast_max")))
        forecast_min = value_at(sheet, row, offset_to_col(anchor_col, offsets.get("forecast_min")))
        intercept_value = value_at(sheet, row, intercept_col)
        slope_value = value_at(sheet, row, slope_col)

        if all(
            is_blank(v)
            for v in (forecast_value, forecast_max, forecast_min, intercept_value, slope_value)
        ):
            continue

        row_key = normalized_key_for_dedupe(
            num_quarters_used,
            forecast_value,
            forecast_max,
            forecast_min,
            intercept_value,
            slope_value,
        )

        # Prevent duplicate final row (and any contiguous duplicates) by value comparison.
        if last_key is not None and row_key == last_key:
            continue
        last_key = row_key

        row_dict = {
            "model": meta["model"],
            "ticker": meta["ticker"],
            "model_period": meta["model_period"],
            "model_date": meta["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,  # TOT FCST w/o SA
            "actual_value": actual_value if not is_blank(actual_value) else None,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": calc_range_width(forecast_max, forecast_min),
            "intercept": intercept_value,
            "slope": slope_value,
            "source_file": source_file,
        }
        rows.append(row_dict)

    return rows


def write_sheet(
    workbook: Workbook,
    title: str,
    columns: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    ws = workbook.create_sheet(title=title)
    ws.append(list(columns))

    for row in rows:
        ws.append([row.get(col) for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(idx)].width = min(max(max_len + 2, 12), 60)


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


def main() -> None:
    if not input_dir.exists():
        raise SystemExit(f"Input directory does not exist: {input_dir.resolve()}")

    output_path = get_output_path(input_dir, output_dir)
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app: Optional[xw.App] = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in iter_source_files(input_dir):
            print(f"Processing file: {file_path.name}")
            workbook: Optional[xw.Book] = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                meta = parse_filename_label(file_path)

                empirical_sheet = get_sheet_by_name(workbook, "Empirical Model")
                if empirical_sheet is None:
                    print(f"Skipped empirical sheet in {file_path.name}: sheet not found")
                else:
                    empirical_rows.extend(
                        extract_empirical_rows(empirical_sheet, meta=meta, source_file=file_path.name)
                    )

                regression_sheet = get_sheet_by_name(workbook, "Regression Model")
                if regression_sheet is None:
                    print(f"Skipped regression sheet in {file_path.name}: sheet not found")
                else:
                    regression_rows.extend(
                        extract_regression_rows(regression_sheet, meta=meta, source_file=file_path.name)
                    )

                processed_files += 1
            except Exception as exc:
                print(f"Skipped file: {file_path.name} ({exc})")
            finally:
                if workbook is not None:
                    close_workbook_no_save(workbook)
    finally:
        if app is not None:
            app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path.resolve()}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

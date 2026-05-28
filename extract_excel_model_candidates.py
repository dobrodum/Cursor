#!/usr/bin/env python3
from __future__ import annotations

import calendar
import re
from datetime import date
from pathlib import Path
from typing import Any, Optional

import xlwings as xw


# ---------------------------
# Runtime configuration
# ---------------------------
input_dir = "./input"
output_dir = "./output"


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


N_QUARTERS = 10
MODEL_DAY = {"early": 5, "mid": 15, "late": 25}
MAX_HEADER_TOKEN = "max"


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def to_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def parse_month_number(month_token: str) -> Optional[int]:
    token = month_token.strip().lower()
    token = token[:3] if len(token) >= 3 else token

    month_map = {
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
    return month_map.get(token)


def parse_file_metadata(file_name: str) -> Optional[dict[str, str]]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ") if part.strip()]

    ticker = ""
    if len(parts) >= 2:
        ticker = re.sub(r"\s+", "", parts[1]).upper()

    period_source = parts[2] if len(parts) >= 3 else stem
    period_match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*([1-2][0-9]{3})",
        period_source,
        flags=re.IGNORECASE,
    )
    if not period_match:
        period_match = re.search(
            r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*([1-2][0-9]{3})",
            stem,
            flags=re.IGNORECASE,
        )
    if not period_match:
        return None

    period_bucket = period_match.group(1).lower()
    month_token = period_match.group(2)
    year_text = period_match.group(3)

    month_number = parse_month_number(month_token)
    if month_number is None:
        return None

    year = int(year_text)
    month_abbr = calendar.month_abbr[month_number]
    day = MODEL_DAY[period_bucket]

    model_period = f"{period_bucket.title()}{month_abbr}_{year}"
    model_date = date(year, month_number, day).isoformat()
    model = f"{ticker}_{model_period}" if ticker else model_period

    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def next_output_path(src_dir: Path, dst_dir: Path) -> Path:
    folder_name = src_dir.name
    base_name = f"{folder_name}_PARAM.xlsx"
    candidate = dst_dir / base_name
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = dst_dir / f"{folder_name}_PARAM.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def find_anchor_cell(sheet: xw.Sheet, token: str = MAX_HEADER_TOKEN) -> tuple[int, int]:
    used = sheet.used_range
    values = to_2d(used.value)
    base_row = used.row
    base_col = used.column
    target = normalize_text(token)

    for row_idx, row_values in enumerate(values):
        for col_idx, cell_value in enumerate(row_values):
            if normalize_text(cell_value) == target:
                return base_row + row_idx, base_col + col_idx

    raise ValueError(f"Could not find '{token}' anchor on sheet '{sheet.name}'.")


def header_map_from_anchor(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    left_span: int = 25,
    right_span: int = 8,
) -> dict[str, int]:
    start_col = max(1, anchor_col - left_span)
    end_col = anchor_col + right_span
    values = sheet.range((anchor_row, start_col), (anchor_row, end_col)).value
    row_values = values if isinstance(values, list) else [values]

    headers: dict[str, int] = {}
    for idx, cell_value in enumerate(row_values):
        key = normalize_text(cell_value)
        if key and key not in headers:
            headers[key] = start_col + idx
    return headers


def pick_column(header_map: dict[str, int], variants: list[str], fallback_col: int) -> int:
    normalized_variants = [normalize_text(variant) for variant in variants]

    for variant in normalized_variants:
        if variant in header_map:
            return header_map[variant]

    for header_key, header_col in header_map.items():
        if any(variant in header_key for variant in normalized_variants):
            return header_col

    return fallback_col


def pick_optional_column(header_map: dict[str, int], variants: list[str]) -> Optional[int]:
    normalized_variants = [normalize_text(variant) for variant in variants]
    for variant in normalized_variants:
        if variant in header_map:
            return header_map[variant]
    for header_key, header_col in header_map.items():
        if any(variant in header_key for variant in normalized_variants):
            return header_col
    return None


def close_workbook_without_saving(workbook: xw.Book) -> None:
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

    try:
        workbook.api.Close(SaveChanges=False)
        return
    except Exception:
        workbook.api.Close(False)


def set_formula2(target_range: xw.Range, formula_r1c1: str) -> None:
    try:
        target_range.formula2 = formula_r1c1
    except Exception:
        target_range.formula = formula_r1c1


def as_int(value: Any, default: Optional[int] = None) -> Optional[int]:
    try:
        if value is None or value == "":
            return default
        return int(round(float(value)))
    except Exception:
        return default


def as_float(value: Any) -> Optional[float]:
    try:
        if value is None or value == "":
            return None
        return float(value)
    except Exception:
        return None


def row_signature(values: tuple[Any, ...]) -> tuple[Any, ...]:
    normalized: list[Any] = []
    for value in values:
        if isinstance(value, float):
            normalized.append(round(value, 10))
        else:
            normalized.append(value)
    return tuple(normalized)


def rc_ref(offset: int) -> str:
    return "RC" if offset == 0 else f"RC[{offset}]"


def get_row_cell(row_values: list[Any], min_col: int, target_col: Optional[int]) -> Any:
    if target_col is None:
        return None
    idx = target_col - min_col
    if idx < 0 or idx >= len(row_values):
        return None
    return row_values[idx]


def empirical_column_map(anchor_col: int, headers: dict[str, int]) -> dict[str, int]:
    return {
        "num_quarters_used": pick_column(
            headers,
            ["num quarters used", "num_quarters_used", "quarters used", "n quarters"],
            anchor_col - 9,
        ),
        "last_quarter_used": pick_column(
            headers,
            ["last quarter used", "last_quarter_used", "last quarter", "last qtr"],
            anchor_col - 8,
        ),
        "avg_penetration_pct": pick_column(
            headers,
            ["avg penetration pct", "average penetration", "avg penetration", "penetration"],
            anchor_col - 7,
        ),
        "quarterly_sales": pick_column(
            headers,
            ["quarterly sales", "quarterly_sales", "q sales"],
            anchor_col - 6,
        ),
        "growth_rate_pct": pick_column(
            headers,
            ["growth rate pct", "growth_rate_pct", "growth rate"],
            anchor_col - 5,
        ),
        "forecast_value": pick_column(
            headers,
            [
                "estimated total sold",
                "forecast value",
                "total forecast",
                "tot fcst",
                "tot fcst w/o sa",
            ],
            anchor_col - 4,
        ),
        "actual_value": pick_column(
            headers,
            ["reported sales", "actual value", "actual sales", "actual"],
            anchor_col - 3,
        ),
        "sales_captured_in_db_pct": pick_column(
            headers,
            ["sales captured in db pct", "sales captured", "captured in db", "db pct"],
            anchor_col - 2,
        ),
        "reported_sales": pick_column(
            headers,
            ["reported sales", "reported_sales", "actual sales"],
            anchor_col - 3,
        ),
        "forecast_max": anchor_col,
        "forecast_min": pick_column(headers, ["min", "forecast min"], anchor_col + 1),
    }


def regression_column_map(anchor_col: int, headers: dict[str, int]) -> dict[str, Optional[int]]:
    return {
        "num_quarters_used": pick_column(
            headers,
            ["num quarters used", "num_quarters_used", "quarters used", "n quarters"],
            anchor_col - 6,
        ),
        "intercept": pick_column(headers, ["intercept"], anchor_col - 4),
        "slope": pick_column(headers, ["slope"], anchor_col - 3),
        "forecast_value": pick_column(
            headers,
            ["tot fcst w/o sa", "forecast value", "total forecast", "tot fcst"],
            anchor_col - 2,
        ),
        "actual_value": pick_optional_column(
            headers, ["actual value", "reported sales", "actual sales", "actual"]
        ),
        "forecast_max": anchor_col,
        "forecast_min": pick_column(headers, ["min", "forecast min"], anchor_col + 1),
    }


def extract_empirical_rows(
    workbook: xw.Book,
    metadata: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    sheet = workbook.sheets["Empirical Model"]
    anchor_row, anchor_col = find_anchor_cell(sheet, MAX_HEADER_TOKEN)
    headers = header_map_from_anchor(sheet, anchor_row, anchor_col)
    col_map = empirical_column_map(anchor_col, headers)

    start_row = anchor_row + 1
    num_q_col = col_map["num_quarters_used"]
    avg_pen_col = col_map["avg_penetration_pct"]
    source_pen_col = col_map["sales_captured_in_db_pct"]

    n_ref_from_avg = rc_ref(num_q_col - avg_pen_col)
    source_col_ref = f"C{source_pen_col}"
    avg_pen_formula = (
        f'=IFERROR(AVERAGE(INDEX({source_col_ref},MAX(1,MATCH(9.9E+307,{source_col_ref})-{n_ref_from_avg}+1)):'
        f'INDEX({source_col_ref},MATCH(9.9E+307,{source_col_ref}))),"")'
    )

    for idx in range(N_QUARTERS):
        row = start_row + idx
        sheet.range((row, num_q_col)).value = idx + 1
        set_formula2(sheet.range((row, avg_pen_col)), avg_pen_formula)

    workbook.app.calculate()

    needed_cols = sorted(
        {
            col_map["num_quarters_used"],
            col_map["last_quarter_used"],
            col_map["forecast_value"],
            col_map["actual_value"],
            col_map["forecast_max"],
            col_map["forecast_min"],
            col_map["avg_penetration_pct"],
            col_map["quarterly_sales"],
            col_map["reported_sales"],
            col_map["growth_rate_pct"],
            col_map["sales_captured_in_db_pct"],
        }
    )

    min_col = min(needed_cols)
    max_col = max(needed_cols)
    values = to_2d(
        sheet.range((start_row, min_col), (start_row + N_QUARTERS - 1, max_col)).value
    )

    while len(values) < N_QUARTERS:
        values.append([None] * (max_col - min_col + 1))

    rows: list[dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        row_values = values[idx]
        num_quarters_used = as_int(
            get_row_cell(row_values, min_col, col_map["num_quarters_used"]), default=idx + 1
        )
        forecast_value = as_float(get_row_cell(row_values, min_col, col_map["forecast_value"]))
        actual_value = as_float(get_row_cell(row_values, min_col, col_map["actual_value"]))
        forecast_max = as_float(get_row_cell(row_values, min_col, col_map["forecast_max"]))
        forecast_min = as_float(get_row_cell(row_values, min_col, col_map["forecast_min"]))
        avg_penetration_pct = as_float(
            get_row_cell(row_values, min_col, col_map["avg_penetration_pct"])
        )
        quarterly_sales = as_float(get_row_cell(row_values, min_col, col_map["quarterly_sales"]))
        reported_sales = as_float(get_row_cell(row_values, min_col, col_map["reported_sales"]))
        growth_rate_pct = as_float(get_row_cell(row_values, min_col, col_map["growth_rate_pct"]))
        sales_captured_in_db_pct = as_float(
            get_row_cell(row_values, min_col, col_map["sales_captured_in_db_pct"])
        )
        last_quarter_used = get_row_cell(row_values, min_col, col_map["last_quarter_used"])

        if (
            forecast_value is None
            and actual_value is None
            and forecast_max is None
            and forecast_min is None
            and avg_penetration_pct is None
        ):
            continue

        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
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
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_rows(
    workbook: xw.Book,
    metadata: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    sheet = workbook.sheets["Regression Model"]
    anchor_row, anchor_col = find_anchor_cell(sheet, MAX_HEADER_TOKEN)
    headers = header_map_from_anchor(sheet, anchor_row, anchor_col)
    col_map = regression_column_map(anchor_col, headers)

    start_row = anchor_row + 1
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    num_q_col = int(col_map["num_quarters_used"])
    intercept_col = int(col_map["intercept"])
    slope_col = int(col_map["slope"])

    n_ref_from_intercept = rc_ref(num_q_col - intercept_col)
    n_ref_from_slope = rc_ref(num_q_col - slope_col)
    y_col_ref = f"C{y_col}"
    x_col_ref = f"C{x_col}"

    intercept_formula = (
        f'=IFERROR(INTERCEPT(INDEX({y_col_ref},MAX(1,MATCH(9.9E+307,{y_col_ref})-{n_ref_from_intercept}+1)):'
        f'INDEX({y_col_ref},MATCH(9.9E+307,{y_col_ref})),'
        f'INDEX({x_col_ref},MAX(1,MATCH(9.9E+307,{x_col_ref})-{n_ref_from_intercept}+1)):'
        f'INDEX({x_col_ref},MATCH(9.9E+307,{x_col_ref}))),"")'
    )
    slope_formula = (
        f'=IFERROR(SLOPE(INDEX({y_col_ref},MAX(1,MATCH(9.9E+307,{y_col_ref})-{n_ref_from_slope}+1)):'
        f'INDEX({y_col_ref},MATCH(9.9E+307,{y_col_ref})),'
        f'INDEX({x_col_ref},MAX(1,MATCH(9.9E+307,{x_col_ref})-{n_ref_from_slope}+1)):'
        f'INDEX({x_col_ref},MATCH(9.9E+307,{x_col_ref}))),"")'
    )

    for idx in range(N_QUARTERS):
        row = start_row + idx
        sheet.range((row, num_q_col)).value = idx + 1
        set_formula2(sheet.range((row, intercept_col)), intercept_formula)
        set_formula2(sheet.range((row, slope_col)), slope_formula)

    workbook.app.calculate()

    needed_cols = sorted(
        col
        for col in {
            col_map["num_quarters_used"],
            col_map["forecast_value"],
            col_map["actual_value"],
            col_map["forecast_max"],
            col_map["forecast_min"],
            col_map["intercept"],
            col_map["slope"],
        }
        if isinstance(col, int)
    )

    min_col = min(needed_cols)
    max_col = max(needed_cols)
    values = to_2d(
        sheet.range((start_row, min_col), (start_row + N_QUARTERS - 1, max_col)).value
    )

    while len(values) < N_QUARTERS:
        values.append([None] * (max_col - min_col + 1))

    rows: list[dict[str, Any]] = []
    prior_signature: Optional[tuple[Any, ...]] = None

    for idx in range(N_QUARTERS):
        row_values = values[idx]
        num_quarters_used = as_int(
            get_row_cell(row_values, min_col, int(col_map["num_quarters_used"])),
            default=idx + 1,
        )
        forecast_value = as_float(
            get_row_cell(row_values, min_col, int(col_map["forecast_value"]))
        )
        actual_value = as_float(get_row_cell(row_values, min_col, col_map["actual_value"]))
        forecast_max = as_float(get_row_cell(row_values, min_col, int(col_map["forecast_max"])))
        forecast_min = as_float(get_row_cell(row_values, min_col, int(col_map["forecast_min"])))
        intercept = as_float(get_row_cell(row_values, min_col, int(col_map["intercept"])))
        slope = as_float(get_row_cell(row_values, min_col, int(col_map["slope"])))

        if (
            forecast_value is None
            and forecast_max is None
            and forecast_min is None
            and intercept is None
            and slope is None
        ):
            continue

        signature = row_signature(
            (
                num_quarters_used,
                forecast_value,
                forecast_max,
                forecast_min,
                intercept,
                slope,
            )
        )

        if idx == N_QUARTERS - 1 and signature == prior_signature:
            continue

        prior_signature = signature

        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

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
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def write_output_sheet(
    app: xw.App,
    sheet: xw.Sheet,
    columns: list[str],
    rows: list[dict[str, Any]],
) -> None:
    sheet.range((1, 1), (1, len(columns))).value = columns
    if rows:
        data = [[row.get(col) for col in columns] for row in rows]
        sheet.range((2, 1)).value = data

    header_range = sheet.range((1, 1), (1, len(columns)))
    header_range.api.Font.Bold = True

    last_row = max(1, len(rows) + 1)
    filter_range = sheet.range((1, 1), (last_row, len(columns)))
    filter_range.api.AutoFilter()

    sheet.activate()
    app.api.ActiveWindow.SplitColumn = 0
    app.api.ActiveWindow.SplitRow = 1
    app.api.ActiveWindow.FreezePanes = True

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row in rows:
            value = row.get(col_name)
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        sheet.range((1, col_idx)).column_width = min(max(max_len + 2, 12), 45)


def write_output_workbook(
    app: xw.App,
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    workbook = app.books.add()
    try:
        empirical_sheet = workbook.sheets[0]
        empirical_sheet.name = "empirical_candidates"
        regression_sheet = workbook.sheets.add("regression_candidates", after=empirical_sheet)

        write_output_sheet(app, empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)
        write_output_sheet(app, regression_sheet, REGRESSION_COLUMNS, regression_rows)

        workbook.save(str(output_path))
    finally:
        close_workbook_without_saving(workbook)


def main() -> None:
    src_dir = Path(input_dir).expanduser().resolve()
    dst_dir = Path(output_dir).expanduser().resolve()

    if not src_dir.exists():
        raise FileNotFoundError(f"Input directory does not exist: {src_dir}")

    dst_dir.mkdir(parents=True, exist_ok=True)
    output_path = next_output_path(src_dir, dst_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(src_dir.iterdir()):
            if not file_path.is_file():
                continue

            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file")
                continue

            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary file")
                continue

            metadata = parse_file_metadata(file_path.name)
            if metadata is None:
                print(
                    f"Skipped {file_path.name}: could not parse ticker/model period from file name"
                )
                continue

            print(f"Processing {file_path.name}")
            workbook: Optional[xw.Book] = None
            processed_files += 1

            try:
                workbook = app.books.open(str(file_path), update_links=False)

                try:
                    empirical_rows.extend(extract_empirical_rows(workbook, metadata, file_path.name))
                except Exception as exc:
                    print(f"Skipped empirical extraction in {file_path.name}: {exc}")

                try:
                    regression_rows.extend(
                        extract_regression_rows(workbook, metadata, file_path.name)
                    )
                except Exception as exc:
                    print(f"Skipped regression extraction in {file_path.name}: {exc}")

            except Exception as exc:
                print(f"Skipped {file_path.name}: failed to open/process workbook ({exc})")
            finally:
                if workbook is not None:
                    close_workbook_without_saving(workbook)

        write_output_workbook(app, output_path, empirical_rows, regression_rows)

    finally:
        app.quit()

    print(f"Output path: {output_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

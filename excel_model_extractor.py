from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

try:
    import xlwings as xw
except ImportError as exc:  # pragma: no cover - runtime dependency check
    raise SystemExit("xlwings is required to run this script.") from exc


# Edit these two paths before running.
input_dir = Path("./input")
output_dir = Path("./output")


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

MONTH_NUM = {
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

PERIOD_DAY = {"early": 5, "mid": 15, "late": 25}
PERIOD_RE = re.compile(
    r"(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(\d{4})",
    re.IGNORECASE,
)

EMPIRICAL_LABELS = {
    "forecast_max": ["max"],
    "forecast_min": ["min"],
    "forecast_value": ["estimated total sold", "est total sold", "total sold", "forecast"],
    "actual_value": ["reported sales", "actual sales", "actual"],
    "quarterly_sales": ["quarterly sales", "qtr sales", "quarter sales"],
    "reported_sales": ["reported sales"],
    "growth_rate_pct": ["growth rate", "growth %", "growth pct"],
    "sales_captured_in_db_pct": ["sales captured in db", "captured in db", "db %"],
    "last_quarter_used": ["last quarter used", "last quarter", "last qtr used", "last qtr"],
    "avg_penetration_source": ["penetration", "captured in db"],
}

EMPIRICAL_FALLBACK_OFFSETS = {
    "forecast_max": 0,
    "forecast_min": 1,
    "forecast_value": -1,
    "actual_value": -2,
    "quarterly_sales": -5,
    "reported_sales": -2,
    "growth_rate_pct": -4,
    "sales_captured_in_db_pct": -3,
    "last_quarter_used": -6,
    "avg_penetration_source": -3,
}

REGRESSION_LABELS = {
    "forecast_total_without_sa": [
        "tot fcst w/o sa",
        "tot fcst wo sa",
        "tot fcst without sa",
        "tot fcst",
    ],
    "forecast_max": ["max"],
    "forecast_min": ["min"],
    "num_quarters_used": ["num quarters used", "num quarters", "quarters used", "n quarters"],
    "actual_value": ["actual", "reported sales"],
}

REGRESSION_FALLBACK_OFFSETS = {
    "forecast_total_without_sa": -1,
    "forecast_max": 0,
    "forecast_min": 1,
}


@dataclass(frozen=True)
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def to_float(value: Any) -> float | None:
    if value in ("", None):
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace(",", "").strip()
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def to_int(value: Any) -> int | None:
    number = to_float(value)
    if number is None:
        return None
    return int(round(number))


def normalize_matrix(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if not isinstance(values[0], list):
        return [values]
    normalized: list[list[Any]] = []
    for row in values:
        if isinstance(row, list):
            normalized.append(row)
        else:
            normalized.append([row])
    return normalized


def normalize_vector(values: Any, expected_len: int) -> list[Any]:
    if isinstance(values, list):
        if values and isinstance(values[0], list):
            flattened: list[Any] = []
            for row in values:
                flattened.extend(row)
        else:
            flattened = list(values)
    else:
        flattened = [values]
    if len(flattened) < expected_len:
        flattened.extend([None] * (expected_len - len(flattened)))
    return flattened[:expected_len]


def matrix_cell(matrix: list[list[Any]], row_idx: int, col_idx: int) -> Any:
    if row_idx < 0 or col_idx < 0 or row_idx >= len(matrix):
        return None
    row = matrix[row_idx]
    if col_idx >= len(row):
        return None
    return row[col_idx]


def row_slice(matrix: list[list[Any]], row_idx: int, start_col_idx: int, width: int) -> list[Any]:
    return [matrix_cell(matrix, row_idx, start_col_idx + i) for i in range(width)]


def find_max_anchor(matrix: list[list[Any]]) -> tuple[int, int] | None:
    candidates: list[tuple[int, int, int]] = []
    for r_idx, row in enumerate(matrix):
        for c_idx, value in enumerate(row):
            label = normalize_text(value)
            if label == "max":
                numeric_right = 0
                for probe in range(c_idx + 1, min(c_idx + 16, len(row))):
                    if to_float(row[probe]) is not None:
                        numeric_right += 1
                candidates.append((numeric_right, r_idx, c_idx))
    if not candidates:
        return None
    candidates.sort(key=lambda item: (item[0], -item[1], -item[2]), reverse=True)
    _, best_row, best_col = candidates[0]
    return best_row, best_col


def metric_row_offsets(
    matrix: list[list[Any]],
    anchor_row_idx: int,
    anchor_col_idx: int,
    labels: dict[str, list[str]],
    fallback_offsets: dict[str, int],
) -> dict[str, int]:
    offsets: dict[str, int] = {}
    row_start = max(0, anchor_row_idx - 80)
    row_end = min(len(matrix), anchor_row_idx + 81)
    label_cols = [anchor_col_idx, anchor_col_idx - 1, anchor_col_idx - 2, anchor_col_idx - 3, 0, 1]
    for metric, tokens in labels.items():
        found_offset: int | None = None
        for row_idx in range(row_start, row_end):
            parts: list[str] = []
            for col_idx in label_cols:
                text = normalize_text(matrix_cell(matrix, row_idx, col_idx))
                if text and text not in parts:
                    parts.append(text)
            if not parts:
                continue
            blob = " | ".join(parts)
            if any(token in blob for token in tokens):
                found_offset = row_idx - anchor_row_idx
                break
        if found_offset is not None:
            offsets[metric] = found_offset
        elif metric in fallback_offsets:
            offsets[metric] = fallback_offsets[metric]
    return offsets


def set_formula2(target_range: Any, formula: str) -> None:
    try:
        target_range.formula2 = formula
    except Exception:
        target_range.formula = formula


def safe_close_source_workbook(wb: Any) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass
    try:
        wb.close(False)
        return
    except Exception:
        pass
    wb.close()


def next_output_path(in_dir: Path, out_dir: Path) -> Path:
    base_name = f"{in_dir.name}_PARAM"
    first_choice = out_dir / f"{base_name}.xlsx"
    if not first_choice.exists():
        return first_choice
    idx = 1
    while True:
        candidate = out_dir / f"{base_name}.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def parse_file_label(file_path: Path) -> FileLabel:
    stem = file_path.stem
    pieces = [piece.strip() for piece in stem.split(" - ") if piece.strip()]
    ticker = pieces[1] if len(pieces) >= 2 else "UNKNOWN"

    period_match = PERIOD_RE.search(stem)
    if period_match:
        period_prefix = period_match.group(1).capitalize()
        period_month = period_match.group(2).capitalize()
        period_year = period_match.group(3)
        model_period = f"{period_prefix}{period_month}_{period_year}"
        month_num = MONTH_NUM[period_month.lower()]
        day_num = PERIOD_DAY[period_prefix.lower()]
        model_date = date(int(period_year), month_num, day_num).isoformat()
    else:
        model_period = "UNKNOWN_PERIOD"
        model_date = ""

    model = f"{ticker}_{model_period}" if ticker != "UNKNOWN" else model_period
    return FileLabel(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def clean_for_output(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, float) and (value != value):
        return None
    if isinstance(value, datetime):
        return value.isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return value


def calc_range_width(max_value: Any, min_value: Any) -> float | None:
    max_num = to_float(max_value)
    min_num = to_float(min_value)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


def find_sheet(workbook: Any, target_name: str) -> Any | None:
    target = target_name.strip().lower()
    for sheet in workbook.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def calculated_avg_penetration(
    sheet: Any,
    workbook: Any,
    source_row_abs: int,
    anchor_col_abs: int,
    n_quarters: int,
) -> list[Any]:
    if source_row_abs < 1:
        return [None] * n_quarters

    used = sheet.used_range
    temp_row = used.last_cell.row + 2
    temp_col = used.last_cell.column + 2
    for idx in range(n_quarters):
        current_col = anchor_col_abs + idx + 1
        start_col = max(1, current_col - idx)
        formula = (
            f'=IFERROR(AVERAGE(R{source_row_abs}C{start_col}:R{source_row_abs}C{current_col}),"")'
        )
        set_formula2(sheet.range((temp_row, temp_col + idx)), formula)

    workbook.app.calculate()
    values = normalize_vector(
        sheet.range((temp_row, temp_col), (temp_row, temp_col + n_quarters - 1)).value,
        n_quarters,
    )
    sheet.range((temp_row, temp_col), (temp_row, temp_col + n_quarters - 1)).clear_contents()
    return values


def collect_pair_rows(
    matrix: list[list[Any]],
    base_row_abs: int,
    base_col_abs: int,
    x_col_abs: int,
    y_col_abs: int,
) -> list[int]:
    x_idx = x_col_abs - base_col_abs
    y_idx = y_col_abs - base_col_abs
    if x_idx < 0 or y_idx < 0:
        return []

    pair_rows: list[int] = []
    for rel_row in range(len(matrix)):
        x_val = to_float(matrix_cell(matrix, rel_row, x_idx))
        y_val = to_float(matrix_cell(matrix, rel_row, y_idx))
        if x_val is not None and y_val is not None:
            pair_rows.append(base_row_abs + rel_row)
    return pair_rows


def calculate_intercept_slope(
    sheet: Any,
    workbook: Any,
    pair_rows: list[int],
    x_col_abs: int,
    y_col_abs: int,
    quarters_used: list[int],
) -> tuple[list[Any], list[Any]]:
    n_rows = len(quarters_used)
    intercepts = [None] * n_rows
    slopes = [None] * n_rows
    if len(pair_rows) < 2:
        return intercepts, slopes

    used = sheet.used_range
    temp_row = used.last_cell.row + 4
    temp_col = used.last_cell.column + 2

    for idx, q_count in enumerate(quarters_used):
        lookback = max(2, min(len(pair_rows), q_count))
        start_row = pair_rows[-lookback]
        end_row = pair_rows[-1]
        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col_abs}:R{end_row}C{y_col_abs},'
            f'R{start_row}C{x_col_abs}:R{end_row}C{x_col_abs}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{start_row}C{y_col_abs}:R{end_row}C{y_col_abs},'
            f'R{start_row}C{x_col_abs}:R{end_row}C{x_col_abs}),"")'
        )
        set_formula2(sheet.range((temp_row, temp_col + idx)), intercept_formula)
        set_formula2(sheet.range((temp_row + 1, temp_col + idx)), slope_formula)

    workbook.app.calculate()
    intercepts = normalize_vector(
        sheet.range((temp_row, temp_col), (temp_row, temp_col + n_rows - 1)).value,
        n_rows,
    )
    slopes = normalize_vector(
        sheet.range((temp_row + 1, temp_col), (temp_row + 1, temp_col + n_rows - 1)).value,
        n_rows,
    )
    sheet.range((temp_row, temp_col), (temp_row + 1, temp_col + n_rows - 1)).clear_contents()
    return intercepts, slopes


def extract_empirical_rows(sheet: Any, workbook: Any, label: FileLabel, source_file: str) -> list[dict[str, Any]]:
    used = sheet.used_range
    matrix = normalize_matrix(used.value)
    if not matrix:
        return []

    anchor = find_max_anchor(matrix)
    if anchor is None:
        print(f"  skipped empirical extraction for {source_file}: max anchor not found")
        return []

    anchor_r, anchor_c = anchor
    anchor_row_abs = used.row + anchor_r
    anchor_col_abs = used.column + anchor_c
    n_quarters = 10

    offsets = metric_row_offsets(
        matrix,
        anchor_r,
        anchor_c,
        EMPIRICAL_LABELS,
        EMPIRICAL_FALLBACK_OFFSETS,
    )
    start_col_idx = anchor_c + 1

    forecast_max_values = row_slice(
        matrix,
        anchor_r + offsets.get("forecast_max", 0),
        start_col_idx,
        n_quarters,
    )
    forecast_min_values = row_slice(
        matrix,
        anchor_r + offsets.get("forecast_min", 1),
        start_col_idx,
        n_quarters,
    )
    forecast_values = row_slice(
        matrix,
        anchor_r + offsets.get("forecast_value", -1),
        start_col_idx,
        n_quarters,
    )
    actual_values = row_slice(
        matrix,
        anchor_r + offsets.get("actual_value", -2),
        start_col_idx,
        n_quarters,
    )
    quarterly_sales_values = row_slice(
        matrix,
        anchor_r + offsets.get("quarterly_sales", -5),
        start_col_idx,
        n_quarters,
    )
    reported_sales_values = row_slice(
        matrix,
        anchor_r + offsets.get("reported_sales", -2),
        start_col_idx,
        n_quarters,
    )
    growth_rate_values = row_slice(
        matrix,
        anchor_r + offsets.get("growth_rate_pct", -4),
        start_col_idx,
        n_quarters,
    )
    sales_captured_values = row_slice(
        matrix,
        anchor_r + offsets.get("sales_captured_in_db_pct", -3),
        start_col_idx,
        n_quarters,
    )
    last_quarter_values = row_slice(
        matrix,
        anchor_r + offsets.get("last_quarter_used", -6),
        start_col_idx,
        n_quarters,
    )

    avg_source_row_abs = anchor_row_abs + offsets.get("avg_penetration_source", -3)
    avg_penetration_values = calculated_avg_penetration(
        sheet=sheet,
        workbook=workbook,
        source_row_abs=avg_source_row_abs,
        anchor_col_abs=anchor_col_abs,
        n_quarters=n_quarters,
    )

    rows: list[dict[str, Any]] = []
    for idx in range(n_quarters):
        num_quarters = idx + 1
        avg_penetration = avg_penetration_values[idx]
        if avg_penetration in ("", None):
            avg_penetration = sales_captured_values[idx]

        forecast_max = forecast_max_values[idx]
        forecast_min = forecast_min_values[idx]
        forecast_value = forecast_values[idx]
        actual_value = actual_values[idx]
        if actual_value in ("", None):
            actual_value = reported_sales_values[idx]

        row = {
            "model": label.model,
            "ticker": label.ticker,
            "model_period": label.model_period,
            "model_date": label.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration,
            "num_quarters_used": num_quarters,
            "last_quarter_used": (
                last_quarter_values[idx]
                if last_quarter_values[idx] not in ("", None)
                else num_quarters
            ),
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": calc_range_width(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration,
            "quarterly_sales": quarterly_sales_values[idx],
            "reported_sales": reported_sales_values[idx],
            "growth_rate_pct": growth_rate_values[idx],
            "sales_captured_in_db_pct": sales_captured_values[idx],
            "source_file": source_file,
        }

        key_values = (
            row["forecast_value"],
            row["forecast_max"],
            row["forecast_min"],
            row["parameter_value"],
        )
        if all(value in ("", None) for value in key_values):
            continue
        rows.append(row)
    return rows


def extract_regression_rows(
    sheet: Any,
    workbook: Any,
    label: FileLabel,
    source_file: str,
) -> list[dict[str, Any]]:
    used = sheet.used_range
    matrix = normalize_matrix(used.value)
    if not matrix:
        return []

    anchor = find_max_anchor(matrix)
    if anchor is None:
        print(f"  skipped regression extraction for {source_file}: max anchor not found")
        return []

    anchor_r, anchor_c = anchor
    anchor_row_abs = used.row + anchor_r
    anchor_col_abs = used.column + anchor_c
    n_quarters = 10
    start_col_idx = anchor_c + 1

    offsets = metric_row_offsets(
        matrix,
        anchor_r,
        anchor_c,
        REGRESSION_LABELS,
        REGRESSION_FALLBACK_OFFSETS,
    )

    forecast_values = row_slice(
        matrix,
        anchor_r + offsets.get("forecast_total_without_sa", -1),
        start_col_idx,
        n_quarters,
    )
    forecast_max_values = row_slice(
        matrix,
        anchor_r + offsets.get("forecast_max", 0),
        start_col_idx,
        n_quarters,
    )
    forecast_min_values = row_slice(
        matrix,
        anchor_r + offsets.get("forecast_min", 1),
        start_col_idx,
        n_quarters,
    )
    actual_values = row_slice(
        matrix,
        anchor_r + offsets.get("actual_value", -2),
        start_col_idx,
        n_quarters,
    )

    raw_quarters = row_slice(
        matrix,
        anchor_r + offsets.get("num_quarters_used", -99),
        start_col_idx,
        n_quarters,
    )
    quarters_used: list[int] = []
    for idx in range(n_quarters):
        parsed = to_int(raw_quarters[idx])
        if parsed is None or parsed < 1:
            parsed = idx + 1
        quarters_used.append(parsed)

    y_col_abs = anchor_col_abs - 7
    x_col_abs = anchor_col_abs - 11
    pair_rows = collect_pair_rows(
        matrix=matrix,
        base_row_abs=used.row,
        base_col_abs=used.column,
        x_col_abs=x_col_abs,
        y_col_abs=y_col_abs,
    )
    intercept_values, slope_values = calculate_intercept_slope(
        sheet=sheet,
        workbook=workbook,
        pair_rows=pair_rows,
        x_col_abs=x_col_abs,
        y_col_abs=y_col_abs,
        quarters_used=quarters_used,
    )

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None
    for idx in range(n_quarters):
        forecast_value = forecast_values[idx]
        forecast_max = forecast_max_values[idx]
        forecast_min = forecast_min_values[idx]
        intercept = intercept_values[idx]
        slope = slope_values[idx]
        actual_value = actual_values[idx]

        row = {
            "model": label.model,
            "ticker": label.ticker,
            "model_period": label.model_period,
            "model_date": label.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": quarters_used[idx],
            "num_quarters_used": quarters_used[idx],
            "forecast_value": forecast_value,
            "actual_value": actual_value if actual_value not in (None, "") else "",
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": calc_range_width(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        signature = (
            row["num_quarters_used"],
            to_float(row["forecast_value"]),
            to_float(row["forecast_max"]),
            to_float(row["forecast_min"]),
            to_float(row["intercept"]),
            to_float(row["slope"]),
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        key_values = (
            row["forecast_value"],
            row["forecast_max"],
            row["forecast_min"],
            row["intercept"],
            row["slope"],
        )
        if all(value in ("", None) for value in key_values):
            continue
        rows.append(row)

    return rows


def write_sheet(ws: Any, headers: list[str], rows: list[dict[str, Any]]) -> None:
    ws.append(headers)
    for row in rows:
        ws.append([clean_for_output(row.get(header)) for header in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            cell_value = ws.cell(row=row_idx, column=idx).value
            if cell_value is None:
                continue
            max_len = max(max_len, len(str(cell_value)))
        ws.column_dimensions[get_column_letter(idx)].width = min(max(max_len + 2, 12), 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    workbook = Workbook()
    empirical_sheet = workbook.active
    empirical_sheet.title = "empirical_candidates"
    write_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)

    regression_sheet = workbook.create_sheet("regression_candidates")
    write_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)

    workbook.save(output_path)


def run() -> None:
    in_dir = Path(input_dir).expanduser().resolve()
    out_dir = Path(output_dir).expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    if not in_dir.exists():
        raise SystemExit(f"input_dir does not exist: {in_dir}")
    if not in_dir.is_dir():
        raise SystemExit(f"input_dir is not a directory: {in_dir}")

    files_processed = 0
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        for file_path in sorted(in_dir.iterdir()):
            if not file_path.is_file():
                continue
            file_name = file_path.name
            if not file_name.lower().endswith(".xlsx"):
                print(f"Skipped {file_name}: not an .xlsx file")
                continue
            if file_name.startswith("~"):
                print(f"Skipped {file_name}: temporary Excel file")
                continue

            print(f"Processing {file_name}")
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                file_label = parse_file_label(file_path)

                empirical_sheet = find_sheet(wb, "Empirical Model")
                if empirical_sheet is None:
                    print(f"  skipped empirical extraction for {file_name}: sheet not found")
                else:
                    empirical_rows.extend(
                        extract_empirical_rows(empirical_sheet, wb, file_label, file_name)
                    )

                regression_sheet = find_sheet(wb, "Regression Model")
                if regression_sheet is None:
                    print(f"  skipped regression extraction for {file_name}: sheet not found")
                else:
                    regression_rows.extend(
                        extract_regression_rows(regression_sheet, wb, file_label, file_name)
                    )
                files_processed += 1
            except Exception as exc:  # pragma: no cover - excel runtime protection
                print(f"Skipped {file_name}: {exc}")
            finally:
                if wb is not None:
                    safe_close_source_workbook(wb)
    finally:
        app.quit()

    output_path = next_output_path(in_dir=in_dir, out_dir=out_dir)
    write_output_workbook(
        output_path=output_path,
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
    )

    print(f"Output: {output_path}")
    print(f"Files processed: {files_processed}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run()

#!/usr/bin/env python3
"""Extract empirical/regression model candidates from .xlsx workbooks.

This script opens each source workbook once, processes both model sheets while
the workbook is open, then closes the source workbook without saving.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
import re
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# Update these paths for your environment.
input_dir = Path("./input")
output_dir = Path("./output")

N_QUARTERS = 10

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


@dataclass(frozen=True)
class FileMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_2d(values: Any) -> List[List[Any]]:
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def to_number(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip()
        if not cleaned:
            return None
        percent = cleaned.endswith("%")
        cleaned = cleaned.replace(",", "").replace("$", "").replace("%", "")
        try:
            parsed = float(cleaned)
        except ValueError:
            return None
        return parsed / 100.0 if percent else parsed
    return None


def value_to_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    return str(value).strip()


def round_key(value: Optional[float]) -> Optional[float]:
    if value is None:
        return None
    return round(value, 10)


def month_to_number(month_token: str) -> int:
    cleaned = month_token.strip().lower()
    if not cleaned:
        raise ValueError("missing month token")

    month_full = [
        "january",
        "february",
        "march",
        "april",
        "may",
        "june",
        "july",
        "august",
        "september",
        "october",
        "november",
        "december",
    ]
    aliases: Dict[str, int] = {}
    for idx, month_name in enumerate(month_full, start=1):
        aliases[month_name] = idx
        aliases[month_name[:3]] = idx
    aliases["sept"] = 9

    if cleaned in aliases:
        return aliases[cleaned]
    short = cleaned[:3]
    if short in aliases:
        return aliases[short]
    raise ValueError(f"unknown month token '{month_token}'")


def parse_file_metadata(file_name: str) -> FileMetadata:
    # Example: MedMiner_Model - AORT - MidJan2026_Send.xlsx
    stem = Path(file_name).stem
    match = re.search(
        r"\s-\s(?P<ticker>[A-Za-z0-9]+)\s-\s(?P<period>(Early|Mid|Late)[A-Za-z]+\d{4})",
        stem,
        re.IGNORECASE,
    )
    if not match:
        raise ValueError("filename does not match expected pattern")

    ticker = match.group("ticker").upper()
    period_token = match.group("period")

    period_match = re.fullmatch(
        r"(Early|Mid|Late)([A-Za-z]+)(\d{4})",
        period_token,
        re.IGNORECASE,
    )
    if not period_match:
        raise ValueError("model period token is invalid")

    timing = period_match.group(1).capitalize()
    month_token = period_match.group(2)
    year = int(period_match.group(3))
    month_num = month_to_number(month_token)
    month_label = datetime(year, month_num, 1).strftime("%b")

    day_map = {"Early": 5, "Mid": 15, "Late": 25}
    model_period = f"{timing}{month_label}_{year}"
    model_date = date(year, month_num, day_map[timing]).isoformat()
    model = f"{ticker}_{model_period}"
    return FileMetadata(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def pick_output_path(in_dir: Path, out_dir: Path) -> Path:
    base_name = f"{in_dir.name}_PARAM.xlsx"
    output_path = out_dir / base_name
    suffix = 1
    while output_path.exists():
        output_path = out_dir / f"{in_dir.name}_PARAM.{suffix}.xlsx"
        suffix += 1
    return output_path


def get_used_values(sheet: xw.Sheet) -> Tuple[int, int, List[List[Any]]]:
    used = sheet.used_range
    return used.row, used.column, normalize_2d(used.value)


def find_text_positions(
    values: List[List[Any]],
    top_row: int,
    left_col: int,
    text_value: str,
) -> List[Tuple[int, int]]:
    target = text_value.strip().lower()
    positions: List[Tuple[int, int]] = []
    for row_idx, row_values in enumerate(values):
        for col_idx, value in enumerate(row_values):
            if isinstance(value, str) and value.strip().lower() == target:
                positions.append((top_row + row_idx, left_col + col_idx))
    return positions


def find_best_max_anchor(sheet: xw.Sheet) -> Tuple[int, int]:
    top_row, left_col, values = get_used_values(sheet)
    max_positions = find_text_positions(values, top_row, left_col, "max")
    if not max_positions:
        raise ValueError("could not find 'max' anchor")

    min_positions = find_text_positions(values, top_row, left_col, "min")

    scored: List[Tuple[int, int, int]] = []
    for row, col in max_positions:
        score = 0
        if to_number(sheet.cells(row, col + 1).value) is not None:
            score += 2
        if to_number(sheet.cells(row + 1, col + 1).value) is not None:
            score += 1
        if any(abs(row - r2) <= 5 and abs(col - c2) <= 3 for r2, c2 in min_positions):
            score += 2
        scored.append((score, row, col))

    scored.sort(key=lambda item: (-item[0], item[1], item[2]))
    _, anchor_row, anchor_col = scored[0]
    return anchor_row, anchor_col


def local_labels(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    row_window: int = 40,
    col_window: int = 30,
) -> List[Tuple[str, int, int]]:
    start_row = max(1, anchor_row - row_window)
    end_row = anchor_row + 8
    start_col = max(1, anchor_col - col_window)
    end_col = anchor_col + 8
    values = normalize_2d(sheet.range((start_row, start_col), (end_row, end_col)).value)

    labels: List[Tuple[str, int, int]] = []
    for row_idx, row_values in enumerate(values):
        for col_idx, value in enumerate(row_values):
            if isinstance(value, str):
                txt = value.strip().lower()
                if txt:
                    labels.append((txt, start_row + row_idx, start_col + col_idx))
    return labels


def resolve_column(
    labels: Sequence[Tuple[str, int, int]],
    keywords: Sequence[str],
    anchor_col: int,
    default_offset: int,
) -> int:
    keyword_set = [kw.lower() for kw in keywords]
    candidates: List[int] = []
    for text, _, col in labels:
        if any(keyword in text for keyword in keyword_set):
            candidates.append(col)
    if candidates:
        candidates.sort(key=lambda col: (abs(col - anchor_col), col))
        return max(1, candidates[0])
    return max(1, anchor_col + default_offset)


def find_numeric_near_label(
    sheet: xw.Sheet,
    labels: Sequence[Tuple[str, int, int]],
    keywords: Sequence[str],
) -> Optional[float]:
    keyword_set = [kw.lower() for kw in keywords]
    for text, row, col in labels:
        if not any(keyword in text for keyword in keyword_set):
            continue
        for row_offset, col_offset in ((0, 1), (1, 0), (0, -1), (1, 1), (-1, 1), (-1, 0)):
            numeric_value = to_number(sheet.cells(row + row_offset, col + col_offset).value)
            if numeric_value is not None:
                return numeric_value
    return None


def resolve_max_min(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    labels: Sequence[Tuple[str, int, int]],
) -> Tuple[Optional[float], Optional[float]]:
    max_value = None
    for row_offset, col_offset in ((0, 1), (1, 1), (0, -1), (1, 0)):
        max_value = to_number(sheet.cells(anchor_row + row_offset, anchor_col + col_offset).value)
        if max_value is not None:
            break

    min_value = find_numeric_near_label(sheet, labels, ["min"])
    if min_value is None:
        for row_offset, col_offset in ((1, 1), (1, 0), (2, 1), (2, 0)):
            min_value = to_number(sheet.cells(anchor_row + row_offset, anchor_col + col_offset).value)
            if min_value is not None:
                break
    return max_value, min_value


def collect_history_rows(
    sheet: xw.Sheet,
    start_row: int,
    end_row: int,
    required_cols: Sequence[int],
    fallback_cols: Optional[Sequence[int]] = None,
) -> List[int]:
    if start_row > end_row:
        return []
    valid_required_cols = [col for col in required_cols if col >= 1]
    if not valid_required_cols:
        return []

    min_col = min(valid_required_cols)
    max_col = max(valid_required_cols)
    values = normalize_2d(sheet.range((start_row, min_col), (end_row, max_col)).value)
    rows: List[int] = []

    for idx, row_values in enumerate(values):
        row_number = start_row + idx
        if all(to_number(row_values[col - min_col]) is not None for col in valid_required_cols):
            rows.append(row_number)

    if rows or not fallback_cols:
        return rows

    valid_fallback_cols = [col for col in fallback_cols if col >= 1]
    if not valid_fallback_cols:
        return rows

    min_col_fb = min(valid_fallback_cols)
    max_col_fb = max(valid_fallback_cols)
    values_fb = normalize_2d(sheet.range((start_row, min_col_fb), (end_row, max_col_fb)).value)
    for idx, row_values in enumerate(values_fb):
        row_number = start_row + idx
        if all(to_number(row_values[col - min_col_fb]) is not None for col in valid_fallback_cols):
            rows.append(row_number)
    return sorted(set(rows))


def close_source_workbook(workbook: Optional[xw.Book]) -> None:
    if workbook is None:
        return
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    for closer in (
        lambda: workbook.close(False),
        lambda: workbook.api.Close(SaveChanges=False),
        lambda: workbook.api.Close(False),
    ):
        try:
            closer()
            return
        except Exception:
            continue


def extract_empirical_rows(
    workbook: xw.Book,
    metadata: FileMetadata,
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = workbook.sheets["Empirical Model"]
    except Exception:
        print(f"skipped file: {source_file} (reason: missing 'Empirical Model' sheet)")
        return []

    anchor_row, anchor_col = find_best_max_anchor(sheet)
    labels = local_labels(sheet, anchor_row, anchor_col)

    quarter_col = resolve_column(labels, ["quarter", "qtr"], anchor_col, -13)
    quarterly_sales_col = resolve_column(labels, ["quarterly sales", "q sales", "sales"], anchor_col, -11)
    reported_sales_col = resolve_column(labels, ["reported sales", "actual sales", "reported"], anchor_col, -9)
    penetration_col = resolve_column(labels, ["penetration"], anchor_col, -7)
    estimated_total_col = resolve_column(labels, ["estimated total sold", "est total sold", "total sold"], anchor_col, -5)

    max_value, min_value = resolve_max_min(sheet, anchor_row, anchor_col, labels)
    range_width = (max_value - min_value) if max_value is not None and min_value is not None else None

    history_rows = collect_history_rows(
        sheet=sheet,
        start_row=max(1, anchor_row - 250),
        end_row=anchor_row - 1,
        required_cols=[penetration_col, reported_sales_col],
        fallback_cols=[penetration_col],
    )
    if not history_rows:
        print(f"skipped file: {source_file} (reason: no empirical history rows)")
        return []

    max_n = min(N_QUARTERS, len(history_rows))
    end_row = history_rows[-1]
    scratch_col = max(sheet.used_range.last_cell.column + 2, anchor_col + 20)
    avg_cell = sheet.cells(anchor_row, scratch_col)
    forecast_cell = sheet.cells(anchor_row + 1, scratch_col)

    rows: List[Dict[str, Any]] = []
    for n_quarters in range(1, max_n + 1):
        start_row = history_rows[-n_quarters]
        avg_cell.formula2 = f"=AVERAGE(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})"
        forecast_cell.formula2 = (
            f'=IFERROR(R{end_row}C{reported_sales_col}/R{avg_cell.row}C{avg_cell.column},"")'
        )
        workbook.app.calculate()

        avg_penetration = to_number(avg_cell.value)
        reported_sales = to_number(sheet.cells(end_row, reported_sales_col).value)
        quarterly_sales = to_number(sheet.cells(end_row, quarterly_sales_col).value)
        forecast_value = to_number(sheet.cells(end_row, estimated_total_col).value)
        if forecast_value is None:
            forecast_value = to_number(forecast_cell.value)

        first_quarter_sales = to_number(sheet.cells(start_row, quarterly_sales_col).value)
        growth_rate_pct = None
        if (
            first_quarter_sales is not None
            and quarterly_sales is not None
            and first_quarter_sales != 0
            and n_quarters > 1
        ):
            growth_rate_pct = ((quarterly_sales / first_quarter_sales) - 1.0) * 100.0

        sales_captured_pct = None
        if reported_sales is not None and forecast_value not in (None, 0):
            sales_captured_pct = (reported_sales / forecast_value) * 100.0

        row = {
            "model": metadata.model,
            "ticker": metadata.ticker,
            "model_period": metadata.model_period,
            "model_date": metadata.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration,
            "num_quarters_used": n_quarters,
            "last_quarter_used": value_to_text(sheet.cells(end_row, quarter_col).value),
            "forecast_value": forecast_value,
            "actual_value": reported_sales,
            "forecast_max": max_value,
            "forecast_min": min_value,
            "range_width": range_width,
            "avg_penetration_pct": avg_penetration,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate_pct,
            "sales_captured_in_db_pct": sales_captured_pct,
            "source_file": source_file,
        }
        rows.append(row)

    sheet.range((anchor_row, scratch_col), (anchor_row + 2, scratch_col)).clear_contents()
    return rows


def pick_forecast_x_row(sheet: xw.Sheet, anchor_row: int, x_col: int, fallback_row: int) -> int:
    for candidate_row in (anchor_row, anchor_row + 1, anchor_row + 2, fallback_row):
        if to_number(sheet.cells(candidate_row, x_col).value) is not None:
            return candidate_row
    return fallback_row


def extract_regression_rows(
    workbook: xw.Book,
    metadata: FileMetadata,
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = workbook.sheets["Regression Model"]
    except Exception:
        print(f"skipped file: {source_file} (reason: missing 'Regression Model' sheet)")
        return []

    anchor_row, anchor_col = find_best_max_anchor(sheet)
    labels = local_labels(sheet, anchor_row, anchor_col)

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if x_col < 1 or y_col < 1:
        print(f"skipped file: {source_file} (reason: invalid regression anchor offsets)")
        return []

    max_value, min_value = resolve_max_min(sheet, anchor_row, anchor_col, labels)
    range_width = (max_value - min_value) if max_value is not None and min_value is not None else None
    actual_value = find_numeric_near_label(sheet, labels, ["actual"])

    history_rows = collect_history_rows(
        sheet=sheet,
        start_row=max(1, anchor_row - 250),
        end_row=anchor_row - 1,
        required_cols=[x_col, y_col],
    )
    if not history_rows:
        print(f"skipped file: {source_file} (reason: no regression history rows)")
        return []

    max_n = min(N_QUARTERS, len(history_rows))
    end_row = history_rows[-1]
    forecast_x_row = pick_forecast_x_row(sheet, anchor_row, x_col, end_row)

    scratch_col = max(sheet.used_range.last_cell.column + 2, anchor_col + 20)
    intercept_cell = sheet.cells(anchor_row, scratch_col)
    slope_cell = sheet.cells(anchor_row + 1, scratch_col)
    forecast_cell = sheet.cells(anchor_row + 2, scratch_col)

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Optional[float], ...]] = None
    for n_quarters in range(1, max_n + 1):
        start_row = history_rows[-n_quarters]
        intercept_cell.formula2 = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_cell.formula2 = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        forecast_cell.formula2 = (
            f"=R{intercept_cell.row}C{intercept_cell.column}"
            f"+R{slope_cell.row}C{slope_cell.column}*R{forecast_x_row}C{x_col}"
        )
        workbook.app.calculate()

        intercept = to_number(intercept_cell.value)
        slope = to_number(slope_cell.value)
        forecast_total = to_number(forecast_cell.value)
        if forecast_total is None:
            x_value = to_number(sheet.cells(forecast_x_row, x_col).value)
            if intercept is not None and slope is not None and x_value is not None:
                forecast_total = intercept + slope * x_value

        signature = (
            round_key(forecast_total),
            round_key(intercept),
            round_key(slope),
            round_key(max_value),
            round_key(min_value),
        )
        if signature == previous_signature:
            continue
        previous_signature = signature

        row = {
            "model": metadata.model,
            "ticker": metadata.ticker,
            "model_period": metadata.model_period,
            "model_date": metadata.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n_quarters,
            "num_quarters_used": n_quarters,
            "forecast_value": forecast_total,
            "actual_value": actual_value,
            "forecast_max": max_value,
            "forecast_min": min_value,
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }
        rows.append(row)

    sheet.range((anchor_row, scratch_col), (anchor_row + 3, scratch_col)).clear_contents()
    return rows


def write_sheet(
    worksheet,
    headers: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    worksheet.append(list(headers))
    for row in rows:
        worksheet.append(
            [("" if row.get(header) is None else row.get(header)) for header in headers]
        )

    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    preview_row_limit = min(worksheet.max_row, 200)
    for col in range(1, worksheet.max_column + 1):
        max_len = 0
        for row in range(1, preview_row_limit + 1):
            value = worksheet.cell(row=row, column=col).value
            value_len = len(str(value)) if value is not None else 0
            if value_len > max_len:
                max_len = value_len
        worksheet.column_dimensions[get_column_letter(col)].width = min(max(max_len + 2, 12), 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    empirical_sheet = workbook.create_sheet("empirical_candidates")
    regression_sheet = workbook.create_sheet("regression_candidates")
    write_sheet(empirical_sheet, EMPIRICAL_HEADERS, empirical_rows)
    write_sheet(regression_sheet, REGRESSION_HEADERS, regression_rows)

    workbook.save(output_path)


def gather_source_files(in_dir: Path) -> List[Path]:
    paths = [path for path in sorted(in_dir.iterdir(), key=lambda p: p.name.lower()) if path.is_file()]
    return paths


def should_skip_file(file_path: Path) -> Optional[str]:
    name = file_path.name
    if name.startswith("~"):
        return "temp file starts with '~'"
    if file_path.suffix.lower() != ".xlsx":
        return "not an .xlsx file"
    if re.search(r"_param(?:\.\d+)?\.xlsx$", name, re.IGNORECASE):
        return "looks like an output workbook"
    return None


def main() -> None:
    in_dir = input_dir.expanduser().resolve()
    out_dir = output_dir.expanduser().resolve()

    if not in_dir.exists():
        raise FileNotFoundError(f"input_dir does not exist: {in_dir}")
    if not in_dir.is_dir():
        raise NotADirectoryError(f"input_dir is not a directory: {in_dir}")

    out_dir.mkdir(parents=True, exist_ok=True)
    output_path = pick_output_path(in_dir, out_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    source_files = gather_source_files(in_dir)
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in source_files:
            skip_reason = should_skip_file(file_path)
            if skip_reason:
                print(f"skipped file: {file_path.name} (reason: {skip_reason})")
                continue

            try:
                metadata = parse_file_metadata(file_path.name)
            except ValueError as exc:
                print(f"skipped file: {file_path.name} (reason: {exc})")
                continue

            workbook: Optional[xw.Book] = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                file_empirical_rows = extract_empirical_rows(workbook, metadata, file_path.name)
                file_regression_rows = extract_regression_rows(workbook, metadata, file_path.name)
                empirical_rows.extend(file_empirical_rows)
                regression_rows.extend(file_regression_rows)
                files_processed += 1
                print(
                    "processed file: "
                    f"{file_path.name} "
                    f"(empirical_rows={len(file_empirical_rows)}, regression_rows={len(file_regression_rows)})"
                )
            except Exception as exc:
                print(f"skipped file: {file_path.name} (reason: processing error: {exc})")
            finally:
                close_source_workbook(workbook)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)
    print(f"output path: {output_path}")
    print(f"number of files processed: {files_processed}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()

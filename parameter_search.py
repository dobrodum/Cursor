#!/usr/bin/env python3
"""Search for the best data-derived parameter setting.

This script helps compare candidate parameters (including proportions/ratios)
against a binary target and finds the best threshold by a chosen metric.
"""

from __future__ import annotations

import argparse
import csv
import math
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Sequence, Tuple


EPSILON = 1e-12


@dataclass
class ScoreResult:
    candidate: str
    threshold: float
    metric_name: str
    metric_value: float
    accuracy: float
    precision: float
    recall: float
    f1: float
    tp: int
    fp: int
    tn: int
    fn: int
    evaluated_rows: int


def safe_float(value: str) -> Optional[float]:
    if value is None:
        return None
    text = value.strip()
    if text == "":
        return None
    try:
        parsed = float(text)
    except ValueError:
        return None
    if math.isnan(parsed) or math.isinf(parsed):
        return None
    return parsed


def safe_binary(value: str, positive_label: str) -> Optional[int]:
    if value is None:
        return None
    text = value.strip()
    if text == "":
        return None
    if text == positive_label:
        return 1
    if text in {"0", "0.0", "false", "False", "FALSE", "no", "No", "NO"}:
        return 0
    if text in {"1", "1.0", "true", "True", "TRUE", "yes", "Yes", "YES"}:
        return 1
    return None


def read_csv(path: str) -> List[Dict[str, str]]:
    with open(path, "r", encoding="utf-8", newline="") as handle:
        reader = csv.DictReader(handle)
        if not reader.fieldnames:
            raise ValueError("CSV file has no header row.")
        return list(reader)


def parse_ratio_specs(specs: Sequence[str]) -> List[Tuple[str, str, str]]:
    parsed: List[Tuple[str, str, str]] = []
    for raw in specs:
        if ":" not in raw or "/" not in raw:
            raise ValueError(
                f"Invalid ratio spec '{raw}'. Expected format: "
                "feature_name:numerator_column/denominator_column"
            )
        name, expr = raw.split(":", 1)
        numerator, denominator = expr.split("/", 1)
        name = name.strip()
        numerator = numerator.strip()
        denominator = denominator.strip()
        if not name or not numerator or not denominator:
            raise ValueError(f"Invalid ratio spec '{raw}'. Empty piece detected.")
        parsed.append((name, numerator, denominator))
    return parsed


def parse_thresholds(raw: str) -> List[float]:
    values: List[float] = []
    for item in raw.split(","):
        value = safe_float(item)
        if value is None:
            raise ValueError(f"Threshold '{item}' is not numeric.")
        values.append(value)
    if not values:
        raise ValueError("No thresholds were provided.")
    unique_sorted = sorted(set(values))
    return unique_sorted


def auto_thresholds(values: Sequence[float], bins: int) -> List[float]:
    sorted_values = sorted(values)
    if not sorted_values:
        return []
    if len(sorted_values) == 1:
        return sorted_values
    picks = set()
    for idx in range(bins + 1):
        pos = int((idx / bins) * (len(sorted_values) - 1))
        picks.add(sorted_values[pos])
    return sorted(picks)


def metrics(tp: int, fp: int, tn: int, fn: int) -> Dict[str, float]:
    total = tp + fp + tn + fn
    accuracy = (tp + tn) / max(total, EPSILON)
    precision = tp / max(tp + fp, EPSILON)
    recall = tp / max(tp + fn, EPSILON)
    f1 = 2 * precision * recall / max(precision + recall, EPSILON)
    return {
        "accuracy": accuracy,
        "precision": precision,
        "recall": recall,
        "f1": f1,
    }


def evaluate_threshold(
    y_true: Sequence[int],
    scores: Sequence[float],
    threshold: float,
    direction: str,
) -> Dict[str, float]:
    tp = fp = tn = fn = 0
    for truth, score in zip(y_true, scores):
        predicted_positive = score >= threshold if direction == "above" else score <= threshold
        if predicted_positive and truth == 1:
            tp += 1
        elif predicted_positive and truth == 0:
            fp += 1
        elif not predicted_positive and truth == 0:
            tn += 1
        else:
            fn += 1
    all_metrics = metrics(tp, fp, tn, fn)
    all_metrics.update({"tp": tp, "fp": fp, "tn": tn, "fn": fn, "n": len(y_true)})
    return all_metrics


def build_candidates(
    rows: Sequence[Dict[str, str]],
    candidate_cols: Sequence[str],
    ratio_specs: Sequence[Tuple[str, str, str]],
) -> Dict[str, List[Optional[float]]]:
    candidates: Dict[str, List[Optional[float]]] = {}

    for col in candidate_cols:
        values = [safe_float(row.get(col, "")) for row in rows]
        candidates[col] = values

    for name, num_col, den_col in ratio_specs:
        ratio_values: List[Optional[float]] = []
        for row in rows:
            num = safe_float(row.get(num_col, ""))
            den = safe_float(row.get(den_col, ""))
            if num is None or den is None or abs(den) < EPSILON:
                ratio_values.append(None)
            else:
                ratio_values.append(num / den)
        candidates[name] = ratio_values
    return candidates


def filter_valid_rows(
    targets: Sequence[Optional[int]],
    candidate: Sequence[Optional[float]],
) -> Tuple[List[int], List[float]]:
    y_true: List[int] = []
    scores: List[float] = []
    for t, c in zip(targets, candidate):
        if t is None or c is None:
            continue
        y_true.append(t)
        scores.append(c)
    return y_true, scores


def find_best_for_candidate(
    candidate_name: str,
    y_true: Sequence[int],
    scores: Sequence[float],
    thresholds: Sequence[float],
    metric_name: str,
    direction: str,
) -> Optional[ScoreResult]:
    if not y_true or not scores or not thresholds:
        return None

    best: Optional[ScoreResult] = None
    for threshold in thresholds:
        result = evaluate_threshold(y_true, scores, threshold, direction)
        candidate_result = ScoreResult(
            candidate=candidate_name,
            threshold=threshold,
            metric_name=metric_name,
            metric_value=result[metric_name],
            accuracy=result["accuracy"],
            precision=result["precision"],
            recall=result["recall"],
            f1=result["f1"],
            tp=result["tp"],
            fp=result["fp"],
            tn=result["tn"],
            fn=result["fn"],
            evaluated_rows=int(result["n"]),
        )
        if best is None:
            best = candidate_result
            continue

        better_metric = candidate_result.metric_value > best.metric_value + EPSILON
        tie_metric = abs(candidate_result.metric_value - best.metric_value) <= EPSILON
        tie_breaker = candidate_result.accuracy > best.accuracy + EPSILON
        if better_metric or (tie_metric and tie_breaker):
            best = candidate_result
    return best


def format_float(value: float) -> str:
    return f"{value:.4f}"


def print_results(results: Sequence[ScoreResult], metric_name: str, top_n: int) -> None:
    ordered = sorted(
        results,
        key=lambda r: (r.metric_value, r.accuracy, r.recall, r.precision),
        reverse=True,
    )
    if top_n > 0:
        ordered = ordered[:top_n]

    secondary_metrics = [name for name in ("accuracy", "precision", "recall", "f1") if name != metric_name]
    headers = ["rank", "candidate", "threshold", metric_name, *secondary_metrics, "n"]

    rows = []
    for idx, item in enumerate(ordered, start=1):
        secondary_values = [format_float(getattr(item, name)) for name in secondary_metrics]
        rows.append(
            [
                str(idx),
                item.candidate,
                format_float(item.threshold),
                format_float(item.metric_value),
                *secondary_values,
                str(item.evaluated_rows),
            ]
        )

    widths = [len(h) for h in headers]
    for row in rows:
        for i, cell in enumerate(row):
            widths[i] = max(widths[i], len(cell))

    def render_line(values: Iterable[str]) -> str:
        return " | ".join(value.ljust(widths[idx]) for idx, value in enumerate(values))

    print(render_line(headers))
    print("-+-".join("-" * width for width in widths))
    for row in rows:
        print(render_line(row))


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Compare candidate parameters from CSV data and find the best threshold "
            "for binary target prediction."
        )
    )
    parser.add_argument("--data", required=True, help="Path to CSV file.")
    parser.add_argument("--target-col", required=True, help="Binary target column.")
    parser.add_argument(
        "--positive-label",
        default="1",
        help="Value in target column treated as positive class (default: 1).",
    )
    parser.add_argument(
        "--candidate-cols",
        default="",
        help="Comma-separated numeric columns to evaluate directly.",
    )
    parser.add_argument(
        "--ratio",
        action="append",
        default=[],
        help=(
            "Ratio/proportion feature definition in form "
            "feature_name:numerator_column/denominator_column. "
            "Can be used multiple times."
        ),
    )
    parser.add_argument(
        "--thresholds",
        default="",
        help="Comma-separated thresholds. If omitted, uses value quantiles.",
    )
    parser.add_argument(
        "--threshold-bins",
        type=int,
        default=20,
        help="Quantile bins for automatic thresholds (default: 20).",
    )
    parser.add_argument(
        "--metric",
        choices=["accuracy", "precision", "recall", "f1"],
        default="f1",
        help="Metric used to choose best setting (default: f1).",
    )
    parser.add_argument(
        "--direction",
        choices=["above", "below"],
        default="above",
        help=(
            "How to convert score to positive class. "
            "'above' means score >= threshold, 'below' means score <= threshold."
        ),
    )
    parser.add_argument(
        "--top-n",
        type=int,
        default=10,
        help="How many top candidates to print (default: 10).",
    )
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    if args.threshold_bins < 1:
        raise ValueError("--threshold-bins must be at least 1.")

    rows = read_csv(args.data)
    candidate_cols = [c.strip() for c in args.candidate_cols.split(",") if c.strip()]
    ratio_specs = parse_ratio_specs(args.ratio)

    if not candidate_cols and not ratio_specs:
        raise ValueError("Add at least one --candidate-cols value or one --ratio spec.")

    targets = [
        safe_binary(row.get(args.target_col, ""), positive_label=args.positive_label)
        for row in rows
    ]

    candidates = build_candidates(rows, candidate_cols, ratio_specs)
    results: List[ScoreResult] = []

    for name, values in candidates.items():
        y_true, scores = filter_valid_rows(targets, values)
        if args.thresholds:
            thresholds = parse_thresholds(args.thresholds)
        else:
            thresholds = auto_thresholds(scores, bins=args.threshold_bins)

        best = find_best_for_candidate(
            candidate_name=name,
            y_true=y_true,
            scores=scores,
            thresholds=thresholds,
            metric_name=args.metric,
            direction=args.direction,
        )
        if best is not None:
            results.append(best)

    if not results:
        raise ValueError(
            "No valid rows available for evaluation. Check missing values and target labels."
        )

    print_results(results, metric_name=args.metric, top_n=args.top_n)
    overall_best = max(results, key=lambda r: (r.metric_value, r.accuracy))
    print()
    print(
        f"Best overall: candidate='{overall_best.candidate}', "
        f"threshold={overall_best.threshold:.4f}, "
        f"{args.metric}={overall_best.metric_value:.4f}"
    )


if __name__ == "__main__":
    main()

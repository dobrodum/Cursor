# Parameter Comparison Utility

This repository now includes `parameter_search.py`, a small script to:

1. Pull candidate parameters from CSV data.
2. Build proportion/ratio parameters (for example: `clicks / impressions`).
3. Compare each parameter by sweeping thresholds.
4. Report which parameter + threshold performs best for a binary target.

## Quick start

Run:

```bash
python3 parameter_search.py \
  --data data.csv \
  --target-col converted \
  --candidate-cols age,score \
  --ratio ctr:clicks/impressions \
  --ratio completion_rate:completed/assigned \
  --metric f1
```

### What this command does

- Evaluates existing numeric columns: `age`, `score`.
- Creates two proportion features:
  - `ctr = clicks / impressions`
  - `completion_rate = completed / assigned`
- For each candidate feature, tests multiple thresholds and picks the best one.
- Ranks candidates by the chosen metric (`f1` above).

## Key options

- `--data`: CSV file path.
- `--target-col`: binary label column.
- `--positive-label`: value treated as positive class (default `1`).
- `--candidate-cols`: comma-separated existing numeric columns.
- `--ratio`: add derived proportion feature in the form
  `feature_name:numerator_column/denominator_column` (repeatable).
- `--metric`: `accuracy`, `precision`, `recall`, or `f1`.
- `--thresholds`: explicit thresholds (e.g. `0.1,0.2,0.3`). If omitted, automatic quantile thresholds are used.
- `--direction`: `above` (score >= threshold means positive) or `below`.
- `--top-n`: number of top candidates to show.

## Expected CSV format

Any CSV with headers is supported, as long as the referenced columns exist.
Example:

```csv
converted,clicks,impressions,completed,assigned,age,score
1,10,100,8,10,29,0.62
0,2,80,3,9,41,0.31
1,6,50,7,8,35,0.58
```

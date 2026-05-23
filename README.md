# Korean Stock Alpha Scoring Tool

A multi-factor alpha scoring model for Korean semiconductor materials/parts/equipment stocks (소부장).

## Overview

Fetches real-time market data via `yfinance` and ranks stocks across three quantitative factors to generate a composite alpha score and buy/neutral/sell signal.

## Factors

| Factor | Weight | Method |
|--------|--------|--------|
| Momentum | 40% | Risk-adjusted momentum: `(12M–1M return) / annualized volatility` |
| Quality | 35% | Average of ROE, operating margin, and EPS trend (YoY improvement) |
| Low Volatility | 25% | Inverse of 52-week daily return standard deviation |

Each factor is cross-sectionally ranked (0–100) across the universe. The composite alpha score is a weighted average of the three factor scores.

## Universe

20 KOSPI/KOSDAQ semiconductor materials, parts & equipment stocks including 한미반도체, ISC, 리노공업, HPSP, 솔브레인, and others.

## Output

Generates a formatted Excel report with:
- Per-stock raw metrics (return, volatility, ROE, operating margin, EPS trend)
- Factor scores and composite alpha score
- Buy / Neutral / Sell signal (threshold: ≥65 buy, ≤35 sell)

## Usage

```bash
pip install yfinance pandas openpyxl numpy
python alpha_score.py
```

Results are printed to terminal (top 10) and saved as `alpha_score_YYYYMMDD_HHMM.xlsx`.

## Requirements

- Python 3.8+
- yfinance, pandas, numpy, openpyxl

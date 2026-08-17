# My First Repo

This is a small practice project for learning the Git and GitHub pull request workflow with Claude Code.

## Getting Started

Clone this repo and start exploring!

## Weekly Stock Portfolio Analyzer

`data/portfolio.xlsx` tracks a small stocks portfolio (holdings + quarterly
dividends). Two scripts keep it current and turn it into a weekly report:

- **`scripts/fetch_prices.py`** - pulls current prices and dividend history
  from Yahoo Finance (via `yfinance`) and updates the spreadsheet's
  "Huidige waarde" / "Dividend %" columns, the current quarter's dividend
  cell, and appends a snapshot to the `History` sheet. Runs weekly via the
  `.github/workflows/weekly-fetch.yml` GitHub Action (Monday 05:00 UTC),
  since GitHub Actions runners have normal internet access.
- **`scripts/analyze_portfolio.py`** - reads the spreadsheet (no network
  needed) and writes a markdown report to `reports/<date>.md` covering
  performance, allocation, week-over-week dividend yield changes, and
  rule-based flags (concentration risk, underwater positions, possible
  dividend cuts, currency-mixing caveats).

Adding a new holding: append a row to the holdings table and the dividends
table in `data/portfolio.xlsx`, and add a ticker -> Yahoo Finance symbol
mapping in `scripts/ticker_map.py` if the plain ticker doesn't resolve
directly (e.g. exchange-specific symbols like ING Groep's `INGA.AS`).

> **Note:** scheduled GitHub Actions workflows only fire on the repository's
> default branch, so `weekly-fetch.yml` won't run on a schedule until this
> is merged into `master` (manual `workflow_dispatch` runs work on any
> branch in the meantime).

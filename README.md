# Pricing Analyzer

Scrapes eBay seller feedback for three accounts (bidallies, directauth, cellfeee), matches each item against
WatchCount for a comp price, and builds a pricing analysis workbook per account.

Everything is in one file: `run.py`.

## Requirements

- Python 3.10+
- Google Chrome installed

`run.py` installs any missing Python packages itself (selenium, beautifulsoup4, openpyxl), so there's no
separate setup step.

## Usage

```
python run.py
```

Runs all three accounts one after another. Other options:

```
python run.py --account bidallies
python run.py --account cellfeee --skip-scrape   # reuse existing feedback.xlsx, redo pricing + analysis only
```

Each account goes through three steps: scrape feedback, look up WatchCount prices, build the analysis
workbook. If one account fails, it prints the error and moves to the next account instead of stopping the
whole run.

Output goes to `output/<account>/`:
- `feedback.xlsx` - scraped feedback with WatchCount price/title columns added
- `<account>_analysis_<date>.xlsx` - the final workbook (this is the one you want)
- `checkpoint.txt`, `seen_feedback_ids.txt`, `watchcount_prices.json` - resume/cache state, safe to delete for a clean re-run

## Heads up: not fully hands-off

The WatchCount lookup step opens a visible Chrome window. WatchCount sometimes throws up a verification
challenge - if that happens, someone needs to click through it in that window. The script waits up to 60
seconds for it to clear and then keeps going on its own. Not a bug, just how WatchCount works.

## Resuming a scrape

Feedback scraping is capped at `--max-pages` (55 by default) and checkpoints its progress, so a big account
can span several runs - just run the same command again and it continues from where it stopped. If a run got
interrupted hours ago, delete that account's `checkpoint.txt` and `seen_feedback_ids.txt` first, since eBay
sorts newest-first and a stale checkpoint can skip feedback that arrived since.

## Optional WatchCount login

Not required, anonymous search works fine. If you want to log in, set `WATCHCOUNT_EMAIL` and
`WATCHCOUNT_PASSWORD` as environment variables before running.

## If auto-install fails

```
pip install selenium beautifulsoup4 openpyxl
```

# Vexwire Pricing Pipeline

For each eBay seller account (`bidallies`, `directauth`, `cellfeee`), scrapes the last month's "Received as
Seller" feedback (Item Title / Item ID / Price), looks up a comparable WatchCount price for each item, and
builds a 2-sheet pricing analysis workbook.

Everything lives in one file: `run.py`.

## Prerequisites

- Python 3.10+
- Google Chrome installed

That's it — `run.py` installs any missing Python packages (selenium, beautifulsoup4, openpyxl) itself the
first time it runs. Selenium also manages its own matching Chromedriver, so no separate driver install is
needed.

## Usage

Clone the repo, then just run:

```
python run.py
```

That processes all three accounts, one after another. Other options:

```
python run.py --account bidallies                  # just one account
python run.py --account cellfeee --skip-scrape      # reuse existing feedback.xlsx, redo pricing + analysis only
```

Each account goes through three steps in order: scrape feedback -> look up WatchCount prices -> build the
analysis workbook. If a step fails, the run stops there with an error instead of silently continuing to the
next account.

Output lands in `output/<account>/`:
- `feedback.xlsx` - raw scraped feedback, with WatchCount Price/Title columns added
- `<account>_analysis_<date>.xlsx` - the final 2-sheet workbook (`Feedback` row-level, `Avg Price by Item` summary) — this is the file you want
- `checkpoint.txt`, `seen_feedback_ids.txt`, `watchcount_prices.json` - resume/cache state (safe to delete to force a clean re-scrape/re-lookup)

## Important: this cannot run fully unattended

The WatchCount price-lookup step opens a **visible** Chrome window. WatchCount occasionally shows a
verification challenge page — when it does, **a human needs to solve it in that window**. The script detects
the challenge, waits up to 60 seconds for it to clear, and continues automatically once it's gone. This is a
property of WatchCount's site, not a bug in this pipeline — just keep an eye on the Chrome window while a
lookup run is going.

## Resuming a long scrape

The feedback scrape fetches a bounded number of pages per run (`--max-pages`, default 55) and checkpoints
its progress, so a large account can be scraped across multiple runs — just re-run the same command and it
picks up where it left off. If a scrape was interrupted more than a few hours ago, delete that account's
`checkpoint.txt` and `seen_feedback_ids.txt` under `output/<account>/` before resuming, since eBay's feed is
sorted newest-first and a stale checkpoint can cause the first page or two of newly-arrived feedback to be
silently skipped.

## Optional: WatchCount login

Anonymous WatchCount search works fine for normal use. If you have WatchCount credentials and want to log in
first, set `WATCHCOUNT_EMAIL` and `WATCHCOUNT_PASSWORD` as environment variables before running — this is
optional.

## If the automatic dependency install fails

`run.py` runs `pip install selenium beautifulsoup4 openpyxl` on your behalf if any are missing. If that fails
(no internet, restricted permissions, etc.), install them yourself and re-run:

```
pip install selenium beautifulsoup4 openpyxl
```

# NEPSE Automated Stock Price Collector

A GitHub Actions + Python project that periodically collects the publicly available NEPSE stock snapshot from the YONEPSE public JSON feed and appends it to CSV and Excel databases.

## Schedule

The workflow runs every 15 minutes from 11:00 through 15:45 Nepal Standard Time (Asia/Kathmandu), Monday through Friday. You can also start a run manually from the Actions tab.

GitHub scheduled workflows are not guaranteed to start at the exact minute. The collection timestamp in each row records when this bot actually fetched the data.

## Important data-source limitation

The configured YONEPSE feed documents hourly updates during market hours. This collector requests the feed every 15 minutes, but multiple collection runs may contain the same underlying market snapshot. `Source Updated At` preserves the upstream timestamp so repeated/stale source data can be identified.

The public feed's documented sample includes LTP, previous close, high, low, volume, turnover, percentage change, trades and last-updated time. It may not include `Open`. The collector leaves missing fields blank; it never invents values.

Data source: https://github.com/Shubhamnpk/yonepse  
Endpoint: https://raw.githubusercontent.com/Shubhamnpk/yonepse/main/data/nepse_data.json

This is an independent community project, not an official NEPSE service. Check the source project's documentation, terms and availability before relying on it.

## Files

- `tracker.py` — fetches the feed and appends each stock snapshot.
- `data/stock_prices.csv` — append-only CSV database, Excel compatible.
- `data/stock_prices.xlsx` — append-only Excel workbook.
- `.github/workflows/collect.yml` — GitHub Actions schedule.
- `requirements.txt` — Python dependencies.

## Create the repository

1. Sign in to GitHub and choose **New repository**.
2. Name it `nepse-automated-tracker`.
3. Choose **Public** or **Private**. Private is fine.
4. Create the repository.
5. Upload the project files and folders, keeping their paths exactly as shown.
6. In **Settings → Actions → General → Workflow permissions**, ensure workflows can have read and write permissions. The workflow also declares `contents: write`.
7. Open **Actions**, select **NEPSE Stock Price Collector**, and choose **Run workflow** to test it.
8. Open the latest run and inspect its logs.
9. After a successful run, open the `data` folder. Both database files should appear after the workflow commits them.

## Downloading the data

Open `data/stock_prices.csv` or `data/stock_prices.xlsx` in the repository and use the GitHub download/raw option. Excel can open the CSV directly.

## Changing the schedule

Edit `.github/workflows/collect.yml`. The current expression:

```yaml
- cron: "*/15 11-15 * * 1-5"
  timezone: "Asia/Kathmandu"
```

means every 15 minutes during hours 11, 12, 13, 14 and 15. It therefore includes the selected hourly times 11:00, 12:00, 13:00, 14:00 and 15:00, plus the quarter-hour runs in between.

For only the five selected hourly times, use:

```yaml
- cron: "0 11-15 * * 1-5"
  timezone: "Asia/Kathmandu"
```

## Reliability notes

- GitHub Actions may delay scheduled starts, especially during high load.
- Runs happen Monday-Friday and may still occur on exchange holidays; check `Source Updated At` to identify stale feed data.
- A failed request fails the workflow instead of silently appending an empty dataset.
- The script appends one row per returned symbol per run. It does not backfill missed runs.
- This is a research/record-keeping tool, not a trading signal or guaranteed real-time market feed.

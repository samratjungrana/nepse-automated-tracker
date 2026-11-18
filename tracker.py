"""
NEPSE market collector
- Live fields from the YONEPSE public JSON feed.
- Daily opening price from the srijankarki07/nepse-data end-of-day archive.
- Stores snapshots in data/stock_prices.csv and data/stock_prices.xlsx.

The daily archive is published after the market session. Therefore, Open can be
blank in intraday snapshots initially and is backfilled automatically on later
runs when that day's official-session OHLC row becomes available.
"""

import csv
import io
import os
import sys
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

TIMEZONE = ZoneInfo("Asia/Kathmandu")
LIVE_URL = "https://raw.githubusercontent.com/Shubhamnpk/yonepse/main/data/nepse_data.json"
DAILY_OHLC_URL = (
    "https://raw.githubusercontent.com/srijankarki07/nepse-data/main/"
    "data/daily/{year}/{date}.csv"
)

DATA_DIR = "data"
CSV_PATH = os.path.join(DATA_DIR, "stock_prices.csv")
XLSX_PATH = os.path.join(DATA_DIR, "stock_prices.xlsx")

HEADERS = [
    "Collection Date",
    "Collection Time",
    "Collection Timestamp",
    "Symbol",
    "LTP",
    "Open",
    "High",
    "Low",
    "Volume",
    "Turnover",
    "% Change",
    "Total Trades",
    "Previous Close",
    "Source Updated At",
]

SESSION = requests.Session()
SESSION.headers.update({"User-Agent": "NEPSE-Price-Collector/1.0"})


def fetch_live_data():
    response = SESSION.get(LIVE_URL, timeout=30)
    response.raise_for_status()
    payload = response.json()
    if not isinstance(payload, list):
        raise ValueError("Unexpected YONEPSE response: expected a list of stocks.")
    return payload


def fetch_daily_open(trading_date):
    """Return {SYMBOL: open} for one session; None if archive is not published."""
    date_text = trading_date.strftime("%Y-%m-%d")
    url = DAILY_OHLC_URL.format(year=trading_date.year, date=date_text)

    try:
        response = SESSION.get(url, timeout=20)
        if response.status_code == 404:
            return None
        response.raise_for_status()
        reader = csv.DictReader(io.StringIO(response.text))
        if not reader.fieldnames or not {"date", "symbol", "open"}.issubset(
            set(reader.fieldnames)
        ):
            raise ValueError(f"Unexpected OHLC columns for {date_text}: {reader.fieldnames}")

        result = {}
        for item in reader:
            symbol = (item.get("symbol") or "").strip().upper()
            open_value = (item.get("open") or "").strip()
            row_date = (item.get("date") or "").strip()
            if symbol and row_date == date_text and open_value:
                result[symbol] = open_value
        return result
    except requests.RequestException as exc:
        print(f"Could not retrieve daily opening prices for {date_text}: {exc}")
        return None


def load_existing_rows():
    if not os.path.exists(CSV_PATH):
        return []

    with open(CSV_PATH, "r", newline="", encoding="utf-8-sig") as file:
        reader = csv.DictReader(file)
        if not reader.fieldnames:
            return []
        rows = list(reader)

    # Keep compatibility with earlier versions if a column was absent.
    for row in rows:
        for header in HEADERS:
            row.setdefault(header, "")
    return rows


def backfill_open_prices(rows, today):
    """Fill blank Open cells for recent recorded dates when the archive is ready."""
    dates_to_check = set()
    for row in rows:
        if not (row.get("Open") or "").strip():
            try:
                day = datetime.strptime(row["Collection Date"], "%Y-%m-%d").date()
                if today - timedelta(days=10) <= day <= today:
                    dates_to_check.add(day)
            except (KeyError, ValueError):
                continue

    changed = 0
    for day in sorted(dates_to_check):
        opening_prices = fetch_daily_open(day)
        if not opening_prices:
            continue
        for row in rows:
            if row.get("Collection Date") == day.strftime("%Y-%m-%d"):
                symbol = (row.get("Symbol") or "").strip().upper()
                if not (row.get("Open") or "").strip() and symbol in opening_prices:
                    row["Open"] = opening_prices[symbol]
                    changed += 1
        print(f"Opening prices checked for {day}: {len(opening_prices)} symbols")

    return changed


def value(stock, *keys):
    for key in keys:
        result = stock.get(key)
        if result is not None and result != "":
            return result
    return ""


def make_snapshot_rows(stocks, now):
    date_text = now.strftime("%Y-%m-%d")
    time_text = now.strftime("%H:%M:%S")
    timestamp = now.strftime("%Y-%m-%d %H:%M:%S")

    rows = []
    for stock in stocks:
        symbol = str(value(stock, "symbol", "Symbol") or "").strip().upper()
        if not symbol:
            continue

        rows.append({
            "Collection Date": date_text,
            "Collection Time": time_text,
            "Collection Timestamp": timestamp,
            "Symbol": symbol,
            "LTP": value(stock, "ltp", "LTP"),
            # YONEPSE does not currently publish Open; daily archive fills it later.
            "Open": value(stock, "open", "Open"),
            "High": value(stock, "high", "High"),
            "Low": value(stock, "low", "Low"),
            "Volume": value(stock, "volume", "Volume"),
            "Turnover": value(stock, "turnover", "Turnover"),
            "% Change": value(stock, "percent_change", "percentage_change", "change_percent", "% Change"),
            "Total Trades": value(stock, "trades", "total_trades", "Total Trades"),
            "Previous Close": value(stock, "previous_close", "previousClose", "Previous Close"),
            "Source Updated At": value(stock, "last_updated", "updated_at", "timestamp"),
        })

    return rows


def write_csv(rows):
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(CSV_PATH, "w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=HEADERS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def write_excel(rows):
    os.makedirs(DATA_DIR, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Stock Prices"
    sheet.append(HEADERS)

    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")

    for row in rows:
        sheet.append([row.get(header, "") for header in HEADERS])

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    for column_number, header in enumerate(HEADERS, start=1):
        values = [len(str(header))]
        values.extend(
            len(str(sheet.cell(row=r, column=column_number).value or ""))
            for r in range(2, min(sheet.max_row, 250) + 1)
        )
        sheet.column_dimensions[get_column_letter(column_number)].width = min(max(max(values) + 2, 12), 28)

    workbook.save(XLSX_PATH)


def main():
    os.makedirs(DATA_DIR, exist_ok=True)
    now = datetime.now(TIMEZONE)
    today = now.date()

    rows = load_existing_rows()

    # Backfill Open values in earlier snapshots once the daily archive is published.
    backfilled = backfill_open_prices(rows, today)

    # Collect a fresh live snapshot.
    stocks = fetch_live_data()
    new_rows = make_snapshot_rows(stocks, now)

    # Avoid duplicate rows if GitHub Actions retries the same collection timestamp.
    existing_keys = {
        (row.get("Collection Timestamp", ""), row.get("Symbol", "").upper())
        for row in rows
    }
    added = 0
    for row in new_rows:
        key = (row["Collection Timestamp"], row["Symbol"])
        if key not in existing_keys:
            rows.append(row)
            existing_keys.add(key)
            added += 1

    # If today's archive is already available, fill Open in the just-added rows too.
    today_open = fetch_daily_open(today)
    if today_open:
        for row in rows:
            if row.get("Collection Date") == today.strftime("%Y-%m-%d"):
                symbol = (row.get("Symbol") or "").strip().upper()
                if not (row.get("Open") or "").strip() and symbol in today_open:
                    row["Open"] = today_open[symbol]

    write_csv(rows)
    write_excel(rows)

    print(f"Collection completed at {now.isoformat()}")
    print(f"Live symbols received: {len(new_rows)}")
    print(f"New snapshot rows added: {added}")
    print(f"Previously blank Open cells backfilled: {backfilled}")
    print(f"Files updated: {CSV_PATH}, {XLSX_PATH}")


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        print(f"Collector failed: {error}", file=sys.stderr)
        raise

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

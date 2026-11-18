import csv
import json
import os
import sys
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

BASE_URL = "https://raw.githubusercontent.com/Shubhamnpk/yonepse/main"
DATA_URL = f"{BASE_URL}/data/nepse_data.json"
CSV_PATH = Path("data/stock_prices.csv")
XLSX_PATH = Path("data/stock_prices.xlsx")
TIMEZONE = ZoneInfo("Asia/Kathmandu")

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

def first_present(record, *keys):
    """Return the first non-empty source value among possible field names."""
    for key in keys:
        value = record.get(key)
        if value is not None and value != "":
            return value
    return ""

def fetch_stocks():
    response = requests.get(
        DATA_URL,
        headers={"User-Agent": "NEPSE-Historical-Collector/1.0"},
        timeout=45,
    )
    response.raise_for_status()
    payload = response.json()

    if not isinstance(payload, list):
        raise ValueError("Unexpected API response: expected a JSON list.")

    stocks = [
        item for item in payload
        if isinstance(item, dict) and first_present(item, "symbol")
    ]
    if not stocks:
        raise ValueError("The API returned no stock records; refusing to write an empty snapshot.")
    return stocks

def make_rows(stocks, captured_at):
    date_text = captured_at.strftime("%Y-%m-%d")
    time_text = captured_at.strftime("%H:%M:%S")
    timestamp_text = captured_at.isoformat(timespec="seconds")

    rows = []
    for stock in stocks:
        rows.append({
            "Collection Date": date_text,
            "Collection Time": time_text,
            "Collection Timestamp": timestamp_text,
            "Symbol": first_present(stock, "symbol"),
            "LTP": first_present(stock, "ltp"),
            # The public feed may not provide today's opening price.
            "Open": first_present(stock, "open", "open_price"),
            "High": first_present(stock, "high"),
            "Low": first_present(stock, "low"),
            "Volume": first_present(stock, "volume"),
            "Turnover": first_present(stock, "turnover"),
            "% Change": first_present(stock, "percent_change", "per_change"),
            "Total Trades": first_present(stock, "trades", "total_trades"),
            "Previous Close": first_present(stock, "previous_close", "prev_close"),
            "Source Updated At": first_present(stock, "last_updated", "updated_at"),
        })

    return sorted(rows, key=lambda row: str(row["Symbol"]))

def append_csv(rows):
    CSV_PATH.parent.mkdir(parents=True, exist_ok=True)
    needs_header = not CSV_PATH.exists() or CSV_PATH.stat().st_size == 0

    with CSV_PATH.open("a", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=HEADERS)
        if needs_header:
            writer.writeheader()
        writer.writerows(rows)

def style_worksheet(ws):
    header_fill = PatternFill("solid", fgColor="17365D")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    widths = {
        "A": 14, "B": 13, "C": 26, "D": 15,
        "E": 13, "F": 13, "G": 13, "H": 13,
        "I": 15, "J": 18, "K": 13, "L": 15,
        "M": 17, "N": 27,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

def append_excel(rows):
    XLSX_PATH.parent.mkdir(parents=True, exist_ok=True)

    if XLSX_PATH.exists():
        wb = load_workbook(XLSX_PATH)
        ws = wb["Stock Prices"] if "Stock Prices" in wb.sheetnames else wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Stock Prices"
        ws.append(HEADERS)

    if ws.max_row == 1 and ws.cell(1, 1).value != HEADERS[0]:
        ws.delete_rows(1, ws.max_row)
        ws.append(HEADERS)

    for row in rows:
        ws.append([row.get(header, "") for header in HEADERS])

    style_worksheet(ws)
    wb.save(XLSX_PATH)

def main():
    captured_at = datetime.now(TIMEZONE)
    stocks = fetch_stocks()
    rows = make_rows(stocks, captured_at)

    append_csv(rows)
    append_excel(rows)

    print(
        f"SUCCESS: saved {len(rows)} stock rows "
        f"for collection time {captured_at.isoformat(timespec='seconds')}"
    )
    print(f"CSV: {CSV_PATH}")
    print(f"Excel: {XLSX_PATH}")

if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"COLLECTION FAILED: {exc}", file=sys.stderr)
        raise

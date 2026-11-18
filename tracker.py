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

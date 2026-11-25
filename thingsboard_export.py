"""ThingsBoard PE raw telemetry exporter (ALL keys; monthly XLSX).

Run: py thingsboard_export_all_keys.py
Requires: pip install requests openpyxl tzdata
"""

import calendar
import csv
import getpass
import hashlib
import json
import os
import re
import sqlite3
import time
import uuid
from datetime import datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

import requests
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font, PatternFill

# ================= CONFIGURATION =================
TB_URL = "https://cranes.asia/"
TIMEZONE = ZoneInfo("Asia/Kuala_Lumpur")
PAGE_SIZE = 500              # readings per key per API response
WINDOW_DAYS = 7              # fewer empty-period checks; use 1 if server queries are slow
REQUEST_DELAY_SECONDS = 0.5 # minimum pause before each GET
MAX_ATTEMPTS = 4
SKIP_COMPLETED_MONTHS = True
RECHECK_NO_DATA_MONTHS = True # recheck in case historical data was added
PROGRESS_INTERVAL_SECONDS = 12
EXCEL_MAX_ROWS = 1_048_576   # including the header
# =================================================


def ask_int(prompt, minimum, maximum):
    while True:
        try:
            value = int(input(prompt).strip())
            if minimum <= value <= maximum:
                return value
        except ValueError:
            pass
        print(f"Please enter a number between {minimum} and {maximum}.")


class ThingsBoard:
    def __init__(self, url, username, password):
        self.url = url.rstrip("/")
        self.username = username
        self.password = password
        self.session = requests.Session()
        self.requests_count = 0
        self.login()

    def login(self):
        response = self.session.post(
            self.url + "/api/auth/login",
            json={"username": self.username, "password": self.password},
            timeout=(10, 30),
        )
        response.raise_for_status()
        self.session.headers["X-Authorization"] = "Bearer " + response.json()["token"]

    def get(self, path, params=None):
        refreshed = False
        for attempt in range(1, MAX_ATTEMPTS + 1):
            time.sleep(REQUEST_DELAY_SECONDS)
            try:
                self.requests_count += 1
                response = self.session.get(self.url + path, params=params, timeout=(10, 60))
                if response.status_code == 401 and not refreshed:
                    self.login()
                    refreshed = True
                    continue
                if response.status_code in (429, 500, 502, 503, 504):
                    if attempt == MAX_ATTEMPTS:
                        response.raise_for_status()
                    retry_header = response.headers.get("Retry-After")
                    try:
                        pause = min(120, max(1, float(retry_header)))
                    except (TypeError, ValueError):
                        pause = min(60, 2 ** attempt)
                    print(f"\nHTTP {response.status_code}; backing off {pause:.0f}s", flush=True)
                    time.sleep(pause)
                    continue
                response.raise_for_status()
                return response.json()
            except (requests.Timeout, requests.ConnectionError):
                if attempt == MAX_ATTEMPTS:
                    raise
                time.sleep(min(60, 2 ** attempt))
        raise RuntimeError("ThingsBoard API failed after maximum retries")


def month_bounds(year, month):
    start = datetime(year, month, 1, tzinfo=TIMEZONE)
    end = datetime(year + (month == 12), month % 12 + 1, 1, tzinfo=TIMEZONE)
    return start, end


def safe_excel(raw):
    if raw is None or isinstance(raw, (int, float, bool)):
        return raw
    if isinstance(raw, (dict, list)):
        raw = json.dumps(raw, ensure_ascii=False)
    val = str(raw)
    # Prevent untrusted device values from becoming spreadsheet formulas.
    if val.lstrip().startswith(("=", "+", "-", "@")) and not re.fullmatch(r"-?\d+(\.\d+)?", val.strip()):
        return "'" + val
    if re.fullmatch(r"-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?", val.strip()):
        # Preserve large IDs and leading-zero strings, which Excel otherwise changes.
        if len(re.sub(r"[^0-9]", "", val)) <= 15:
            try:
                return float(val) if any(c in val.lower() for c in (".", "e")) else int(val)
            except ValueError:
                pass
    return val


def save_xlsx(db, keys, year, month, output):
    count = db.execute("SELECT COUNT(DISTINCT ts) FROM readings").fetchone()[0]
    if not count:
        return 0
    book = Workbook(write_only=True)
    cursor = db.execute("SELECT ts, key, value FROM readings ORDER BY ts, key")
    sheet = None
    sheet_number = 0
    rows_in_sheet = EXCEL_MAX_ROWS
    previous_ts = None
    row_values = {}

    def create_sheet():
        nonlocal sheet, sheet_number, rows_in_sheet
        sheet_number += 1
        sheet = book.create_sheet(f"{calendar.month_abbr[month]}_{sheet_number}")
        title_cells = []
        for value in ["Timestamp (Malaysia)", "Timestamp (Unix ms)"] + keys:
            cell = WriteOnlyCell(sheet, value=value)
            cell.font = Font(color="FFFFFF", bold=True)
            cell.fill = PatternFill("solid", fgColor="176B74")
            title_cells.append(cell)
        sheet.append(title_cells)
        rows_in_sheet = 1

    def write_record(ts, values):
        nonlocal rows_in_sheet
        if rows_in_sheet >= EXCEL_MAX_ROWS:
            create_sheet()
        local = datetime.fromtimestamp(ts / 1000, tz=TIMEZONE).replace(tzinfo=None)
        date_cell = WriteOnlyCell(sheet, value=local)
        date_cell.number_format = "yyyy-mm-dd hh:mm:ss.000"
        sheet.append([date_cell, ts] + [safe_excel(values.get(key)) for key in keys])
        rows_in_sheet += 1

    create_sheet()
    for ts, key, value in cursor:
        if previous_ts is not None and ts != previous_ts:
            write_record(previous_ts, row_values)
            row_values = {}
        previous_ts = ts
        row_values[key] = json.loads(value)
    if previous_ts is not None:
        write_record(previous_ts, row_values)
    temp_path = output.with_name(output.name + ".partial.xlsx")
    try:
        book.save(temp_path)
        os.replace(temp_path, output)
    finally:
        temp_path.unlink(missing_ok=True)
    return count


def export_month(tb, device_id, year, month, keys, out_dir, label):
    start, end = month_bounds(year, month)
    out_file = out_dir / f"TB_{device_id[:8]}_{year}_{month:02d}.xlsx"
    db_path = out_dir / f".{year}_{month:02d}.inprogress.sqlite"
    if db_path.exists():
        db_path.unlink()  # restart this month cleanly, never claim partial data is complete
    db = sqlite3.connect(db_path)
    db.execute("CREATE TABLE readings (ts INTEGER, key TEXT, value TEXT, PRIMARY KEY(ts,key))")
    downloaded = 0
    page_count = 0
    start_requests = tb.requests_count
    t0 = time.monotonic()
    total_windows = (end - start).days // WINDOW_DAYS + ((end - start).days % WINDOW_DAYS != 0)
    total_jobs = total_windows * len(keys)
    finished_jobs = 0
    last_print = 0
    try:
        window_start = start
        win = 0
        while window_start < end:
            win += 1
            window_end = min(window_start + timedelta(days=WINDOW_DAYS), end)
            begin = int(window_start.timestamp() * 1000)
            finish = int(window_end.timestamp() * 1000) - 1
            for ki, key in enumerate(keys, start=1):
                cursor = begin
                key_pages = 0
                while cursor <= finish:
                    payload = tb.get(
                        f"/api/plugins/telemetry/DEVICE/{device_id}/values/timeseries",
                        {"keys": key, "startTs": cursor, "endTs": finish,
                         "limit": PAGE_SIZE, "agg": "NONE", "orderBy": "ASC"},
                    )
                    points = payload.get(key, [])
                    if not points:
                        break
                    ts_list = [int(p["ts"]) for p in points]
                    if ts_list != sorted(ts_list) or ts_list[0] < cursor or ts_list[-1] > finish:
                        raise RuntimeError(f"Unexpected API timestamp ordering for {key}")
                    if len(points) >= PAGE_SIZE and ts_list.count(ts_list[-1]) > 1:
                        raise RuntimeError(f"Unsafe page boundary (same timestamp repeated): {key}")
                    db.executemany(
                        "INSERT OR REPLACE INTO readings(ts,key,value) VALUES (?,?,?)",
                        [(int(p["ts"]), key, json.dumps(p.get("value"), ensure_ascii=False)) for p in points],
                    )
                    downloaded += len(points)
                    page_count += 1
                    key_pages += 1
                    next_cursor = ts_list[-1] + 1
                    if next_cursor <= cursor:
                        raise RuntimeError(f"Pagination did not advance: {key}")
                    cursor = next_cursor
                    now = time.monotonic()
                    if now - last_print >= PROGRESS_INTERVAL_SECONDS:
                        pct = 100 * finished_jobs / total_jobs
                        print(f"  {calendar.month_name[month]} ~{pct:5.1f}% | window {win}/{total_windows} "
                              f"| key {ki}/{len(keys)}: {key[:35]} | "
                              f"pages {key_pages} | values {downloaded:,} "
                              f"| API {tb.requests_count - start_requests:,} "
                              f"| {int(now-t0)}s", flush=True)
                        last_print = now
                    if len(points) < PAGE_SIZE:
                        break
                finished_jobs += 1
                if ki % 10 == 0 or ki == len(keys):
                    db.commit()
            window_start = window_end
            db.commit()
            print(f"  {calendar.month_name[month]} ~{100 * finished_jobs / total_jobs:5.1f}% "
                  f"| window {win}/{total_windows} done | values {downloaded:,} "
                  f"| API {tb.requests_count-start_requests:,} | {int(time.monotonic()-t0)}s", flush=True)
        count = db.execute("SELECT COUNT(DISTINCT ts) FROM readings").fetchone()[0]
        if count == 0:
            return {"status": "NO DATA", "rows": 0, "values": 0, "requests": tb.requests_count-start_requests,
                    "seconds": round(time.monotonic()-t0, 1), "details": "All device keys queried; no readings"}
        rows = save_xlsx(db, keys, year, month, out_file)
        return {"status": "SUCCESS", "rows": rows, "values": downloaded,
                "requests": tb.requests_count-start_requests,
                "seconds": round(time.monotonic()-t0, 1), "details": str(out_file)}
    finally:
        db.close()
        db_path.unlink(missing_ok=True)


def save_summary(out_dir, results):
    path = out_dir / "MONTHLY_SUMMARY.csv"
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(["Month", "Status", "Timestamp rows", "Raw readings", "API requests", "Seconds", "Details"])
        for month in sorted(results):
            item = results[month]
            writer.writerow([calendar.month_name[month], item["status"], item.get("rows", 0),
                             item.get("values", 0), item.get("requests", 0), item.get("seconds", 0),
                             item.get("details", "")])


def main():
    print("=" * 62)
    print("THINGSBOARD RAW TELEMETRY EXPORT - ALL KEYS")
    print("=" * 62)
    while True:
        try:
            device_id = str(uuid.UUID(input("Device ID: ").strip()))
            break
        except ValueError:
            print("Enter the full device UUID.")
    year = ask_int("Year (YYYY): ", 2000, 2100)
    while True:
        first = ask_int("Start month (1-12): ", 1, 12)
        last = ask_int("End month (1-12): ", 1, 12)
        if first <= last:
            break
        print("End month must be greater than or equal to start month.")
    print(f"Selected period: {calendar.month_name[first]} - {calendar.month_name[last]} {year}")
    username = os.getenv("TB_USERNAME") or input("ThingsBoard email: ").strip()
    password = os.getenv("TB_PASSWORD") or getpass.getpass("ThingsBoard password: ")
    tb = ThingsBoard(TB_URL, username, password)
    try:
        all_keys = sorted(tb.get(f"/api/plugins/telemetry/DEVICE/{device_id}/keys/timeseries"))
        if not all_keys:
            raise SystemExit("No telemetry keys available or insufficient permissions.")
        keys = all_keys
        print(f"\nAutomatically exporting ALL {len(keys)} telemetry keys.")
        signature = hashlib.sha256(json.dumps(keys, ensure_ascii=False).encode("utf-8")).hexdigest()[:8]
        out_dir = Path("exports") / f"TB_{device_id[:8]}_{year}_{signature}"
        out_dir.mkdir(parents=True, exist_ok=True)
        (out_dir / "ALL_KEYS.txt").write_text("\n".join(keys) + "\n", encoding="utf-8")
        print(f"\nOutput: {out_dir.resolve()}")
        print(f"{len(keys)} total keys | {PAGE_SIZE} readings/page | "
              f"{WINDOW_DAYS}-day windows | {REQUEST_DELAY_SECONDS}s delay | 1 request at a time")
        summary_path = out_dir / "MONTHLY_SUMMARY.csv"
        results = {}
        if summary_path.exists():
            with summary_path.open(encoding="utf-8-sig", newline="") as file:
                for row in csv.DictReader(file):
                    try:
                        m = list(calendar.month_name).index(row["Month"])
                        results[m] = {"status": row["Status"], "rows": int(row["Timestamp rows"]),
                                      "values": int(row["Raw readings"]), "requests": int(row["API requests"]),
                                      "seconds": float(row["Seconds"]), "details": row["Details"]}
                    except (KeyError, ValueError):
                        continue
        total = last-first+1
        for i, month in enumerate(range(first, last+1), start=1):
            old = results.get(month)
            existing_xlsx = out_dir / f"TB_{device_id[:8]}_{year}_{month:02d}.xlsx"
            if (SKIP_COMPLETED_MONTHS and old and old["status"] == "SUCCESS" and existing_xlsx.exists()):
                print(f"\n[{i:02d}/{total:02d}] {calendar.month_name[month]}: SKIPPED (previous SUCCESS)")
                continue
            if (not RECHECK_NO_DATA_MONTHS and old and old["status"] == "NO DATA"):
                print(f"\n[{i:02d}/{total:02d}] {calendar.month_name[month]}: SKIPPED (previous NO DATA)")
                continue
            print(f"\n[{i:02d}/{total:02d}] {calendar.month_name[month]}: START", flush=True)
            try:
                result = export_month(tb, device_id, year, month, keys, out_dir, i)
            except KeyboardInterrupt:
                result = {"status": "INTERRUPTED", "rows": 0, "values": 0, "requests": 0,
                          "seconds": 0, "details": "Stopped by Ctrl+C; incomplete month not saved"}
                results[month] = result
                save_summary(out_dir, results)
                print("\nInterrupted. Previously completed months remain saved.")
                break
            except Exception as error:
                result = {"status": "FAILED", "rows": 0, "values": 0, "requests": 0,
                          "seconds": 0, "details": f"{type(error).__name__}: {error}"}
                print("  ERROR:", result["details"])
            results[month] = result
            save_summary(out_dir, results)
            print(f"[{i:02d}/{total:02d}] {calendar.month_name[month]}: "
                  f"{result['status']} | {result['rows']:,} rows | {result['values']:,} values", flush=True)
        print("\nMONTHLY STATUS:")
        for month in range(first, last+1):
            info = results.get(month, {"status": "NOT RUN", "rows": 0})
            print(f"  {calendar.month_name[month]:10} {info['status']:12} {info['rows']:>10,} rows")
        print("Summary:", summary_path.resolve())
    finally:
        tb.session.close()


if __name__ == "__main__":
    main()

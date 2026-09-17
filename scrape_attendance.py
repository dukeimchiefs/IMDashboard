#!/usr/bin/env python3
"""
Duke IM Resident Dashboard — Attendance Importer
=================================================
Downloads the email-free attendance export through Cloudflare Access and
appends any new records into the "Attendance" sheet of the shared
Point_Spreadsheet.xlsx workbook. Run refresh_data.py afterward to fold these
into the dashboard.

SETUP (one-time)
----------------
1. Install dependencies:
       pip3 install -r requirements.txt

2. Credentials are read from environment variables first. On the dashboard
   Mac, they fall back to the three macOS Keychain entries documented in the
   README. Never put real credentials in a committed file.

3. Run it whenever you want to sync attendance:
       python3 scrape_attendance.py

4. Run refresh_data.py afterward and commit the updated data.js.
"""

import json
import os
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit('openpyxl not found. Install it with:  pip3 install -r requirements.txt')

try:
    from dotenv import load_dotenv
except ImportError:
    sys.exit('python-dotenv not found. Install it with:  pip3 install -r requirements.txt')

try:
    import requests
except ImportError:
    sys.exit('requests not found. Install it with:  pip3 install -r requirements.txt')

import workbook_io

# ============================================================
# CONFIGURATION — edit these to match your setup
# ============================================================

ATTENDANCE_URL = 'https://imresidentdashboardapp.pages.dev/export'

# Same workbook refresh_data.py reads from.
EXCEL_FILE = Path('/Users/nbrazeau/Library/CloudStorage/OneDrive-SharedLibraries-DukeUniversity/Duke Chiefs 2024-2026 - Documents/zChief_Gamification/Point_Spreadsheet.xlsx')

ATTENDANCE_SHEET_NAME = 'AttendancePoints'

# The daily launchd job fires at 09:00, when the Mac may have only just woken
# and Wi-Fi/DNS may not be up yet. A single failed request used to abort the
# whole sync (sync_and_publish.sh runs under `set -e`), losing the day's
# publish over a momentary blip. Retry across roughly four minutes instead:
# delays double from 2s up to a 20s ceiling, so 15 attempts span ~3m50s.
MAX_ATTEMPTS        = 15
RETRY_BACKOFF_START = 2   # seconds to wait before the 2nd attempt
RETRY_BACKOFF_CAP   = 20  # seconds — ceiling on the doubling
REQUEST_TIMEOUT     = 30  # seconds per individual attempt

# ============================================================


def _retry_window(attempts, start, cap):
    """Total seconds spent sleeping across `attempts` tries, excluding request time."""
    total, delay = 0, start
    for _ in range(attempts - 1):
        total += delay
        delay = min(delay * 2, cap)
    return total


# Derived, not hardcoded, so the error message stays accurate if the retry
# constants above are tuned.
TOTAL_RETRY_WINDOW = _retry_window(MAX_ATTEMPTS, RETRY_BACKOFF_START, RETRY_BACKOFF_CAP)


KEYCHAIN_ACCOUNT = 'nbrazeau'
KEYCHAIN_SERVICES = {
    'ADMIN_EXPORT_KEY': 'imresidentdashboardapp-admin-export-key',
    'CF_ACCESS_CLIENT_ID': 'imresidentdashboard-access-client-id',
    'CF_ACCESS_CLIENT_SECRET': 'imresidentdashboard-access-client-secret',
}

EVENT_LABELS = {
    'noon_conference': 'Noon Conference',
    'learning_session': 'Learning Session',
    'grand_rounds': 'Grand Rounds',
    'welcome': 'Welcome',
}

# Row-count watermark. Kept off the OneDrive path for the same reason the
# workbook backups are: it has to survive the workbook being reverted.
#
# The workbook is written by this script and edited by hand in Excel on a synced
# shared library, and a stale Excel copy can land on top of a save that already
# verified clean — that is how ten attendance rows disappeared on 2026-08-20 and
# eight more the night before. The scrape itself never notices: it re-reads the
# full export every run and simply re-appends whatever is missing, so
# "appended 10 new" and "appended 2 new" read identically in the log.
#
# Counting distinct records at open time and comparing against the previous run
# catches it exactly, with none of the false positives a date heuristic would
# raise for a check-in filed after the day's last scheduled run.
STATE_DIR = Path.home() / 'Library' / 'Application Support' / 'IMResidentDashboard'
ROW_WATERMARK = STATE_DIR / 'attendance-row-watermark'

# The export, snapshotted for refresh_data.py. See write_export_cache().
ATTENDANCE_CACHE = STATE_DIR / 'attendance-export.json'

# Distinct from 1 so sync_and_publish.sh can tell "the workbook lost rows and
# this run put them back" apart from "the scrape failed and published nothing".
# They need different wording, or the alert sends you chasing the wrong thing.
EXIT_WORKBOOK_REGRESSED = 3


def read_watermark():
    """Distinct AttendancePoints records at the end of the last successful run."""
    try:
        return int(ROW_WATERMARK.read_text().strip())
    except (OSError, ValueError):
        return None


def write_watermark(count):
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    ROW_WATERMARK.write_text(str(count))


def write_export_cache(rows):
    """Snapshot the export so refresh_data.py never reads attendance back out
    of the workbook.

    The workbook sits on a OneDrive shared library that people also edit in
    Excel, and a stale Excel save can silently drop rows from it. While
    refresh_data.py sourced attendance from the AttendancePoints sheet, that
    published fewer points than the database actually held — ten rows on
    2026-08-19/20. Sourcing the dashboard from this file instead means no Excel
    save can cost anyone attendance points; the sheet stays as the
    human-readable mirror and nothing downstream depends on it.

    Written before the workbook append, so a snapshot lands even on a run where
    the guarded save is refused. A local file rather than a live fetch on
    purpose: refresh_data.py stays runnable when the export is unreachable,
    which is the property sync_and_publish.sh relies on to publish hand-entered
    points through a network blip.
    """
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    payload = {
        'fetched': datetime.now(timezone.utc).isoformat(),
        'rows': [list(row) for row in rows],
    }
    tmp = ATTENDANCE_CACHE.with_name(ATTENDANCE_CACHE.name + '.tmp')
    tmp.write_text(json.dumps(payload))
    os.replace(tmp, ATTENDANCE_CACHE)


def keychain_value(service):
    """Read a credential from macOS Keychain without echoing it."""
    if sys.platform != 'darwin':
        return None
    result = subprocess.run(
        ['security', 'find-generic-password', '-s', service, '-a', KEYCHAIN_ACCOUNT, '-w'],
        capture_output=True,
        text=True,
        check=False,
    )
    return result.stdout.strip() if result.returncode == 0 else None


def get_credentials():
    load_dotenv()
    credentials = {}
    missing = []
    for name, service in KEYCHAIN_SERVICES.items():
        value = os.environ.get(name) or keychain_value(service)
        if value:
            credentials[name] = value
        else:
            missing.append(name)
    if missing:
        sys.exit(
            '\nERROR: Missing attendance export credentials: '
            + ', '.join(missing)
            + '.\nSet them as environment variables or install the documented macOS Keychain entries.\n'
        )
    return credentials


def fetch_export(credentials):
    """GET the protected export, retrying transient failures.

    Retried: connection-level errors (DNS, refused, TLS, timeout — no response
    came back at all) and 5xx responses, both of which routinely fix themselves
    within seconds. Not retried: any other non-200. A 401/403 means the
    credentials or Cloudflare Access policy are wrong and a 404 means the URL
    is, and neither resolves by asking 15 more times.
    """
    headers = {
        'Accept': 'application/json',
        'X-Admin-Key': credentials['ADMIN_EXPORT_KEY'],
        'CF-Access-Client-Id': credentials['CF_ACCESS_CLIENT_ID'],
        'CF-Access-Client-Secret': credentials['CF_ACCESS_CLIENT_SECRET'],
    }

    delay = RETRY_BACKOFF_START
    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            response = requests.get(ATTENDANCE_URL, headers=headers, timeout=REQUEST_TIMEOUT)
        except requests.RequestException as error:
            reason = f'{type(error).__name__}'
        else:
            if response.status_code == 200:
                if attempt > 1:
                    print(f'  [export] succeeded on attempt {attempt} of {MAX_ATTEMPTS}.', flush=True)
                return response
            if response.status_code < 500:
                raise RuntimeError(f'attendance export returned HTTP {response.status_code}')
            reason = f'HTTP {response.status_code}'

        if attempt == MAX_ATTEMPTS:
            raise RuntimeError(
                f'attendance export unreachable after {MAX_ATTEMPTS} attempts '
                f'over ~{TOTAL_RETRY_WINDOW // 60}m{TOTAL_RETRY_WINDOW % 60:02d}s — last failure: {reason}'
            )

        print(
            f'  [export] attempt {attempt}/{MAX_ATTEMPTS} failed ({reason}) — retrying in {delay}s.',
            flush=True,
        )
        time.sleep(delay)
        delay = min(delay * 2, RETRY_BACKOFF_CAP)


def scrape_attendance(credentials):
    """Download the protected export and return (date, name, event) tuples."""
    response = fetch_export(credentials)

    try:
        payload = response.json()
    except requests.JSONDecodeError as error:
        raise ValueError('attendance export did not return JSON') from error

    if payload.get('ok') is not True or not isinstance(payload.get('rows'), list):
        raise ValueError('attendance export returned an unexpected response')

    rows = []
    for item in payload['rows']:
        if not isinstance(item, dict):
            raise ValueError('attendance export contained an invalid row')
        date = str(item.get('event_date', '')).strip()
        name = str(item.get('name', '')).strip()
        event_type = str(item.get('event_type', '')).strip()
        if not date or not name or event_type not in EVENT_LABELS:
            raise ValueError('attendance export contained a missing or unknown field')
        rows.append((date, name, EVENT_LABELS[event_type]))
    return rows


def append_to_workbook(path, sheet_name, scraped_rows):
    """Append new (date, name, event) rows to the Attendance sheet, skipping duplicates.

    Saving strips the cached values of the workbook's in-sheet formulas, so on
    the (common) run where the export holds nothing new, the file is left
    untouched instead of being rewritten identically.
    """
    signature = workbook_io.file_signature(path)
    wb = workbook_io.load_workbook(path)

    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        existing = set()
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row and len(row) >= 3:
                existing.add((str(row[0]).strip(), str(row[1]).strip(), str(row[2]).strip()))
    else:
        ws = wb.create_sheet(sheet_name)
        ws.append(['Date', 'Name', 'Event'])
        existing = set()

    # Captured before anything is appended: what the workbook actually held when
    # we opened it, and the only reliable way to notice it lost rows it used to
    # have (see the watermark check in __main__).
    present_before = len(existing)

    appended = 0
    skipped = 0
    for date, name, event in scraped_rows:
        key = (date, name, event)
        if key in existing:
            skipped += 1
            continue
        ws.append([date, name, event])
        existing.add(key)
        appended += 1

    if not appended:
        wb.close()
        return appended, skipped, present_before

    workbook_io.save_workbook(wb, path, expect_signature=signature)
    return appended, skipped, present_before


if __name__ == '__main__':
    if not EXCEL_FILE.exists():
        sys.exit(
            f'\nERROR: Excel file not found:\n  {EXCEL_FILE}\n\n'
            f'Edit EXCEL_FILE in scrape_attendance.py to point to your workbook.\n'
        )

    credentials = get_credentials()

    try:
        scraped_rows = scrape_attendance(credentials)
    except Exception as e:
        sys.exit(f'\nERROR downloading attendance export: {e}\n')

    if not scraped_rows:
        sys.exit('No attendance rows found on the page — nothing to append.\n')

    # Before the workbook append: the dashboard reads this, and it should be
    # current even if the guarded save below is refused.
    write_export_cache(scraped_rows)

    try:
        appended, skipped, present_before = append_to_workbook(
            EXCEL_FILE, ATTENDANCE_SHEET_NAME, scraped_rows
        )
    except workbook_io.WorkbookError as error:
        sys.exit(
            f'\nERROR: attendance not saved — {error}\n'
            f'No data was lost; the export is re-read in full on every run.\n'
        )

    print(
        f'Scraped {len(scraped_rows)} row(s) — '
        f'appended {appended} new, skipped {skipped} already present.'
    )

    # Written before the regression is reported, so a workbook that was reverted
    # once does not re-alert on every run for the rest of the day.
    watermark = read_watermark()
    write_watermark(present_before + appended)

    if watermark is not None and present_before < watermark:
        lost = watermark - present_before
        print(
            f'\nWARNING: the workbook held {present_before} attendance record(s) when this '
            f'run opened it, down from {watermark} at the end of the last run — {lost} had '
            f'been removed since.\nThis run restored them from the export, so no points are '
            f'lost, but something overwrote the workbook. Check whether it was open in '
            f'Excel, and see OneDrive version history.\n',
            file=sys.stderr,
        )
        sys.exit(EXIT_WORKBOOK_REGRESSED)

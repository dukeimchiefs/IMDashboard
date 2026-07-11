#!/usr/bin/env python3
"""
Duke IM Resident Dashboard — Attendance Scraper
=================================================
Logs into the password-protected attendance roster and appends any new
records into the "Attendance" sheet of the shared Point_Spreadsheet.xlsx
workbook. Run refresh_data.py afterward to fold these into the dashboard.

SETUP (one-time)
----------------
1. Install dependencies:
       pip3 install -r requirements.txt
       playwright install chromium

2. Copy .env.example to .env and fill in ATTENDANCE_PASSWORD.
   .env is gitignored — never commit it.

3. Run it whenever you want to sync attendance:
       python3 scrape_attendance.py

4. Run refresh_data.py afterward and commit the updated data.js.
"""

import os
import sys
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
    from playwright.sync_api import sync_playwright
except ImportError:
    sys.exit(
        'playwright not found. Install it with:\n'
        '  pip3 install -r requirements.txt\n'
        '  playwright install chromium'
    )

# ============================================================
# CONFIGURATION — edit these to match your setup
# ============================================================

ATTENDANCE_URL = 'https://imresidentdashboardapp.pages.dev/attendance'

# Same workbook refresh_data.py reads from.
EXCEL_FILE = Path('/Users/nbrazeau/Library/CloudStorage/OneDrive-SharedLibraries-DukeUniversity/Duke Chiefs 2024-2026 - Documents/zChief_Gamification/Point_Spreadsheet.xlsx')

ATTENDANCE_SHEET_NAME = 'AttendancePoints'

# ============================================================


def get_password():
    load_dotenv()
    password = os.environ.get('ATTENDANCE_PASSWORD')
    if not password:
        sys.exit(
            '\nERROR: ATTENDANCE_PASSWORD not set.\n'
            'Copy .env.example to .env and fill in the password.\n'
        )
    return password


def scrape_attendance(password):
    """Log into ATTENDANCE_URL and return a list of (date, name, event) tuples."""
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()
        page.goto(ATTENDANCE_URL)

        password_input = page.locator('input[type="password"]').first
        password_input.wait_for(timeout=15000)
        password_input.fill(password)

        submit_button = page.locator(
            'button[type="submit"], input[type="submit"], button:has-text("Unlock"), button:has-text("Submit"), button:has-text("Login"), button:has-text("Log in")'
        ).first
        if submit_button.count() > 0:
            submit_button.click()
        else:
            password_input.press('Enter')

        page.locator('table').first.wait_for(timeout=15000)

        headers = [
            h.strip().lower()
            for h in page.locator('table thead th, table tr:first-child th, table tr:first-child td').all_inner_texts()
        ]

        def col(name):
            try:
                return headers.index(name)
            except ValueError:
                raise ValueError(f'Column "{name}" not found in attendance table. Headers detected: {headers}')

        date_i = col('date')
        name_i = col('name')
        event_i = col('event')

        rows = []
        body_rows = page.locator('table tbody tr')
        if body_rows.count() == 0:
            body_rows = page.locator('table tr').filter(has_not=page.locator('th'))

        for i in range(body_rows.count()):
            cells = body_rows.nth(i).locator('td').all_inner_texts()
            if not cells or len(cells) <= max(date_i, name_i, event_i):
                continue
            rows.append((cells[date_i].strip(), cells[name_i].strip(), cells[event_i].strip()))

        browser.close()
        return rows


def append_to_workbook(path, sheet_name, scraped_rows):
    """Append new (date, name, event) rows to the Attendance sheet, skipping duplicates."""
    wb = openpyxl.load_workbook(path)

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

    wb.save(path)
    return appended, skipped


if __name__ == '__main__':
    if not EXCEL_FILE.exists():
        sys.exit(
            f'\nERROR: Excel file not found:\n  {EXCEL_FILE}\n\n'
            f'Edit EXCEL_FILE in scrape_attendance.py to point to your workbook.\n'
        )

    password = get_password()

    try:
        scraped_rows = scrape_attendance(password)
    except Exception as e:
        sys.exit(f'\nERROR scraping attendance page: {e}\n')

    if not scraped_rows:
        sys.exit('No attendance rows found on the page — nothing to append.\n')

    appended, skipped = append_to_workbook(EXCEL_FILE, ATTENDANCE_SHEET_NAME, scraped_rows)
    print(
        f'Scraped {len(scraped_rows)} row(s) — '
        f'appended {appended} new, skipped {skipped} already present.'
    )

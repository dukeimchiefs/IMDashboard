#!/usr/bin/env python3
"""
Duke IM Resident Dashboard — Weekly Data Refresh
=================================================
Reads the team points Excel workbook and regenerates data.js for the dashboard.

SETUP (one-time)
----------------
1. Install dependency:
       pip3 install openpyxl

2. Set EXCEL_FILE below to the path of your workbook.
   For a file synced via OneDrive / Box / Dropbox, use the local sync path —
   the file itself never leaves your machine.

3. Run it manually whenever the spreadsheet is updated:
       python3 refresh_data.py

4. Commit the updated data.js to push live changes to GitHub Pages.

EXCEL WORKBOOK FORMAT
---------------------
Sheet tab named "Points" (or first sheet), headers in row 1:

    Date | Team | Category | Points | Notes

    Date:     Any Excel date cell, or text as YYYY-MM-DD or M/D/YYYY
    Team:     Must match a key in TEAM_COLORS below  (e.g. "Gold", "Blue")
    Category: Must match a key in CATEGORY_COLORS below
    Points:   A number
    Notes:    Optional — ignored by the dashboard
"""

import json
import sys
from datetime import datetime
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit('openpyxl not found. Install it with:  pip3 install openpyxl')

# ============================================================
# CONFIGURATION — edit these to match your setup
# ============================================================

# Full path to the shared Excel workbook.
# Use the local OneDrive / Box / Dropbox sync path if the file is shared that way.
EXCEL_FILE = Path('/Users/nbrazeau/Library/CloudStorage/OneDrive-SharedLibraries-DukeUniversity/Duke Chiefs 2024-2026 - Documents/zChief_Gamification/Point_Spreadsheet.xlsx')


# Sheet tab name with the points log. None = use the first (active) sheet.
SHEET_NAME = None

# Team colors — keys must match the Team column values exactly (case-sensitive).
TEAM_COLORS = {
    'Bezoars':  '#012169',
    'Cultures': '#C84E00',
    'PEEPs':    '#339898',
    'Loops':    '#D97706',
    'Blasts':   '#DC2626',
    'Dispos':   '#16A34A',
}

# Category colors — keys must match Category column values exactly.
# 'All Points' is used automatically when no Category column exists.
CATEGORY_COLORS = {
    'Mentoring':      '#012169',
    'Teaching':       '#C84E00',
    'Diagnosing':     '#339898',
    'Grittiness':     '#D97706',
    'Follow-Through': '#DC2626',
    'Wellness':       '#16A34A',
    'Learning':       '#6366F1',
    'Attendance':     '#EC4899',
    'All Points':     '#012169',
}

# Academic year starts in July.
ACADEMIC_MONTHS = ['Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun']

# Output path — always the data.js sitting next to this script.
DATA_JS = Path(__file__).parent / 'data.js'

# ============================================================


def parse_date(value):
    if isinstance(value, datetime):
        return value
    if hasattr(value, 'year'):
        return datetime(value.year, value.month, value.day)
    s = str(value).strip()
    for fmt in ('%Y-%m-%d', '%m/%d/%Y', '%m/%d/%y', '%-m/%-d/%Y', '%-m/%-d/%y'):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    raise ValueError(f'Cannot parse date: {value!r}')


def read_events(path, sheet_name):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)

    if sheet_name and sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb.active

    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if len(rows) < 2:
        return []

    headers = [str(h).strip().lower() if h is not None else '' for h in rows[0]]

    def col(name, required=True):
        try:
            return headers.index(name)
        except ValueError:
            if required:
                raise ValueError(f'Column "{name}" not found. Headers detected: {list(rows[0])}')
            return None

    ci = {
        'date':     col('date'),
        'team':     col('team'),
        'category': col('category', required=False),  # optional
        'points':   col('points'),
    }

    events = []
    for row in rows[1:]:
        try:
            date = parse_date(row[ci['date']])
            team = str(row[ci['team']] or '').strip()
            cat  = str(row[ci['category']] or '').strip() if ci['category'] is not None else 'All Points'
            pts  = float(row[ci['points']] or 0)
        except (TypeError, ValueError, IndexError):
            continue
        if team and pts:
            events.append({'date': date, 'team': team, 'category': cat, 'points': pts})

    return events


def aggregate(events):
    team_monthly = {}
    cat_totals   = {}
    seen_months  = set()

    for e in events:
        month = e['date'].strftime('%b')
        seen_months.add(month)

        t = e['team']
        team_monthly.setdefault(t, {})
        team_monthly[t][month] = team_monthly[t].get(month, 0) + e['points']

        c = e['category']
        cat_totals[c] = cat_totals.get(c, 0) + e['points']

    months = [m for m in ACADEMIC_MONTHS if m in seen_months]

    teams = []
    for name, by_month in team_monthly.items():
        monthly = [int(by_month.get(m, 0)) for m in months]
        teams.append({
            'name':    'Team ' + name,
            'color':   TEAM_COLORS.get(name, '#6B7280'),
            'total':   sum(monthly),
            'monthly': monthly,
        })
    teams.sort(key=lambda t: t['total'], reverse=True)

    categories = [
        {
            'label': label,
            'value': int(v),
            'color': CATEGORY_COLORS.get(label, '#6B7280'),
        }
        for label, v in cat_totals.items()
    ]

    return {'months': months, 'teams': teams, 'categories': categories}


def write_data_js(data):
    ts   = datetime.now().strftime('%Y-%m-%d %H:%M')
    blob = json.dumps(data, indent=4)
    content = (
        f'// Auto-generated {ts} by refresh_data.py — do not edit manually.\n'
        f'// To refresh: run python3 refresh_data.py then commit data.js\n'
        f'const SAMPLE_DATA = {blob};\n\n'
        f'async function loadDashboardData() {{\n'
        f'    return SAMPLE_DATA;\n'
        f'}}\n'
    )
    DATA_JS.write_text(content, encoding='utf-8')

    n_teams  = len(data['teams'])
    n_months = len(data['months'])
    n_pts    = sum(t['total'] for t in data['teams'])
    print(f'[{ts}] data.js updated — {n_teams} teams, {n_months} months, {n_pts:,} total pts.')


def diagnose(path, sheet_name):
    """Print sheet names, row count, headers, and first 5 data rows to help debug."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    print(f'\n--- DIAGNOSTIC ---')
    print(f'Sheet tabs found:  {wb.sheetnames}')
    ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active
    print(f'Reading sheet:     "{ws.title}"')
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    print(f'Total rows read:   {len(rows)}')
    for i, row in enumerate(rows[:6], start=1):
        label = 'headers' if i == 1 else 'data   '
        print(f'  Row {i} ({label}): {list(row)}')
    print(f'------------------\n')


if __name__ == '__main__':
    if not EXCEL_FILE.exists():
        sys.exit(
            f'\nERROR: Excel file not found:\n  {EXCEL_FILE}\n\n'
            f'Edit EXCEL_FILE in refresh_data.py to point to your workbook.\n'
        )

    try:
        events = read_events(EXCEL_FILE, SHEET_NAME)
    except ValueError as e:
        diagnose(EXCEL_FILE, SHEET_NAME)
        sys.exit(f'\nERROR reading workbook: {e}\n')
    except Exception as e:
        diagnose(EXCEL_FILE, SHEET_NAME)
        sys.exit(f'\nERROR: {e}\n')

    if not events:
        diagnose(EXCEL_FILE, SHEET_NAME)
        sys.exit(
            'ERROR: No valid data rows found.\n'
            'Check the diagnostic output above — look for:\n'
            '  - Required headers: Date, Team, Points  (Category is optional)\n'
            '  - Date values parseable as dates (e.g. 7/1/2024 or 2024-07-01)\n'
            '  - Points column containing numbers\n'
            '  - Team values matching TEAM_COLORS keys in refresh_data.py\n'
        )

    data = aggregate(events)
    write_data_js(data)

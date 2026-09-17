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

    Date | Resident | Team | Category | Points | Notes

    Date:     Any Excel date cell, or text as YYYY-MM-DD or M/D/YYYY
    Resident: Optional. Full name matching a member in teams.html's ROSTER —
              if given and Team is left blank, the team is looked up
              automatically. Use this instead of an in-sheet lookup formula
              (e.g. XLOOKUP): openpyxl never evaluates formulas, and it also
              strips any cached formula value on save (write_attendance_summary
              below resaves this workbook every run), so formula-based Team
              cells silently go blank.
    Team:     Must match a key in TEAM_COLORS below (e.g. "Gold", "Blue").
              Required only if Resident is blank (e.g. team-wide bonus rows).
    Category: Must match a key in CATEGORY_COLORS below
    Points:   A number
    Notes:    Optional — ignored by the dashboard

Attendance points are folded in from the export snapshot that
scrape_attendance.py writes — NOT from the workbook's AttendancePoints sheet,
which is kept only as a human-readable mirror. See read_attendance and
EVENT_POINTS below. Run scrape_attendance.py first to refresh the snapshot,
then run this script.
"""

import hashlib
import json
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit('openpyxl not found. Install it with:  pip3 install openpyxl')

import workbook_io
# Imported rather than restated: a second copy of this path would fail silently
# by falling back to the workbook, and this repo has enough hand-mirrored pairs.
from scrape_attendance import ATTENDANCE_CACHE

# ============================================================
# CONFIGURATION — edit these to match your setup
# ============================================================

# Full path to the shared Excel workbook.
# Use the local OneDrive / Box / Dropbox sync path if the file is shared that way.
EXCEL_FILE = Path('/Users/nbrazeau/Library/CloudStorage/OneDrive-SharedLibraries-DukeUniversity/Duke Chiefs 2024-2026 - Documents/zChief_Gamification/Point_Spreadsheet.xlsx')


# Sheet tab name with the points log. None = use the first (active) sheet —
# only safe with a single-sheet workbook. Now that the workbook has multiple
# tabs (OtherPoints, Residents, Teams, Categories, AttendancePoints), hardcode
# this so it doesn't depend on whichever tab Excel last had "active" when saved.
SHEET_NAME = 'OtherPoints'

# Sheet tab name with the scraped attendance roster (written by scrape_attendance.py).
# Skipped entirely if this sheet doesn't exist yet.
ATTENDANCE_SHEET_NAME = 'AttendancePoints'

# Sheet tab (re)written each run with one row per resident's cumulative
# attendance points, normally sourced from the local attendance-export snapshot.
ATTENDANCE_SUMMARY_SHEET_NAME = 'Attendance Summary'

# Points awarded per attendance event type. Add new event types here as needed.
EVENT_POINTS = {
    'Noon Conference':  20,
    'Learning Session': 10,
    'Welcome':          20,
}

# teams.html holds the single source of truth for resident team membership
# (the ROSTER array) — parsed at runtime instead of duplicated here.
TEAMS_HTML = Path(__file__).parent / 'teams.html'

# Team colors — keys must match the Team column values exactly (case-sensitive).
TEAM_COLORS = {
    'Creatininjas':    '#2a78d6',
    'Scopetrotters':   '#1baf7a',
    'PEEPs':           '#eda100',
    'Karius':          '#008300',
    'Hemoglobbers':    '#4a3aa7',
    'Stentinels':      '#e34948',
    'Jointventurers':  '#e87ba4',
    'Codeblazers':     '#eb6834',
    'Glandiators':     '#0891b2',
    'Remissionaries':  '#9d174d',
}

# Category colors — keys must match Category column values exactly.
# 'All Points' is used automatically when no Category column exists.
#
# THE ORDER OF THIS DICT IS THE DOUGHNUT'S SLICE ORDER (see aggregate()), and it
# is load-bearing: the palette is validated on *adjacent* slice pairs, so
# reordering these keys can put two hard-to-separate hues side by side. Add or
# reorder only after re-running the validator on the new sequence.
#
# One hue family per slot (pink/green/blue/orange/cyan/violet/yellow) so slices
# stay matchable against the legend. Duke Blue (#012169) and Teal (#339898) from
# the brand palette are deliberately absent: as data colors the first falls below
# the lightness band and the second below the chroma floor (it reads gray). They
# remain the correct choices for page chrome — see CLAUDE.md.
#
# Validated light-mode, white surface: lightness band, chroma floor, CVD
# separation (worst adjacent ΔE 16.9), normal-vision floor (23.2) all pass.
# #eda100 sits at 2.11:1 against white, which is legal only because the legend
# in index.html prints a text label and value beside every swatch — do not drop
# those labels without re-checking.
CATEGORY_COLORS = {
    'Attendance':           '#EC4899',  # pink
    'Safety First':         '#008300',  # green
    'Residency Engagement': '#2563EB',  # blue
    'Teaching':             '#C84E00',  # orange
    'Caring Colleague':     '#0891b2',  # cyan
    'Got Catch ‘Em All': '#4a3aa7',  # violet — note the curly apostrophe
    'Report Rockstar':      '#eda100',  # yellow
    # Explicitly retain the neutral color previously supplied by the unknown-
    # category fallback. This removes the warning without changing the chart.
    'Community Engagement': '#6B7280',  # neutral gray
    'All Points':           '#2563EB',
}

# Academic year starts in July.
ACADEMIC_MONTHS = ['Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun']

# Academic year start date — the first day that counts toward the season.
ACADEMIC_YEAR_START = datetime(2026, 7, 1)

# Weeks run Monday through Sunday. ACADEMIC_YEAR_START is a Wednesday, and
# anchoring the buckets directly to it made every academic week run Wednesday
# to Tuesday, which split a normal work week across two buckets — Monday and
# Tuesday conference landed in the week that was already closing. Anchor on the
# Monday on or before the year start (2026-06-29) instead, so week 1 still
# contains 1 Jul and no event can fall outside weeks 1-52.
# index.html duplicates this calculation for its "Week X of 52" readout; the two
# must be changed together or the dashboard reads the wrong weekly column.
ACADEMIC_WEEK_ANCHOR = ACADEMIC_YEAR_START - timedelta(days=ACADEMIC_YEAR_START.weekday())

# Output path — always the data.js sitting next to this script.
DATA_JS = Path(__file__).parent / 'data.js'

# Page that loads data.js. GitHub Pages serves data.js with max-age=600 and the
# filename never changes, so a browser that has once cached it keeps showing
# stale standings. write_data_js stamps a content hash onto the script tag here
# (data.js?v=<hash>) so the URL changes exactly when the numbers do.
INDEX_HTML = Path(__file__).parent / 'index.html'

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


def normalize_name(name):
    """Lowercase and collapse whitespace for roster matching."""
    return re.sub(r'\s+', ' ', str(name or '')).strip().lower()


def build_name_index(roster_map):
    """normalized name -> (canonical roster name, team).

    Names in the spreadsheet are hand-typed, so 'maxwell sumner' and
    'Sean Lafata' still have to match the roster's 'Maxwell Sumner' /
    'Sean LaFata'.
    """
    return {normalize_name(n): (n, t) for n, t in roster_map.items()}


def read_events(path, sheet_name, roster_map=None):
    """Read point events from the OtherPoints sheet.

    The Team column is optional per-row: if it's blank but a Resident name is
    given, the team is looked up from roster_map (parsed from teams.html) —
    this avoids relying on an in-sheet formula (e.g. XLOOKUP), whose cached
    value openpyxl can't compute and will blank out on any resave.
    """
    roster_map = roster_map or {}
    name_index = build_name_index(roster_map)
    wb = workbook_io.load_workbook(path, read_only=True, data_only=True)

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
        'team':     col('team', required=False),      # optional if 'resident' is given
        'resident': col('resident', required=False),  # optional — used to derive team
        'category': col('category', required=False),  # optional
        'points':   col('points'),
    }

    events = []
    for row in rows[1:]:
        try:
            date     = parse_date(row[ci['date']])
            team     = str(row[ci['team']] or '').strip() if ci['team'] is not None else ''
            resident = str(row[ci['resident']] or '').strip() if ci['resident'] is not None else ''
            cat      = str(row[ci['category']] or '').strip() if ci['category'] is not None else 'All Points'
            pts      = float(row[ci['points']] or 0)
        except (TypeError, ValueError, IndexError):
            continue

        if not team and resident:
            _, team = name_index.get(normalize_name(resident), (None, ''))
            if not team and pts:
                print(f'  [points] skipping unmapped resident: {resident!r}')

        if team and pts:
            events.append({'date': date, 'team': team, 'category': cat, 'points': pts})

    return events


def load_roster_from_teams_html(path):
    """Parse the ROSTER array out of teams.html and return {full_name: team_name}."""
    text = path.read_text(encoding='utf-8')
    roster = {}
    for team_match in re.finditer(r"team:\s*'([^']+)'.*?members:\s*\[([^\]]*)\]", text, re.DOTALL):
        team_name = team_match.group(1)
        members_blob = team_match.group(2)
        for name_match in re.finditer(r"'([^']+)'", members_blob):
            roster[name_match.group(1)] = team_name
    return roster


def _attendance_rows_from_workbook(path, sheet_name):
    """(date, name, event) tuples read straight off the AttendancePoints sheet.

    Fallback only — see read_attendance for why this is no longer the source.
    """
    wb = workbook_io.load_workbook(path, read_only=True, data_only=True)

    if sheet_name not in wb.sheetnames:
        wb.close()
        return []

    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if len(rows) < 2:
        return []

    headers = [str(h).strip().lower() if h is not None else '' for h in rows[0]]

    def col(name):
        try:
            return headers.index(name)
        except ValueError:
            raise ValueError(f'Column "{name}" not found in Attendance sheet. Headers detected: {list(rows[0])}')

    date_i = col('date')
    name_i = col('name')
    event_i = col('event')

    out = []
    for row in rows[1:]:
        try:
            out.append((row[date_i], row[name_i], row[event_i]))
        except IndexError:
            continue
    return out


def _attendance_rows_from_export():
    """(date, name, event) tuples from the snapshot scrape_attendance.py writes.

    Returns None rather than [] when there is no snapshot, so the caller can
    tell "nothing has ever been fetched here" apart from "the export is empty".
    """
    try:
        payload = json.loads(ATTENDANCE_CACHE.read_text())
    except (OSError, ValueError):
        return None
    rows = payload.get('rows')
    if not isinstance(rows, list):
        return None
    return [tuple(row[:3]) for row in rows if isinstance(row, (list, tuple)) and len(row) >= 3]


def read_attendance(path, sheet_name, roster_map):
    """Read attendance and return (events, resident_totals).

      - events: list of {'date', 'team', 'category': 'Attendance', 'points'} for aggregate()
      - resident_totals: {name: cumulative_attendance_points} for every resident in
        roster_map (0 for residents with no attendance rows yet)

    Sourced from the export snapshot, NOT from the workbook. The workbook lives
    on a OneDrive shared library that people also edit in Excel, and a stale
    Excel save can silently drop rows from the AttendancePoints sheet. While
    this function read that sheet, such a save published fewer points than the
    database actually held — ten rows went that way on 2026-08-19/20, and the
    dashboard reported them as simply not having happened. The sheet is still
    written and is still the human-readable record; nothing downstream depends
    on it any more, so an Excel save can no longer cost a resident points.

    Falls back to the sheet only when no snapshot exists, so a checkout that has
    never run scrape_attendance.py still produces a dashboard.
    """
    resident_totals = {name: 0 for name in roster_map}
    name_index = build_name_index(roster_map)

    rows = _attendance_rows_from_export()
    if rows is None:
        print('  [attendance] no export snapshot yet — reading the workbook sheet instead.')
        rows = _attendance_rows_from_workbook(path, sheet_name)

    events = []
    for raw_date, raw_name, raw_event in rows:
        try:
            date = parse_date(raw_date)
            name = str(raw_name or '').strip()
            event = str(raw_event or '').strip()
        except (TypeError, ValueError):
            continue

        canonical, team = name_index.get(normalize_name(name), (None, None))
        if not team:
            print(f'  [attendance] skipping unmapped name: {name!r}')
            continue

        points = EVENT_POINTS.get(event)
        if not points:
            print(f'  [attendance] skipping unknown event type: {event!r}')
            continue

        events.append({'date': date, 'team': team, 'category': 'Attendance', 'points': points})
        resident_totals[canonical] = resident_totals.get(canonical, 0) + points

    return events, resident_totals


def write_attendance_summary(path, summary_sheet_name, resident_totals, roster_map):
    """(Re)write the per-resident cumulative attendance points sheet.

    Returns (row_count, wrote). Saving this workbook strips the cached values
    of its in-sheet formulas, so when the summary is already correct we leave
    the file alone entirely rather than rewrite it identically.
    """
    signature = workbook_io.file_signature(path)
    wb = workbook_io.load_workbook(path)

    rows = [['Name', 'Team', 'Attendance Points']]
    for name in sorted(roster_map, key=lambda n: (roster_map[n], n)):
        rows.append([name, roster_map[name], resident_totals.get(name, 0)])

    if summary_sheet_name in wb.sheetnames:
        current = [
            list(row) for row in wb[summary_sheet_name].iter_rows(values_only=True)
            if any(cell is not None for cell in row)
        ]
        if current == rows:
            wb.close()
            return len(roster_map), False
        del wb[summary_sheet_name]

    ws = wb.create_sheet(summary_sheet_name, 0)
    for row in rows:
        ws.append(row)

    # The summary is rebuilt from scratch each run, so it is the one sheet
    # allowed to come out shorter than it went in.
    workbook_io.save_workbook(
        wb, path,
        expect_signature=signature,
        allow_shrink=(summary_sheet_name,),
    )
    return len(roster_map), True


def week_number(date):
    """1-52 Monday-based week index, matching the frontend's calc."""
    delta_days = (date - ACADEMIC_WEEK_ANCHOR).days
    return min(52, max(1, delta_days // 7 + 1))


def week_month_label(week_num):
    """Calendar month abbreviation that a given academic week falls in.

    Week 1 starts on 29 Jun, so label it by the year start rather than by its
    own Monday — otherwise the chart's first x-axis tick reads "Jun" for a
    season that begins in July.
    """
    week_start = ACADEMIC_WEEK_ANCHOR + timedelta(weeks=week_num - 1)
    return max(week_start, ACADEMIC_YEAR_START).strftime('%b')


def aggregate(events, today=None):
    """Roll events up into the shape data.js expects.

    Each team also gets `rank` (current standing) and `prevRank` (standing as of
    the end of yesterday, computed from events dated strictly before today) so
    the dashboard can show a movement arrow. `prevRank` is None when there are
    no events before today at all — i.e. the season just started and there is no
    prior standing to compare against.
    """
    today = (today or datetime.now()).replace(hour=0, minute=0, second=0, microsecond=0)

    team_monthly = {}
    team_weekly  = {}
    cat_totals   = {}
    seen_months  = set()
    seen_weeks   = set()

    for e in events:
        month = e['date'].strftime('%b')
        seen_months.add(month)

        wk = week_number(e['date'])
        seen_weeks.add(wk)

        t = e['team']
        team_monthly.setdefault(t, {})
        team_monthly[t][month] = team_monthly[t].get(month, 0) + e['points']

        team_weekly.setdefault(t, {})
        team_weekly[t][wk] = team_weekly[t].get(wk, 0) + e['points']

        c = e['category']
        cat_totals[c] = cat_totals.get(c, 0) + e['points']

    months = [m for m in ACADEMIC_MONTHS if m in seen_months]
    weeks  = sorted(seen_weeks)
    week_months = [week_month_label(w) for w in weeks]

    teams = []
    for name in set(team_monthly) | set(team_weekly):
        monthly = [int(team_monthly.get(name, {}).get(m, 0)) for m in months]
        weekly  = [int(team_weekly.get(name, {}).get(w, 0)) for w in weeks]
        teams.append({
            'name':    'Team ' + name,
            'color':   TEAM_COLORS.get(name, '#6B7280'),
            'total':   sum(monthly),
            'monthly': monthly,
            'weekly':  weekly,
        })
    # Tie-break by name so equal totals rank in a stable order run-to-run —
    # otherwise tied teams could swap places on a rerun and fake a rank change.
    teams.sort(key=lambda t: (-t['total'], t['name']))
    for i, t in enumerate(teams):
        t['rank'] = i + 1

    prior_totals = {}
    for e in events:
        if e['date'] < today:
            key = 'Team ' + e['team']
            prior_totals[key] = prior_totals.get(key, 0) + e['points']

    if prior_totals:
        prior_order = sorted(
            (t['name'] for t in teams),
            key=lambda n: (-prior_totals.get(n, 0), n),
        )
        prior_rank = {name: i + 1 for i, name in enumerate(prior_order)}
        for t in teams:
            t['prevRank'] = prior_rank[t['name']]
    else:
        for t in teams:
            t['prevRank'] = None

    # Emit in CATEGORY_COLORS order, not the order rows happen to appear in the
    # sheet. The doughnut draws slices in this order, and the palette is only
    # validated for the hues that end up adjacent — sheet order would shuffle
    # that on any edit. Unknown categories land at the end in fallback gray.
    known = [c for c in CATEGORY_COLORS if c in cat_totals]
    unknown = sorted(c for c in cat_totals if c not in CATEGORY_COLORS)
    for label in unknown:
        print(f'  [categories] no color for {label!r} — falling back to gray. '
              f'Add it to CATEGORY_COLORS in refresh_data.py.')

    categories = [
        {
            'label': label,
            'value': int(cat_totals[label]),
            'color': CATEGORY_COLORS.get(label, '#6B7280'),
        }
        for label in known + unknown
    ]

    return {
        'months': months,
        'weeks': weeks,
        'weekMonths': week_months,
        'teams': teams,
        'categories': categories,
    }


def stamp_index_html(version):
    """Point index.html's data.js script tag at ?v=<version>.

    Returns True if the file changed. The tag is matched with or without an
    existing ?v= so this is idempotent across runs.
    """
    if not INDEX_HTML.exists():
        print(f'  [cache] WARNING: {INDEX_HTML.name} not found — script tag not stamped.')
        return False

    text = INDEX_HTML.read_text(encoding='utf-8')
    stamped, n = re.subn(
        r'src="data\.js(?:\?v=[^"]*)?"',
        f'src="data.js?v={version}"',
        text,
    )
    if not n:
        print(f'  [cache] WARNING: no data.js script tag found in {INDEX_HTML.name}.')
        return False
    if stamped == text:
        return False

    INDEX_HTML.write_text(stamped, encoding='utf-8')
    return True


def write_data_js(data):
    """Write data.js and stamp index.html, but only when the numbers changed.

    The header carries a generation timestamp, so rewriting unconditionally
    made data.js differ on every run even when no points moved — which meant
    sync_and_publish.sh's "nothing to commit" check could never fire and the
    history filled with no-op commits. The payload version below is recorded
    in the header so an unchanged run can be detected and skipped entirely.
    """
    blob = json.dumps(data, indent=4)

    # Hash the payload, not the finished file — the file embeds a timestamp that
    # changes every run, and a version that churns on unchanged data would bust
    # every visitor's cache daily for nothing.
    version = hashlib.sha256(blob.encode('utf-8')).hexdigest()[:8]

    n_teams  = len(data['teams'])
    n_months = len(data['months'])
    n_pts    = sum(t['total'] for t in data['teams'])

    current = DATA_JS.read_text(encoding='utf-8') if DATA_JS.exists() else ''
    if f'// payload v={version}\n' in current:
        print(f'[data.js] already current — {n_teams} teams, {n_months} months, {n_pts:,} total pts.')
    else:
        ts = datetime.now().strftime('%Y-%m-%d %H:%M')
        DATA_JS.write_text(
            f'// Auto-generated {ts} by refresh_data.py — do not edit manually.\n'
            f'// To refresh: run python3 refresh_data.py then commit data.js\n'
            f'// payload v={version}\n'
            f'const SAMPLE_DATA = {blob};\n\n'
            f'async function loadDashboardData() {{\n'
            f'    return SAMPLE_DATA;\n'
            f'}}\n',
            encoding='utf-8',
        )
        print(f'[{ts}] data.js updated — {n_teams} teams, {n_months} months, {n_pts:,} total pts.')

    # Always reconciled, even when data.js was left alone: the stamp could be
    # missing or stale from a hand-edit or a half-finished earlier run.
    restamped = stamp_index_html(version)
    action = 'restamped' if restamped else 'already current —'
    print(f'[cache] {action} index.html script tag at data.js?v={version}')


def diagnose(path, sheet_name):
    """Print sheet names, row count, headers, and first 5 data rows to help debug."""
    wb = workbook_io.load_workbook(path, read_only=True, data_only=True)
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

    roster_map = load_roster_from_teams_html(TEAMS_HTML) if TEAMS_HTML.exists() else {}

    try:
        events = read_events(EXCEL_FILE, SHEET_NAME, roster_map)
    except workbook_io.WorkbookError as e:
        sys.exit(f'\nERROR: {e}\n')
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
            '  - Required headers: Date, Points  (Category optional; either Team or a\n'
            '    Resident name matching teams.html is required)\n'
            '  - Date values parseable as dates (e.g. 7/1/2024 or 2024-07-01)\n'
            '  - Points column containing numbers\n'
            '  - Team values matching TEAM_COLORS keys in refresh_data.py\n'
        )

    if TEAMS_HTML.exists():
        attendance_events, resident_totals = read_attendance(EXCEL_FILE, ATTENDANCE_SHEET_NAME, roster_map)
        if attendance_events:
            print(f'[attendance] merged {len(attendance_events)} attendance event(s).')
        events += attendance_events

        try:
            n_residents, wrote = write_attendance_summary(
                EXCEL_FILE, ATTENDANCE_SUMMARY_SHEET_NAME, resident_totals, roster_map)
        except workbook_io.WorkbookError as error:
            # data.js is still regenerated below — only the workbook write is
            # skipped, and the dashboard does not depend on the summary sheet.
            print(f'[attendance] WARNING: {error}')
        else:
            action = 'wrote' if wrote else 'already current —'
            print(f'[attendance] {action} "{ATTENDANCE_SUMMARY_SHEET_NAME}" sheet — {n_residents} resident row(s).')

    data = aggregate(events)
    write_data_js(data)

#!/usr/bin/env python3
"""
Duke IM Resident Dashboard — Texted Points Importer
====================================================
Downloads the ledger of points that chiefs and attendings award by text
message (the House Point Hourglass, a separate Cloudflare Worker) and
snapshots it for refresh_data.py. Also mirrors it into two sheets of the
shared Point_Spreadsheet.xlsx so the chiefs can read it where they already
look:

    TextPoints         one row per award — who, how many, from whom, and why
    TextPoints Inbox   texts that did NOT end in an award and need a look

Like attendance, the dashboard is published from the snapshot and never from
those sheets; see write_export_cache in scrape_attendance.py for why.

SETUP (one-time)
----------------
This script does nothing until both of these exist, as environment variables
or as macOS Keychain entries (account `nbrazeau`):

    HOURGLASS_EXPORT_URL   housepointhourglass-export-url   the Worker's /export URL
    HOURGLASS_EXPORT_KEY   housepointhourglass-export-key   its X-Admin-Key

    security add-generic-password -s housepointhourglass-export-url -a "$USER" -w
    security add-generic-password -s housepointhourglass-export-key -a "$USER" -w

Run it whenever you want to pull texted points:
       python3 scrape_text_points.py
then run refresh_data.py and commit the updated data.js — or just run
./sync_and_publish.sh, which does all of it.
"""

import json
import os
import sys
import time
from datetime import datetime, timezone

try:
    from dotenv import load_dotenv
except ImportError:
    sys.exit('python-dotenv not found. Install it with:  pip3 install -r requirements.txt')

try:
    import requests
except ImportError:
    sys.exit('requests not found. Install it with:  pip3 install -r requirements.txt')

try:
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
except ImportError:
    sys.exit('openpyxl not found. Install it with:  pip3 install -r requirements.txt')

import workbook_io
# Imported rather than restated, for the reason refresh_data.py gives: a second
# copy of the workbook path or the state directory would drift silently.
from scrape_attendance import EXCEL_FILE, STATE_DIR, keychain_value

# ============================================================
# CONFIGURATION — edit these to match your setup
# ============================================================

KEYCHAIN_SERVICES = {
    'HOURGLASS_EXPORT_URL': 'housepointhourglass-export-url',
    'HOURGLASS_EXPORT_KEY': 'housepointhourglass-export-key',
}

TEXT_POINTS_SHEET_NAME = 'TextPoints'
TEXT_INBOX_SHEET_NAME = 'TextPoints Inbox'

# The export, snapshotted for refresh_data.py — the texted-points counterpart
# of ATTENDANCE_CACHE, kept beside it and off the OneDrive path for the same
# reason.
TEXT_POINTS_CACHE = STATE_DIR / 'text-points-export.json'

# Far shorter than the attendance scrape's window on purpose. That scrape runs
# first and has already waited out a Mac that woke with no network; by the time
# this one starts, a failure is the Worker's, and four tries is enough to ride
# out a blip without holding up the publish.
MAX_ATTEMPTS        = 4
RETRY_BACKOFF_START = 2   # seconds, doubling
REQUEST_TIMEOUT     = 30  # seconds per individual attempt

LEDGER_HEADER = ['ID', 'Date', 'Resident', 'Category', 'Points', 'Awarded By', 'Note', 'Status']
INBOX_HEADER  = ['Received', 'From', 'Text', 'What happened', 'Details']

# The Worker's outcome codes, in words a chief can act on. An outcome missing
# from here is shown as its raw code rather than hidden.
OUTCOME_WORDS = {
    'not_found':       'No resident matched the name — nothing recorded',
    'rejected_sender': 'Sender is not on the list — a chief can text: ADD <number> <name>',
    'asked':           'The name was unclear (several residents, or only a near miss) — nothing recorded; see Details',
    'partial':         'Some awards recorded; the sentences listed in Details were skipped',
    'over_limit':      'Over the point or daily limit — nothing recorded',
    'not_understood':  'Could not be read as an award',
    'refused':         'Refused (a deduction, a zero, or a chief-only command)',
    'error':           'The roster was unavailable — nothing recorded',
    'processing':      'Failed partway through — nothing recorded',
    'ambiguous':       'Too many residents matched — nothing recorded',
}

# ============================================================


def get_credentials():
    """The export URL and key, or None when texted points are simply not set up.

    Neither present is the normal state on a machine that has never been
    connected to the Hourglass, and is not an error. One without the other is
    a half-finished setup and is reported, because skipping it quietly would
    look exactly like "nobody has texted any points".
    """
    load_dotenv()
    credentials = {}
    for name, service in KEYCHAIN_SERVICES.items():
        value = os.environ.get(name) or keychain_value(service)
        if value:
            credentials[name] = value.strip()
    if not credentials:
        return None
    missing = [name for name in KEYCHAIN_SERVICES if name not in credentials]
    if missing:
        sys.exit(
            '\nERROR: texted points are half configured — missing '
            + ', '.join(missing)
            + '.\nSet both as environment variables or install both Keychain entries (see the README).\n'
        )
    return credentials


def fetch_export(credentials):
    """GET the export, retrying connection failures and 5xx responses.

    Same split as scrape_attendance.fetch_export: a 401 or 404 means the key or
    the URL is wrong, and asking again will not fix either.
    """
    headers = {
        'Accept': 'application/json',
        'X-Admin-Key': credentials['HOURGLASS_EXPORT_KEY'],
    }

    delay = RETRY_BACKOFF_START
    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            response = requests.get(credentials['HOURGLASS_EXPORT_URL'], headers=headers, timeout=REQUEST_TIMEOUT)
        except requests.RequestException as error:
            reason = f'{type(error).__name__}'
        else:
            if response.status_code == 200:
                return response
            if response.status_code < 500:
                raise RuntimeError(f'texted-points export returned HTTP {response.status_code}')
            reason = f'HTTP {response.status_code}'

        if attempt == MAX_ATTEMPTS:
            raise RuntimeError(
                f'texted-points export unreachable after {MAX_ATTEMPTS} attempts — last failure: {reason}'
            )

        print(
            f'  [text-points] attempt {attempt}/{MAX_ATTEMPTS} failed ({reason}) — retrying in {delay}s.',
            flush=True,
        )
        time.sleep(delay)
        delay *= 2


def parse_export(payload):
    """Validate the export and return (rows, inbox).

    Strict on purpose. A row this script cannot read is points it would
    otherwise drop without a word, so anything unexpected fails the whole run
    and leaves the previous snapshot in place.
    """
    if not isinstance(payload, dict) or payload.get('ok') is not True or not isinstance(payload.get('rows'), list):
        raise ValueError('texted-points export returned an unexpected response')

    rows = []
    for item in payload['rows']:
        if not isinstance(item, dict):
            raise ValueError('texted-points export contained an invalid row')
        try:
            row = {
                'id':         int(item['id']),
                'award_date': str(item['award_date']).strip(),
                'resident':   str(item['resident']).strip(),
                'category':   str(item['category']).strip(),
                'points':     int(item['points']),
                'note':       str(item.get('note') or '').strip(),
                'awarded_by': str(item.get('awarded_by') or '').strip(),
                'voided':     bool(item.get('voided')),
            }
            datetime.strptime(row['award_date'], '%Y-%m-%d')
        except (KeyError, TypeError, ValueError) as error:
            raise ValueError('texted-points export contained a missing or malformed field') from error
        if not row['resident'] or not row['category'] or row['points'] <= 0:
            raise ValueError('texted-points export contained a missing or malformed field')
        rows.append(row)

    if len({row['id'] for row in rows}) != len(rows):
        raise ValueError('texted-points export contained a duplicate award id')

    inbox = []
    for item in payload.get('needs_attention') or []:
        if not isinstance(item, dict):
            continue
        inbox.append({
            'received_at': str(item.get('received_at') or ''),
            'from':        str(item.get('from') or ''),
            'body':        str(item.get('body') or ''),
            'outcome':     str(item.get('outcome') or ''),
            'detail':      str(item.get('detail') or ''),
        })

    return rows, inbox


def cached_row_count():
    """How many awards the last snapshot held, or None if there isn't one."""
    try:
        rows = json.loads(TEXT_POINTS_CACHE.read_text()).get('rows')
    except (OSError, ValueError, AttributeError):
        return None
    return len(rows) if isinstance(rows, list) else None


def write_export_cache(rows):
    """Snapshot the ledger for refresh_data.py. See scrape_attendance.write_export_cache."""
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    payload = {
        'fetched': datetime.now(timezone.utc).isoformat(),
        'rows': rows,
    }
    tmp = TEXT_POINTS_CACHE.with_name(TEXT_POINTS_CACHE.name + '.tmp')
    tmp.write_text(json.dumps(payload))
    os.replace(tmp, TEXT_POINTS_CACHE)


def local_time(iso):
    """'2026-10-02T01:03:21.034Z' -> '2026-10-01 21:03' in this Mac's time zone."""
    try:
        return datetime.fromisoformat(iso.replace('Z', '+00:00')).astimezone().strftime('%Y-%m-%d %H:%M')
    except ValueError:
        return iso


def text_cell(value):
    """A string that is safe to put in a cell exactly as written.

    These values are whatever somebody typed into a text message, and two
    things a person can type are trouble in a cell. openpyxl stores any string
    beginning with '=' as a formula, so one that starts that way is quoted and
    stays text. And it refuses outright to write a control character — the
    append raises, and since the offending text is re-sent in every export
    until it ages out, one stray ESC from a stranger's phone would stop these
    sheets updating for two weeks. The Hourglass strips those on the way in;
    this is the same check on the way out, for anything already stored.
    """
    value = ILLEGAL_CHARACTERS_RE.sub('', value)
    return f"'{value}" if value.startswith('=') else value


def inbox_rows(inbox):
    return [INBOX_HEADER] + [
        [
            local_time(item['received_at']),
            text_cell(item['from']),
            text_cell(item['body']),
            OUTCOME_WORDS.get(item['outcome'], item['outcome']),
            text_cell(item['detail']),
        ]
        for item in inbox
    ]


def mirror_to_workbook(path, rows, inbox):
    """Bring the two TextPoints sheets up to date. Returns (appended, restatused, inbox_rewritten).

    The ledger sheet is append-only, keyed on the award ID: a new award gets a
    new row, and one that was voided has VOIDED written into the Status cell of
    the row it already has. Nothing is ever deleted from it — openpyxl row
    deletion has already cost this workbook a sheet once.

    The inbox is a rolling 14-day view, so it is rebuilt whole, the same way
    the Attendance Summary is, and is the one sheet here allowed to shrink.

    Saving strips the cached values of the workbook's in-sheet formulas, so
    when nothing has changed the file is left untouched.
    """
    signature = workbook_io.file_signature(path)
    wb = workbook_io.load_workbook(path)
    changed = False

    if TEXT_POINTS_SHEET_NAME in wb.sheetnames:
        ws = wb[TEXT_POINTS_SHEET_NAME]
    else:
        ws = wb.create_sheet(TEXT_POINTS_SHEET_NAME)
        ws.append(LEDGER_HEADER)
        changed = True

    status_column = LEDGER_HEADER.index('Status') + 1
    row_of = {}
    for cells in ws.iter_rows(min_row=2, max_col=1):
        try:
            row_of[int(cells[0].value)] = cells[0].row
        except (TypeError, ValueError):
            continue

    appended = 0
    restatused = 0
    for row in rows:
        status = 'VOIDED' if row['voided'] else None
        if row['id'] not in row_of:
            ws.append([
                row['id'], row['award_date'], row['resident'], row['category'],
                row['points'], text_cell(row['awarded_by']), text_cell(row['note']) or None, status,
            ])
            appended += 1
            continue
        cell = ws.cell(row=row_of[row['id']], column=status_column)
        if (cell.value or None) != status:
            cell.value = status
            restatused += 1

    target = inbox_rows(inbox)
    if TEXT_INBOX_SHEET_NAME in wb.sheetnames:
        current = [
            [cell if cell is not None else '' for cell in row]
            for row in wb[TEXT_INBOX_SHEET_NAME].iter_rows(values_only=True)
            if any(cell is not None for cell in row)
        ]
    else:
        # No sheet and nothing to put in one: don't create it just to hold a header.
        current = None if inbox else target

    inbox_rewritten = current != target
    if inbox_rewritten:
        if TEXT_INBOX_SHEET_NAME in wb.sheetnames:
            del wb[TEXT_INBOX_SHEET_NAME]
        inbox_ws = wb.create_sheet(TEXT_INBOX_SHEET_NAME)
        for row in target:
            inbox_ws.append(row)

    if not (changed or appended or restatused or inbox_rewritten):
        wb.close()
        return 0, 0, False

    workbook_io.save_workbook(
        wb, path,
        expect_signature=signature,
        allow_shrink=(TEXT_INBOX_SHEET_NAME,),
    )
    return appended, restatused, inbox_rewritten


if __name__ == '__main__':
    credentials = get_credentials()
    if credentials is None:
        print('[text-points] not configured on this machine — skipping. (See scrape_text_points.py to turn it on.)')
        sys.exit(0)

    try:
        rows, inbox = parse_export(fetch_export(credentials).json())
    except Exception as e:
        sys.exit(f'\nERROR downloading texted-points export: {e}\n')

    # The ledger only grows — an award that is taken back is flagged, not
    # removed — so fewer rows than last time means the wrong database or a
    # wiped one, not a quiet week. Publishing from it would silently take points
    # off the board, so keep the last snapshot and fail where it will be seen.
    previous = cached_row_count()
    if previous is not None and len(rows) < previous:
        sys.exit(
            f'\nERROR: the texted-points export holds {len(rows)} award(s), down from {previous} '
            f'at the last sync. The ledger never shrinks, so this is not a real change.\n'
            f'Kept the previous snapshot; nothing was published from the smaller one.\n'
        )

    # Before the workbook mirror: the dashboard reads this, and it should be
    # current even if the guarded save below is refused.
    write_export_cache(rows)

    voided = sum(1 for row in rows if row['voided'])
    print(f'[text-points] {len(rows)} texted award(s) in the ledger — {len(rows) - voided} live, {voided} voided.')

    if inbox:
        print(f'[text-points] {len(inbox)} text(s) in the last two weeks did not end in an award:')
        for item in inbox:
            print(f'    {local_time(item["received_at"])}  {item["from"]}: {item["body"]!r} — '
                  f'{OUTCOME_WORDS.get(item["outcome"], item["outcome"])}')

    if not EXCEL_FILE.exists():
        print('[text-points] WARNING: sheets not updated —' + workbook_io.missing_workbook_message(EXCEL_FILE).rstrip())
        sys.exit(0)

    try:
        appended, restatused, inbox_rewritten = mirror_to_workbook(EXCEL_FILE, rows, inbox)
    except workbook_io.WorkbookError as error:
        # Not a failure of the sync. The snapshot above is what gets published;
        # the sheets are a courtesy copy and catch up on the next run.
        print(f'[text-points] WARNING: workbook sheets not updated this run — {error}')
    except Exception as error:
        # Anything else is a bug in the mirror rather than a busy workbook, and
        # would recur on every run, so it exits nonzero and gets reported. The
        # publish is unaffected either way: the snapshot is already written.
        sys.exit(
            f'\nERROR: texted points were fetched and will be published, but the workbook '
            f'sheets could not be updated — {type(error).__name__}: {error}\n'
        )
    else:
        print(
            f'[text-points] "{TEXT_POINTS_SHEET_NAME}" sheet — appended {appended} new, '
            f'marked {restatused} voided; "{TEXT_INBOX_SHEET_NAME}" '
            f'{"rewritten" if inbox_rewritten else "already current"}.'
        )

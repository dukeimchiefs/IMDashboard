#!/usr/bin/env python3
"""
Duke IM Resident Dashboard — Guarded Workbook Writes
====================================================
Both scrape_attendance.py and refresh_data.py save the shared
Point_Spreadsheet.xlsx, which lives on a OneDrive shared-library sync path.
A plain openpyxl `wb.save(path)` there is riskier than it looks:

  - openpyxl truncates and rewrites the file in place, so a crash or a
    interrupted write leaves a corrupt workbook with no way back.
  - If someone has the workbook open in Excel, or OneDrive is mid-sync, the
    cloud copy can win the resulting conflict and silently revert the save.
    That is how the August attendance rows went missing.
  - Every resave strips the cached values of in-sheet formulas (the XLOOKUP
    in the OtherPoints Team column), so needless writes actively degrade the
    workbook.

Every save goes through save_workbook() below, which skips writes that would
change nothing, refuses to write while Excel holds the file or when the file
changed underneath us, keeps a rolling local backup, swaps the new file in
atomically, and restores the backup if a sheet lost rows.

Backups live outside the synced folder and never leave the machine.
"""

import os
import shutil
import stat
import tempfile
from zipfile import BadZipFile
from datetime import datetime
from pathlib import Path

import openpyxl

# Backups deliberately live off the OneDrive sync path — the workbook must
# never end up anywhere it could be published.
BACKUP_DIR = Path.home() / 'Library' / 'Application Support' / 'IMResidentDashboard' / 'workbook-backups'

KEEP_BACKUPS = 10


class WorkbookError(RuntimeError):
    """The workbook could not be safely read or written."""


class WorkbookUnavailable(WorkbookError):
    """The workbook exists but its contents are not currently available."""


class WorkbookInvalid(WorkbookError):
    """The workbook was downloaded but is not a valid XLSX archive."""


class WorkbookBusy(WorkbookError):
    """The workbook is open in Excel — writing now risks losing the save."""


class WorkbookConflict(WorkbookError):
    """The file changed on disk since we read it — refuse to clobber it."""


class WorkbookVerificationError(WorkbookError):
    """The workbook lost rows during the save; the backup was restored."""


def lock_file(path):
    """Path of the '~$' owner file Excel creates while a workbook is open."""
    return Path(path).parent / f'~${Path(path).name}'


def is_open_in_excel(path):
    return lock_file(path).exists()


def file_signature(path):
    """(mtime_ns, size) — captured before a read, rechecked before the write."""
    st = Path(path).stat()
    return (st.st_mtime_ns, st.st_size)


def is_dataless(path):
    """Whether macOS File Provider has evicted the file's local contents."""
    flags = getattr(Path(path).stat(), 'st_flags', 0)
    # Python 3.13 exposes SF_DATALESS on macOS. Keep the documented Darwin
    # value as a fallback so this check still works with older interpreters.
    return bool(flags & getattr(stat, 'SF_DATALESS', 0x40000000))


def materialize(path):
    """Force File Provider to fetch an evicted file's contents before a read.

    Being dataless is not a failure. macOS downloads a placeholder on demand
    the moment something reads it, so the flag on its own only means "nobody
    has touched this lately" — the provider reclaims space from idle files by
    design. Reading one byte triggers that download and blocks until it lands
    (a few seconds for this workbook), which is exactly what the caller wants.
    Only a genuine fetch failure — offline, signed out, EDEADLK — raises.
    """
    with open(path, 'rb') as handle:
        handle.read(1)


def load_workbook(path, **kwargs):
    """Open an XLSX with useful errors for OneDrive/File Provider failures.

    An online-only File Provider placeholder has a normal path and advertised
    size, so exists()/stat() both succeed. Reading it can instead raise EDEADLK;
    zipfile then often replaces that useful error with the misleading
    ``BadZipFile: File is not a zip file``. Materialize the placeholder before
    the read and normalize both failure modes for callers.

    This used to refuse the read outright whenever is_dataless() was true, which
    turned an ordinary eviction into a failed run: on 2026-09-17 both scripts
    aborted on a workbook that a plain open() then fetched in under four
    seconds. Attempt the download and report unavailability only when the
    attempt actually fails.
    """
    path = Path(path)
    if is_dataless(path):
        try:
            materialize(path)
        except OSError as error:
            raise WorkbookUnavailable(
                f'{path.name} is online-only and OneDrive could not download it '
                f'({error}). In Finder, right-click it (or its folder), choose '
                '"Always Keep on This Device", and wait for the solid green '
                'checkmark.'
            ) from error

    try:
        return openpyxl.load_workbook(path, **kwargs)
    except BadZipFile as error:
        underlying = error.__context__
        if is_dataless(path) or isinstance(underlying, OSError):
            raise WorkbookUnavailable(
                f'{path.name} became unavailable while it was being read. OneDrive '
                'may still be downloading it; wait for the solid green checkmark '
                'and re-run.'
            ) from error
        raise WorkbookInvalid(
            f'{path.name} is downloaded but is not a valid Excel workbook. '
            'Check OneDrive version history or a local workbook backup.'
        ) from error
    except OSError as error:
        raise WorkbookUnavailable(
            f'{path.name} could not be read ({error}). OneDrive may still be '
            'downloading or syncing it; wait for the solid green checkmark and re-run.'
        ) from error


def backup_workbook(path, keep=KEEP_BACKUPS):
    """Copy the workbook into BACKUP_DIR, pruning all but the newest `keep`."""
    path = Path(path)
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)

    stamp = datetime.now().strftime('%Y%m%d-%H%M%S')
    target = BACKUP_DIR / f'{path.stem}.{stamp}{path.suffix}'
    shutil.copy2(path, target)

    existing = sorted(BACKUP_DIR.glob(f'{path.stem}.*{path.suffix}'))
    for stale in existing[:-keep]:
        stale.unlink()

    return target


def sheet_row_counts(path):
    """{sheet name: count of non-empty rows} — used to prove nothing was lost.

    Read with data_only=False so formula cells count by their formula rather
    than their cached value. The cache is present right after Excel saves and
    absent right after openpyxl does, so counting cached values would make the
    700 XLOOKUP rows in OtherPoints appear and vanish on their own and trip the
    row-loss check on a perfectly good save.
    """
    wb = load_workbook(path, read_only=True, data_only=False)
    try:
        counts = {}
        for name in wb.sheetnames:
            counts[name] = sum(
                1 for row in wb[name].iter_rows(values_only=True)
                if any(cell is not None for cell in row)
            )
        return counts
    finally:
        wb.close()


def save_workbook(wb, path, expect_signature=None, allow_shrink=(), keep_backups=KEEP_BACKUPS):
    """Back up, atomically replace, then verify no sheet lost rows.

    expect_signature: the file_signature() taken before the workbook was read.
                      If the file has changed since, raise WorkbookConflict
                      rather than overwrite someone else's edit.
    allow_shrink:     sheet names that are legitimately rewritten and may end
                      up with fewer rows than before.
    """
    path = Path(path)

    if is_open_in_excel(path):
        raise WorkbookBusy(
            f'{path.name} is currently open in Excel ({lock_file(path).name} present). '
            'Close it and re-run — writing now risks the save being reverted.'
        )

    if expect_signature is not None and file_signature(path) != expect_signature:
        raise WorkbookConflict(
            f'{path.name} changed on disk while this script was running (another '
            'edit, or a OneDrive sync). Nothing was written — re-run to pick up '
            'the current version.'
        )

    before = sheet_row_counts(path)
    backup = backup_workbook(path, keep=keep_backups)

    # Write to a temp file on the same filesystem, then swap it in with a single
    # atomic rename, so the workbook is never left half-written.
    fd, tmp_name = tempfile.mkstemp(dir=str(path.parent), prefix=f'.{path.stem}.', suffix=path.suffix)
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        wb.save(tmp)
        # copymode, NOT copystat. copystat also carries the *old* file's mtime
        # onto the replacement, so every guarded save left the workbook looking
        # untouched since whenever a human last opened it in Excel — the backup
        # set shows six consecutive script saves all stamped 2026-08-18 14:11:02.
        # OneDrive resolves conflicts against that timestamp, which is how a save
        # that verified clean locally kept losing to a stale cloud copy: ten
        # attendance rows went that way on 2026-08-20. Only the permission bits
        # need to survive the atomic swap; the new mtime is the point.
        shutil.copymode(path, tmp)
        os.replace(tmp, path)
    except BaseException:
        tmp.unlink(missing_ok=True)
        raise

    after = sheet_row_counts(path)
    shrunk = {
        name: (before[name], after.get(name, 0))
        for name in before
        if name not in allow_shrink and after.get(name, 0) < before[name]
    }
    if shrunk:
        shutil.copy2(backup, path)
        detail = ', '.join(f'{n}: {b} -> {a}' for n, (b, a) in shrunk.items())
        raise WorkbookVerificationError(
            f'Save lost rows ({detail}). Restored the backup from {backup}.'
        )

    return backup

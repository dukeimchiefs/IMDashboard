#!/usr/bin/env bash
# Duke IM Resident Dashboard — Full Attendance Sync + Publish
# =============================================================
# 1. Scrapes the attendance roster into the AttendancePoints sheet.
# 2. Regenerates data.js (team totals + per-resident Attendance Summary sheet).
# 3. Commits data.js if it changed.
# 4. Pushes to GitHub, which republishes the live dashboard via GitHub Pages.
#
# Step 1 is the only step that needs the network. If it fails, steps 2-4 still
# run so that points already in the workbook get published, and the script then
# exits 1 so the failure is still visible in launchd's last exit code.
#
# Usage:
#   ./sync_and_publish.sh
#
# Requires: .env with ATTENDANCE_PASSWORD set (see .env.example), and
# dependencies installed (pip3 install -r requirements.txt; playwright install chromium).

set -euo pipefail
cd "$(dirname "$0")"

echo "===================================================="
echo "==> Run started $(date +'%Y-%m-%d %H:%M:%S %Z')"

# launchd runs with a minimal PATH and resolves a different python3 than an
# interactive shell does — /opt/homebrew/bin/python3, which has none of the
# dependencies. Pin the interpreter that actually has them. Override with
# PYTHON=/path/to/python3 ./sync_and_publish.sh when running by hand.
PYTHON="${PYTHON:-/opt/homebrew/Caskroom/miniconda/base/bin/python3}"
if [ ! -x "$PYTHON" ]; then
    PYTHON="$(command -v python3 || true)"
fi
if [ -z "$PYTHON" ]; then
    echo "ERROR: no python3 interpreter found." >&2
    exit 1
fi

# Fail loudly and early rather than midway through a partial sync.
if ! "$PYTHON" -c 'import openpyxl, dotenv, requests' 2>/dev/null; then
    echo "ERROR: $PYTHON is missing dependencies. Install them with:" >&2
    echo "    $PYTHON -m pip install -r requirements.txt" >&2
    exit 1
fi
echo "==> Using $PYTHON"

# Step 1 talks to the network; step 2 reads the local workbook and does not
# consume step 1's output. Under `set -e` a failed scrape used to abort the
# whole run, so a momentary network blip also blocked hand-entered points from
# ever reaching the site. Record the failure and carry on instead — nothing is
# lost, because scrape_attendance.py re-reads the full export every run and
# skips duplicates, so a missed day self-heals tomorrow.
scrape_failed=0

echo "==> [1/4] Scraping attendance..."
if ! "$PYTHON" scrape_attendance.py; then
    echo "WARNING: attendance sync failed — publishing workbook data anyway." >&2
    scrape_failed=1
fi

echo "==> [2/4] Refreshing dashboard data..."
"$PYTHON" refresh_data.py

echo "==> [3/4] Committing data.js..."
# index.html carries the data.js?v=<hash> cache-busting stamp that refresh_data.py
# rewrites, so it has to ship in the same commit or browsers keep the stale file.
git add data.js index.html
if git diff --cached --quiet; then
    echo "    No changes to data.js — nothing to commit."
else
    git commit -m "Automated attendance sync $(date +'%Y-%m-%d %H:%M')"
fi

echo "==> [4/4] Pushing to GitHub..."
git push

# Publishing anyway must not make a broken scrape look like a clean run: exit
# nonzero so `launchctl print ... | grep 'last exit code'` still surfaces it.
if [ "$scrape_failed" -eq 1 ]; then
    echo "==> Done — BUT the attendance sync failed; published from the workbook only." >&2
    exit 1
fi

echo "==> Done."

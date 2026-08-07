#!/usr/bin/env bash
# Duke IM Resident Dashboard — Full Attendance Sync + Publish
# =============================================================
# 1. Scrapes the attendance roster into the AttendancePoints sheet.
# 2. Regenerates data.js (team totals + per-resident Attendance Summary sheet).
# 3. Commits data.js if it changed.
# 4. Pushes to GitHub, which republishes the live dashboard via GitHub Pages.
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

echo "==> [1/4] Scraping attendance..."
"$PYTHON" scrape_attendance.py

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

echo "==> Done."

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

echo "==> [1/4] Scraping attendance..."
python3 scrape_attendance.py

echo "==> [2/4] Refreshing dashboard data..."
python3 refresh_data.py

echo "==> [3/4] Committing data.js..."
git add data.js
if git diff --cached --quiet; then
    echo "    No changes to data.js — nothing to commit."
else
    git commit -m "Automated attendance sync $(date +'%Y-%m-%d %H:%M')"
fi

echo "==> [4/4] Pushing to GitHub..."
git push

echo "==> Done."

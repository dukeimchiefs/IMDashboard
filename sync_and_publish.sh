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
# Run by launchd five times a day (09:00, 12:00, 15:00, 18:00, 21:00 local) via
# ~/Library/LaunchAgents/com.nbrazeau.imresidentdashboard.sync-and-publish.plist.
# Every step is idempotent — the scrape skips rows already present, and a run
# with nothing to publish makes no commit — so the extra runs cost nothing and
# bound how long a check-in, a missed schedule, or a workbook that was open in
# Excel can hold the site stale.
#
# Usage:
#   ./sync_and_publish.sh
#
# Requires: .env with ATTENDANCE_PASSWORD set (see .env.example), and
# dependencies installed (pip3 install -r requirements.txt; playwright install chromium).

set -euo pipefail
cd "$(dirname "$0")"

# --- Failure alerting -------------------------------------------------------
# A nonzero exit is recorded by launchd and read by nobody: the 2026-08-06 run
# died on a missing dependency and sat unnoticed for a week while the schedule
# kept firing. Every failing exit path now routes through report_exit, including
# the `set -e` aborts that previously left nothing behind but a truncated log.
#
# The Notification Center banner always fires. Email is additionally sent if —
# and only if — both Keychain entries below exist, so this stays inert until
# they are installed:
#   security add-generic-password -s imresidentdashboard-resend-key  -a "$USER" -w
#   security add-generic-password -s imresidentdashboard-alert-email -a "$USER" -w
KEYCHAIN_ACCOUNT="${KEYCHAIN_ACCOUNT:-nbrazeau}"
ALERT_FROM='noreply@nicholasbrazeau.com'   # the domain already verified in Resend
LOG_DIR="$HOME/Library/Logs/IMResidentDashboard"
# Five runs a day means a sustained outage would send five alerts a day, which
# trains you to ignore them. The banner is free and fires every time; email is
# rate-limited to one per window so it keeps meaning something.
ALERT_EMAIL_THROTTLE_SECONDS=43200         # 12 hours
ALERT_STAMP="$HOME/Library/Application Support/IMResidentDashboard/last-alert-email"

failed_line=''
failure_reason=''
trap 'failed_line=$LINENO' ERR

keychain_value() {
    security find-generic-password -s "$1" -a "$KEYCHAIN_ACCOUNT" -w 2>/dev/null || true
}

send_alert_email() {
    local subject="$1" body="$2"
    local key recipient now last

    key="$(keychain_value imresidentdashboard-resend-key)"
    recipient="$(keychain_value imresidentdashboard-alert-email)"
    if [ -z "$key" ] || [ -z "$recipient" ]; then
        return 0
    fi

    now="$(date +%s)"
    last=0
    if [ -f "$ALERT_STAMP" ]; then
        last="$(cat "$ALERT_STAMP" 2>/dev/null || echo 0)"
    fi
    if [ $((now - last)) -lt "$ALERT_EMAIL_THROTTLE_SECONDS" ]; then
        echo "    (alert email suppressed — one was already sent within the throttle window)" >&2
        return 0
    fi

    # subject and body are built below from our own text and never contain a
    # double quote or a newline, so interpolating them into the JSON directly is
    # safe and saves depending on jq or on a Python that may be what just failed.
    # --fail matters as much as the request does: without it curl exits 0 on an
    # HTTP 401 or 422, so a revoked key would leave this reporting success
    # forever — exactly when the alert is the only thing that would tell us.
    if ! curl -sS --fail --max-time 30 -X POST 'https://api.resend.com/emails' \
        -H "Authorization: Bearer $key" \
        -H 'Content-Type: application/json' \
        -d "{\"from\":\"$ALERT_FROM\",\"to\":[\"$recipient\"],\"subject\":\"$subject\",\"text\":\"$body\"}" \
        >/dev/null 2>&1; then
        echo "    (alert email rejected or failed to send)" >&2
        return 0
    fi

    mkdir -p "$(dirname "$ALERT_STAMP")"
    printf '%s' "$now" > "$ALERT_STAMP"
}

report_exit() {
    local status=$?
    if [ "$status" -eq 0 ]; then
        return 0
    fi

    local detail
    if [ -n "$failure_reason" ]; then
        detail="$failure_reason"
    elif [ -n "$failed_line" ]; then
        detail="failed at line $failed_line"
    else
        detail='failed'
    fi

    local when
    when="$(date +'%Y-%m-%d %H:%M:%S %Z')"
    echo "==> FAILED (exit $status) — $detail. $when" >&2

    osascript -e "display notification \"$detail (exit $status)\" with title \"IM Dashboard sync failed\"" \
        >/dev/null 2>&1 || true
    send_alert_email \
        "IM Dashboard sync failed (exit $status)" \
        "sync_and_publish.sh on $(hostname -s) exited $status at $when — $detail. Logs: $LOG_DIR"
}
trap report_exit EXIT
# ---------------------------------------------------------------------------

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
    failure_reason='no python3 interpreter found'
    exit 1
fi

# Fail loudly and early rather than midway through a partial sync.
if ! "$PYTHON" -c 'import openpyxl, dotenv, requests' 2>/dev/null; then
    echo "ERROR: $PYTHON is missing dependencies. Install them with:" >&2
    echo "    $PYTHON -m pip install -r requirements.txt" >&2
    # This is what broke on 2026-08-06. An explicit exit trips neither `set -e`
    # nor the ERR trap, so the alert has no line number to report unless the
    # reason is named here.
    failure_reason="$PYTHON is missing dependencies (openpyxl/dotenv/requests)"
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
    # A failed scrape is reached through `if !`, which trips neither `set -e` nor
    # the ERR trap, so name it here or the alert would report a bare exit 1.
    failure_reason='attendance scrape failed; published from the workbook only'
    exit 1
fi

echo "==> Done."

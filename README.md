# IM Resident Dashboard

A static site tracking resident team points, published via GitHub Pages.
`index.html` (Statistics), `teams.html` (Team Assignments), and `rules.html`
(Rules) all read from `data.js`, a plain JS file regenerated from a private
Excel workbook — the workbook itself never leaves your machine or gets
committed.

## Data flow

```
Point_Spreadsheet.xlsx (local, OneDrive-synced, gitignored)
  ├─ OtherPoints        one row per manual point event (Date/Resident/Team/Category/Points/Notes)
  ├─ AttendancePoints   one row per attendance record (Date/Name/Event) — see below
  ├─ Residents          resident directory (name/team/contact info)
  ├─ Teams              team name list
  └─ Categories         point-category list
        │
        │  refresh_data.py
        ▼
  ├─ Attendance Summary   (re)written each run — one row per resident,
  │                       cumulative attendance points
  │
  ▼
data.js  →  committed & pushed  →  GitHub Pages serves the live dashboard
```

Resident → team lookups are parsed at runtime straight out of `teams.html`'s
`ROSTER` array, so team membership has a single source of truth.

## Scripts

| Script | What it does |
| --- | --- |
| `scrape_attendance.py` | Logs into the password-protected attendance roster at `imresidentdashboardapp.pages.dev/attendance`, and appends any new (Date, Name, Event) rows into the `AttendancePoints` sheet — skipping rows already recorded. |
| `refresh_data.py` | Reads `OtherPoints` + `AttendancePoints`, converts attendance events to points by event type (`Noon Conference` = 20, `Learning Session` = 10), aggregates everything into team/category totals, regenerates `data.js`, and rewrites the `Attendance Summary` sheet. |
| `sync_and_publish.sh` | Runs both of the above in order, commits `data.js` if it changed, and pushes to GitHub — the one command to run for a full attendance sync + live publish. |

## One-time setup

```bash
pip3 install -r requirements.txt
playwright install chromium
cp .env.example .env   # then fill in ATTENDANCE_PASSWORD yourself — never commit .env
```

## Day-to-day use

```bash
./sync_and_publish.sh
```

Or run the pieces individually — `python3 scrape_attendance.py` to sync
attendance only, `python3 refresh_data.py` to just regenerate `data.js` from
whatever's currently in the workbook.

## Privacy

`Point_Spreadsheet.xlsx` and `.env` are both gitignored — the underlying
points data and credentials never get committed or published. Only the
derived, aggregated `data.js` becomes public.

# IM Resident Dashboard

A static site tracking resident team points, published via GitHub Pages.
`index.html` (Statistics), `teams.html` (Team Assignments), and `rules.html`
(Rules) all read from `data.js`, a plain JS file regenerated from a private
Excel workbook — the workbook itself never leaves your machine or gets
committed.

## Data flow

```
protected attendance export
        │  scrape_attendance.py
        ├─→ AttendancePoints (human-readable mirror in the workbook)
        └─→ attendance-export.json (local, outside OneDrive and git) ─┐
                                                                        │
points awarded by text message (House Point Hourglass export)           │
        │  scrape_text_points.py                                        │
        ├─→ TextPoints / TextPoints Inbox (mirrors in the workbook)     │
        └─→ text-points-export.json (local, outside OneDrive and git) ──┤
                                                                        │
Point_Spreadsheet.xlsx (local, OneDrive-synced, gitignored)              │
  ├─ OtherPoints (manual point events) ───────────────────────────────────┤
  ├─ Residents / Teams / Categories                                    │
  └─ Attendance Summary ←────────────────────────────────────┐    │
                                                      │    │
                                              refresh_data.py ─┘
                                                      │
                                                      └─→ data.js
                                                           │
                                                           └─→ GitHub Pages
```

Resident → team lookups are parsed at runtime straight out of `teams.html`'s
`ROSTER` array, so team membership has a single source of truth.

## Scripts

| Script | What it does |
| --- | --- |
| `scrape_attendance.py` | Downloads the email-free `/export` feed through Cloudflare Access Service Auth, and appends any new (Date, Name, Event) rows into the `AttendancePoints` sheet — skipping rows already recorded. |
| `scrape_text_points.py` | Downloads the ledger of points that chiefs and attendings award by text message (the House Point Hourglass, a separate Worker), snapshots it for `refresh_data.py`, and mirrors it into the workbook: `TextPoints` holds one row per award, and `TextPoints Inbox` lists recent texts that did **not** end in an award — a misspelled name, a sender nobody has added — so a chief can follow up. Does nothing, successfully, until its two credentials are installed. |
| `refresh_data.py` | Reads `OtherPoints` plus the local attendance-export snapshot, converts attendance events to points by event type (`Noon Conference` = 20, `Learning Session` = 10), aggregates everything into team/category totals, regenerates `data.js`, and rewrites the `Attendance Summary` sheet. It falls back to `AttendancePoints` only before the first snapshot has been created. Texted points are folded in from their snapshot the same way, each with the category and value it was awarded with. |
| `sync_and_publish.sh` | Runs all of the above in order, commits `data.js` if it changed, and pushes to GitHub — the one command to run for a full attendance sync + live publish. |

## One-time setup

```bash
pip3 install -r requirements.txt
```

The sync requires `ADMIN_EXPORT_KEY`, `CF_ACCESS_CLIENT_ID`, and
`CF_ACCESS_CLIENT_SECRET`. Environment variables take precedence. On the
dashboard Mac, the script automatically reads the credentials from these
macOS Keychain services (account `nbrazeau`):

- `imresidentdashboardapp-admin-export-key`
- `imresidentdashboard-access-client-id`
- `imresidentdashboard-access-client-secret`

For another machine, copy `.env.example` to `.env` and supply the three values
locally. `.env` is gitignored and must never be committed.

Texted points are optional and off until both of these exist, as environment
variables (`HOURGLASS_EXPORT_URL`, `HOURGLASS_EXPORT_KEY`) or as Keychain
services under the same account:

- `housepointhourglass-export-url` — the Hourglass Worker's `/export` URL
- `housepointhourglass-export-key` — the key it expects as `X-Admin-Key`

With neither installed the sync skips that step and says so. With only one, it
fails, so a half-finished setup cannot pass for "nobody texted any points".

## Day-to-day use

```bash
./sync_and_publish.sh
```

Or run the pieces individually — `python3 scrape_attendance.py` to sync
attendance only, `python3 refresh_data.py` to just regenerate `data.js` from
whatever's currently in the workbook.

## Privacy

`Point_Spreadsheet.xlsx` and `.env` are both gitignored — the underlying
points data and credentials never get committed or published. The attendance
export contains names, event types, and dates, but no email addresses. The
texted-points export contains resident names, who awarded the points, and the
reason they gave; it and the snapshot made from it stay on this machine. Only
the derived, aggregated `data.js` becomes public.

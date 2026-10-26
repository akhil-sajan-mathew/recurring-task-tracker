# Setup Guide

How to get the code in `apps-script/` running in a Google Sheet and prove it works.
Allow about 20 minutes.

## What's in the folder

| File | What it does |
|---|---|
| `apps-script/Logic.gs` | Pure date, recurrence and SLA rules (working days, working hours, due/alert/escalation times) |
| `apps-script/Config.gs` | Sheet names, column headers, config defaults, the simulated clock |
| `apps-script/Data.gs` | Sheet read/write, the script lock, audit logging |
| `apps-script/Engine.gs` | The 15-minute engine: create cycles → statuses → reminders → L1 → L2 → weekly digest |
| `apps-script/Complete.gs` | Employee completion (Ticket ID / daily update) and manager actions, plus the onEdit handler |
| `apps-script/Views.gs` | Builds the "My Tasks – …" tabs and the Manager_Dashboard from the ledger |
| `apps-script/Setup.gs` | Builds the workbook, demo data, triggers, protections, and the **Allo Ops** menu |
| `apps-script/Tests.gs` | 43 tests (42 automatic + 1 manual) covering eight scenarios; results go to `Test_Results` |
| `apps-script/appsscript.json` | Manifest: time zone Asia/Kolkata, V8 runtime, permissions |
| `local-test/` | Runs the same test suite on your PC against a fake Google Sheets. **Don't upload it.** |

## 1. Create the sheet and paste the code

1. Use a **separate Google account** with no personal data in it. A fresh Gmail account is fine.
2. Create a new Google Sheet and name it something like **Allo – Recurring Task Tracker**.
3. Open **Extensions → Apps Script**.
4. In **Project Settings** (the gear icon):
   - Tick **"Show 'appsscript.json' manifest file in editor"**.
   - Set the time zone to **(GMT+05:30) India Standard Time**.
5. Back in the editor, open `appsscript.json` and replace its contents with `apps-script/appsscript.json`.
6. Rename the default `Code.gs` to `Logic.gs` and paste `Logic.gs` into it.
7. Then add each remaining file with **+ → Script**, named exactly like the file without `.gs`: `Config`, `Data`, `Engine`, `Complete`, `Views`, `Setup`, `Tests`. Paste in each file's contents.
8. Click **Save** (the disk icon).

> Alternative: if you use [clasp](https://github.com/google/clasp), run `clasp create --type sheets --rootDir apps-script` and then `clasp push`.

## 2. Build the workbook

1. Reload the Google Sheet tab. An **Allo Ops** menu appears after a few seconds.
2. Click **Allo Ops → 1. Set up workbook**.
   - Google asks for permission. Because this is your own unpublished script, you'll see "Google hasn't verified this app". Click **Advanced → Go to (project name) (unsafe) → Allow**.
   - The script needs permission to edit this spreadsheet, send email as you, and create its own triggers.
3. Click **Allo Ops → 2. Load demo data (this week)**.
   - This creates Employees A–D, tasks T01–T06, the four **My Tasks** tabs and the **Manager_Dashboard**.
   - Every email address is a **+alias of your own address** (`you+empa@gmail.com`, `you+manager@…`, `you+leadership@…`), so every email lands in your own inbox.

The engine now runs by itself every 15 minutes (**Apps Script → Triggers** shows `scheduledRun` and `onEditInstalled`).

## 3. Run the automated tests

1. Click **Allo Ops → Run all tests**. If it times out, which can happen on a slow day, use **Run tests 1–4 only** and then **Run tests 5–8 only**.
2. Open **Test_Results**. You should see 42 PASS and 1 MANUAL (2.6, protection, which needs a second Google account to show).

The tests use the simulated clock and dry-run email, so nothing real is sent. When they finish, the workbook is reset to fresh demo data for the current week.

On your PC you can run the same suite without Google at all:
```bash
node local-test/run.js
```

## 4. Demo walkthrough

Start with **Allo Ops → 2b. Load demo on the test calendar**. The schedule then starts Mon 2 Nov 2026 and the clock is set to Mon 10:00.
Use **Allo Ops → Set test clock…** for each step; every clock change also runs the engine. Emails go to your own inbox (+aliases), so keep Gmail open in another tab.

| # | Clock | Do this | Point out |
|---|---|---|---|
| 1 | Mon 10:00 | Show **Task_Master**, **Config**, and the four **My Tasks** tabs | One row per recurring task. Every SLA value is a setting. Each employee sees only their own tasks. |
| 2 | `2026-11-02 10:15` | On **My Tasks – Employee A**: type an update in ✎, click **Share to WhatsApp**, then tick **Posted** | The text is the proof; the row turns grey with Completed At. |
| 3 | `2026-11-02 15:00` | On **My Tasks – Employee B**: type `PTA-10234` on Portal Task A. Then type `abc` on Portal Task A2. | One entry completes the task. `abc` is rejected with a cell note, and `Audit_Log` records both. (Your inbox also gets the morning "due today" reminders for B, C and D.) |
| 4 | `2026-11-02 17:31` | Open your inbox | One same-day email to `+manager` listing Employee C (Portal Task B) and Employee D (Portal Task C). |
| 5 | (same) | On **Manager_Dashboard**: type a Note on C's row, then choose **Reminded employee**. Leave D's row alone. | Acknowledged gets stamped. Choosing the action *is* the acknowledgement. |
| 6 | `2026-11-03 12:35` | Open your inbox | A leadership email: D's task, "Manager did not acknowledge" (4 working hours after 17:31 = Tue 12:31). |
| 7 | `2026-11-03 17:01` | Open your inbox | C's task, "Acknowledged but unresolved". The resolution deadline (1 working day) has passed. The same digest includes A's Tuesday daily update, which nobody completed or acted on. |
| 8 | (same) | **Run engine now** three times, then show `Notification_Log` and `Cycles` | No new rows and no new emails. |
| 9 | `2026-11-07 10:00` | Open **My Tasks – Employee B** | The next Portal Task A cycle (due Mon 9 Nov) appeared by itself. |
| 10 | — | Show **Config** and walk through the SLA table in the design note | The timing logic and why it's set that way. |
| 11 | — | Show **Test_Results** | All tests green. |

Finish with **Allo Ops → Use real time (clear test clock)**.

## 5. Share a copy

1. Use **File → Make a copy**. The copy includes the script.
2. In the copy, run **Allo Ops → 1. Set up workbook** and then **2. Load demo data**, because triggers aren't copied.
3. Share the copy as **Editor** only with people you trust: editors can open and change the bound script, which runs with your permissions.

## Things to know

- **Who completed it.** On personal Gmail accounts, Google often doesn't tell the script which editor made an edit. The script therefore records the tab's owner (each tab can be protected to one employee) and adds the editor's email whenever Google provides it.
- **Email quota.** A personal Gmail account can send about 100 emails a day from scripts. The engine sends digests (one email per manager per run), so the demo uses only a handful.
- **Protection** only does something when employees are real, separate Google accounts. With +aliases everything is you, the owner, so you can edit everything.
- **If something breaks,** `Run_Log` shows the error, and Google emails you when a scheduled run fails.

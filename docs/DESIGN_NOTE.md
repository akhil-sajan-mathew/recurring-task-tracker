# Recurring Task Tracking, Pending Reports and Escalation: Design Note

---

## 1. What I built

A Google Sheets system, automated with Apps Script, that:

- creates every recurring task occurrence on schedule;
- shows each employee only their own pending work, in one tab where they also record completion;
- alerts the manager the same day a task misses its time;
- records the manager's response; and
- escalates to leadership when the miss isn't handled in time.

The automation can run any number of times without creating duplicate tasks, completions or emails.

**Design rule: one ledger, many views.** Every task occurrence (a "cycle") is one row in a protected `Cycles` sheet. Its ID is fixed by the task and its due date (`T02|2026-11-09`), so creating the same cycle twice is impossible.

The employee tabs and the Manager Dashboard are generated from that ledger. Whatever an employee or manager types on a view is validated, written to the ledger, and the view is rebuilt. That's how "update once, from one place" is guaranteed rather than hoped for.

## 2. Sheets

| Sheet | Who uses it | Purpose |
|---|---|---|
| **My Tasks – Employee X** | Employee | Pending and overdue tasks first, then the last 5 completed. The only place an employee works. |
| **Manager_Dashboard** | Manager | Per-employee summary (open, overdue, on-time %, awaiting action), plus an action list for overdue items |
| Task_Master | Admin | One row per recurring task: assignee, frequency, cycle days, start date, due time, lead days, proof type, ticket format |
| Employees / Holidays / Config | Admin | People and their managers, non-working days, every SLA value |
| Cycles | Script only | The ledger, with one row per occurrence: due, status, proof, who, when, alert, acknowledgement and escalation stamps |
| Notification_Log / Audit_Log / Run_Log | Read-only | Every email (and why), every edit including rejected ones, every engine run |
| Test_Results | Read-only | Output of the built-in automated test suite |

## 3. Recurrence

- **Fixed schedule.** Cycle *k* is due on `Start Date + k × Cycle Days`. A daily task occurs every working day.
- **Late completion doesn't shift the schedule.** If the weekly portal task is due Mondays and is done on a Wednesday, the next one is still due next Monday. Recurring portal work is usually a compliance cadence, and a drifting schedule would quietly erode it.
- **Lead time.** Each task appears on the employee's tab a set number of working days before it's due (7-day tasks 1 day early, 15-day tasks 2, 30-day tasks 3). Employees never calculate a date.
- **Unfinished work never disappears.** When a new cycle opens before the old one is done, both show and the old one keeps escalating. Silently replacing a missed cycle is the worst failure mode, because the miss vanishes.
- **Non-working days.** A due date on a Sunday or holiday moves to the next working day. "Every 30 days" means 30 calendar days, not the same date each month.

## 4. Completion: single point, with proof

**Portal tasks.** The employee types the Ticket ID in their tab, and that one action completes the cycle. The system checks that:
- the ID matches the task's format (for example `PTA-` followed by 4–8 digits);
- the ID hasn't already closed another cycle;
- the cycle is open.

It records who completed it, when, and the ticket. Changing a ticket later counts as a logged correction, not a second completion. Clearing a completed task is blocked.

**Daily task: "write it here, share from here."** WhatsApp group messages can't be read by a free API, so the system can't watch the group. Instead the sheet becomes where the update starts:

1. The employee types today's update in their tab.
2. A **Share to WhatsApp** link next to it opens WhatsApp with the text already filled in. They pick the ops group and send.
3. They tick **Posted**. The text and time are stored as the proof.

Why this design:
- **Minimal effort.** The employee types the update once, where they already work. There's no screenshot, form or second place to update.
- **Readable proof.** The manager can read every update in the sheet, and the history is searchable.
- **Honest limit.** The tick is self-reported. Guards: a minimum length, a block on re-posting yesterday's text, and manager spot checks against the group.
- **Upgrade path.** The WhatsApp Business API could post the update directly from the sheet. The system would then know the message was sent, and the tick would go away.

## 5. Alert, acknowledgement and escalation SLA

Working hours are Mon–Sat, 10:00–19:00 IST. SLA clocks pause outside them. All values are settings in `Config`.

| Stage | Daily Operations Update | Portal tasks (7/15/30 days) |
|---|---|---|
| Visible to employee | 10:00 that day | 1/2/3 working days before due |
| Reminder to employee | 10:30 on the due day | 10:30 on the due day |
| **Due** | **11:00** | **17:00** |
| **Manager alert (L1, same day)** | 11:30 | 17:30 |
| Manager must choose an action | within 4 working hours (≈15:30) | within 4 working hours (≈12:30 next working day) |
| Task must be done | 19:00 same day | 17:00 next working day |
| **Leadership escalation (L2)** | Manager took no action in time, or chose "Escalate now" | Manager took no action in time, the task is still not done by its deadline, or "Escalate now" |
| Pattern | 3+ late or missed daily updates in a week → flagged in the Monday leadership digest | — |

**Why these numbers**
- **30 minutes of grace before alerting** filters out "entered it a few minutes late". Managers who get noise learn to ignore alerts.
- **The alert is always the same working day.** It's capped at 19:00 even when the task is due late in the day.
- **Four *working* hours to act.** A 17:30 alert shouldn't escalate overnight; the clock picks up at 10:00 the next working day.
- **Two different escalation reasons:**
  - *"Manager did not acknowledge"* is a supervision gap.
  - *"Acknowledged but unresolved"* is an execution gap.
  - Leadership needs to know which one it is, so the email says so.
- **Daily misses go to a weekly pattern review instead of an immediate leadership email,** once the manager is handling them. One late daily update doesn't need leadership; three in a week does.
- **No repeated emails:**
  - Each cycle gets at most one L1 and one L2, enforced by a log.
  - Each recipient gets one digest per run, never one email per task.
  - Nothing is sent on non-working days (pending emails go out the next working day).
  - Leadership hears once when an escalated item is finally closed.

**Manager acknowledgement.** The manager's expected response to an alert is to contact the employee and then choose one Action on the dashboard:

| Action | Effect |
|---|---|
| Reminded employee | Records the acknowledgement; the deadlines still apply. |
| Extend deadline | Up to 3 working days. Moves the "must be done" deadline. |
| Waive | A written reason is required. Closes the cycle. |
| Escalate to leadership now | Sends the escalation on the next run. |

Choosing an action **is** the acknowledgement. A bare "seen" tick would record that the manager looked without recording what they decided.

The system can always tell the four outcomes apart:

| Acknowledged? | Done? | Outcome |
|---|---|---|
| Yes | Yes | Closed, no escalation |
| Yes | No (deadline passed) | Escalated: "Acknowledged but unresolved" |
| No | Yes (before the action deadline) | Closed, no escalation |
| No | No (action deadline passed) | Escalated: "Manager did not acknowledge" |

## 6. Reliability: why running it twice is safe

| Risk | Guard |
|---|---|
| A cycle created twice | Fixed cycle ID, checked before every insert |
| A completion recorded twice | An entry on a completed cycle is either no change or a logged correction |
| An email sent twice | Every alert has a log key (`L1\|cycle`, `L2\|cycle`), written *before* sending; failed sends are retried |
| Two runs at once (timer + an edit) | A script lock serialises all writes |
| People editing the wrong thing | Ledger, config and logs are protected; each employee tab's input column is limited to that employee |
| The automation silently stopping | The dashboard shows "last engine run x min ago" in red after an hour; Google emails the owner when a run fails |
| Formula injection from typed text | User text starting with `=` `+` `-` `@` is stored as plain text |

## 7. Testing

- **The simulated clock.** A `TEST_NOW` setting makes the whole system believe it's any date and time, so a week of behavior replays in minutes.
- **The automated suite.** **Allo Ops → Run all tests** runs 43 tests covering eight core scenarios (recurrence, single-point completion, daily completion, same-day alert, acknowledgement, escalation, SLA timing, no duplicates), plus edge cases: holidays, stacked cycles, wrong and duplicate tickets, copy-pasted daily updates, waivers without a reason, extensions beyond the limit, weekend SLA arithmetic, and re-running setup. It writes PASS/FAIL to `Test_Results`; the result is 42 PASS and 1 MANUAL.
- **Mutation checks.** I also ran the suite against deliberately broken versions of the code: removing the duplicate check, removing the grace period, letting the schedule drift, allowing duplicate tickets. In each case the relevant tests failed, so they catch real errors rather than passing trivially.

## 8. Deliberately *not* automated

- **Whether a ticket exists on the portal.** There's no portal API. The format and uniqueness checks catch typos and reuse; spot checks catch the rest.
- **Waivers and extensions.** These need human judgment, so they're recorded with a reason, never automatic.
- **Changes to Task_Master.** Admin only. Changing a schedule is a business decision, not a workflow step.
- **AI.** ChatGPT isn't in the decision path. Due dates, statuses and escalations must be deterministic and auditable. Good uses would be a weekly plain-English summary of patterns for leadership, or drafting the email wording.

## 9. Known limits and next steps

- **Identity on personal Gmail.** Google often doesn't reveal which editor made an edit. The tab's owner (protected per employee) is the identity. On Google Workspace the editor's email is reliable and is recorded automatically.
- **Email quota.** About 100 recipients a day on personal Gmail. Digests keep usage to a handful a day.
- **Next steps:**
  - a WhatsApp Business API integration for the daily update;
  - a one-click "Acknowledge" link in the alert email (an Apps Script web app);
  - an optional completion-anchored schedule for tasks that shouldn't stack;
  - a Google Chat/Slack alert channel.

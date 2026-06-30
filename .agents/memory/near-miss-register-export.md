---
name: Near Miss Register export
description: How the Near Miss Register export works and why it differs from the Actions report.
---

The Near Miss Register exports EVERY near-miss card, not just the ones that became
tracked actions (the Actions report only shows actioned items — that was the user's
original complaint).

**Where:** `server/near-miss-register-export.ts` (HTML/CSV/Markdown/Word) built on the
same Paged.js A4 foundation as the meeting minutes & Actions report, but with an
ORANGE (#ea580c) brand theme. Endpoints: `/api/generate-near-miss-register-{html,
csv,markdown,word}` in server/routes.ts.

**Data flow:** client (actions.tsx) posts `nearMissItems = meetingItems.filter(
type === 'Near Miss')` — i.e. ALL near-miss cards. Server's `enrichNearMissForRegister`
loads each card's full investigation via `storage.getNearMissInvestigation(id)` and
JSON-parses `hazards` / `resultingActions` (array-or-JSON-string safe parse). Output
has an overview table + per-item detail blocks (event details, risk assessment,
hazards table, resulting actions, dual sign-off).

**Why a separate engine from actions-export:** different data shape (investigation
fields) and a register-style layout, though it mirrors the actions-export structure
and shareable-URL/htmlCache pattern.

**How to apply:** keep the 4 formats in lockstep (see meeting-export-lockstep.md). Risk
palette: Extreme #000, High #ef4444, Moderate #eab308 (black text), Low #22c55e.

## Actions page shows ALL near misses (not just actioned ones)

The Actions page (`actions.tsx`) is otherwise a worklist that hides items with no
action data, but Near Misses are an exception: EVERY near miss shows as a permanent
safety record. `actionItems` includes `item.type === 'Near Miss' || hasActionData(item)`.

- `isNearMissRecord(item)` = near miss with no action data ("old" ones). These render a
  grey **"Closed"** badge (no "Investigate" badge) and are excluded from the live
  `stats.open` count (they have no due date/priority so they're already out of
  overdue/highPriority).
- Near misses BYPASS the open/archived status toggle (`filterStatus`) and the
  closed-item hide rule, so the full register is always on the page. Non-near-miss
  worklist behaviour (closed hidden, status toggle, deep-link focus) is unchanged.

**Why:** the source SharePoint list returns all near misses fine (no filtering bug) —
the app reads them LIVE, no sync needed; the user wanted the page itself to mirror the
register, with old ones marked closed but still visible for the record.

**Note:** "Closed" here is display-only; it does NOT write a Closed status back to
SharePoint. A real bulk write-back would be a separate, confirmed operation.

## Optional From/To date band on the register export

The register export modal has an optional From/To date band. The client filters
`nearMissItems` by `getDateGroupKey(item.submittedDate)` (a NZ-aligned yyyy-mm-dd key)
vs the date-input strings via plain string compare — inclusive of both ends and
timezone-proof. Empty band = all near misses. Filtered `registerItems` (not
`nearMissItems`) drive the modal count and the buttons' disabled state.

Server: HTML/Markdown/Word generators take an optional `dateRangeLabel` and render a
"Period covered" line; routes build it from dateFrom/dateTo via buildNearMissRangeLabel.
**CSV is intentionally label-free** (raw tabular data) — the date band still applies to
CSV because filtering happens client-side before POST. So the "4 formats in lockstep"
rule has ONE exception: the period label is on 3 of 4, the filtered data is on all 4.

**Why:** all formats receive the already-filtered `items` array, so data stays
consistent; only the human-readable period caption differs by format.

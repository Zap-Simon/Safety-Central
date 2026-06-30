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

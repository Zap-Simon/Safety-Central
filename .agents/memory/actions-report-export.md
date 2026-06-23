---
name: Actions report export engine
description: How the Actions page export is built and how it relates to the meeting-minutes export engine
---

# Actions report export

The Actions page export is action-centric (one block per tracked action), built on
the SAME professional foundation as the meeting minutes: Cranfield-branded, A4
print-ready HTML with the inlined Paged.js polyfill ("Page X of Y" footers,
repeating table headers, break-inside protection), plus matching CSV / Markdown /
Word outputs.

**Why a separate module:** the meeting minutes are organised by meeting; the
Actions report is organised by action and carries data the meeting export does not
(per-action activity history, Near Miss investigation details, On Hold revisit date,
Ready to Close, due-date analytics). Keeping it separate avoids overloading the
meeting generator while still sharing the look-and-feel and Paged.js setup.

**How to apply:**
- Generators live in `server/actions-export.ts`; endpoints `/api/generate-actions-{html,csv,markdown,word}` in `server/routes.ts`.
- The client (`client/src/pages/actions.tsx`) posts `sortedItems` (already
  filtered/sorted on the page) + the on-page `stats`. What the user sees is what
  exports — never re-filter server-side.
- The server enriches each item: activity via `storage.getActivityLog(listType, id)`
  and (Near Miss only) `storage.getNearMissInvestigation(id)`. listType map:
  `'Near Miss' -> 'NearMiss'`, else `type.replace(' ', '')`.
- HTML reuses the same `global.htmlCache` + `/api/view-html/:shareId` shareable-URL
  mechanism as the meeting HTML export.
- Type colours kept in lockstep everywhere: Business Ideas #2563eb, Safety Ideas
  #dc2626, Near Miss #ea580c. Due-date badges are suppressed for archived
  (Completed / Ready to Close) actions, matching the page.

---
name: Paged.js split-table column widths
description: Why table columns drift between page 1 and later pages in the meeting-minutes HTML export, and the reliable fix.
---

# Paged.js drops <colgroup> on split table fragments

The meeting-minutes export ("items-table") uses `table-layout: fixed` plus a
`<colgroup>` (40/30/30) and is paginated by inlined Paged.js.

**Symptom:** the Agenda Item column is wider on page 1 than on page 2+, and the
final column narrows to compensate — i.e. column widths are inconsistent across
pages.

**Why:** when Paged.js splits a table across pages it rebuilds a fresh table
fragment per page and does **not** reliably copy the `<colgroup>` to the
continuation fragments. Those pages then fall back to content-based sizing even
though `table-layout: fixed` is set, so widths drift.

**How to apply:** don't rely on `<colgroup>` alone. Pin each column width on the
cells themselves by position, e.g. `.items-table th/td:nth-child(1){width:40%}`,
`:nth-child(2){30%}`, `:nth-child(3){30%}`. These rules apply to every fragment
regardless of colgroup propagation. Keep the colgroup too (defense in depth).
If columns are added/reordered, update the nth-child mapping in the same change.

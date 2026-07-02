---
name: Paged.js oversized table rows must be allowed to split
description: Blanket break-inside avoid on table rows strands section headers above empty tables when one row is taller than a page
---

**Rule:** In the Paged.js meeting-minutes export, `.items-table tr { break-inside: avoid }` keeps normal items whole, but any row taller than the remaining page space gets pushed wholesale to the next page — stranding the section header ("II. Meeting Minutes") + thead above an empty table. Rows whose longest cell exceeds ~1200 chars are server-flagged `class="row-splittable"` (break-inside: auto on tr AND td) so they split across the boundary; the repeated thead covers continuation pages.

**Why:** A detailed near-miss write-up (submission + hazard-register discussion) produced a near-blank page with just the heading, which read as broken formatting.

**How to apply:** When adding new columns or content to export table rows, include their text length in the `longestCellChars` heuristic (it counts submission, discussion, and action-required lines). Never re-blanket `break-inside: avoid` without the splittable exception. The threshold is heuristic, not layout-aware — tune from real exports if misclassification shows up.

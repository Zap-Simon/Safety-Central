---
name: Export table colgroup/CSS width lockstep
description: Fixed-layout export tables define widths twice (colgroup + CSS nth-child); colgroup wins, so change both.
---

The A4 report export tables (near-miss register, actions report) use `table-layout:fixed`
AND declare column widths in TWO places:
- an inline `<colgroup><col style="width:..%">` row in the HTML string, and
- `.items-table th/td:nth-child(n){width:..%}` rules in the `<style>` block.

**Rule:** any column-width change must update BOTH in lockstep. With `table-layout:fixed`
the `<colgroup>` is authoritative in browsers/Paged.js, so editing only the CSS nth-child
rules has NO visible effect (headers still clip/wrap, dates still wrap).

**Why:** a layout fix to widen the "Investigation" column and stop date wrapping edited
only the CSS and appeared to do nothing because the stale colgroup kept the old widths.

**How to apply:** when fixing column widths or wrap issues in server/*-export.ts, grep for
`colgroup` in the same file and update the `<col style="width:...">` list to match the CSS.
For single-line headers use `th{white-space:nowrap}` and put `white-space:nowrap` on the
specific date column td (via nth-child) since body cells otherwise use `overflow-wrap:anywhere`.

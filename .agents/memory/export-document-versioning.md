---
name: Export document versioning (document control)
description: All human-readable exports carry version/issued/review via one shared helper; CSV excluded.
---

Every human-readable export (meeting minutes, actions report, near-miss register) shows a
"Document control" footer: version, issued date (= generation date), and next-review date
(NZ today + 1 year). This is a quality/compliance requirement, not decoration.

**Single source:** `server/meeting-export-shared.ts` exports `buildDocumentVersionInfo`,
`documentVersionFooterHTML`, and `documentVersionFooterMarkdown`. Review date is computed
from NZ "today" (Pacific/Auckland) then formatted in UTC so it never drifts a day around
midnight. Default version is "1.0".

**Wired into:** HTML — routes.ts (meeting), actions-export.ts, near-miss-register-export.ts
(footer sits before the inlined Paged.js `<script>`, before `</body>`). Markdown — the same
three engines + markdown-generator.ts (meeting). Word — each docx builder pushes a bordered
centered paragraph before `new Document(...)` (template-engine.ts for meeting).

**Why:** CSV is intentionally SKIPPED (data files, same decision as the period-covered label)
— adding a version row would corrupt the tabular format.

**How to apply:** any NEW export format/engine should import and render these helpers so it
stays in lockstep; keep CSV exempt. Change the shared helper (not per-engine copies) to bump
the version or reword the footer.

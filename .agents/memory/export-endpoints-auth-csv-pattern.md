---
name: Export endpoints auth & CSV escaping pattern
description: Why the report-export endpoints have no auth guard, trust client-posted items, and don't formula-escape CSV — and that this is intentional and app-wide.
---

The branded report-export endpoints (meeting minutes, Actions report, Near Miss
Register — `/api/generate-*-{html,csv,markdown,word}`) all share three deliberate
traits:

1. **No per-route auth/authz guard** — they're registered as plain `app.post(...)`
   with no session/token check.
2. **They trust the client-posted `items` payload** (`rawItems: any[]` + cast), then
   enrich server-side from local DB by id.
3. **CSV uses only quote-doubling** (`"${cell.replace(/"/g,'""')}"`), no
   formula-leading (`= + - @`) neutralization.

**Why:** This is the established project-wide export pattern, not an oversight. The
near miss data already lives client-side (cards are loaded via the same authed
queries) and enrichment data is reachable via existing investigation endpoints the
client already uses. Every export endpoint follows this same shape for consistency.

**How to apply:** If you add a new export format/endpoint, match this pattern — do
NOT add a divergent auth guard or CSV formula-escaping to just one endpoint. If
hardening is genuinely required, change ALL export endpoints together (the formats
are kept in lockstep; see meeting-export-lockstep.md / actions-report-export.md), and
confirm the scope with the user first since it's a cross-cutting security decision.

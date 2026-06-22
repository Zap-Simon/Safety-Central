---
name: Large-file decomposition convention
description: Stop growing the oversized god-files; ship new features as new modules and carve structure out over time.
---

# Large-file decomposition convention

Two files have grown into hard-to-maintain god-files:
- `server/routes.ts` (~6.5k lines) — every API route in one file.
- `client/src/pages/meeting-history.tsx` (~4.6k lines) — one giant page component.

**The rule:** Do NOT add new features into these files. New backend endpoints go into a
new domain-scoped router module (e.g. group by area: sharepoint, meetings, actions,
near-miss) registered from `routes.ts`. New UI goes into focused sub-components/hooks,
not more lines in `meeting-history.tsx`. When you touch a cohesive slice of an existing
god-file for another reason, opportunistically extract that slice into its own module so
structure forms incrementally rather than via one risky big-bang refactor.

**Why:** At this size the files are hard to navigate, risky to edit (unrelated things
break), and slow to load into context. The user explicitly wants future features shipped
*out* of these files with a real folder structure forming over time. Keeping additions
modular also avoids forcing context compression when an agent has to read them.

**How to apply:** Before implementing a feature that would land in `routes.ts` or
`meeting-history.tsx`, create/choose a smaller module instead and wire it in. Prefer many
small files over appending to the big ones. Treat any edit to these files as a chance to
extract, never to grow.

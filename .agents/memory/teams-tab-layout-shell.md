---
name: Teams personal-tab layout shell
description: Fixed app-shell + internal-scroll pattern that keeps the Submit and Orders Teams tabs consistent and scroll-free.
---

# Teams personal-tab layout shell

The two Teams personal tabs (Submit, Orders) share one layout contract. Keep them in lockstep when editing either.

- The router shell (App.tsx `TeamsRouterContent`) is a **fixed-height** column: `h-screen overflow-hidden` with `paddingBottom: calc(3.5rem + safe-area)` reserving space for the `fixed` bottom nav. The active tab is wrapped in `flex flex-col flex-1 min-h-0 overflow-hidden`.
- Each tab's authenticated root is `flex flex-col h-full min-h-0`; a slim `shrink-0` header sits at top; the content/list region is `flex-1 min-h-0 overflow-y-auto` so **only the content scrolls**, never the page.
- Submit centers its single card with an inner `min-h-full flex flex-col justify-center` inside the scroll region (so it's balanced when short, scrolls when tall).
- Full-screen states (loading / unauthenticated / done) must use `h-full` (NOT `min-h-screen`).

**Why:** `min-h-screen` (=100vh) on a tab while it lives inside the shell that already pads for the fixed bottom nav makes the tab taller than its viewport slot → phantom page scroll and inputs pinned awkwardly at the top. The user explicitly disliked both ("text inputs too high up", "unsure if we need scroll").

**How to apply:** Never reintroduce `min-h-screen` inside a Teams tab. Both tabs use matching slim headers (icon chip + "Hi {first}" + one-line subtitle), Submit=blue / Orders=purple accents. No gradient headers (replit.md user preference). Orders is a shared list visible to all signed-in staff (GET /api/orders is public) — its header subtitle / empty state / count footer must keep saying it's shared with the team.

---
name: Teams personal-tab layout shell
description: Fixed app-shell + top segmented switcher pattern for the Submit/Orders Teams tabs, and what NOT to reintroduce.
---

# Teams personal-tab layout shell

The two Teams personal tabs (Submit, Orders) share one layout contract. Keep them in lockstep when editing either.

## Chrome belongs to Teams, not us
- **No in-app header bar on either tab.** Teams already shows the app name in its own chrome, so a second app header reads as duplicate. (User explicitly asked to remove them.)
- **Navigation is a quiet top segmented toggle** (`TeamsTabSwitcher` in App.tsx), NOT a bottom nav bar. A bottom app-nav stacks right above Teams' own bottom navigation and looks like a competing second nav. The toggle is a rounded-full pill group, max-w-[240px], active segment = white/gray-900 thumb + blue(Submit)/purple(Orders) text + shadow-sm; it lives in the shell above the content and outside the scroll region so it stays visible while content scrolls.

## Height / scroll contract (keyboard-safe)
- Router shell (App.tsx `TeamsRouterContent`) is a **fixed-height** column: `h-screen overflow-hidden`, padding only for `env(safe-area-inset-*)` (NO bottom-nav height reservation — there is no bottom nav). Renders `<TeamsTabSwitcher />` then the active tab wrapped in `flex flex-col flex-1 min-h-0 overflow-hidden`.
- **Both tabs use the same keyboard-safe pattern**: authenticated root `flex flex-col h-full min-h-0`; any header-ish/greeting element is **pinned `shrink-0` at top** (does NOT scroll); the primary input + content lives in a single `flex-1 min-h-0 overflow-y-auto` region, **top-aligned (no `my-auto`/`justify-center`)**. This keeps the focused input near the top so the on-screen keyboard never hides it, and only the content scrolls — never the page.
- Full-screen states (loading / unauthenticated / done) use `h-full`, never `min-h-screen`.

**Why keyboard-safe matters:** when Submit centered its card with `my-auto`, the focused textarea sat mid-viewport and the on-screen keyboard pushed/hid it. Orders never had this problem because its add bar is pinned `shrink-0` at top. Submit was restructured to mirror Orders. (User explicitly compared the two and wanted Submit to behave like Orders.)

## Per-tab specifics
- **Submit**: greeting is a **prominent, purposeful heading** (`text-2xl font-bold`, gray-900 light / white dark), reads `Hi {first} 👋`, pinned `shrink-0` at top. (It was briefly a faint ambient watermark — user reversed that and asked for it to be clearly visible with the wave emoji.) Card sits top-aligned in the scroll region below.
- **Orders**: NO greeting. Starts directly with the quick-add bar (pinned `shrink-0`). It is a **shared list visible to all signed-in staff** (GET /api/orders is public) — keep the "shared with your team" messaging in the empty state and the count footer.
- **Toggle** (`TeamsTabSwitcher`): give it real top breathing room (`pt-5`); rounded-full track with a subtle `ring-1`, `font-semibold` segments. It should feel modern and in-content, never like a second nav bar.

**Why:** `min-h-screen` inside a tab makes it taller than its viewport slot → phantom page scroll and inputs pinned awkwardly high; the user disliked both. A bottom app-nav duplicated Teams' own nav; a second app header duplicated Teams' app-name chrome.

**How to apply:** Never reintroduce `min-h-screen` inside a tab, a fixed bottom nav, a per-tab app header, or a `my-auto`/centered input region (breaks keyboard safety). Submit greeting stays a prominent pinned heading; Orders stays greeting-free and shared-list-labelled.

---
name: Teams personal-tab layout shell
description: Fixed app-shell + top segmented switcher pattern for the Submit/Orders Teams tabs, and what NOT to reintroduce.
---

# Teams personal-tab layout shell

The two Teams personal tabs (Submit, Orders) share one layout contract. Keep them in lockstep when editing either.

## Chrome belongs to Teams, not us
- **No in-app header bar on either tab.** Teams already shows the app name in its own chrome, so a second app header reads as duplicate. (User explicitly asked to remove them.)
- **Navigation is a quiet top segmented toggle** (`TeamsTabSwitcher` in App.tsx), NOT a bottom nav bar. A bottom app-nav stacks right above Teams' own bottom navigation and looks like a competing second nav. The toggle is a rounded-full pill group, max-w-[240px], active segment = white/gray-900 thumb + blue(Submit)/purple(Orders) text + shadow-sm; it lives in the shell above the content and outside the scroll region so it stays visible while content scrolls.

## Height / scroll contract
- Router shell (App.tsx `TeamsRouterContent`) is a **fixed-height** column: `h-screen overflow-hidden`, padding only for `env(safe-area-inset-*)` (NO bottom-nav height reservation anymore — there is no bottom nav). Renders `<TeamsTabSwitcher />` then the active tab wrapped in `flex flex-col flex-1 min-h-0 overflow-hidden`.
- Each tab's authenticated root is `flex flex-col h-full min-h-0`; the content/list region is `flex-1 min-h-0 overflow-y-auto` so **only the content scrolls**, never the page.
- Full-screen states (loading / unauthenticated / done) use `h-full`, never `min-h-screen`.

## Per-tab specifics
- **Submit**: greeting is an **ambient background element**, not a header — soft large text (`text-3xl font-extrabold`, gray-300 light / gray-700 dark) in normal flow at the top of the content, with the card `my-auto` centered below it. Only rendered when `userName` exists. (User wanted it to "feel part of the app", not a header bar.)
- **Orders**: NO greeting. Starts directly with the quick-add bar. It is a **shared list visible to all signed-in staff** (GET /api/orders is public) — keep the "shared with your team" messaging in the empty state and the count footer.

**Why:** `min-h-screen` inside a tab makes it taller than its viewport slot → phantom page scroll and inputs pinned awkwardly high; the user disliked both. A bottom app-nav duplicated Teams' own nav; a second app header duplicated Teams' app-name chrome.

**How to apply:** Never reintroduce `min-h-screen` inside a tab, a fixed bottom nav, or a per-tab app header. Submit greeting stays ambient; Orders stays greeting-free and shared-list-labelled.

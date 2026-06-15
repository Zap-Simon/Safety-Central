---
name: Mobile-friendly tab navigation pattern
description: How to keep shadcn Tabs from overlapping their own labels on phones in this app
---

# Mobile tab overlap gotcha

The base shadcn `TabsTrigger` (client/src/components/ui/tabs.tsx) sets `whitespace-nowrap`.
When a `TabsList` uses a fixed multi-column grid (e.g. `grid-cols-4`) with long text labels,
the labels cannot wrap and overflow into neighboring columns — they visually overlap on narrow
phone screens.

**The fix / convention used across tabbed pages (home, equipment-maintenance, staff-training):**
- TabsList: responsive grid that reduces columns on mobile (home uses `grid-cols-2 sm:grid-cols-4`),
  plus `h-auto gap-1`.
- Each TabsTrigger: `flex flex-col sm:flex-row items-center justify-center gap-1 sm:gap-2 h-auto py-2
  whitespace-normal text-center text-xs sm:text-sm leading-tight` (add `px-2 sm:px-3` for tighter
  3-column layouts). `whitespace-normal` overrides the base nowrap; `flex-col` stacks icon over label
  on mobile; `h-auto` lets the row grow.

**Why:** `cn()` uses tailwind-merge, so the later `whitespace-normal`/`h-auto` correctly override the
base `whitespace-nowrap`/`h-10`.

**How to apply:** any new tabbed page or any `grid-cols-N` TabsList with multi-word labels must follow
this pattern, otherwise labels overlap on mobile.

---
name: action_items duplicate rows from listType drift
description: Why local action_items can have >1 row per SharePoint item and how readers must collapse them deterministically.
---

# action_items duplicate rows from listType naming drift

The local `action_items` table can hold MORE THAN ONE row for the same
`sharepoint_item_id`, because `list_type` was written differently across app
versions (legacy spaced "Safety Ideas"/"Business Ideas" vs current unspaced
"SafetyIdeas"/"BusinessIdeas", produced by `item.type.replace(' ', '')`).
`upsertActionItem` matches on (listType, sharePointItemId), so a rename creates a
new row instead of updating the old one.

**Rule:** any reader that merges local action overrides onto SharePoint items
(e.g. `buildMergedMeetingItems`) MUST collapse duplicates by keeping the row with
the greatest `updatedAt` — newest row wins wholesale. Never just `map.set()` per
row keyed on sharePointItemId alone.

**Why:** a plain `map.set()` lets whichever duplicate is iterated LAST win
non-deterministically, so a field the user just saved (famously a "Low" priority)
silently reverts on reopen. "Latest row wins wholesale" is deterministic, mirrors
single-row behavior, and keeps intentional clears working (a null on the newest
row wins, then falls back to the SharePoint value via `??`).

**Don't** do "latest non-null value per field" — that resurrects stale values a
newer save intentionally cleared to null (clears appear to fail).

**Note:** legacy duplicate rows still sit in prod; the read fix makes them
harmless. A one-time dedupe is optional cleanup, not required for correctness.

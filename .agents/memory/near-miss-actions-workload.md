---
name: Near miss visibility on Actions page
description: Which near misses count as "open to manage" vs permanent record on the Actions page
---

# Near miss visibility on the Actions page

The Actions page keeps EVERY near miss on the page as a permanent safety record, but must not let finished ones clutter the live "Open (to manage)" view (the default filter is `'open'`).

**Rule:** a near miss is OUT of the live workload (`isNearMissOutOfWorkload`) when it is formally Closed (`item.status === 'Closed'`), its action is finished (Completed / Ready to Close), or it is record-only (no action data). Such near misses only appear under the "All" status filter; every other filter (incl. default "open") hides them.

**Why:** users reported closed near misses showing under "Open (to manage)". The old code let near misses bypass the status filter entirely (`item.type !== 'Near Miss'`), so closed/finished ones always showed. This mirrors how the dashboard `stats.open` count already excluded record-only near misses — the list just wasn't consistent with it.

**How to apply:** near misses must respect the status filter using out-of-workload as their "archived/finished" flag; never re-add a blanket `item.type !== 'Near Miss'` bypass. Discoverability of closed near misses lives under the "All" filter and the Near Miss Register export.

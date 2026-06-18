---
name: Meeting date grouping must use one UTC key everywhere
description: Why item-per-meeting counts/filters must use getDateGroupKey, not isSameDay
---

Meeting items are grouped into meetings by **UTC day**. Two date helpers in
`shared/dateUtils.ts` do NOT agree:
- `getDateGroupKey()` uses **UTC** date components — this is what the meeting view
  grouping and the server export filter (`new Date(d).toISOString().split('T')[0]`) use.
- `isSameDay()` compares with `Date.toDateString()` — **local/browser timezone**.

**Trap:** items WITHOUT an actual SharePoint MeetingDate get a *calculated* meeting date
whose time component is arbitrary, so it can fall on a different day under local vs UTC.
When such items (often Safety / Near Miss, which more often lack an explicit MeetingDate
than Business Ideas) are grouped by UTC but counted/filtered by `isSameDay` (local), the
count is short — e.g. a meeting shows (17) instead of (25). The items are still grouped
and exported correctly; only the local-timezone count was wrong.

**Rule:** any per-meeting count or filter MUST compare with
`getDateGroupKey(a) === getDateGroupKey(b)` so it matches the grouping and the server
export. Do not use `isSameDay` for meeting membership.

---
name: Meeting re-surfacing aggregation (Ready to Close / On Hold)
description: How action items are pulled into the nearest upcoming meeting, and the date-key rules that keep them from being miscounted or duplicated.
---

# Re-surfacing action items into a meeting

Meeting-history.tsx renders one section per meeting group and, inside that loop, has
"aggregation containers" that pull items in from OTHER meeting groups:
- "Ready to Close – Group Review" (emerald): every item whose actionStatus === 'Ready to Close'.
- "On Hold – Time to Reconsider" (orange): On Hold items whose `reconsiderDate` has arrived.

**Rule: aggregation must target the single nearest upcoming meeting, not every upcoming group.**
Gating only on `getMeetingStatus(meetingDate).isUpcoming` duplicates the same item into
every future-dated meeting group. Compute the earliest upcoming group key once
(`groupedByMeetingAndCategory.map(getDateGroupKey).filter(isUpcoming).sort()[0]`) and render
the aggregated container only when the current group equals that key.
**Why:** items can have future meetingDates, creating several "upcoming" groups; without
scoping, a due item shows up multiple times.

**Rule: all date comparisons here use `getDateGroupKey` (UTC), never raw `toISOString()`
or `isSameDay` (local).** A reconsider date is stored as a `YYYY-MM-DD` string; compare
`getDateGroupKey(reconsiderDate) <= getDateGroupKey(new Date())` for "due", and
`getDateGroupKey(reconsiderDate) === getDateGroupKey(meetingDate)` for past-meeting placement.
**Why:** mixing local-midnight (`getMeetingStatus`) with UTC keys drifts items by a day around
the NZ/UTC boundary. Staying on the UTC key basis matches how items are grouped. (Companion:
meeting-date-grouping-consistency.md.)

# Auto-scheduled ("synthetic") next meeting
When NO upcoming meeting group exists but items need review (any Ready to Close, or On Hold
whose reconsiderDate is due), meeting-history unshifts a placeholder upcoming group so the
aggregation containers have a home: default date = latest real meeting + 7 days rolled forward
until upcoming, overridable by a shared server-side planned date (GET/POST
/api/planned-meeting-date, app_settings KV table). The header reschedule pencil detects the
synthetic meeting and saves the planned date instead of rewriting SharePoint items.
**Rules:** skip synthetic creation in search mode; when seeding the default date, filter group
keys to `^\d{4}-\d{2}-\d{2}$` first — the "unknown-meeting" bucket otherwise produces an
Invalid Date that throws at render. Planned-date endpoints are deliberately unauthenticated,
matching the meeting-locks convention.

# On Hold deferral data
`actionItems.reconsiderDate` (text ISO date) holds the revisit date. It round-trips through
POST /api/action-items (normalize empty→null) and the /api/meeting-history merge block, same as
the other action fields. ProcessedMeetingItem also carries it.

# ActionStatusWorkflow component
Single clickable stepper Not Started → In Progress → Ready to Close → Completed, plus a
separate On Hold branch that reveals the reconsider-date picker. It exposes ONE atomic
`onChange({actionStatus?, reconsiderDate?})` callback (not two) so leaving On Hold clears the
date in the same save. **Why:** two back-to-back updateActionFields calls race — the second
reads stale `item` from before the first's cache update and re-writes the old reconsiderDate.

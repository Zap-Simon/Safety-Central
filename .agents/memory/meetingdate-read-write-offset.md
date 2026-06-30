---
name: MeetingDate +1-read / -1-write compensation
description: All SharePoint MeetingDate WRITE paths must subtract 1 day, because the READ path adds 1 day; missing it shifts/cascades dates forward.
---

`processMeetingDate()` adds **+1 day** when reading SharePoint `MeetingDate` (to reconcile the raw UTC API value with the NZ MS Lists UI date staff see). Therefore every path that WRITES a user-facing/display meeting date MUST subtract **1 day** first, or items read back one day late.

**Why:** A create flow (web "big blue button") wrote `MeetingDate` raw without the -1 compensation, so new items showed up a day late. Worse, it cascaded: each subsequent add targeted the newly mis-dated "nearest upcoming" meeting group, pushing dates forward one day per add (e.g. items meant for the 30th landed on 1st, 2nd, 3rd, 4th July). `updateItemMeetingDate()` already did the -1; `createListItem()` did not.

**How to apply (server/sharepoint-lists-service.ts):**
- `createListItem` and `updateItemMeetingDate` → subtract 1 day before `toISOString()`.
- `updateItemFields` meetingDate branch is currently un-compensated but UNUSED for date edits (UI routes date changes through `/api/sharepoint/move-item-to-meeting` → `updateItemMeetingDate`). If ever wired to a display-date edit, add the -1.
- `moveItem` cross-list copy writes `sourceData.MeetingDate` raw→raw unchanged — correct, do NOT compensate (it's not a display date).
- Rule of thumb: compensate when the input is a display/NZ date from the client; don't when copying a raw stored value.

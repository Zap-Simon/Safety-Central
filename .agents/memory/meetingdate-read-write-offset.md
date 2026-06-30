---
name: MeetingDate read/write convention (canonical NZ-midnight)
description: READ adds +1 day to every stored MeetingDate; WRITES must store NZ-midnight "(D-1)T12:00:00Z" so the app AND the SharePoint Lists UI both show date D.
---

`processMeetingDate()` (server/sharepoint-lists-service.ts) adds **+1 day** to every stored `MeetingDate` on read (reconciling the raw UTC API value with the NZ MS Lists UI date staff see), then normalizes to `<D>T10:00:00.000Z`.

**Canonical storage rule:** every WRITE of a display/NZ meeting date must store **NZ-midnight of D = `(D-1)T12:00:00.000Z`** via the shared helper `formatMeetingDateForWrite(displayDate)`. Then:
- app read: parse `(D-1)T12:00Z`, +1 day → UTC date D → shows D.
- Lists UI: `(D-1)T12:00Z` is 00:00 (NZST/+12) or 01:00 (NZDT/+13) on D in NZ → shows D.
- 12:00Z (not 00:00Z) is deliberate: the +12/+13 NZ offset never crosses a day boundary, so it's correct year-round including DST and month/year edges.

**Why this convention (not the old "-1 day keeping time"):** an earlier fix just did `setDate(getDate()-1)`, keeping the input's `T10:00Z` time. That made the **app** correct but stored `(D-1)T10:00Z`, which the NZ Lists UI renders as D-1 — so SharePoint was then a day BEHIND. The bulk of correct live data (the Forms/Power Automate flow, ~116/156 Business items) is already stored at NZ-midnight (`(D-1)T11:00Z` summer / `(D-1)T12:00Z` winter). Writing NZ-midnight makes app writes match that dominant convention so BOTH surfaces agree.

**Live data has three conventions (survey via Graph app-only):**
- `T11:00Z` / `T12:00Z` = NZ-midnight → correct under +1 read (the good bulk).
- `T00:00Z` (midnight UTC) and `T10:00Z` = off-by-one: app shows D+1 while Lists shows D. ~57 such items existed (Business 40, Safety 14, Near Miss 3). The `T10` ones were the app's own old writes.
- Fixing an off-by-one item = rewrite to `(UTC-date-of-stored − 1 day)T12:00:00Z`; this LEAVES the SharePoint-displayed date unchanged and only stops the app adding the extra day.

**How to apply (server/sharepoint-lists-service.ts):**
- `createListItem`, `updateItemMeetingDate`, and `updateItemFields` meetingDate branch → all route through `formatMeetingDateForWrite()`.
- `moveItem` cross-list copy writes `sourceData.MeetingDate` raw→raw unchanged — correct, do NOT compensate (it's a stored value, not a display date).
- Rule of thumb: use `formatMeetingDateForWrite` when the input is a display/NZ date from the client; never when copying an already-stored raw value.

**Graph app-only access for read/inspection & targeted data fixes:** env `AZURE_TENANT_ID`/`AZURE_CLIENT_ID`/`AZURE_CLIENT_SECRET`, scope `https://graph.microsoft.com/.default`, Sites.ReadWrite.All. SharePoint REST app-only is REJECTED ("Unsupported app only token") — must use Graph. Sites: root = `/sites/cranfieldglass.sharepoint.com`; Near Miss lives on `/sites/cranfieldglass.sharepoint.com:/sites/IncidentsReports`. Provenance of a bare stored value can't distinguish buggy-create from move-hack, so blanket auto-migration is unsafe — prefer targeted fixes + the forward-only code fix.

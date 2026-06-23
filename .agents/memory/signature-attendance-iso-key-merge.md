---
name: Signature/attendance cross-ISO key merge
description: Why meeting signatures/attendance must be merged by date-group key, not read by a single raw ISO.
---

Meeting signatures and attendance are stored keyed by a **raw meeting ISO string**, not a date key. The admin meeting-history page and the Teams personal Sign tab each independently pick a "representative ISO" for a calendar day as the *first item's ISO* from their own (differently-ordered) item lists. So the same person/day can be written under two different ISO keys — and a signature collected in Teams silently fails to appear on the admin Meeting Sign-Off / attendance UI.

**Rule:** any reader of signatures/attendance must merge every ISO entry that shares a `getDateGroupKey` (YYYY-MM-DD) bucket, never read by a single raw ISO.
- Attendance merge = union of names.
- Signature merge = keep the most recent `signedAt` per name on conflict.
- In `isAttending`, a person's **signature status is the source of truth** (`absent` ⇒ not attending; `signed`/`remote` ⇒ attending); fall back to the attendance union only when no signature exists. This also sidesteps the union "can't uncheck" edge, since signed people aren't governed by the attendance union.
- Toggle/derive-current-state logic must use the same merged source the checkbox renders from, or the toggle flips the wrong direction.
- HTML export (`/api/generate-meeting-html`) fetches signatures from the DB and must merge **all** same-day keys, not first-match; it uses the client-sent `meetingAttendance` directly, so the client must re-key its merged buckets onto each representative ISO the export expects.

**Why:** storage was never normalized to a date key, and the two representative-ISO pickers diverge. Normalizing on read is the contained fix; changing the storage key format would touch the storage layer, Teams GET/POST, exports, and need a data migration.

---
name: Teams locked-minutes read access
description: Who may read locked meeting minutes in the Improve+ Teams SignTab, and why read is decoupled from sign/attendance.
---

Any signed-in Cranfield Glass staff member may READ locked/closed meeting
minutes in the Teams personal app (SignTab) — no attendee-roster match and no
allowlist. Reading is fully decoupled from signing.

**Why:** User-confirmed product rule (e.g. a remote worker like Debra Craig is
never on the roster and never signs, but must still be able to read minutes).
Tying read access to attendance silently excludes rostered-but-absent staff too.

**How to apply:**
- Sign permission (POST /api/teams/sign) STILL requires roster match — do not relax that.
- Read endpoints (GET /api/teams/minutes/meetings and /:dateKey) only validate a
  valid Microsoft identity via resolveSignerFromRequest; roster match is intentionally NOT checked.
- Read access is still gated to locked/closed meetings only (buildMeetingLockMap)
  and to >= SIGN_VISIBLE_FROM_KEY, grouped/merged by getDateGroupKey (UTC).
- In SignTab, the readable-minutes list loads for EVERY authenticated user (not
  just non-roster ones), so absent rostered staff can browse minutes they didn't attend.
- Minutes HTML is rendered via generateTeamsMinutesHTML (lightweight, no
  Paged.js/@page/print button — distinct from generateMeetingMinutesHTML) inside
  a sandboxed <iframe srcDoc> so its styles can't leak into the Teams chrome.

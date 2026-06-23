---
name: Teams locked-minutes read access
description: Who may read locked meeting minutes in the Improve+ Teams SignTab, and why read is decoupled from sign/attendance.
---

Any signed-in Cranfield Glass staff member may READ locked/closed meeting
minutes in the Teams personal app (SignTab) — no attendee-roster match and no
allowlist. Reading is fully decoupled from signing.

**Why:** User-confirmed product rule (e.g. a remote worker who is never on the
roster and never signs must still be able to read minutes). Tying read access to
attendance silently excludes rostered-but-absent staff too.

**How to apply:**
- Sign permission STILL requires roster match — do not relax that. Only the READ
  path drops the roster requirement (identity validation only).
- Read access is still gated to locked/closed meetings only and grouped by
  getDateGroupKey (UTC).
- The readable-minutes list must load for EVERY authenticated user (not just
  non-roster ones), so absent rostered staff can browse minutes they didn't attend.
- A Teams read-only minutes renderer MUST reuse the official export's content +
  attendance/signature semantics (signed/remote => present; attendance ticks
  otherwise; no ticks => default present; present-without-signature => pending;
  absent = not-present or explicit absent). Merge BOTH signatures AND attendance
  across same-day ISO keys, or present-without-signature attendees vanish.

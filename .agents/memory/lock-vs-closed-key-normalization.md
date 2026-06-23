---
name: Meeting lock vs closed key normalization mismatch
description: Why any reader of meetingLocks must re-normalize keys, because closed-meeting writes don't.
---

The `meetingLocks` table stores BOTH the locked and closed flags, but the two
admin writers normalize the date key inconsistently:

- The lock writer normalizes to UTC `YYYY-MM-DD` (via `normaliseLockDate` /
  `getDateGroupKey`) before persisting.
- The closed-meeting writer persists the date as given (often a raw ISO string),
  with NO normalization.

**Why:** a row written by the closed path can therefore sit under a raw-ISO key
while the lock path uses the normalized key. A reader that looks up by normalized
key alone will MISS a closed row and treat a closed meeting as still open.

**How to apply:** any code that reads `meetingLocks` to decide if a meeting is
locked/closed must re-normalize every row key with `getDateGroupKey(row.meetingDate)`
and OR the flags together (a locked/closed meeting must never slip through). Don't
trust the stored key to already be normalized. The Teams self-sign endpoints do
this; replicate the pattern anywhere else that gates on lock/closed state.

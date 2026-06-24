---
name: Teams Meetings tab visual hierarchy
description: The 3-tier priority model the Teams "Meetings" (SignTab) list is built around and why.
---

The Teams Meetings tab (`client/src/pages/teams/SignTab.tsx`, shared cards in `MeetingCards.tsx`) is organised into three deliberate visual tiers, ordered by business importance, NOT by meeting chronology:

1. **Priority zone (HeroCard)** — signing is the highest-value action (attendance is compliance evidence). When any open meeting still needs the user's signature, it gets the SOLID-brand hero; the next-meeting agenda then sits below as the brand-TINT hero. When nothing needs signing, the agenda is promoted to the solid hero. Only the single most-recent unsigned meeting becomes the hero; extra unsigned ones fall to an "Also ready to sign" list.
2. **Your attendance** — open-but-already-signed meetings (awaiting admin lock) + locked meetings the user attended. Success-tinted icon chips.
3. **Minutes archive** — quiet, neutral-toned read-only list, de-duplicated against everything shown above via `shownKeys`.

**Why:** the boss values agendas, compliance values minutes, but signing is the one thing that *must not be missed*, so it out-weighs everything when pending. Keep this ordering if you add sections.

**How to apply:** `HeroCard` has two tones — `solid` (top action) and `tint` (secondary). Don't give two solid heroes at once; the solid slot is reserved for the single most important thing on screen. Open meetings are sorted newest-first so the hero is deterministic regardless of backend order.

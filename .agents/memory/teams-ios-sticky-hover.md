---
name: Teams iOS sticky-hover on Fluent Card
description: Why interactive Fluent v9 Cards get "stuck" looking focused on iOS Teams, and how to avoid it.
---

Fluent v9 `Card`, when given `onClick`, becomes interactive and applies its own
`:hover` rules that recolor BOTH text (`colorNeutralForeground1Hover`) and
background (`colorNeutralBackground1Hover`). Those rules are NOT gated behind
`@media (hover: hover)`.

**Symptom:** on iOS WebKit (Teams mobile), `:hover` sticks to the last-tapped
screen position after the DOM swaps. So after tapping a card to open a sub-view
and returning (the list remounts), the card re-rendered at that position gets the
stuck hover → text looks dark/"black", background changed, until the next tap.
Reads to the user as the card being stuck in "focus mode".

**Fix (the reliable one):** make the Fluent `Card` NON-interactive — do not pass
`onClick` to `Card`. Move click + keyboard semantics to a wrapper `div`
(`role="button"`, `tabIndex={0}`, `onKeyDown` for Enter/Space). With no `onClick`,
Fluent attaches none of its interactive hover/focus restyle. Also gate any of your
OWN custom `:hover` styles behind `@media (hover: hover)` so touch never triggers
them either. Keep `:active` (transform scale) ungated — it gives tap feedback and
does not stick like `:hover`.

**Why not just override colors:** the per-card base bg/text differ (hero solid vs
tint vs default card), so a single global `@media (hover: none)` reset can't
restore each card's base. Inline-pinning works but kills desktop hover. The
non-interactive-Card + wrapper approach preserves desktop hover AND keyboard a11y.

**How to apply:** any new tappable Fluent Card in the Teams tabs should use the
`tapWrap` pattern in `client/src/pages/teams/MeetingCards.tsx`, not `Card onClick`.

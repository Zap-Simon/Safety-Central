---
name: Teams swipe-to-delete (framer-motion)
description: How admin order cards are deleted by horizontal swipe in the Teams Orders tab, and the gotcha that lets it coexist with vertical list scroll.
---

# Swipe-sideways-to-delete on Teams cards

Admin item cards in the Orders tab are deleted by an iOS-style left-swipe (no
visible trash button). Built with `framer-motion` (already a dependency) — Fluent
UI v9 has NO built-in swipe/swipeable-row primitive, so this must come from
framer-motion (or similar), not Fluent.

Structure per row: a `swipeWrap` (position relative, overflow hidden, rounded) holds
a red `deleteLayer` (absolute, full-bleed, trash icon right-aligned) BEHIND a
`motion.div` card. Drag the card left to reveal the red zone; release past threshold
deletes.

Key props on the `motion.div`:
- `drag="x"`, `dragConstraints={{ left: -96, right: 0 }}`, `dragElastic={{ left: 0.15, right: 0 }}`
- `dragSnapToOrigin` — always springs back to 0; on a real delete the row just
  unmounts when the query invalidates, so no manual exit animation needed.
- `onDragEnd` fires delete when `info.offset.x < -72 || info.velocity.x < -500`
  (distance OR flick). Guard with the per-item `deleting` flag so a second swipe
  during an in-flight delete is ignored.

**Why the scroll gotcha matters:** a horizontal drag inside a vertically-scrolling
list will hijack vertical scroll unless you (1) set `dragDirectionLock` so framer
commits to ONE axis based on initial finger movement, and (2) set
`touchAction: "pan-y"` on the draggable element so the browser still owns vertical
panning. Omit either and mobile users can't scroll the list past a card.

**How to apply:** non-admin users get a plain non-draggable card (just
`<div>{cardInner}</div>`) — only render the motion/swipe wrapper when the user is an
order admin. Deletes are server-side soft-archives, so an accidental swipe is
recoverable.

**Fluent v9 Card stroke gotcha:** `<Card>` defaults to `appearance="filled"`, which
has NO visible border — just a faint fill. If the swipe rows look like they lost their
"container" outline, set `appearance="outline"` on the Card to restore the stroke.

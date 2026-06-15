---
name: Global mobile CSS quirks
description: Site-wide CSS rules in index.css that silently distort small components on mobile
---

# Global mobile CSS quirks (client/src/index.css)

Inside `@media (max-width: 768px)` there is a global rule:

```css
button, a, .nav-link { min-height: 44px; }
```

**Why it matters:** Any small square `<button>` (e.g. avatar/initials circles sized
~36px) gets force-stretched to 44px tall while width stays ~36px → renders as a
**vertical oval on mobile only**. Looks fine on desktop.

**How to apply:** When making a small round/square button, override with a
`min-h-[…]` utility (a class beats the bare `button` selector by specificity).
Inline `style` width/height alone does NOT fix it — `min-height` still wins.
If the element doesn't need to be a button, use a `div` to dodge the rule entirely.

# iOS safe-area insets
- Notch (top) is painted blue by giving the fixed header `paddingTop: env(safe-area-inset-top)`
  (header bg is the blue gradient, so the padding region shows blue).
- Requires `viewport-fit=cover` in the viewport meta (client/index.html).
- Content offset: pages use `pt-16` to clear the fixed header; a mobile-only
  `.pt-16 { padding-top: calc(4rem + env(safe-area-inset-top)) }` override keeps content clear.
- Bottom home-indicator strip kept white via a `body::after` fixed white bar of height
  `env(safe-area-inset-bottom)`.

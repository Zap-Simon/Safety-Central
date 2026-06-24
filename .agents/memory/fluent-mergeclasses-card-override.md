---
name: Fluent v9 mergeClasses vs Card defaults
description: Why combining makeStyles classes with string concatenation silently fails to override Fluent Card styles, and the two-part fix.
---

# Fluent v9: combine makeStyles classes with mergeClasses, never string concat

When styling Fluent UI v9 (`@fluentui/react-components`) components — especially
`<Card>` — combine `makeStyles` (Griffel) classes with `mergeClasses(a, b, ...)`,
NOT plain template-string concatenation like `` `${styles.a} ${styles.b}` ``.

**Why:** Griffel classes carry a leading `___seq_xxxx` lookup token that
`mergeClasses` uses to dedupe/order atomic classes by property. Concatenating two
Griffel outputs as a string produces TWO `___` tokens, which breaks the merge —
the console warns "a passed string contains multiple identifiers of atomic
classes … concatenated in a wrong way." The practical symptom: your conflicting
properties (background, flex-direction, color) do NOT win against the component's
internal defaults. Fluent `Card` internally does `mergeClasses(internal, className)`
so a *properly merged* className wins; a raw concatenated string does not.

Two real bugs this caused (Teams Meetings tab):
- Solid hero card's brand background was overridden by Card's default white bg →
  white-on-white text in LIGHT mode. (Dark "worked" by accident: Card's default
  dark bg kept white text readable, masking the same broken override.)
- Minutes card stacked vertically (chevron on its own line) because Card defaults
  to `flex-direction: column` and the row layout never applied.

**How to apply:**
1. Use `mergeClasses(styles.base, cond && styles.variant, "global-anim-class")`.
   Falsy operands are safely ignored; plain global (non-Griffel) class names like
   `animate-fade-in-up` are preserved.
2. An UNSET property can never win the merge. Fluent `Card` defaults to
   `display:flex; flex-direction:column`. If you want a horizontal row you must
   set `flexDirection: "row"` EXPLICITLY in your makeStyles — mergeClasses alone
   won't fix layout if the property isn't declared.
3. Combining ONE Griffel class with ONE global class via string concat is fine
   (no double-`___` token); the bug only bites when concatenating 2+ Griffel
   classes, or relying on an undeclared property to override a component default.

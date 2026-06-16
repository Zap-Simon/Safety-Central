---
name: Teams manifest icons & color override
description: Why the Teams nav icon shows our literal color instead of the Teams theme tint, and how the two manifest icons differ.
---

# Teams manifest icons (color.png vs outline.png)

Teams personal apps ship two icons in `teams-app/manifest.json` and use them in different places:
- **color.png** — 192×192, full color. Shown in the app store, the "..." apps list, larger surfaces. May be a filled colored icon (ours: blue `#2563EB` bg + white modern-plus glyph).
- **outline.png** — 32×32, shown in the **left nav rail** once the app is pinned/saved. MUST be a **transparent background with a single solid color (white) silhouette**. Teams then tints it with its own theme/selected color (purple selected, grey unselected).

## The gotcha (the override only works with a proper outline)
If `outline.png` is a flat, fully-colored, non-transparent PNG (e.g. a solid blue square saved as plain RGB, no alpha), Teams has nothing transparent/white to tint, so it renders your **literal color** instead of applying its theme override. Symptom reported: "Teams applies a logo then overrides it with our zipped square / the selected color reverts to our blue." Fix = outline.png as 32×32 transparent + white glyph (verify color type is RGBA/gray+alpha, not RGB).

**Why:** Teams recolors the outline glyph; it cannot recolor opaque pixels. A colored/opaque outline defeats the override.

**How to apply:** generated with ImageMagick (`magick` available; `convert` too). Design at 8× then downscale for crisp edges. `xc:none` = transparent canvas. Always re-check the PNG color type after export — RGB (colortype 2) means no alpha and the tint will fail.

## App display name drives the loading placeholder
Before the manifest icon loads (e.g. when reopening Teams), Teams shows a generated placeholder tile using the **app name initials** — not anything from our files. App short name is **"Improve+"** (was "Safety & Ideas").

## Re-publishing a manifest change
Bump `manifest.json` `version` (Teams won't pick up changes otherwise) and re-zip `manifest.json` + `color.png` + `outline.png` at the zip ROOT. No `zip` binary in this env — use Python `zipfile`.

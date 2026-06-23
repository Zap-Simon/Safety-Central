---
name: Teams srcDoc iframe blank on iOS
description: Why the Teams minutes srcDoc iframe must use sandbox="allow-same-origin"
---
A `srcDoc` iframe that is sandboxed WITHOUT `allow-same-origin` renders completely BLANK on iOS WebKit (the engine behind Teams mobile on iPhone). Desktop Chromium/WebView2 renders it fine, so the bug only shows on iOS.

**Why:** The Teams read-only meeting-minutes viewer (SignTab) injects server-rendered, escaped HTML/CSS (no scripts) into a srcDoc iframe. Using `sandbox=""` gave it an opaque origin and iOS WebKit refused to paint it — the reported "blank meeting minutes page" on Teams mobile.

**How to apply:** For any srcDoc iframe that must render on iOS/Teams mobile, set `sandbox="allow-same-origin"`. Deliberately OMIT `allow-scripts` — the content has no JS, so scripts stay blocked and the only change is the origin. CSP needs no relaxation (inline styles already covered by `style-src 'unsafe-inline'`, signature images by `img-src data:`). Never reach for `blob:` URLs as an alternative: prod CSP `frame-src` does not include `blob:`.

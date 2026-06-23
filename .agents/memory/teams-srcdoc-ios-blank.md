---
name: Teams srcDoc iframe blank on iOS
description: Why the Teams minutes srcDoc iframe must use sandbox="allow-same-origin"
---
A `srcDoc` iframe that is sandboxed WITHOUT `allow-same-origin` renders completely BLANK on iOS WebKit (the engine behind Teams mobile on iPhone). Desktop Chromium/WebView2 renders it fine, so the bug only shows on iOS.

**Why:** The Teams read-only meeting-minutes viewer (SignTab) injects server-rendered, escaped HTML/CSS (no scripts) into a srcDoc iframe. Using `sandbox=""` gave it an opaque origin and iOS WebKit refused to paint it — the reported "blank meeting minutes page" on Teams mobile.

**How to apply:** For any srcDoc iframe that must render on iOS/Teams mobile, set `sandbox="allow-same-origin"`. Deliberately OMIT `allow-scripts` — the content has no JS, so scripts stay blocked and the only change is the origin. CSP needs no relaxation (inline styles already covered by `style-src 'unsafe-inline'`, signature images by `img-src data:`). Never reach for `blob:` URLs as an alternative: prod CSP `frame-src` does not include `blob:`.

## iOS iframe-in-flex zero-height (intermittent blank)
Even with `allow-same-origin`, an iframe sized purely as a flex item (`flex-grow:1; min-height:0`) intermittently computes to ZERO height on iOS WebKit and only paints after an unrelated reflow (switching Teams apps, re-entering the card). Symptom: minutes show "after clicking around Teams a few times."

**Fix:** don't size the iframe by flex intrinsic sizing. Wrap it in a `position:relative; flex-grow:1; min-height:0` box and make the iframe `position:absolute; inset:0; width:100%; height:100%`. The concrete positioned box gives it a real pixel size at first layout, so it renders immediately and reliably.

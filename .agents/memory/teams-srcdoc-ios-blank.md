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

## iOS post-resume blank (sign-out/in) — deterministic iframe remount
Even with `allow-same-origin` and the positioned wrapper, the minutes go BLANK again specifically AFTER a sign-out/sign-in (a WebView navigation + resume). The freshly-mounted iframe isn't composited until an unrelated reflow; users' reliable workaround was tapping another Teams app and coming back (a visibility change forces the paint).

**Fix:** automate that nudge. The strongest lever is to fully RE-CREATE the iframe, not just nudge a repaint — bump a `nonce` state used in the iframe `key`. Do it (a) once shortly after the minutes open (setTimeout ~250ms) for the stubborn post-auth case, and (b) on `visibilitychange`(visible)/`focus`/`pageshow` to mirror the app-switch workaround. Keep nonce OUT of the effect deps so bumping it remounts the iframe without re-running the effect (no loop). An immediate offsetHeight-read + transform `translateZ(0)`→`''` repaint nudge handles the fast path flicker-free before the remount lands.

**Why:** transform/offsetHeight nudges alone are heuristic and flaky on WKWebView resume; a brand-new DOM element forces WebKit to lay out and paint from scratch.

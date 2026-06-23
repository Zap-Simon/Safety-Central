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

## RESOLUTION: abandon the srcDoc iframe entirely — render minutes via Shadow DOM
All of the iframe workarounds above (allow-same-origin, positioned wrapper, repaint nudges, deterministic remount on visibility/focus/pageshow) STILL left the minutes blank on iOS Teams after sign-out/sign-in. The root cause is the separate iframe compositor getting stuck un-painted on a WKWebView auth resume — no amount of nudging is reliable. The durable fix is to NOT use an iframe at all: render the server minutes inline in a **Shadow DOM** root.

**How to apply:** keep a `<div ref>` host; in an effect, `host.shadowRoot ?? host.attachShadow({mode:"open"})` (attachShadow only once per element — reuse on re-render), parse the server's full-document HTML with `DOMParser` (INERT — nothing executes/loads), create a `<style>` via `textContent` containing a `:host{...}` rule that restates the doc body's background/typography (the doc CSS targets `body`, which maps onto `:host`) + the trusted `<head>` `<style>` text, then move sanitized body nodes into the shadow root with `document.importNode`+`appendChild`. Cleanup: `shadow.replaceChildren()`. This paints reliably because it's an ordinary in-document element, not a second compositor.

**Why:** Shadow DOM keeps the minutes' CSS isolated from the Fluent/Teams chrome (same benefit the iframe gave) WITHOUT a separate iframe compositor that iOS can leave blank.

**SECURITY — mandatory when inlining:** the old `sandbox` iframe (no allow-scripts) was a hard script barrier; rendering inline removes it, so server HTML must be sanitized before it touches the LIVE tree or stored fields (e.g. `signatureData` in `<img src=...>`) become a stored-XSS vector. Sanitize the INERT DOMParser body first: remove script/iframe/object/embed/base/link/meta/form/style/svg, strip every `on*` attribute, and drop src/href/xlink:href/srcset whose value isn't `https:`/`data:image/`/`mailto:`/`#`/`/`. NEVER assign the raw body string to `shadow.innerHTML`; use importNode so no string is re-parsed in the live context. Lift styles only from `doc.head` (not arbitrary body `<style>`).

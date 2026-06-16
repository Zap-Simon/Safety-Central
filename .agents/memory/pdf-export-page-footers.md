---
name: PDF/HTML export page-number footers
description: How to get reliable "Page X of Y" footers in the meeting-minutes HTML export (downloaded standalone file printed in-browser).
---

# Page-number footers in the HTML meeting-minutes export

The meeting minutes export is a standalone `.html` file the user downloads and
opens from `file://`, then prints/saves-as-PDF from the browser.

**Rule:** CSS `@page { @bottom-center/@bottom-right { content: counter(page) ... } }`
margin boxes do **not** render in Chrome/Firefox/Safari print engines. A
`position: fixed` footer repeats on every printed page but **cannot** show a
dynamic page number (`counter(page)` only works inside `@page` context).

**Why:** The requirement "page number footers on each page" cannot be met with
pure CSS when printing HTML from a normal browser.

**How to apply:** Inline **Paged.js** (the paged-media polyfill) into the export.
It paginates the body into page boxes and renders `@page` margin boxes reliably,
offline, in any browser. Specifics that matter:
- Bundle the library as base64 in a `.ts` module (`server/assets/pagedjs.ts`) and
  decode once at module load — this guarantees it survives esbuild bundling into
  `dist/` (a `server/assets/*.js` file would NOT be copied by esbuild). CDN is
  unreliable (offline files + dev-server CSP blocks external scripts).
- Escape any literal `</script>` in the decoded library before inlining.
- Put branding in `@bottom-center`, page numbers in `@bottom-right`; do NOT also
  add a fixed-position footer (duplicates in engines that support both).
- Paged.js paginates whatever is in `<body>`, so keep the floating Print button
  OUT of the source flow — re-inject it in `window.PagedConfig.after`, plus a
  `setTimeout` fallback in case Paged.js never finishes.

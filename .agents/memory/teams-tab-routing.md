---
name: Teams tab routing & manifest coupling
description: The Teams Personal Tab route, its manifest, and the perceived-performance choices must stay in lockstep.
---

# Teams Personal Tab — routing & manifest

The "Safety & Ideas" Teams Personal Tab is served at an intentionally
**unlisted/obscure** route so it is not discoverable from the main app.

**Rule:** The route string in `client/src/App.tsx`, the `contentUrl`/`websiteUrl`
in `teams-app/manifest.json`, and the published domain must always match.

**Why:** Teams loads the tab via the manifest URL. If the route changes but the
manifest is not updated in lockstep, the tab silently 404s inside Teams with no
visible error to staff. The route is deliberately not linked from any main-app
nav so it stays effectively hidden on the public domain.

**How to apply:** When changing the tab path, update both files and the
`validDomains` + `webApplicationInfo.resource` entries in the manifest. The
manifest targets the published deployment domain, not the dev `.replit.dev` URL.

## Perceived-performance choices (Teams tab)
- Classification (`/api/ai-classify`) runs **eagerly in the background** while the
  user types (debounced), cached + de-duped by input text, so pressing Continue is
  usually instant.
- `/api/sharepoint/create-item` accepts a `deferTitle` flag: when set it returns as
  soon as the item is persisted and generates the AI title fire-and-forget. Default
  (main-app) path stays synchronous — do not remove the flag gate.
- Classification uses `gpt-4o-mini` (not gpt-4o) for speed; classification is simple.

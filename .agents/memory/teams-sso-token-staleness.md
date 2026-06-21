---
name: Teams SSO token staleness
description: Why Teams tabs must fetch a fresh SSO token per request instead of caching the init-time token
---

# Teams SSO token must be re-fetched per request, not cached at init

Teams SSO tokens (`microsoftTeams.authentication.getAuthToken()`) expire after ~1 hour. If a tab captures the token once at provider init and reuses it for every API call, an idle tab will send a stale token; the server OBO exchange then fails and surfaces a bogus "Your Microsoft sign-in has expired. Please sign in again." even though the user is still signed in. Reloading "fixes" it only because reload re-runs init.

**Rule:** the Teams auth provider exposes a `getToken()` that calls `getAuthToken()` fresh right before every outgoing request. The SDK caches/refreshes internally, so this is cheap on the happy path and transparently returns a new token after expiry.

**Why:** the SDK is the source of truth for token freshness; React state is not. Caching the token in state turns a self-healing SDK behaviour into a manual-reload bug.

**How to apply:**
- Use `await getToken()` for the Authorization header in each fetch (both tabs, queries and mutations).
- Keep the stored `teamsToken` only as a coarse "auth initialized" signal for queryKey/enabled gating; do NOT gate actions on it — gate on `authState === "authenticated"` and let `getToken()` be the source of truth.
- Clear `authError` on a successful `getToken()` so stale error text doesn't linger.

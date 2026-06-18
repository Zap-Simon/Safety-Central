---
name: Auth 401 returns a benign {data:[]} body
description: Why data-list React Query calls must check res.ok or items silently vanish after a deploy
---

Auth-protected data endpoints (e.g. `/api/meeting-history`) return **HTTP 401 with a
body of `{ authenticated:false, data:[] }`** in two cases: (1) no/invalid Bearer header,
and (2) Microsoft Graph/SharePoint rejects an expired token (server maps those to 401 in
its catch). The body is intentionally shaped like a success payload.

**The trap:** a query that does `.then(res => res.json())` WITHOUT checking `res.ok`
parses the 401 body as success → renders an empty list → every item appears to have
vanished, with no retry. This recurs right after deploys because the page reloads and the
SharePoint token isn't valid for a beat.

**Rule:** every React Query call hitting these endpoints must `if (!res.ok) throw` so the
401 becomes an error (enabling retry + an error UI) instead of silent empty data.

**Retry:** retry transient HTTP/network failures (token-not-ready self-heals), but do NOT
retry interaction-required / "No authenticated accounts" / popup-cancelled errors — that
just re-triggers MSAL sign-in popups in a loop. Gate with a `retry: (failureCount, error)`
predicate.

**Re-auth UX:** on persistent failure show a "session expired, sign in again" banner that
calls `authService.signIn()` then `refetch()`, rather than leaving an empty list. Note the
app already has a global sign-in gate, so these endpoints only fail for *expired* tokens of
already-signed-in users, not first-load "not signed in".

**Note:** `meeting-history.tsx` uses `getSharePointToken()`; `actions.tsx` uses
`getAccessToken()` — both hit the same endpoint. Keep their res.ok/retry handling in sync.

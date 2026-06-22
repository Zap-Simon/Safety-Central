---
name: MSAL stale token 401 after deploy
description: Why authenticated fetches return 401 after code changes/deploys, and the fix pattern.
---

## The rule
Any `authenticatedFetch` wrapper that calls `authService.getAccessToken()` must include a **forceRefresh retry** on 401. Without it, stale MSAL cache tokens cause blank pages after every deploy.

**Why:** After a server restart MSAL's browser-side cache still holds the previous access token. `acquireTokenSilent` returns that cached token silently — no network call to Microsoft. SharePoint/server rejects it with 401. Retrying with the same token (even multiple times) always fails.

**How to apply:** Two-step pattern in every `authenticatedFetch`:
1. Try with normal token.
2. If response is 401, call `getAccessToken(forceRefresh: true)` and retry once.

Also: `authService.getAccessToken(forceRefresh = false)` must pass `forceRefresh` into the MSAL `SilentRequest` — `{ scopes, account, forceRefresh }`.

```ts
// authService.ts
async getAccessToken(forceRefresh = false): Promise<string> {
  const silentRequest: SilentRequest = { scopes: loginRequest.scopes, account: accounts[0], forceRefresh };
  ...
}

// page component
const authenticatedFetch = async (url, options = {}) => {
  const makeRequest = async (forceRefresh: boolean) => {
    const token = await authService.getAccessToken(forceRefresh);
    return fetch(url, { ...options, headers: { ...options.headers, Authorization: `Bearer ${token}` } });
  };
  const response = await makeRequest(false);
  if (response.status === 401) return makeRequest(true);
  return response;
};
```

**TanStack Query retry:** Once 401 is handled at the fetch level, stop the query from also retrying on 401 — add `401` to the no-retry pattern:
```ts
retry: (failureCount, error) => {
  const msg = error instanceof Error ? error.message : String(error);
  if (/401|interaction|popup|cancel|No authenticated accounts/i.test(msg)) return false;
  return failureCount < 2;
}
```

**Fallback UX:** If the forced refresh still fails (genuine auth expiry), show a visible "Retry" button — never leave the user looking at an empty page with no way to recover without a full reload.

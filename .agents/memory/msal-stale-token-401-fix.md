---
name: SharePoint endpoints need getSharePointToken, not getAccessToken
description: 401s on SharePoint-backed API endpoints come from sending the wrong MSAL token scope; and why forceRefresh popups are a trap.
---

## The rule
Any client fetch to a SharePoint-backed endpoint (e.g. `/api/meeting-history`, `/api/sharepoint/*`) MUST use `authService.getSharePointToken()`, NOT `authService.getAccessToken()`.

**Why:** `getAccessToken()` requests `loginRequest.scopes` (login/Graph). `getSharePointToken()` requests `sharePointRequest.scopes` (SharePoint resource). The server validates the token audience against SharePoint, so a login/Graph token is rejected with 401 even though the user is fully signed in. The reference implementation that works is `meeting-history.tsx`; mirror its `authenticatedFetch` exactly when adding a new SharePoint-backed page.

**How to apply:** copy this shape (no forceRefresh, no popup path):
```ts
const authenticatedFetch = async (url, options = {}) => {
  const token = await authService.getSharePointToken();
  return fetch(url, { ...options, headers: { ...options.headers, Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' } });
};
```
Query: check `res.ok` before `.json()` (401 returns a benign `{data:[]}` body that would silently empty the list); retry transient faults a few times but NOT `interaction|popup|cancel|No authenticated accounts`.

## The forceRefresh popup trap (do NOT repeat)
Do not "fix" a 401 by calling `getAccessToken(forceRefresh: true)` and retrying. `forceRefresh` makes `acquireTokenSilent` fail more readily, which falls through to `acquireTokenPopup`. In production the security headers block the popup, producing an endless flood of `Cross-Origin-Opener-Policy policy would block the window.closed call` and the 401 never clears. The 401 was never a stale-token problem — it was the wrong token scope. Fix the scope, never force interactive popups inside a background data fetch.

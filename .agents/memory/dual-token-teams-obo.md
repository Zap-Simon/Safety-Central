---
name: Dual-token backend (Teams OBO + main-app MSAL)
description: Why shared API endpoints must detect token audience and either OBO-exchange or pass through, and what must stay consistent.
---

# Dual-token auth on shared endpoints

The Teams personal-tab and the main browser app hit the SAME backend endpoints, but
they authenticate differently:

- **Main (non-Teams) app** uses browser MSAL and sends a ready-made downstream Graph /
  SharePoint access token as the Bearer.
- **Teams tabs** send a single Teams SSO token (audience = this app's clientId or its
  `api://` URI), which CANNOT call Graph/SharePoint directly. The backend exchanges it
  via On-Behalf-Of (OBO) for the downstream token.

The backend broker inspects the unverified `aud` claim ONLY to choose a path: Teams-audience
→ OBO exchange; anything else → pass the token through unchanged. Trust still lives with
Azure AD / the downstream Microsoft API (a forged token fails OBO or fails the Graph/SharePoint
call), so audience-routing is not a security boundary.

**Why:** the browser MSAL flow in Teams (ssoSilent/popup/redirect) was fragile inside the
Teams iframe — `interaction_in_progress`, `monitor_window_timeout`. OBO moves the token dance
server-side; the Teams client only needs `getAuthToken()`. The pass-through keeps the main
app working without touching its MSAL code.

**How to apply:**
- When adding a new endpoint that BOTH the Teams tabs and the main app call, resolve the token
  through the broker (graph vs sharepoint resource) — never trust the raw Bearer for Teams tabs.
- OBO downstream scopes MUST mirror what the browser MSAL flow already requests, or consent is
  missing and OBO returns AADSTS65001.
- Do NOT reintroduce browser MSAL (msalConfig/authService) into the Teams tabs, and do NOT
  remove it from the main app — they are intentionally separate.
- Error mapping in the broker: consent/AADSTS65001/interaction → 403 (admin must consent);
  expired/revoked assertion (invalid_grant/expired) → 401 (user signs in again).

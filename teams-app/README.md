# Cranfield Glass — Improve+ Teams App

This folder contains the Teams Personal Tab app package for the Cranfield Glass staff engagement system.

## What the app contains

The personal tab has two screens, switched via a small **segmented toggle at the top of the content** (Submit / Orders). It is deliberately *not* a bottom navigation bar — Teams already provides its own navigation along the bottom of the window, so an in-content toggle keeps the two clearly separate and makes the screen feel like part of the app rather than a second nav layer.

| Tab | Route | Purpose |
|-----|-------|---------|
| **Submit** | `/teams-submit-cg7k2x9m` (manifest) / `/teams-tab` (toggle) | Describe a near miss, safety observation, supply request, or improvement idea. AI classifies it and routes it to the correct SharePoint list. Shows an ambient "Hi {first name}" greeting. |
| **Orders** | `/teams-tab/orders` (toggle) | Shared ordering whiteboard — every signed-in staff member sees the same list; staff add items they need ordered and admins mark items as ordered. |

The Teams manifest registers a single static tab pointing to the Submit screen. The Orders screen is reached via the in-app toggle inside the same iframe — no second manifest entry is needed.

## Folder contents

- `manifest.json` — Teams app manifest (schema v1.16)
- `bump-version.js` — helper to increment the manifest version before re-packaging
- `color.png` — 192×192 full-colour icon (required by Teams)
- `outline.png` — 32×32 transparent outline icon (required by Teams)
- `README.md` — this file

## How to package and upload

1. Zip only the manifest and the two icons (not the folder, and not this README or the helper script):
   ```
   cd teams-app
   zip ../cranfield-safety-ideas.zip manifest.json color.png outline.png
   ```

2. Go to **Teams Admin Center → Teams apps → Manage apps → Upload new app**.

3. Upload `cranfield-safety-ideas.zip`.

4. Once approved, users find it under **Apps → Built for your org** and pin it as a personal tab.

## Before packaging for a new domain

If the app is ever moved to a custom domain or a new Replit deployment, update these four fields in `manifest.json`:

- `staticTabs[0].contentUrl`
- `staticTabs[0].websiteUrl`
- `validDomains[0]`
- `webApplicationInfo.resource`

## Authentication — On-Behalf-Of (OBO)

This app uses the **OAuth 2.0 On-Behalf-Of flow**. The browser never holds a Microsoft Graph or SharePoint token — it only ever holds a Teams SSO token, and the **server** swaps that for the downstream tokens it needs.

### How it works

Both tabs sign in the **same silent way**. There is no popup, no redirect, and no interactive fallback — this is an internal-staff app that only runs inside Teams.

1. **Teams hands the tab a single SSO token** via `microsoftTeams.authentication.getAuthToken()`. The audience of that token is *this* Azure AD app (not Graph or SharePoint), so the browser cannot call those APIs with it directly.
2. **The tab sends that token to our backend** as a normal `Authorization: Bearer <token>` header on API requests that need Microsoft 365 access.
3. **The backend exchanges it On-Behalf-Of the signed-in user** (`server/teams-obo-auth.ts`, using `@azure/msal-node`'s `acquireTokenOnBehalfOf`) for the downstream access token each endpoint actually needs — Microsoft Graph or SharePoint.
4. **Downstream tokens are cached server-side** (in memory, keyed by a hash of the assertion + resource) until shortly before they expire, so repeat requests don't re-exchange. The raw token is never logged.

Because routing is decided by the incoming token's audience, the **main (non-Teams) website is unaffected**: it still uses browser MSAL and sends ready-made Graph/SharePoint tokens, which the backend detects and passes through unchanged. The same API endpoints serve both clients.

> Outside Teams (e.g. opening the route in a plain browser) `getAuthToken()` throws *"No Parent window found"* and the tab shows a **Sign in required** screen. That is expected — there is no Teams parent to issue the SSO token.

### Server environment variables (required)

The OBO exchange runs with the app's **confidential client** credentials. Set these as server secrets — never expose them to the browser:

| Variable | What it is |
|----------|-----------|
| `AZURE_CLIENT_ID` | The Azure AD app (client) ID |
| `AZURE_CLIENT_SECRET` | A client secret generated under *Certificates & secrets* |
| `AZURE_TENANT_ID` | The Azure AD tenant (directory) ID |

### Downstream delegated permissions (must be consented)

The OBO exchange requests these delegated scopes. Admin consent must already be granted for the tenant (the existing browser MSAL flow uses the same scopes, so consent is normally already in place):

- **Microsoft Graph:** `User.Read`, `Sites.Read.All`
- **SharePoint** (`https://cranfieldglass.sharepoint.com`): `AllSites.FullControl`

If a scope has not been consented the backend returns a clear `403` asking an administrator to grant the app's delegated permissions. An expired/stale SSO token returns `401` (sign in again).

### Azure AD app registration — one-time setup (for this app or a fork)

The Teams tab shares the existing Azure AD app. No second registration is needed because the main site and the Teams tab are served from the same domain.

For Teams SSO + OBO to work, the app registration needs:

**1. Application ID URI** (*Expose an API*):
```
api://staff-engagement-system.replit.app/dad12e09-3e7f-42fa-86e8-a0378bdd2699
```
For a fork, use `api://<your-domain>/<your-client-id>` and update `webApplicationInfo.resource` in `manifest.json` to match.

**2. The `access_as_user` scope** (*Expose an API → Add a scope*):
- Name: `access_as_user`
- Who can consent: Admins and users
- State: Enabled

**3. Authorize the Teams client applications** (*Expose an API → Authorized client applications*) — add both so Teams can issue the SSO token without prompting:
- `1fec8e78-bce4-4aaf-ab1b-5451cc387264` (Teams desktop & mobile)
- `5e3ce6c0-2b1f-4285-8d4b-75ee78787346` (Teams web)

**4. A client secret** (*Certificates & secrets*) → store it on the server as `AZURE_CLIENT_SECRET`.

**5. API permissions** — add and admin-consent the delegated scopes listed above (Graph `User.Read`, `Sites.Read.All`; SharePoint `AllSites.FullControl`).

### Reusing this auth setup in a new (forked) project

1. Copy `server/teams-obo-auth.ts` and call `getDownstreamToken(req, 'graph' | 'sharepoint')` inside your route handlers.
2. Register (or reuse) an Azure AD app and complete steps 1–5 above.
3. Set `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET`, and `AZURE_TENANT_ID` as server secrets.
4. In `teams-obo-auth.ts`, update `TEAMS_APP_AUDIENCE_MARKERS` (client id + `api://` URI), `SHAREPOINT_RESOURCE`, and the scope lists to match your tenant and domain.
5. Point `manifest.json`'s `webApplicationInfo`, `validDomains`, and `staticTabs` URLs at your domain.

## Icons

The current icons are simple placeholders. Replace before rolling out to all staff:
- `color.png` — 192×192 px, full colour, PNG
- `outline.png` — 32×32 px, transparent background, white/light icon only, PNG

# Cranfield Glass — Safety & Ideas Teams App

This folder contains the Teams Personal Tab app package for the Cranfield Glass staff engagement system.

## What the app contains

The personal tab has two screens, switched via a bottom navigation bar:

| Tab | Route | Purpose |
|-----|-------|---------|
| **Submit** | `/teams-submit-cg7k2x9m` | Describe a near miss, safety observation, supply request, or improvement idea. AI classifies it and routes it to the correct SharePoint list. |
| **Orders** | `/teams-tab/orders` (in-app nav) | Shared ordering whiteboard — staff add items they need ordered; admins mark items as ordered. |

The Teams manifest registers a single static tab pointing to the Submit screen. The Orders screen is reached via the bottom nav bar inside the same iframe — no second manifest entry is needed.

## Folder contents

- `manifest.json` — Teams app manifest (schema v1.16)
- `color.png` — 192×192 full-colour icon (required by Teams)
- `outline.png` — 32×32 transparent outline icon (required by Teams)

## How to package and upload

1. Zip only the three files (not the folder):
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

## Authentication

### How it works

**Submit tab — Teams SSO (silent, no prompt)**

1. Teams hands the tab a short-lived SSO token via `microsoftTeams.authentication.getAuthToken()`.
2. MSAL exchanges it for full Graph and SharePoint tokens using `ssoSilent()`.
3. If that fails, the tab shows a "Not available here — try refreshing" message. No interactive fallback is intentional for this internal-staff tab.

**Orders tab — Teams SSO with MSAL fallback**

1. Same Teams SSO path as above.
2. If SSO requires consent or isn't available (e.g. first-time use), MSAL falls back to a popup (desktop) or redirect (Teams mobile).
3. After a redirect sign-in, the app automatically returns to the Orders screen rather than the main site — a `sessionStorage` key preserves the target route across the redirect.

### Azure AD app registration — one-time setup required for Teams SSO

The Teams tab shares the existing Azure AD app (`dad12e09-3e7f-42fa-86e8-a0378bdd2699`). No second registration is needed because both the main site and the Teams tab are served from the same domain (`staff-engagement-system.replit.app`).

However, Teams SSO requires these additions to the existing app registration (do this once in the Azure Portal):

**1. Set the Application ID URI**
In *App registrations → Expose an API*, set the Application ID URI to:
```
api://staff-engagement-system.replit.app/dad12e09-3e7f-42fa-86e8-a0378bdd2699
```

**2. Add the `access_as_user` scope**
Still on *Expose an API*, add a scope:
- Name: `access_as_user`
- Who can consent: Admins and users
- State: Enabled

**3. Authorize Teams client applications**
On the same page under *Authorized client applications*, add both Teams client IDs:
- `1fec8e78-bce4-4aaf-ab1b-5451cc387264` (Teams desktop and mobile)
- `5e3ce6c0-2b1f-4285-8d4b-75ee78787346` (Teams web)

Once these three steps are done, the Submit tab will sign users in silently with no prompt.

## Icons

The current icons are simple placeholders. Replace before rolling out to all staff:
- `color.png` — 192×192 px, full colour, PNG
- `outline.png` — 32×32 px, transparent background, white/light icon only, PNG

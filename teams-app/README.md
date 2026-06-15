# Cranfield Glass — Safety & Ideas Teams App

This folder contains the Teams Personal Tab app package for the Safety & Ideas submission system.

## Contents

- `manifest.json` — Teams app manifest (schema v1.16)
- `color.png` — 192×192 full-colour icon (required by Teams)
- `outline.png` — 32×32 transparent outline icon (required by Teams)

## Before packaging

1. **Update the manifest if deploying to production** — The manifest currently uses the Replit dev domain. If you deploy to a custom domain or a Replit Deployment, update:
   - `staticTabs[0].contentUrl`
   - `staticTabs[0].websiteUrl`
   - `validDomains[0]`
   - `webApplicationInfo.resource`

2. **Replace the placeholder icons** — The icons in this folder are simple placeholder images. Replace them with branded icons before uploading to Teams Admin Center:
   - `color.png` — 192×192 px, full colour, PNG format
   - `outline.png` — 32×32 px, transparent background, white/light icon only, PNG format

## How to package and upload

1. Zip the three files together (manifest.json + color.png + outline.png — **not** the folder, just the files):
   ```
   cd teams-app
   zip ../cranfield-safety-ideas.zip manifest.json color.png outline.png
   ```

2. Go to **Teams Admin Center** → Teams apps → Manage apps → Upload new app.

3. Upload `cranfield-safety-ideas.zip`.

4. Once approved, users can find it under **Apps → Built for your org** and add it as a personal tab.

## Authentication notes

The tab uses MSAL popup/redirect authentication (same as the main app). Inside a Teams iframe:

- **Popup auth works** on Teams desktop and Teams web — MSAL will open a popup for login.
- **Teams mobile** may block popups; Teams SSO (via `microsoftTeams.authentication.getAuthToken`) can be added later as a progressive enhancement.
- The CSP `frame-ancestors` header already includes `*.teams.microsoft.com` and `*.sharepoint.com` to allow Teams to embed the tab.
- No additional Azure AD app registration changes are needed — the existing app registration at `dad12e09-3e7f-42fa-86e8-a0378bdd2699` covers the required scopes.

---
name: Teams tabs share one auth provider
description: Why Teams SSO auth must live in a provider above the tab routes, not per-tab.
---

# Teams SSO auth belongs in a provider above the tabs

The Teams personal tab has multiple screens reached by route changes (Submit /
Orders). Each screen is a separate component that mounts/unmounts on navigation.

**Rule:** run the Teams SSO sign-in (`microsoftTeams.authentication.getAuthToken()`)
ONCE in a context provider mounted above the tab routes, and have each tab consume
the shared `{ authState, teamsToken, userName, authError, retry }` via a hook. Do not
let each tab run its own `initAuth` on mount.

**Why:** when each tab did its own auth on mount, switching tabs remounted the tab and
re-ran sign-in, flashing the "Signing you in…" loader on every quick switch even
though the SSO token is identical for both tabs. A provider above the routes persists
auth state across tab remounts, so the loader only ever shows on first load.

**How to apply:** the provider sits inside the Teams theme provider but above the
route-switching content. Removing per-tab auth also lets the shell read `userName`
from the same context for the greeting — drop any `onUser`/lifted-state prop plumbing
that previously bubbled the name up. On auth failure clear the token too
(`setTeamsToken(null)`) so a failed re-auth can't leave a stale token behind.

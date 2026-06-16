import { ConfidentialClientApplication } from '@azure/msal-node';
import crypto from 'crypto';
import type { Request } from 'express';

/**
 * Server-side On-Behalf-Of (OBO) token broker for the Microsoft Teams tabs.
 *
 * The Teams tabs acquire a single SSO token via microsoftTeams.authentication
 * .getAuthToken(). That token's audience is THIS Azure AD app, so it cannot be
 * used to call Microsoft Graph or SharePoint directly. Instead, the backend
 * exchanges it (on behalf of the signed-in user) for the downstream access
 * token each endpoint needs.
 *
 * The main (non-Teams) app still uses browser MSAL and sends ready-made Graph /
 * SharePoint tokens. Those are detected by audience and passed through
 * untouched, so this broker is fully backward compatible.
 */

const CLIENT_ID = process.env.AZURE_CLIENT_ID;
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;
const TENANT_ID = process.env.AZURE_TENANT_ID;

// Markers that identify an SSO token minted by Teams for THIS app. A Teams SSO
// token's audience is the app's client id or its api:// app-id-uri.
const TEAMS_APP_AUDIENCE_MARKERS = [
  CLIENT_ID || 'dad12e09-3e7f-42fa-86e8-a0378bdd2699',
  'api://staff-engagement-system.replit.app',
];

const SHAREPOINT_RESOURCE = 'https://cranfieldglass.sharepoint.com';

// Downstream delegated scopes — these mirror exactly what the working browser
// MSAL flow already requests, so the user/admin consent is already in place.
const GRAPH_SCOPES = ['User.Read', 'https://graph.microsoft.com/Sites.Read.All'];
const SHAREPOINT_SCOPES = [`${SHAREPOINT_RESOURCE}/AllSites.FullControl`];

export type DownstreamResource = 'graph' | 'sharepoint';

/** Error carrying an HTTP status so route handlers can translate it cleanly. */
export class AuthError extends Error {
  status: number;
  constructor(status: number, message: string) {
    super(message);
    this.status = status;
    this.name = 'AuthError';
  }
}

let cca: ConfidentialClientApplication | null = null;
function getClient(): ConfidentialClientApplication {
  if (!CLIENT_ID || !CLIENT_SECRET || !TENANT_ID) {
    throw new AuthError(
      500,
      'Server is missing Azure AD credentials (AZURE_CLIENT_ID / AZURE_CLIENT_SECRET / AZURE_TENANT_ID).',
    );
  }
  if (!cca) {
    cca = new ConfidentialClientApplication({
      auth: {
        clientId: CLIENT_ID,
        authority: `https://login.microsoftonline.com/${TENANT_ID}`,
        clientSecret: CLIENT_SECRET,
      },
    });
  }
  return cca;
}

/**
 * Read the `aud` claim from a JWT without verifying the signature. Used only to
 * route the token (Teams SSO vs. ready-made downstream token). The real trust
 * boundary is Azure AD itself — a forged token simply fails the OBO exchange.
 */
function readAudience(token: string): string {
  try {
    const payload = token.split('.')[1];
    if (!payload) return '';
    const json = Buffer.from(
      payload.replace(/-/g, '+').replace(/_/g, '/'),
      'base64',
    ).toString('utf8');
    const claims = JSON.parse(json);
    return typeof claims.aud === 'string' ? claims.aud : '';
  } catch {
    return '';
  }
}

function isTeamsSsoToken(token: string): boolean {
  const aud = readAudience(token);
  return TEAMS_APP_AUDIENCE_MARKERS.some((m) => aud.includes(m));
}

// In-memory OBO cache keyed by a hash of (resource + assertion). Entries expire
// 5 minutes before the real token expiry so we never hand out a token that dies
// mid-request. The raw assertion is never stored or logged — only its hash.
interface CacheEntry {
  token: string;
  expiresAt: number;
}
const oboCache = new Map<string, CacheEntry>();

function cacheKey(assertion: string, resource: DownstreamResource): string {
  return crypto
    .createHash('sha256')
    .update(`${resource}:${assertion}`)
    .digest('hex');
}

function pruneExpired(): void {
  const now = Date.now();
  for (const [key, entry] of oboCache) {
    if (now >= entry.expiresAt) oboCache.delete(key);
  }
}

async function exchangeOnBehalfOf(
  assertion: string,
  resource: DownstreamResource,
): Promise<string> {
  const key = cacheKey(assertion, resource);
  const cached = oboCache.get(key);
  if (cached && Date.now() < cached.expiresAt) {
    return cached.token;
  }

  const scopes = resource === 'graph' ? GRAPH_SCOPES : SHAREPOINT_SCOPES;
  let result;
  try {
    result = await getClient().acquireTokenOnBehalfOf({
      oboAssertion: assertion,
      scopes,
    });
  } catch (err: any) {
    const code = err?.errorCode || '';
    const desc = err?.errorMessage || err?.message || 'unknown error';
    const blob = `${code} ${desc}`;
    // AADSTS65001 / interaction_required / consent → the downstream delegated
    // permission has not been granted for this user/tenant. Only an admin (or
    // the user) consenting fixes this, so surface an actionable 403.
    if (/consent|interaction_required|AADSTS65001/i.test(blob)) {
      throw new AuthError(
        403,
        `Microsoft 365 admin consent is required for ${resource} access. Please ask your administrator to grant the app's delegated permissions.`,
      );
    }
    // Expired / revoked / otherwise unusable Teams SSO assertion → the user's
    // sign-in is stale. They just need to sign in again (401), not chase an admin.
    if (/invalid_grant|expired|AADSTS50013|AADSTS500133|AADSTS700082/i.test(blob)) {
      throw new AuthError(401, 'Your Microsoft sign-in has expired. Please sign in again.');
    }
    throw new AuthError(502, `Failed to exchange Teams sign-in for ${resource} access.`);
  }

  if (!result?.accessToken) {
    throw new AuthError(502, `On-behalf-of exchange returned no ${resource} token.`);
  }

  if (oboCache.size > 200) {
    pruneExpired();
    // Hard cap: if pruning didn't free enough, evict oldest entries (Map keeps
    // insertion order) so the cache can never grow without bound.
    while (oboCache.size > 200) {
      const oldest = oboCache.keys().next().value;
      if (oldest === undefined) break;
      oboCache.delete(oldest);
    }
  }
  const realExpiry = result.expiresOn
    ? result.expiresOn.getTime()
    : Date.now() + 3600_000;
  oboCache.set(key, { token: result.accessToken, expiresAt: realExpiry - 5 * 60_000 });
  return result.accessToken;
}

/**
 * Resolve the correct downstream access token for an incoming Authorization
 * header. Teams SSO tokens are redeemed via OBO for the requested resource;
 * ready-made downstream tokens (from the main-app browser MSAL flow) pass
 * through unchanged. Throws AuthError on missing header / failed exchange.
 */
export async function resolveDownstreamToken(
  authHeader: string | undefined,
  resource: DownstreamResource,
): Promise<string> {
  if (!authHeader || !authHeader.startsWith('Bearer ')) {
    throw new AuthError(401, 'Authentication required');
  }
  const incoming = authHeader.substring(7);
  if (isTeamsSsoToken(incoming)) {
    return exchangeOnBehalfOf(incoming, resource);
  }
  return incoming;
}

/** Convenience wrapper that reads the Authorization header off an Express request. */
export async function getDownstreamToken(
  req: Request,
  resource: DownstreamResource,
): Promise<string> {
  return resolveDownstreamToken(req.headers.authorization, resource);
}

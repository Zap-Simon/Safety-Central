import { createContext, useContext, useState, useEffect, useCallback } from "react";
import type { ReactNode } from "react";
import * as microsoftTeams from "@microsoft/teams-js";

type AuthState = "loading" | "unauthenticated" | "authenticated";

interface TeamsAuth {
  authState: AuthState;
  teamsToken: string | null;
  userName: string;
  authError: string;
  retry: () => void;
  /**
   * Always returns a *fresh* Teams SSO token. The Teams JS SDK caches and
   * silently refreshes the token internally, so call this right before every
   * API request rather than reusing the token captured at first load —
   * otherwise a token that expired while the tab sat idle gets sent to the
   * server, the OBO exchange fails, and the user sees a bogus
   * "sign-in expired" error even though they're still signed in.
   */
  getToken: () => Promise<string>;
}

// Minimal, dependency-free JWT payload decode (browser-safe, no validation —
// the server validates the token; we only read display fields from it).
function decodeJwtPayload(token: string): Record<string, any> {
  try {
    const base64 = token.split(".")[1].replace(/-/g, "+").replace(/_/g, "/");
    const json = decodeURIComponent(
      atob(base64)
        .split("")
        .map((c) => "%" + ("00" + c.charCodeAt(0).toString(16)).slice(-2))
        .join(""),
    );
    return JSON.parse(json);
  } catch {
    return {};
  }
}

const TeamsAuthContext = createContext<TeamsAuth | null>(null);

// Auth runs ONCE here at the Teams router level and is shared by both tabs.
// Because the SSO token is identical for Submit and Orders, switching tabs no
// longer remounts a per-tab sign-in — so the "Signing you in…" loader only ever
// shows on the very first load, not on every quick tab switch.
export function TeamsAuthProvider({ children }: { children: ReactNode }) {
  const [authState, setAuthState] = useState<AuthState>("loading");
  const [teamsToken, setTeamsToken] = useState<string | null>(null);
  const [userName, setUserName] = useState("");
  const [authError, setAuthError] = useState("");

  const initAuth = useCallback(async () => {
    setAuthState("loading");
    setAuthError("");
    try {
      await microsoftTeams.app.initialize();
      const ssoToken = await microsoftTeams.authentication.getAuthToken();
      const payload = decodeJwtPayload(ssoToken);
      setUserName(payload.name || payload.preferred_username || payload.upn || "");
      setTeamsToken(ssoToken);
      setAuthState("authenticated");
    } catch (err: any) {
      const msg = `Teams sign-in failed: ${err?.message || String(err)}`;
      console.error(msg, err);
      setAuthError(msg);
      setUserName("");
      setTeamsToken(null);
      setAuthState("unauthenticated");
    }
  }, []);

  // Fetch a fresh Teams SSO token for an outgoing request. The SDK handles its
  // own caching/refresh, so this is cheap on the happy path and transparently
  // returns a new token once the old one expires. We refresh the stored token
  // (and any derived display name) so the rest of the UI stays in sync.
  const getToken = useCallback(async (): Promise<string> => {
    try {
      await microsoftTeams.app.initialize();
      const ssoToken = await microsoftTeams.authentication.getAuthToken();
      const payload = decodeJwtPayload(ssoToken);
      setUserName(payload.name || payload.preferred_username || payload.upn || "");
      setTeamsToken(ssoToken);
      setAuthError("");
      setAuthState("authenticated");
      return ssoToken;
    } catch (err: any) {
      const msg = `Teams sign-in failed: ${err?.message || String(err)}`;
      console.error(msg, err);
      setAuthError(msg);
      setUserName("");
      setTeamsToken(null);
      setAuthState("unauthenticated");
      throw new Error(msg);
    }
  }, []);

  useEffect(() => {
    initAuth();
  }, [initAuth]);

  return (
    <TeamsAuthContext.Provider
      value={{ authState, teamsToken, userName, authError, retry: initAuth, getToken }}
    >
      {children}
    </TeamsAuthContext.Provider>
  );
}

export function useTeamsAuth(): TeamsAuth {
  const ctx = useContext(TeamsAuthContext);
  if (!ctx) throw new Error("useTeamsAuth must be used within a TeamsAuthProvider");
  return ctx;
}

import { useState, useRef, useEffect } from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import { msalInstance, loginRequest } from "@/auth/msalConfig";
import { InteractionRequiredAuthError, BrowserAuthError, SsoSilentRequest } from "@azure/msal-browser";
import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import {
  ShoppingCart,
  Plus,
  CheckCircle2,
  Clock,
  Loader2,
  Package,
  User,
  LogIn,
  Shield,
} from "lucide-react";
import type { OrderItem } from "@shared/schema";

function timeAgo(dateStr: string | Date): string {
  const date = typeof dateStr === "string" ? new Date(dateStr) : dateStr;
  const seconds = Math.floor((Date.now() - date.getTime()) / 1000);
  if (seconds < 60) return "just now";
  const minutes = Math.floor(seconds / 60);
  if (minutes < 60) return `${minutes}m ago`;
  const hours = Math.floor(minutes / 60);
  if (hours < 24) return `${hours}h ago`;
  return `${Math.floor(hours / 24)}d ago`;
}

function decodeJwtPayload(token: string): Record<string, any> {
  try {
    const payload = token.split(".")[1];
    return JSON.parse(atob(payload.replace(/-/g, "+").replace(/_/g, "/")));
  } catch {
    return {};
  }
}

function isTeamsMobile(): boolean {
  try {
    return window.parent !== window && /Teams|SkypeSpaces/i.test(navigator.userAgent);
  } catch {
    return false;
  }
}

interface OrdersTabProps {
  userName?: string;
}

export default function OrdersTab({ userName: propUserName = "" }: OrdersTabProps) {
  const qc = useQueryClient();
  const [authState, setAuthState] = useState<"loading" | "unauthenticated" | "authenticated">("loading");
  const [graphToken, setGraphToken] = useState<string | null>(null);
  const [userName, setUserName] = useState(propUserName);
  const [itemText, setItemText] = useState("");
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    initAuth();
  }, []);

  // ─── Auth (mirrors SubmitTab — Teams SSO → MSAL silent → popup/redirect) ──
  async function initAuth() {
    await msalInstance.initialize();

    const redirectResponse = await msalInstance.handleRedirectPromise();
    if (redirectResponse) {
      setUserName(redirectResponse.account?.name || "");
      setGraphToken(redirectResponse.accessToken);
      setAuthState("authenticated");
      return;
    }

    try {
      await microsoftTeams.app.initialize();
      await authenticateViaTeamsSSO();
      return;
    } catch {
      // Not in Teams — fall through to MSAL silent
    }

    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      setUserName(accounts[0].name || "");
      await acquireTokenSilently(accounts[0]);
      return;
    }

    setAuthState("unauthenticated");
  }

  async function authenticateViaTeamsSSO() {
    const teamsToken = await microsoftTeams.authentication.getAuthToken();
    const payload = decodeJwtPayload(teamsToken);
    const loginHint = payload.upn || payload.preferred_username || payload.unique_name || "";
    const displayName = payload.name || "";
    setUserName(displayName);

    const ssoRequest: SsoSilentRequest = { loginHint, scopes: loginRequest.scopes };
    try {
      const graphResp = await msalInstance.ssoSilent(ssoRequest);
      setGraphToken(graphResp.accessToken);
      setUserName(graphResp.account?.name || displayName);
      setAuthState("authenticated");
    } catch (err) {
      if (err instanceof InteractionRequiredAuthError) {
        await signInWithPopupOrRedirect(loginHint);
      } else {
        setAuthState("unauthenticated");
      }
    }
  }

  async function acquireTokenSilently(account: any) {
    try {
      const graphResp = await msalInstance.acquireTokenSilent({ scopes: loginRequest.scopes, account });
      setGraphToken(graphResp.accessToken);
      setAuthState("authenticated");
    } catch (err) {
      if (err instanceof InteractionRequiredAuthError) {
        await signInWithPopupOrRedirect(account.username);
      } else {
        setAuthState("unauthenticated");
      }
    }
  }

  async function signInWithPopupOrRedirect(loginHint?: string) {
    const request = { ...loginRequest, ...(loginHint ? { loginHint } : {}) };
    try {
      if (isTeamsMobile()) {
        await msalInstance.loginRedirect(request);
        return;
      }
      const resp = await msalInstance.loginPopup(request);
      setUserName(resp.account?.name || "");
      await acquireTokenSilently(resp.account!);
    } catch (err) {
      if (
        err instanceof BrowserAuthError &&
        (err.errorCode === "popup_window_error" || err.errorCode === "empty_window_error")
      ) {
        try {
          await msalInstance.loginRedirect(request);
        } catch {
          setAuthState("unauthenticated");
        }
      } else {
        setAuthState("unauthenticated");
      }
    }
  }

  // ─── Server-side admin check via Bearer token (server verifies via Graph /me) ─
  const { data: adminData } = useQuery<{ success: boolean; isAdmin: boolean }>({
    queryKey: ["/api/orders/is-admin", graphToken ? "token" : "none"],
    enabled: authState === "authenticated" && !!graphToken,
    staleTime: 5 * 60 * 1000,
    queryFn: () =>
      fetch("/api/orders/is-admin", {
        headers: { Authorization: `Bearer ${graphToken}` },
      }).then((r) => r.json()),
  });

  const userIsAdmin = adminData?.isAdmin === true;

  // ─── Fetch active items ───────────────────────────────────────────────────
  const { data, isLoading } = useQuery<{ success: boolean; items: OrderItem[] }>({
    queryKey: ["/api/orders"],
    refetchInterval: 15000,
    enabled: authState === "authenticated",
  });

  const items = data?.items ?? [];

  // ─── Add item ─────────────────────────────────────────────────────────────
  const addMutation = useMutation({
    mutationFn: async (payload: { itemName: string }) => {
      const res = await fetch("/api/orders", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${graphToken}`,
        },
        body: JSON.stringify(payload),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Failed to add item");
      return data;
    },
    onMutate: async (payload) => {
      await qc.cancelQueries({ queryKey: ["/api/orders"] });
      const prev = qc.getQueryData<{ success: boolean; items: OrderItem[] }>(["/api/orders"]);
      const optimistic: OrderItem = {
        id: Date.now(),
        itemName: payload.itemName,
        addedBy: userName || "You",
        addedAt: new Date() as any,
        status: "active",
        orderedAt: null,
        orderedBy: null,
      };
      qc.setQueryData<{ success: boolean; items: OrderItem[] }>(["/api/orders"], (old) => ({
        success: true,
        items: [...(old?.items ?? []), optimistic],
      }));
      return { prev };
    },
    onError: (_err, _payload, ctx) => {
      if (ctx?.prev) qc.setQueryData(["/api/orders"], ctx.prev);
    },
    onSettled: () => {
      qc.invalidateQueries({ queryKey: ["/api/orders"] });
    },
  });

  // ─── Mark ordered — Authorization header carries the MSAL graph token;
  //     server verifies identity + role via Graph /me + staff table ─────────
  const orderMutation = useMutation({
    mutationFn: async ({ id }: { id: number }) => {
      const res = await fetch(`/api/orders/${id}`, {
        method: "PATCH",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${graphToken}`,
        },
        body: JSON.stringify({ status: "ordered" }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Failed to update item");
      return data;
    },
    onMutate: async ({ id }) => {
      await qc.cancelQueries({ queryKey: ["/api/orders"] });
      const prev = qc.getQueryData(["/api/orders"]);
      qc.setQueryData<{ success: boolean; items: OrderItem[] }>(["/api/orders"], (old) => ({
        success: true,
        items: (old?.items ?? []).filter((i) => i.id !== id),
      }));
      return { prev };
    },
    onError: (_err, _vars, ctx) => {
      if (ctx?.prev) qc.setQueryData(["/api/orders"], ctx.prev);
    },
    onSettled: () => {
      qc.invalidateQueries({ queryKey: ["/api/orders"] });
    },
  });

  function handleAdd() {
    const text = itemText.trim();
    if (!text) return;
    addMutation.mutate({ itemName: text });
    setItemText("");
    inputRef.current?.focus();
  }

  function handleKeyDown(e: React.KeyboardEvent) {
    if (e.key === "Enter") handleAdd();
  }

  // ─── Loading state ────────────────────────────────────────────────────────
  if (authState === "loading") {
    return (
      <div className="flex items-center justify-center py-16 text-gray-400">
        <Loader2 className="h-5 w-5 animate-spin mr-2" />
        <span className="text-sm">Signing you in…</span>
      </div>
    );
  }

  // ─── Unauthenticated state ────────────────────────────────────────────────
  if (authState === "unauthenticated") {
    return (
      <div className="flex flex-col items-center justify-center min-h-[60vh] px-6 gap-4 text-center">
        <div className="w-14 h-14 rounded-2xl bg-purple-600 flex items-center justify-center shadow-lg">
          <LogIn className="h-7 w-7 text-white" />
        </div>
        <div>
          <h2 className="text-lg font-bold text-gray-900">Sign In Required</h2>
          <p className="text-sm text-gray-500 mt-1">
            Sign in with your Cranfield Glass account to add and manage orders.
          </p>
        </div>
        <Button
          onClick={() => signInWithPopupOrRedirect()}
          className="bg-purple-600 hover:bg-purple-700 text-white w-full max-w-xs"
        >
          <LogIn className="h-4 w-4 mr-2" />
          Sign in with Microsoft
        </Button>
      </div>
    );
  }

  // ─── Authenticated ────────────────────────────────────────────────────────
  return (
    <div className="flex flex-col min-h-0 h-full">
      {/* ── Sticky add bar ── */}
      <div className="sticky top-0 z-10 bg-white border-b border-gray-100 px-4 py-3 flex gap-2">
        <Input
          ref={inputRef}
          placeholder="Add item to order…"
          value={itemText}
          onChange={(e) => setItemText(e.target.value)}
          onKeyDown={handleKeyDown}
          className="flex-1 text-base border-gray-200 focus:border-purple-400"
          autoFocus
        />
        <Button
          onClick={handleAdd}
          disabled={!itemText.trim() || addMutation.isPending}
          className="bg-purple-600 hover:bg-purple-700 active:scale-95 text-white shrink-0"
        >
          {addMutation.isPending ? (
            <Loader2 className="h-4 w-4 animate-spin" />
          ) : (
            <Plus className="h-4 w-4" />
          )}
          <span className="ml-1 hidden sm:inline">Add</span>
        </Button>
      </div>

      {/* ── Admin badge ── */}
      {userIsAdmin && (
        <div className="px-4 pt-2 pb-0">
          <div className="inline-flex items-center gap-1.5 text-xs text-purple-700 bg-purple-50 border border-purple-200 rounded-full px-2.5 py-1">
            <Shield className="h-3 w-3" />
            Admin — you can mark items as ordered
          </div>
        </div>
      )}

      {/* ── Item list ── */}
      <div className="flex-1 overflow-y-auto px-4 py-3 space-y-2 pb-6">
        {isLoading && (
          <div className="flex items-center justify-center py-12 text-gray-400">
            <Loader2 className="h-5 w-5 animate-spin mr-2" />
            <span className="text-sm">Loading…</span>
          </div>
        )}

        {!isLoading && items.length === 0 && (
          <div className="flex flex-col items-center justify-center py-16 text-center gap-3">
            <div className="w-14 h-14 rounded-2xl bg-gray-100 flex items-center justify-center">
              <Package className="h-7 w-7 text-gray-400" />
            </div>
            <p className="text-gray-500 text-sm font-medium">Nothing to order yet</p>
            <p className="text-gray-400 text-xs">Type an item above and tap Add.</p>
          </div>
        )}

        {items.map((item) => (
          <div
            key={item.id}
            className="flex items-center gap-3 bg-white rounded-xl border border-gray-200 px-4 py-3 shadow-sm"
          >
            <ShoppingCart className="h-4 w-4 text-purple-400 shrink-0" />
            <div className="flex-1 min-w-0">
              <p className="text-sm font-medium text-gray-900 truncate">{item.itemName}</p>
              <div className="flex items-center gap-1.5 mt-0.5">
                <User className="h-3 w-3 text-gray-400" />
                <span className="text-xs text-gray-500">{item.addedBy}</span>
                <span className="text-gray-300">·</span>
                <Clock className="h-3 w-3 text-gray-400" />
                <span className="text-xs text-gray-400">{timeAgo(item.addedAt)}</span>
              </div>
            </div>
            {userIsAdmin && (
              <Button
                size="sm"
                variant="outline"
                onClick={() => orderMutation.mutate({ id: item.id })}
                disabled={orderMutation.isPending && orderMutation.variables?.id === item.id}
                className="shrink-0 border-green-200 text-green-700 hover:bg-green-50 hover:border-green-300 active:scale-95 text-xs px-2.5"
              >
                {orderMutation.isPending && orderMutation.variables?.id === item.id ? (
                  <Loader2 className="h-3.5 w-3.5 animate-spin" />
                ) : (
                  <>
                    <CheckCircle2 className="h-3.5 w-3.5 mr-1" />
                    Ordered
                  </>
                )}
              </Button>
            )}
          </div>
        ))}

        {!isLoading && items.length > 0 && (
          <p className="text-center text-xs text-gray-400 pt-2">
            {items.length} item{items.length !== 1 ? "s" : ""} to order
          </p>
        )}
      </div>
    </div>
  );
}

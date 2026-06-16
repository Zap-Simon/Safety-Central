import { useState, useRef, useEffect } from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { useTeamsTheme } from "@/hooks/useTeamsTheme";
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

interface OrdersTabProps {
  userName?: string;
}

export default function OrdersTab({ userName: propUserName = "" }: OrdersTabProps) {
  const qc = useQueryClient();
  const { isDark } = useTeamsTheme();
  const [authState, setAuthState] = useState<"loading" | "unauthenticated" | "authenticated">("loading");
  const [authError, setAuthError] = useState<string>("");
  const [teamsToken, setTeamsToken] = useState<string | null>(null);
  const [userName, setUserName] = useState(propUserName);
  const [itemText, setItemText] = useState("");
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    initAuth();
  }, []);

  // ─── Auth — Teams SSO only ────────────────────────────────────────────────
  // A single SSO token from getAuthToken() is sent to the backend, which
  // exchanges it (on behalf of the user) for the Graph access it needs. No MSAL,
  // popups, or redirects in the browser — so no iframe SSO races.
  async function initAuth() {
    setAuthState("loading");
    setAuthError("");
    try {
      await microsoftTeams.app.initialize();
      const ssoToken = await microsoftTeams.authentication.getAuthToken();
      const payload = decodeJwtPayload(ssoToken);
      setUserName(payload.name || payload.preferred_username || payload.upn || propUserName);
      setTeamsToken(ssoToken);
      setAuthState("authenticated");
    } catch (err: any) {
      const msg = `Teams sign-in failed: ${err?.message || String(err)}`;
      console.error(msg, err);
      setAuthError(msg);
      setAuthState("unauthenticated");
    }
  }

  // ─── Admin check ──────────────────────────────────────────────────────────
  const { data: adminData } = useQuery<{ success: boolean; isAdmin: boolean }>({
    queryKey: ["/api/orders/is-admin", teamsToken ? "token" : "none"],
    enabled: authState === "authenticated" && !!teamsToken,
    staleTime: 5 * 60 * 1000,
    queryFn: () =>
      fetch("/api/orders/is-admin", {
        headers: { Authorization: `Bearer ${teamsToken}` },
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
          Authorization: `Bearer ${teamsToken}`,
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

  // ─── Mark ordered ─────────────────────────────────────────────────────────
  const orderMutation = useMutation({
    mutationFn: async ({ id }: { id: number }) => {
      const res = await fetch(`/api/orders/${id}`, {
        method: "PATCH",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${teamsToken}`,
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

  // ─── Loading ──────────────────────────────────────────────────────────────
  if (authState === "loading") {
    return (
      <div className={`h-full flex items-center justify-center animate-fade-in ${
        isDark ? "bg-gray-900" : "bg-white"
      }`}>
        <div className="text-center">
          <div className="relative w-12 h-12 mx-auto mb-4">
            <div className="absolute inset-0 rounded-full bg-purple-600 flex items-center justify-center">
              <ShoppingCart className="h-6 w-6 text-white" />
            </div>
            <div className="absolute -inset-1 rounded-full border-2 border-purple-600/30 border-t-purple-600 animate-spin" />
          </div>
          <p className={`text-sm ${isDark ? "text-gray-400" : "text-gray-500"}`}>
            Signing you in automatically…
          </p>
        </div>
      </div>
    );
  }

  // ─── Unauthenticated ──────────────────────────────────────────────────────
  if (authState === "unauthenticated") {
    return (
      <div className={`h-full flex items-center justify-center p-4 animate-fade-in ${
        isDark ? "bg-gray-900" : "bg-white"
      }`}>
        <div className="w-full max-w-sm text-center animate-scale-in">
          <div className={`w-14 h-14 rounded-2xl flex items-center justify-center mx-auto mb-4 ${
            isDark ? "bg-purple-900/40" : "bg-purple-100"
          }`}>
            <LogIn className={`h-7 w-7 ${isDark ? "text-purple-400" : "text-purple-500"}`} />
          </div>
          <h1 className={`text-xl font-bold mb-2 ${isDark ? "text-white" : "text-gray-900"}`}>
            Sign in required
          </h1>
          <p className={`text-sm mb-4 ${isDark ? "text-gray-400" : "text-gray-500"}`}>
            Sign in with your Cranfield Glass Microsoft account to add and manage orders.
          </p>
          <Button
            onClick={() => initAuth()}
            className="w-full bg-purple-600 hover:bg-purple-700 text-white mb-3"
          >
            <LogIn className="h-4 w-4 mr-2" />
            Try again
          </Button>
          {authError && (
            <div className="mt-2 p-3 rounded-lg bg-red-50 border border-red-200 text-left">
              <p className="text-xs font-mono text-red-700 break-all">{authError}</p>
            </div>
          )}
        </div>
      </div>
    );
  }

  // ─── Authenticated ────────────────────────────────────────────────────────
  return (
    <div className={`flex flex-col min-h-0 h-full ${isDark ? "bg-gray-900" : "bg-white"}`}>
      {/* ── Quick add bar ── */}
      <div className={`shrink-0 border-b px-4 py-3 flex gap-2 ${
        isDark ? "bg-gray-900 border-gray-800" : "bg-white border-gray-100"
      }`}>
        <Input
          ref={inputRef}
          placeholder="Add item to order…"
          value={itemText}
          onChange={(e) => setItemText(e.target.value)}
          onKeyDown={handleKeyDown}
          className={`flex-1 text-base focus-visible:ring-0 focus-visible:ring-offset-0 ${
            isDark
              ? "bg-gray-800 border-gray-600 text-white placeholder:text-gray-500 focus:border-purple-500"
              : "border-gray-200 focus:border-purple-400"
          }`}
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
          <div className={`inline-flex items-center gap-1.5 text-xs rounded-full px-2.5 py-1 border ${
            isDark
              ? "text-purple-300 bg-purple-900/30 border-purple-700/40"
              : "text-purple-700 bg-purple-50 border-purple-200"
          }`}>
            <Shield className="h-3 w-3" />
            Admin — you can mark items as ordered
          </div>
        </div>
      )}

      {/* ── Item list ── */}
      <div className="flex-1 min-h-0 overflow-y-auto px-4 py-3 space-y-2 pb-6">
        {isLoading && (
          <div className="flex items-center justify-center py-12 text-gray-400 dark:text-gray-500">
            <Loader2 className="h-5 w-5 animate-spin mr-2" />
            <span className="text-sm">Loading…</span>
          </div>
        )}

        {!isLoading && items.length === 0 && (
          <div className="flex flex-col items-center justify-center py-16 text-center gap-3">
            <div className={`w-14 h-14 rounded-2xl flex items-center justify-center ${
              isDark ? "bg-gray-700" : "bg-gray-100"
            }`}>
              <Package className={`h-7 w-7 ${isDark ? "text-gray-500" : "text-gray-400"}`} />
            </div>
            <p className={`text-sm font-medium ${isDark ? "text-gray-400" : "text-gray-500"}`}>
              Nothing to order yet
            </p>
            <p className={`text-xs ${isDark ? "text-gray-500" : "text-gray-400"}`}>
              Add one above — the whole team will see it here.
            </p>
          </div>
        )}

        {items.map((item) => (
          <div
            key={item.id}
            className={`flex items-center gap-3 rounded-xl border px-4 py-3 shadow-sm ${
              isDark ? "bg-gray-800 border-gray-700" : "bg-white border-gray-200"
            }`}
          >
            <ShoppingCart className="h-4 w-4 text-purple-400 shrink-0" />
            <div className="flex-1 min-w-0">
              <p className={`text-sm font-medium truncate ${isDark ? "text-white" : "text-gray-900"}`}>
                {item.itemName}
              </p>
              <div className="flex items-center gap-1.5 mt-0.5">
                <User className={`h-3 w-3 ${isDark ? "text-gray-500" : "text-gray-400"}`} />
                <span className={`text-xs ${isDark ? "text-gray-400" : "text-gray-500"}`}>{item.addedBy}</span>
                <span className={isDark ? "text-gray-600" : "text-gray-300"}>·</span>
                <Clock className={`h-3 w-3 ${isDark ? "text-gray-500" : "text-gray-400"}`} />
                <span className={`text-xs ${isDark ? "text-gray-500" : "text-gray-400"}`}>{timeAgo(item.addedAt)}</span>
              </div>
            </div>
            {userIsAdmin && (
              <Button
                size="sm"
                variant="outline"
                onClick={() => orderMutation.mutate({ id: item.id })}
                disabled={orderMutation.isPending && orderMutation.variables?.id === item.id}
                className={`shrink-0 text-xs px-2.5 active:scale-95 ${
                  isDark
                    ? "border-green-700/50 text-green-400 hover:bg-green-900/30 hover:border-green-600 bg-transparent"
                    : "border-green-200 text-green-700 hover:bg-green-50 hover:border-green-300"
                }`}
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
          <p className={`text-center text-xs pt-2 ${isDark ? "text-gray-500" : "text-gray-400"}`}>
            {items.length} item{items.length !== 1 ? "s" : ""} to order · shared with your team
          </p>
        )}
      </div>
    </div>
  );
}

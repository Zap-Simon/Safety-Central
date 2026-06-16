import { useState, useRef } from "react";
import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import {
  Button,
  Input,
  Card,
  Spinner,
  Text,
  Dialog,
  DialogTrigger,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogContent,
  DialogActions,
  FluentProvider,
  teamsLightTheme,
  teamsDarkTheme,
  teamsHighContrastTheme,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import type { Theme } from "@fluentui/react-components";
import { useTeamsTheme } from "@/hooks/useTeamsTheme";
import { useTeamsAuth } from "@/hooks/useTeamsAuth";
import {
  Cart20Regular,
  Cart16Regular,
  Add20Regular,
  Clock16Regular,
  Person16Regular,
  Delete16Regular,
  Delete20Regular,
} from "@fluentui/react-icons";
import type { OrderItem } from "@shared/schema";
import { TeamsPage, TeamsPinned, TeamsScroll, TeamsCenter, TeamsFullScreen } from "./TeamsPageShell";

// Berry-purple theme for the Orders tab — overrides brand tokens so all
// Fluent primitives (primary buttons, input focus ring, icons, spinner) are
// purple rather than the default Teams blue. Derived from the *active* base
// theme (light/dark/contrast) so dark + high-contrast modes still work — the
// berry palette tokens already carry mode-appropriate values.
function makeBerryTheme(base: Theme): Theme {
  return {
    ...base,
    colorBrandBackground: base.colorPaletteBerryForeground2,
    colorBrandBackgroundHover: base.colorPaletteBerryForeground1,
    colorBrandBackgroundPressed: base.colorPaletteBerryForeground3,
    colorBrandBackgroundSelected: base.colorPaletteBerryForeground2,
    colorBrandForeground1: base.colorPaletteBerryForeground1,
    colorBrandForeground2: base.colorPaletteBerryForeground2,
    colorBrandForegroundLink: base.colorPaletteBerryForeground1,
    colorBrandForegroundLinkHover: base.colorPaletteBerryForeground2,
    colorBrandStroke1: base.colorPaletteBerryBorderActive,
    colorBrandStroke2: base.colorPaletteBerryBorder1,
    colorCompoundBrandBackground: base.colorPaletteBerryForeground2,
    colorCompoundBrandBackgroundHover: base.colorPaletteBerryForeground1,
    colorCompoundBrandBackgroundPressed: base.colorPaletteBerryForeground3,
  };
}

export const berryThemes: Record<string, Theme> = {
  default: makeBerryTheme(teamsLightTheme),
  dark: makeBerryTheme(teamsDarkTheme),
  contrast: makeBerryTheme(teamsHighContrastTheme),
};

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

// Griffel blocks `border*` 4-side shorthands; set each side longhand.
const allBorderColor = (c: string) => ({
  borderTopColor: c,
  borderRightColor: c,
  borderBottomColor: c,
  borderLeftColor: c,
});

const useStyles = makeStyles({
  addBar: {
    display: "flex",
    gap: tokens.spacingHorizontalS,
    paddingLeft: tokens.spacingHorizontalL,
    paddingRight: tokens.spacingHorizontalL,
    paddingTop: tokens.spacingVerticalM,
    paddingBottom: tokens.spacingVerticalM,
    borderBottomWidth: tokens.strokeWidthThin,
    borderBottomStyle: "solid",
    borderBottomColor: tokens.colorNeutralStroke2,
  },
  input: { flexGrow: 1 },
  adminStrip: {
    display: "flex",
    alignItems: "center",
    justifyContent: "flex-end",
    paddingLeft: tokens.spacingHorizontalL,
    paddingRight: tokens.spacingHorizontalL,
    paddingTop: tokens.spacingVerticalXS,
    paddingBottom: tokens.spacingVerticalXS,
    backgroundColor: tokens.colorPaletteRedBackground1,
    borderBottomWidth: tokens.strokeWidthThin,
    borderBottomStyle: "solid",
    borderBottomColor: tokens.colorPaletteRedBorder1,
  },
  clearBtn: {
    color: tokens.colorPaletteRedForeground1,
    ...allBorderColor(tokens.colorPaletteRedBorder1),
  },
  list: {
    paddingLeft: tokens.spacingHorizontalL,
    paddingRight: tokens.spacingHorizontalL,
    paddingTop: tokens.spacingVerticalM,
    paddingBottom: tokens.spacingVerticalXL,
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
  },
  loadingRow: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    paddingTop: tokens.spacingVerticalXXXL,
    paddingBottom: tokens.spacingVerticalXXXL,
  },
  itemCard: {
    display: "flex",
    flexDirection: "row",
    alignItems: "center",
    gap: tokens.spacingHorizontalM,
    padding: tokens.spacingHorizontalM,
  },
  itemIcon: { color: tokens.colorBrandForeground1, flexShrink: 0 },
  itemBody: { flexGrow: 1, minWidth: 0 },
  metaRow: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalXS,
    marginTop: "2px",
    color: tokens.colorNeutralForeground3,
  },
  metaIcon: { width: "12px", height: "12px", flexShrink: 0 },
  dot: { color: tokens.colorNeutralForeground4 },
  footer: { paddingTop: tokens.spacingVerticalS },
  removeBtn: { color: tokens.colorPaletteRedForeground1, flexShrink: 0 },
});

export default function OrdersTab() {
  const styles = useStyles();
  const { theme } = useTeamsTheme();
  const berryTheme = berryThemes[theme] ?? berryThemes.default;
  const qc = useQueryClient();
  // Auth is shared across both tabs (see TeamsAuthProvider) so switching tabs
  // never re-triggers the "Signing you in…" loader.
  const { authState, teamsToken, userName, authError, retry } = useTeamsAuth();
  const [itemText, setItemText] = useState("");
  const inputRef = useRef<HTMLInputElement>(null);

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

  // ─── Remove a single item (admin) ─────────────────────────────────────────
  const removeMutation = useMutation({
    mutationFn: async ({ id }: { id: number }) => {
      const res = await fetch(`/api/orders/${id}`, {
        method: "DELETE",
        headers: { Authorization: `Bearer ${teamsToken}` },
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Failed to remove item");
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

  // ─── Clear the whole list (admin) ─────────────────────────────────────────
  const clearMutation = useMutation({
    mutationFn: async () => {
      const res = await fetch("/api/orders/clear", {
        method: "POST",
        headers: { Authorization: `Bearer ${teamsToken}` },
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Failed to clear list");
      return data;
    },
    onMutate: async () => {
      await qc.cancelQueries({ queryKey: ["/api/orders"] });
      const prev = qc.getQueryData(["/api/orders"]);
      qc.setQueryData<{ success: boolean; items: OrderItem[] }>(["/api/orders"], {
        success: true,
        items: [],
      });
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
      <FluentProvider theme={berryTheme} style={{ display: "contents" }}>
        <TeamsCenter className="animate-fade-in">
          <Spinner size="large" label="Signing you in automatically…" />
        </TeamsCenter>
      </FluentProvider>
    );
  }

  // ─── Unauthenticated ──────────────────────────────────────────────────────
  if (authState === "unauthenticated") {
    return (
      <FluentProvider theme={berryTheme} style={{ display: "contents" }}>
        <TeamsFullScreen
          icon={<Cart20Regular />}
          title="Sign in required"
          description="Sign in with your Cranfield Glass Microsoft account to add and manage orders."
          actionLabel="Try again"
          onAction={retry}
          error={authError || undefined}
          accent={{
            bg: tokens.colorPaletteBerryBackground1,
            fg: tokens.colorPaletteBerryForeground1,
          }}
        />
      </FluentProvider>
    );
  }

  // ─── Authenticated ────────────────────────────────────────────────────────
  return (
    <FluentProvider theme={berryTheme} style={{ display: "contents" }}>
    <TeamsPage>
      <TeamsPinned>
        {/* ── Admin clear strip — very top, red tint, only when admin + items exist ── */}
        {userIsAdmin && items.length > 0 && (
          <div className={styles.adminStrip}>
            <Dialog>
              <DialogTrigger disableButtonEnhancement>
                <Button
                  size="small"
                  appearance="outline"
                  className={styles.clearBtn}
                  icon={clearMutation.isPending ? <Spinner size="tiny" /> : <Delete20Regular />}
                  disabled={clearMutation.isPending}
                >
                  Clear list
                </Button>
              </DialogTrigger>
              <DialogSurface>
                <DialogBody>
                  <DialogTitle>Clear the whole order list?</DialogTitle>
                  <DialogContent>
                    This removes all {items.length} item{items.length !== 1 ? "s" : ""} for
                    everyone on the team. Items are archived so nothing is permanently lost.
                  </DialogContent>
                  <DialogActions>
                    <DialogTrigger disableButtonEnhancement>
                      <Button appearance="secondary">Cancel</Button>
                    </DialogTrigger>
                    <DialogTrigger disableButtonEnhancement>
                      <Button appearance="primary" onClick={() => clearMutation.mutate()}>
                        Clear list
                      </Button>
                    </DialogTrigger>
                  </DialogActions>
                </DialogBody>
              </DialogSurface>
            </Dialog>
          </div>
        )}

        {/* ── Quick-add bar (keyboard-safe — never scrolls away) ── */}
        <div className={styles.addBar}>
          <Input
            className={styles.input}
            input={{ ref: inputRef }}
            placeholder="Add item to order…"
            value={itemText}
            onChange={(_, d) => setItemText(d.value)}
            onKeyDown={handleKeyDown}
            size="large"
          />
          <Button
            appearance="primary"
            size="large"
            onClick={handleAdd}
            disabled={!itemText.trim() || addMutation.isPending}
            icon={addMutation.isPending ? <Spinner size="tiny" /> : <Add20Regular />}
          >
            Add
          </Button>
        </div>
      </TeamsPinned>

      {/* ── The single scroll region: the shared item list ── */}
      <TeamsScroll className={styles.list}>
        {isLoading && (
          <div className={styles.loadingRow}>
            <Spinner size="small" label="Loading…" />
          </div>
        )}

        {items.map((item) => (
          <Card key={item.id} className={styles.itemCard}>
            <Cart16Regular className={styles.itemIcon} />
            <div className={styles.itemBody}>
              <Text size={300} weight="semibold" truncate wrap={false} block>
                {item.itemName}
              </Text>
              <div className={styles.metaRow}>
                <Person16Regular className={styles.metaIcon} />
                <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
                  {item.addedBy}
                </Text>
                <span className={styles.dot}>·</span>
                <Clock16Regular className={styles.metaIcon} />
                <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
                  {timeAgo(item.addedAt)}
                </Text>
              </div>
            </div>
            {userIsAdmin && (
              <Button
                size="small"
                appearance="subtle"
                className={styles.removeBtn}
                aria-label={`Remove ${item.itemName}`}
                title="Remove item"
                onClick={() => removeMutation.mutate({ id: item.id })}
                disabled={removeMutation.isPending && removeMutation.variables?.id === item.id}
                icon={
                  removeMutation.isPending && removeMutation.variables?.id === item.id ? (
                    <Spinner size="tiny" />
                  ) : (
                    <Delete16Regular />
                  )
                }
              />
            )}
          </Card>
        ))}

        {!isLoading && items.length > 0 && (
          <Text size={200} align="center" className={styles.footer} style={{ color: tokens.colorNeutralForeground3 }}>
            {items.length} item{items.length !== 1 ? "s" : ""} to order · shared with your team
          </Text>
        )}
      </TeamsScroll>
    </TeamsPage>
    </FluentProvider>
  );
}

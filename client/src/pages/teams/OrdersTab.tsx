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
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { useTeamsAuth } from "@/hooks/useTeamsAuth";
import {
  Cart20Regular,
  Cart16Regular,
  Add20Regular,
  Clock16Regular,
  Person16Regular,
  Delete20Regular,
} from "@fluentui/react-icons";
import { motion } from "framer-motion";
import type { OrderItem } from "@shared/schema";
import { TeamsPage, TeamsPinned, TeamsScroll, TeamsCenter, TeamsFullScreen } from "./TeamsPageShell";

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
  clearRow: {
    display: "flex",
    justifyContent: "flex-end",
  },
  // Keep the confirm dialog comfortably inset from the screen edges on mobile.
  dialogSurface: {
    maxWidth: "min(420px, calc(100vw - 48px))",
  },
  clearBtn: {
    color: tokens.colorNeutralForeground3,
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
  swipeWrap: {
    position: "relative",
    borderTopLeftRadius: tokens.borderRadiusMedium,
    borderTopRightRadius: tokens.borderRadiusMedium,
    borderBottomLeftRadius: tokens.borderRadiusMedium,
    borderBottomRightRadius: tokens.borderRadiusMedium,
    overflow: "hidden",
  },
  deleteLayer: {
    position: "absolute",
    top: 0,
    right: 0,
    bottom: 0,
    left: 0,
    display: "flex",
    alignItems: "center",
    justifyContent: "flex-end",
    paddingRight: tokens.spacingHorizontalXL,
    backgroundColor: tokens.colorPaletteRedBackground3,
    color: tokens.colorNeutralForegroundOnBrand,
    borderTopLeftRadius: tokens.borderRadiusMedium,
    borderTopRightRadius: tokens.borderRadiusMedium,
    borderBottomLeftRadius: tokens.borderRadiusMedium,
    borderBottomRightRadius: tokens.borderRadiusMedium,
  },
  swipeFront: {
    position: "relative",
    touchAction: "pan-y",
    cursor: "grab",
  },
});

export default function OrdersTab() {
  const styles = useStyles();
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
      <TeamsCenter className="animate-fade-in">
        <Spinner size="large" label="Signing you in automatically…" />
      </TeamsCenter>
    );
  }

  // ─── Unauthenticated ──────────────────────────────────────────────────────
  if (authState === "unauthenticated") {
    return (
      <TeamsFullScreen
        icon={<Cart20Regular />}
        title="Sign in required"
        description="Sign in with your Cranfield Glass Microsoft account to add and manage orders."
        actionLabel="Try again"
        onAction={retry}
        error={authError || undefined}
        accent={{
          bg: tokens.colorBrandBackground2,
          fg: tokens.colorBrandForeground1,
        }}
      />
    );
  }

  // ─── Authenticated ────────────────────────────────────────────────────────
  return (
    <TeamsPage>
      <TeamsPinned>
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
        {/* ── Admin clear — minimal, under the input, before the first item ── */}
        {userIsAdmin && items.length > 0 && (
          <div className={styles.clearRow}>
            <Dialog>
              <DialogTrigger disableButtonEnhancement>
                <Button
                  size="small"
                  appearance="subtle"
                  className={styles.clearBtn}
                  icon={clearMutation.isPending ? <Spinner size="tiny" /> : <Delete20Regular />}
                  disabled={clearMutation.isPending}
                >
                  Clear list
                </Button>
              </DialogTrigger>
              <DialogSurface className={styles.dialogSurface}>
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

        {isLoading && (
          <div className={styles.loadingRow}>
            <Spinner size="small" label="Loading…" />
          </div>
        )}

        {items.map((item) => {
          const deleting =
            removeMutation.isPending && removeMutation.variables?.id === item.id;
          const cardInner = (
            <Card appearance="outline" className={styles.itemCard}>
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
            </Card>
          );

          if (!userIsAdmin) {
            return <div key={item.id}>{cardInner}</div>;
          }

          return (
            <div key={item.id} className={styles.swipeWrap}>
              <div className={styles.deleteLayer} aria-hidden>
                {deleting ? <Spinner size="tiny" /> : <Delete20Regular />}
              </div>
              <motion.div
                className={styles.swipeFront}
                drag="x"
                dragDirectionLock
                dragConstraints={{ left: -96, right: 0 }}
                dragElastic={{ left: 0.15, right: 0 }}
                dragSnapToOrigin
                onDragEnd={(_, info) => {
                  if (!deleting && (info.offset.x < -72 || info.velocity.x < -500)) {
                    removeMutation.mutate({ id: item.id });
                  }
                }}
              >
                {cardInner}
              </motion.div>
            </div>
          );
        })}

        {!isLoading && items.length > 0 && (
          <Text size={200} align="center" className={styles.footer} style={{ color: tokens.colorNeutralForeground3 }}>
            {items.length} item{items.length !== 1 ? "s" : ""} to order · shared with your team
          </Text>
        )}
      </TeamsScroll>
    </TeamsPage>
  );
}

import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import {
  Button,
  Card,
  Spinner,
  Text,
  Badge,
  Accordion,
  AccordionItem,
  AccordionHeader,
  AccordionPanel,
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
  Cart16Regular,
  Clock16Regular,
  Person16Regular,
  Delete20Regular,
} from "@fluentui/react-icons";
import { motion } from "framer-motion";
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

const useStyles = makeStyles({
  wrap: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
  },
  headerRow: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalS,
  },
  countBadge: { flexShrink: 0 },
  panel: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
    paddingTop: tokens.spacingVerticalS,
  },
  clearRow: { display: "flex", justifyContent: "flex-end" },
  clearBtn: { color: tokens.colorNeutralForeground3 },
  loadingRow: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    paddingTop: tokens.spacingVerticalL,
    paddingBottom: tokens.spacingVerticalL,
  },
  empty: {
    paddingTop: tokens.spacingVerticalM,
    paddingBottom: tokens.spacingVerticalM,
    color: tokens.colorNeutralForeground3,
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
    backgroundColor: tokens.colorNeutralBackground1,
    borderTopLeftRadius: tokens.borderRadiusMedium,
    borderTopRightRadius: tokens.borderRadiusMedium,
    borderBottomLeftRadius: tokens.borderRadiusMedium,
    borderBottomRightRadius: tokens.borderRadiusMedium,
  },
});

/**
 * The shared team order list, surfaced beneath the Submit box. It only displays
 * items and admin controls (clear-all + swipe-to-delete). Items are added via
 * the AI order-confirm flow in SubmitTab, not a manual quick-add bar.
 */
export default function OrdersList() {
  const styles = useStyles();
  const qc = useQueryClient();
  const { authState, teamsToken, getToken } = useTeamsAuth();

  const { data: adminData } = useQuery<{ success: boolean; isAdmin: boolean }>({
    queryKey: ["/api/orders/is-admin", teamsToken ? "token" : "none"],
    enabled: authState === "authenticated" && !!teamsToken,
    staleTime: 5 * 60 * 1000,
    queryFn: async () => {
      const token = await getToken();
      return fetch("/api/orders/is-admin", {
        headers: { Authorization: `Bearer ${token}` },
      }).then((r) => r.json());
    },
  });

  const userIsAdmin = adminData?.isAdmin === true;

  const { data, isLoading } = useQuery<{ success: boolean; items: OrderItem[] }>({
    queryKey: ["/api/orders"],
    refetchInterval: 15000,
    enabled: authState === "authenticated",
  });

  const items = data?.items ?? [];

  const removeMutation = useMutation({
    mutationFn: async ({ id }: { id: number }) => {
      const token = await getToken();
      const res = await fetch(`/api/orders/${id}`, {
        method: "DELETE",
        headers: { Authorization: `Bearer ${token}` },
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

  const clearMutation = useMutation({
    mutationFn: async () => {
      const token = await getToken();
      const res = await fetch("/api/orders/clear", {
        method: "POST",
        headers: { Authorization: `Bearer ${token}` },
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

  return (
    <Accordion collapsible className={styles.wrap}>
      <AccordionItem value="orders">
        <AccordionHeader icon={<Cart16Regular />} expandIconPosition="end">
          <span className={styles.headerRow}>
            <Text size={300} weight="semibold">
              Order list
            </Text>
            {items.length > 0 && (
              <Badge
                appearance="filled"
                color="brand"
                shape="rounded"
                size="medium"
                className={styles.countBadge}
              >
                {items.length}
              </Badge>
            )}
          </span>
        </AccordionHeader>
        <AccordionPanel>
          <div className={styles.panel}>
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
                  <DialogSurface style={{ maxWidth: "min(400px, calc(100vw - 48px))" }}>
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

            {!isLoading && items.length === 0 && (
              <Text size={200} align="center" className={styles.empty} block>
                No items to order yet. When you describe something to restock, we'll offer to add it here.
              </Text>
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
          </div>
        </AccordionPanel>
      </AccordionItem>
    </Accordion>
  );
}

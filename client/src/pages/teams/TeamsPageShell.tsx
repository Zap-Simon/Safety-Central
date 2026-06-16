import { useEffect, type ReactNode, type RefObject } from "react";
import {
  Button,
  Text,
  MessageBar,
  MessageBarBody,
  makeStyles,
  mergeClasses,
  tokens,
} from "@fluentui/react-components";
import { ArrowCounterclockwise20Regular } from "@fluentui/react-icons";

/**
 * Teams personal-tab page shell.
 *
 * ONE scrolling strategy for every tab so keyboard/focus behaviour stays
 * identical everywhere. The contract:
 *
 *   <TeamsPage>                         // fixed-height flex column (no page scroll)
 *     <TeamsPinned>…</TeamsPinned>      // 0..n pinned regions — NEVER scroll
 *     <TeamsScroll>…</TeamsScroll>      // exactly ONE scroll region (optional)
 *   </TeamsPage>
 *
 * Keyboard safety rule: the input the user is most likely to focus (a search
 * box, the primary textarea) must live in a `TeamsPinned`, NOT inside
 * `TeamsScroll`. A focused input inside a shrinkable+scrollable region gets
 * auto-scrolled into view when the on-screen keyboard opens, which shifts the
 * whole page. Pinned regions don't shrink, so they stay put. Pair this with
 * `useKeyboardSafeFocus` for any input that is auto-focused on mount.
 */

const useShellStyles = makeStyles({
  page: {
    display: "flex",
    flexDirection: "column",
    height: "100%",
    minHeight: 0,
    overflow: "hidden",
  },
  pinned: { flexShrink: 0 },
  scroll: { flexGrow: 1, minHeight: 0, overflowY: "auto" },
  center: {
    height: "100%",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    paddingLeft: tokens.spacingHorizontalXXL,
    paddingRight: tokens.spacingHorizontalXXL,
    boxSizing: "border-box",
  },
  fsInner: { width: "100%", maxWidth: "340px", textAlign: "center" },
  fsChip: {
    width: "56px",
    height: "56px",
    borderTopLeftRadius: tokens.borderRadiusXLarge,
    borderTopRightRadius: tokens.borderRadiusXLarge,
    borderBottomLeftRadius: tokens.borderRadiusXLarge,
    borderBottomRightRadius: tokens.borderRadiusXLarge,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    marginLeft: "auto",
    marginRight: "auto",
    marginBottom: tokens.spacingVerticalL,
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground1,
  },
  fsDesc: {
    color: tokens.colorNeutralForeground3,
    marginTop: tokens.spacingVerticalS,
    marginBottom: tokens.spacingVerticalL,
  },
  fsBtn: { width: "100%" },
  fsErr: { marginTop: tokens.spacingVerticalM, textAlign: "left" },
});

interface SlotProps {
  children: ReactNode;
  className?: string;
}

export function TeamsPage({ children, className }: SlotProps) {
  const s = useShellStyles();
  return <div className={mergeClasses(s.page, className)}>{children}</div>;
}

export function TeamsPinned({ children, className }: SlotProps) {
  const s = useShellStyles();
  return <div className={mergeClasses(s.pinned, className)}>{children}</div>;
}

export function TeamsScroll({ children, className }: SlotProps) {
  const s = useShellStyles();
  return <div className={mergeClasses(s.scroll, className)}>{children}</div>;
}

export function TeamsCenter({ children, className }: SlotProps) {
  const s = useShellStyles();
  return <div className={mergeClasses(s.center, className)}>{children}</div>;
}

/**
 * Auto-focus an input without the page jumping. `focus({ preventScroll: true })`
 * suppresses the scroll-into-view that mobile webviews do for a lower-positioned
 * input. Gate `active` on the state where focus should happen (e.g. authenticated
 * + correct step).
 *
 * CAUTION: tab components remount on every Teams tab switch, so don't gate this on
 * a default step — it will re-fire and pop the mobile keyboard each switch. Only use
 * it for focus driven by a deliberate in-tab transition, never plain mount.
 */
export function useKeyboardSafeFocus(
  ref: RefObject<HTMLElement | null>,
  active: boolean,
) {
  useEffect(() => {
    if (active) ref.current?.focus({ preventScroll: true });
  }, [active, ref]);
}

interface TeamsFullScreenProps {
  icon: ReactNode;
  title: string;
  description?: string;
  actionLabel?: string;
  onAction?: () => void;
  error?: string;
  /** Accent for the icon chip — defaults to the Teams brand tint. */
  accent?: { bg: string; fg: string };
}

/**
 * Shared full-screen state (sign-in required / fatal error) so both tabs read as
 * one app. Only the accent colour differs per tab.
 */
export function TeamsFullScreen({
  icon,
  title,
  description,
  actionLabel,
  onAction,
  error,
  accent,
}: TeamsFullScreenProps) {
  const s = useShellStyles();
  return (
    <TeamsCenter className="animate-fade-in">
      <div className={`${s.fsInner} animate-scale-in`}>
        <div
          className={s.fsChip}
          style={accent ? { backgroundColor: accent.bg, color: accent.fg } : undefined}
        >
          {icon}
        </div>
        <Text as="h1" size={500} weight="bold" block>
          {title}
        </Text>
        {description && (
          <Text size={300} block className={s.fsDesc}>
            {description}
          </Text>
        )}
        {actionLabel && (
          <Button
            appearance="primary"
            size="large"
            className={s.fsBtn}
            icon={<ArrowCounterclockwise20Regular />}
            onClick={onAction}
          >
            {actionLabel}
          </Button>
        )}
        {error && (
          <MessageBar intent="error" className={s.fsErr}>
            <MessageBarBody>{error}</MessageBarBody>
          </MessageBar>
        )}
      </div>
    </TeamsCenter>
  );
}

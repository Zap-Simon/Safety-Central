import { ReactNode } from "react";
import { Card, Text, makeStyles, tokens } from "@fluentui/react-components";
import { ChevronRight20Regular } from "@fluentui/react-icons";

// Shared presentational pieces for the Meetings tab lists. Every list (rostered
// open/attended/minutes and the non-rostered minutes list) renders through these
// so the card and section-header markup lives in one place instead of being
// copy-pasted, keeping SignTab.tsx from growing as the tab gains more lists.
const useStyles = makeStyles({
  sectionHeader: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalS,
    paddingLeft: tokens.spacingHorizontalXS,
    paddingBottom: tokens.spacingVerticalXXS,
  },
  card: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalM,
    padding: tokens.spacingHorizontalL,
  },
  clickable: {
    cursor: "pointer",
    transitionProperty: "background-color, transform",
    transitionDuration: tokens.durationFaster,
    ":hover": { backgroundColor: tokens.colorNeutralBackground1Hover },
    ":active": { transform: "scale(0.99)" },
  },
  icon: {
    width: "40px",
    height: "40px",
    flexShrink: 0,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    borderRadius: tokens.borderRadiusCircular,
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground1,
  },
  body: { flexGrow: 1, minWidth: 0, display: "flex", flexDirection: "column", gap: "2px" },
  chevron: { flexShrink: 0, color: tokens.colorNeutralForeground4 },
});

export function SectionHeader({ icon, label }: { icon: ReactNode; label: string }) {
  const styles = useStyles();
  return (
    <div className={styles.sectionHeader}>
      {icon}
      <Text size={200} weight="semibold" style={{ color: tokens.colorNeutralForeground3, textTransform: "uppercase", letterSpacing: "0.04em" }}>
        {label}
      </Text>
    </div>
  );
}

export function MeetingCard({
  icon,
  title,
  subtitle,
  onClick,
}: {
  icon: ReactNode;
  title: string;
  // A string is rendered as muted helper text; pass a node (e.g. a Badge) for
  // status chips so each list keeps its own trailing affordance.
  subtitle?: ReactNode;
  onClick?: () => void;
}) {
  const styles = useStyles();
  return (
    <Card
      className={`${styles.card}${onClick ? ` ${styles.clickable}` : ""} animate-fade-in-up`}
      onClick={onClick}
    >
      <div className={styles.icon}>{icon}</div>
      <div className={styles.body}>
        <Text size={300} weight="semibold" truncate block>
          {title}
        </Text>
        {typeof subtitle === "string" ? (
          <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
            {subtitle}
          </Text>
        ) : (
          subtitle
        )}
      </div>
      {onClick && <ChevronRight20Regular className={styles.chevron} />}
    </Card>
  );
}

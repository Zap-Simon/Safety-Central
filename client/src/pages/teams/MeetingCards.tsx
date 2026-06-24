import { ReactNode } from "react";
import { Card, Text, makeStyles, tokens } from "@fluentui/react-components";
import { ArrowRight20Regular, ChevronRight20Regular } from "@fluentui/react-icons";

// Shared presentational pieces for the Meetings tab lists. Every list (rostered
// open/attended/minutes and the non-rostered minutes list) renders through these
// so the card and section-header markup lives in one place instead of being
// copy-pasted, keeping SignTab.tsx from growing as the tab gains more lists.
//
// The tab is built around three visual tiers so importance reads at a glance:
//   1. HeroCard  — the single most important thing right now (sign / agenda)
//   2. MeetingCard (brand/success icon) — signing & attendance rows
//   3. MeetingCard (neutral icon)       — the quieter minutes archive
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
  },
  iconBrand: {
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground1,
  },
  iconSuccess: {
    backgroundColor: tokens.colorPaletteGreenBackground2,
    color: tokens.colorPaletteGreenForeground1,
  },
  iconNeutral: {
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground3,
  },
  body: { flexGrow: 1, minWidth: 0, display: "flex", flexDirection: "column", gap: "2px" },
  trailing: {
    flexShrink: 0,
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalS,
  },
  chevron: { flexShrink: 0, color: tokens.colorNeutralForeground4 },

  // ─── Hero (tier 1) ─────────────────────────────────────────────────────────
  hero: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
    paddingTop: tokens.spacingVerticalL,
    paddingBottom: tokens.spacingVerticalL,
    paddingLeft: tokens.spacingHorizontalL,
    paddingRight: tokens.spacingHorizontalL,
    borderRadius: tokens.borderRadiusXLarge,
    cursor: "pointer",
    transitionProperty: "background-color, transform, box-shadow",
    transitionDuration: tokens.durationFaster,
    ":active": { transform: "scale(0.99)" },
  },
  // Solid brand — reserved for the #1 action (signing). It should out-shout
  // everything else on the page.
  heroSolid: {
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
    boxShadow: tokens.shadow8,
    ":hover": { backgroundColor: tokens.colorBrandBackgroundHover },
  },
  // Brand tint — the agenda hero (boss priority) when nothing needs signing, or
  // the secondary hero when a signing action is already taking the solid slot.
  heroTint: {
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorNeutralForeground1,
    ":hover": { backgroundColor: tokens.colorBrandBackground2Hover },
  },
  heroTop: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalM,
  },
  heroChip: {
    width: "48px",
    height: "48px",
    flexShrink: 0,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    borderRadius: tokens.borderRadiusCircular,
  },
  heroChipSolid: {
    backgroundColor: "rgba(255, 255, 255, 0.22)",
    color: tokens.colorNeutralForegroundOnBrand,
  },
  heroChipTint: {
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
  },
  heroHeadings: { minWidth: 0, display: "flex", flexDirection: "column", gap: "2px" },
  heroEyebrow: { textTransform: "uppercase", letterSpacing: "0.06em" },
  // Footer action affordance. We render a non-interactive pill (the whole card is
  // the click target) so there is no nested button double-firing the handler.
  heroAction: {
    alignSelf: "flex-start",
    display: "inline-flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalXS,
    paddingTop: tokens.spacingVerticalS,
    paddingBottom: tokens.spacingVerticalS,
    paddingLeft: tokens.spacingHorizontalL,
    paddingRight: tokens.spacingHorizontalM,
    borderRadius: tokens.borderRadiusCircular,
    fontWeight: tokens.fontWeightSemibold,
  },
  heroActionSolid: {
    backgroundColor: tokens.colorNeutralForegroundOnBrand,
    color: tokens.colorBrandForeground1,
  },
  // On the tint hero the action reads as a link rather than a filled pill, so the
  // solid signing hero stays visually dominant.
  heroActionTint: {
    paddingLeft: 0,
    paddingRight: 0,
    color: tokens.colorBrandForeground1,
  },
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
  iconTone = "brand",
  title,
  subtitle,
  trailing,
  onClick,
}: {
  icon: ReactNode;
  // Tints the round icon chip so a row's purpose reads instantly: brand = needs
  // attention / signable, success = attended/signed, neutral = read-only archive.
  iconTone?: "brand" | "success" | "neutral";
  title: string;
  // A string is rendered as muted helper text; pass a node for richer content.
  subtitle?: ReactNode;
  // Optional trailing slot (e.g. a status Badge) shown before the chevron.
  trailing?: ReactNode;
  onClick?: () => void;
}) {
  const styles = useStyles();
  const toneClass =
    iconTone === "success" ? styles.iconSuccess : iconTone === "neutral" ? styles.iconNeutral : styles.iconBrand;
  return (
    <Card
      className={`${styles.card}${onClick ? ` ${styles.clickable}` : ""} animate-fade-in-up`}
      onClick={onClick}
    >
      <div className={`${styles.icon} ${toneClass}`}>{icon}</div>
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
      {trailing && <div className={styles.trailing}>{trailing}</div>}
      {onClick && <ChevronRight20Regular className={styles.chevron} />}
    </Card>
  );
}

/**
 * Tier-1 feature card. Two tones create the page's top-of-hierarchy:
 *  - "solid" — solid brand fill, used for the single most important action
 *    (signing your attendance). It deliberately out-weighs everything below.
 *  - "tint"  — brand tint, used for the next-meeting agenda hero.
 * The whole card is the click target; the footer pill/link is a visual cue.
 */
export function HeroCard({
  tone,
  icon,
  eyebrow,
  title,
  subtitle,
  actionLabel,
  onClick,
}: {
  tone: "solid" | "tint";
  icon: ReactNode;
  eyebrow: string;
  title: string;
  subtitle?: string;
  actionLabel: string;
  onClick: () => void;
}) {
  const styles = useStyles();
  const isSolid = tone === "solid";
  const mutedColor = isSolid ? tokens.colorNeutralForegroundOnBrand : tokens.colorNeutralForeground3;
  return (
    <Card
      className={`${styles.hero} ${isSolid ? styles.heroSolid : styles.heroTint} animate-fade-in-up`}
      onClick={onClick}
    >
      <div className={styles.heroTop}>
        <div className={`${styles.heroChip} ${isSolid ? styles.heroChipSolid : styles.heroChipTint}`}>
          {icon}
        </div>
        <div className={styles.heroHeadings}>
          <Text
            size={200}
            weight="semibold"
            className={styles.heroEyebrow}
            style={{ color: isSolid ? tokens.colorNeutralForegroundOnBrand : tokens.colorBrandForeground1, opacity: isSolid ? 0.85 : 1 }}
          >
            {eyebrow}
          </Text>
          <Text size={500} weight="bold" truncate block style={{ color: "inherit" }}>
            {title}
          </Text>
          {subtitle && (
            <Text size={200} block style={{ color: mutedColor, opacity: isSolid ? 0.85 : 1 }}>
              {subtitle}
            </Text>
          )}
        </div>
      </div>
      <span className={`${styles.heroAction} ${isSolid ? styles.heroActionSolid : styles.heroActionTint}`}>
        <Text size={300} weight="semibold" style={{ color: "inherit" }}>
          {actionLabel}
        </Text>
        <ArrowRight20Regular />
      </span>
    </Card>
  );
}

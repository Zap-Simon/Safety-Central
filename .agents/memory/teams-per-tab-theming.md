---
name: Teams per-tab brand theming (Fluent v9)
description: How to recolour ONE Fluent tab/section away from the global brand without breaking dark/high-contrast. NOTE — no longer applied; both Teams tabs now share one blue brand.
---

# CURRENT STATE (durable)

The project deliberately uses ONE shared default-Teams-blue brand for BOTH tabs
(Submit + Orders). The earlier purple "berry" Orders accent was removed at the
user's request — they wanted blue throughout. App.tsx applies a single
`fluentTheme` (light/dark/contrast) at the shell; OrdersTab has NO per-tab or
per-button FluentProvider wrappers. Do not re-introduce a second brand unless the
user explicitly asks. The technique below is kept only as reference if they do.

# Per-tab brand colour in Fluent v9 (reference only — not currently used)

To give one tab/section a different brand accent than the rest of the app:

- **Tab indicator + hover/press flash**: the selected underline lives on an inner
  `span.fui-Tab__tabIndicator::before`, NOT the `<Tab>` root. A Griffel `::after`
  override on the Tab root silently misses it. Instead set the CSS custom property
  directly on the `<Tab>` element via `style` — e.g.
  `style={{ "--colorBrandForeground1": base.colorPaletteBerryForeground1 } as React.CSSProperties}`.
  CSS vars cascade into the inner span and its pseudo-element, recolouring indicator
  and interaction states in one shot.

- **Section content (buttons, icons, input focus, spinner)**: wrap the section in a
  nested `<FluentProvider theme={...} style={{ display: "contents" }}>`. `display:contents`
  adds no layout box, so it won't disturb a flex/scroll shell contract.

**Why:** brand-coloured Fluent primitives all read CSS vars set by the nearest
FluentProvider; overriding the brand tokens once recolours everything without
per-component style props.

**How to apply — critical:** derive the nested theme from the *active* base theme
(light/dark/contrast), never hard-code `teamsLightTheme`. Hard-coding light forces
light mode for users running Teams dark/high-contrast (real a11y regression). Build a
`makeBerryTheme(base: Theme)` factory and pick the base from `useTeamsTheme()`'s
`theme` value. The `colorPaletteX*` tokens already carry mode-appropriate values per
base theme, so overriding `colorBrand*` with `base.colorPaletteX*` stays correct in
every mode.

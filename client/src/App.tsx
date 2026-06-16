import { Switch, Route, useLocation, Link } from "wouter";
import { queryClient } from "./lib/queryClient";
import { QueryClientProvider } from "@tanstack/react-query";
import { Toaster } from "@/components/ui/toaster";
import { TooltipProvider } from "@/components/ui/tooltip";
import AuthWrapper from "@/components/auth-wrapper";
import Home from "@/pages/home";
import MeetingHistory from "@/pages/meeting-history";
import Actions from "@/pages/actions";
import StaffTraining from "@/pages/staff-training";
import EquipmentMaintenance from "@/pages/equipment-maintenance";
import TeamsIntegrationPlan from "@/pages/teams-integration-plan";
import HealthSafetyPolicy from "@/pages/health-safety-policy";
import EnvironmentPolicy from "@/pages/environment-policy";
import QualityPolicy from "@/pages/quality-policy";
import SubmitTab from "@/pages/teams/SubmitTab";
import OrdersTab from "@/pages/teams/OrdersTab";
import NotFound from "@/pages/not-found";
import { TeamsThemeProvider, useTeamsTheme } from "@/hooks/useTeamsTheme";
import {
  FluentProvider,
  teamsLightTheme,
  teamsDarkTheme,
  teamsHighContrastTheme,
  makeStyles,
  mergeClasses,
  tokens,
} from "@fluentui/react-components";

const TEAMS_PATHS = ["/teams-submit-cg7k2x9m", "/teams-tab", "/teams-tab/orders"];

// Griffel blocks the `border*` 4-side shorthands (it can't safely expand them),
// so we set each side longhand via these helpers.
const thinBorderSides = {
  borderTopWidth: tokens.strokeWidthThin,
  borderRightWidth: tokens.strokeWidthThin,
  borderBottomWidth: tokens.strokeWidthThin,
  borderLeftWidth: tokens.strokeWidthThin,
  borderTopStyle: "solid",
  borderRightStyle: "solid",
  borderBottomStyle: "solid",
  borderLeftStyle: "solid",
} as const;
const allBorderColor = (c: string) => ({
  borderTopColor: c,
  borderRightColor: c,
  borderBottomColor: c,
  borderLeftColor: c,
});

const useShellStyles = makeStyles({
  shell: {
    display: "flex",
    flexDirection: "column",
    height: "100dvh",
    overflow: "hidden",
    backgroundColor: tokens.colorNeutralBackground1,
  },
  content: {
    display: "flex",
    flexDirection: "column",
    flexGrow: 1,
    minHeight: 0,
    overflow: "hidden",
  },
});

const useSwitcherStyles = makeStyles({
  bar: {
    flexShrink: 0,
    display: "flex",
    gap: tokens.spacingHorizontalS,
    paddingLeft: tokens.spacingHorizontalL,
    paddingRight: tokens.spacingHorizontalL,
    paddingTop: tokens.spacingVerticalM,
    paddingBottom: tokens.spacingVerticalS,
  },
  pill: {
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    paddingLeft: tokens.spacingHorizontalL,
    paddingRight: tokens.spacingHorizontalL,
    paddingTop: tokens.spacingVerticalXS,
    paddingBottom: tokens.spacingVerticalXS,
    borderRadius: tokens.borderRadiusCircular,
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightMedium,
    fontFamily: tokens.fontFamilyBase,
    lineHeight: tokens.lineHeightBase200,
    cursor: "pointer",
    textDecorationLine: "none",
    ...thinBorderSides,
    ...allBorderColor(tokens.colorNeutralStroke2),
    backgroundColor: tokens.colorTransparentBackground,
    color: tokens.colorNeutralForeground3,
    transitionProperty: "background-color, color, border-color",
    transitionDuration: tokens.durationNormal,
    ":hover": {
      backgroundColor: tokens.colorNeutralBackground1Hover,
      color: tokens.colorNeutralForeground2,
    },
  },
  activeSubmit: {
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground1,
    ...allBorderColor(tokens.colorTransparentStroke),
    ":hover": {
      backgroundColor: tokens.colorBrandBackground2Hover,
      color: tokens.colorBrandForeground1,
    },
  },
  activeOrders: {
    backgroundColor: tokens.colorPaletteBerryBackground1,
    color: tokens.colorPaletteBerryForeground1,
    ...allBorderColor(tokens.colorTransparentStroke),
    ":hover": {
      backgroundColor: tokens.colorPaletteBerryBackground1,
      color: tokens.colorPaletteBerryForeground1,
    },
  },
});

// A quiet in-content segmented toggle — deliberately not a bottom app-nav bar,
// so it reads as part of our content and stays visually separate from the
// Teams navigation chrome that already lives along the bottom of the window.
function TeamsTabSwitcher() {
  const [location] = useLocation();
  const s = useSwitcherStyles();
  const isOrders = location === "/teams-tab/orders";

  return (
    <div className={s.bar}>
      <Link
        href="/teams-tab"
        aria-current={!isOrders ? "page" : undefined}
        className={mergeClasses(s.pill, !isOrders && s.activeSubmit)}
      >
        Submit
      </Link>
      <Link
        href="/teams-tab/orders"
        aria-current={isOrders ? "page" : undefined}
        className={mergeClasses(s.pill, isOrders && s.activeOrders)}
      >
        Orders
      </Link>
    </div>
  );
}

function TeamsRouterContent() {
  const [location] = useLocation();
  const { theme } = useTeamsTheme();
  const styles = useShellStyles();
  const isOrders = location === "/teams-tab/orders";

  const fluentTheme =
    theme === "dark"
      ? teamsDarkTheme
      : theme === "contrast"
      ? teamsHighContrastTheme
      : teamsLightTheme;

  return (
    <FluentProvider
      theme={fluentTheme}
      className={styles.shell}
      style={{
        paddingTop: "env(safe-area-inset-top)",
        paddingBottom: "env(safe-area-inset-bottom)",
      }}
    >
      <TeamsTabSwitcher />
      <div className={styles.content}>
        {isOrders ? <OrdersTab /> : <SubmitTab />}
      </div>
    </FluentProvider>
  );
}

function TeamsRouter() {
  return (
    <TeamsThemeProvider>
      <TeamsRouterContent />
    </TeamsThemeProvider>
  );
}

function MainRouter() {
  return (
    <AuthWrapper>
      <Switch>
        <Route path="/" component={Home} />
        <Route path="/meeting-history" component={MeetingHistory} />
        <Route path="/actions" component={Actions} />
        <Route path="/staff-training" component={StaffTraining} />
        <Route path="/equipment-maintenance" component={EquipmentMaintenance} />
        <Route path="/teams-integration-plan" component={TeamsIntegrationPlan} />
        <Route path="/policy/health-safety" component={HealthSafetyPolicy} />
        <Route path="/health-safety-policy" component={HealthSafetyPolicy} />
        <Route path="/environment-policy" component={EnvironmentPolicy} />
        <Route path="/quality-policy" component={QualityPolicy} />
        <Route component={NotFound} />
      </Switch>
    </AuthWrapper>
  );
}

function Router() {
  const [location] = useLocation();

  if (TEAMS_PATHS.includes(location)) {
    return <TeamsRouter />;
  }
  return <MainRouter />;
}

function App() {
  return (
    <QueryClientProvider client={queryClient}>
      <TooltipProvider>
        <Toaster />
        <Router />
      </TooltipProvider>
    </QueryClientProvider>
  );
}

export default App;

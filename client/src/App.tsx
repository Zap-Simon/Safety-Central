import { useState } from "react";
import { Switch, Route, useLocation } from "wouter";
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
  TabList,
  Tab,
  Text,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import type { SelectTabEvent, SelectTabData } from "@fluentui/react-components";

const TEAMS_PATHS = ["/teams-submit-cg7k2x9m", "/teams-tab", "/teams-tab/orders"];

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
    alignItems: "center",
    gap: tokens.spacingHorizontalXXL,
    paddingLeft: tokens.spacingHorizontalM,
    paddingRight: tokens.spacingHorizontalM,
    paddingTop: tokens.spacingVerticalXS,
    borderBottomWidth: tokens.strokeWidthThin,
    borderBottomStyle: "solid",
    borderBottomColor: tokens.colorNeutralStroke2,
  },
  // Greeting lives on the left of the tab row; the tabs start to its right.
  greeting: {
    flexShrink: 1,
    minWidth: 0,
    overflow: "hidden",
    whiteSpace: "nowrap",
    textOverflow: "ellipsis",
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
    paddingLeft: tokens.spacingHorizontalXS,
  },
  tabs: { flexShrink: 0 },
  // The selected-tab indicator is a `::after` bar coloured by Fluent's brand
  // token. We recolour it per tab so the blue=Submit / purple=Orders wayfinding
  // survives the move to native Fluent tabs. Only applied to the selected tab,
  // so the bar shape still comes from Fluent's own selected styles.
  submitInd: { "::after": { backgroundColor: tokens.colorBrandStroke1 } },
  ordersInd: { "::after": { backgroundColor: tokens.colorPaletteBerryForeground1 } },
});

// A quiet in-content segmented toggle — deliberately not a bottom app-nav bar,
// so it reads as part of our content and stays visually separate from the
// Teams navigation chrome that already lives along the bottom of the window.
function TeamsTabSwitcher({
  userName,
  showGreeting,
}: {
  userName: string;
  showGreeting: boolean;
}) {
  const [location, navigate] = useLocation();
  const s = useSwitcherStyles();
  const selected = location === "/teams-tab/orders" ? "/teams-tab/orders" : "/teams-tab";

  const onTabSelect = (_e: SelectTabEvent, data: SelectTabData) => {
    navigate(data.value as string);
  };

  return (
    <div className={s.bar}>
      {showGreeting && userName && (
        <Text className={s.greeting}>
          Hi {userName.split(" ")[0]} <span style={{ marginLeft: "2px" }}>👋</span>
        </Text>
      )}
      <TabList className={s.tabs} selectedValue={selected} onTabSelect={onTabSelect} size="large">
        <Tab
          value="/teams-tab"
          className={selected === "/teams-tab" ? s.submitInd : undefined}
        >
          Submit
        </Tab>
        <Tab
          value="/teams-tab/orders"
          className={selected === "/teams-tab/orders" ? s.ordersInd : undefined}
        >
          Orders
        </Tab>
      </TabList>
    </div>
  );
}

function TeamsRouterContent() {
  const [location] = useLocation();
  const { theme } = useTeamsTheme();
  const styles = useShellStyles();
  const [teamsUser, setTeamsUser] = useState("");
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
      <TeamsTabSwitcher userName={teamsUser} showGreeting={!isOrders} />
      <div className={styles.content}>
        {isOrders ? <OrdersTab /> : <SubmitTab onUser={setTeamsUser} />}
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

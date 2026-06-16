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

const TEAMS_PATHS = ["/teams-submit-cg7k2x9m", "/teams-tab", "/teams-tab/orders"];

// A quiet in-content segmented toggle — deliberately not a bottom app-nav bar,
// so it reads as part of our content and stays visually separate from the
// Teams navigation chrome that already lives along the bottom of the window.
function TeamsTabSwitcher() {
  const [location] = useLocation();
  const { isDark } = useTeamsTheme();
  const isOrders = location === "/teams-tab/orders";

  const base =
    "inline-flex items-center justify-center rounded-full px-4 py-1 text-[13px] font-medium transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-offset-1 focus-visible:ring-blue-500/50";
  const inactive = isDark
    ? "border border-gray-700 text-gray-400"
    : "border border-gray-200 text-gray-600";

  return (
    <div className="shrink-0 flex gap-2 px-4 pt-3 pb-2">
      <Link
        href="/teams-tab"
        aria-current={!isOrders ? "page" : undefined}
        className={`${base} ${
          !isOrders
            ? isDark
              ? "bg-blue-500/15 text-blue-300"
              : "bg-blue-50 text-blue-700"
            : inactive
        }`}
      >
        Submit
      </Link>
      <Link
        href="/teams-tab/orders"
        aria-current={isOrders ? "page" : undefined}
        className={`${base} ${
          isOrders
            ? isDark
              ? "bg-purple-500/15 text-purple-300"
              : "bg-purple-50 text-purple-700"
            : inactive
        }`}
      >
        Orders
      </Link>
    </div>
  );
}

function TeamsRouterContent() {
  const [location] = useLocation();
  const { isDark } = useTeamsTheme();
  const isOrders = location === "/teams-tab/orders";

  return (
    <div
      className={`flex flex-col h-[100dvh] overflow-hidden ${isDark ? "dark bg-gray-900" : "bg-white"}`}
      style={{
        paddingTop: "env(safe-area-inset-top)",
        paddingBottom: "env(safe-area-inset-bottom)",
      }}
    >
      <TeamsTabSwitcher />
      <div className="flex flex-col flex-1 min-h-0 overflow-hidden">
        {isOrders ? <OrdersTab /> : <SubmitTab />}
      </div>
    </div>
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

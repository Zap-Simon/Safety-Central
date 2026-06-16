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
import { Send, ShoppingCart } from "lucide-react";
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
    "flex-1 flex items-center justify-center gap-1.5 rounded-full py-2 text-sm font-semibold transition-all focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-offset-1 focus-visible:ring-blue-500/60";
  const inactive = isDark ? "text-gray-400" : "text-gray-500";

  return (
    <div className="shrink-0 px-4 pt-5 pb-2">
      <div className={`mx-auto flex max-w-[260px] rounded-full p-1 ${isDark ? "bg-gray-800 ring-1 ring-white/5" : "bg-gray-100 ring-1 ring-gray-200"}`}>
        <Link
          href="/teams-tab"
          aria-current={!isOrders ? "page" : undefined}
          className={`${base} ${
            !isOrders
              ? isDark
                ? "bg-gray-900 text-blue-400 shadow-sm"
                : "bg-white text-blue-600 shadow-sm"
              : inactive
          }`}
        >
          <Send className="h-4 w-4" />
          Submit
        </Link>
        <Link
          href="/teams-tab/orders"
          aria-current={isOrders ? "page" : undefined}
          className={`${base} ${
            isOrders
              ? isDark
                ? "bg-gray-900 text-purple-400 shadow-sm"
                : "bg-white text-purple-600 shadow-sm"
              : inactive
          }`}
        >
          <ShoppingCart className="h-4 w-4" />
          Orders
        </Link>
      </div>
    </div>
  );
}

function TeamsRouterContent() {
  const [location] = useLocation();
  const { isDark } = useTeamsTheme();
  const isOrders = location === "/teams-tab/orders";

  return (
    <div
      className={`flex flex-col h-screen overflow-hidden ${isDark ? "dark bg-gray-900" : "bg-gray-50"}`}
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

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

function TeamsBottomNav() {
  const [location] = useLocation();
  const { isDark } = useTeamsTheme();
  const isOrders = location === "/teams-tab/orders";

  return (
    <nav
      className={`fixed bottom-0 left-0 right-0 border-t flex z-50 ${
        isDark ? "bg-gray-900 border-gray-700" : "bg-white border-gray-200"
      }`}
      style={{ paddingBottom: "env(safe-area-inset-bottom)" }}
    >
      <Link
        href="/teams-tab"
        className={`flex-1 flex flex-col items-center justify-center py-2 gap-0.5 text-xs font-medium transition-colors ${
          !isOrders
            ? "text-blue-500"
            : isDark
            ? "text-gray-500 hover:text-gray-300"
            : "text-gray-400 hover:text-gray-600"
        }`}
      >
        <Send className="h-5 w-5" />
        <span>Submit</span>
      </Link>
      <Link
        href="/teams-tab/orders"
        className={`flex-1 flex flex-col items-center justify-center py-2 gap-0.5 text-xs font-medium transition-colors ${
          isOrders
            ? "text-purple-500"
            : isDark
            ? "text-gray-500 hover:text-gray-300"
            : "text-gray-400 hover:text-gray-600"
        }`}
      >
        <ShoppingCart className="h-5 w-5" />
        <span>Orders</span>
      </Link>
    </nav>
  );
}

function TeamsRouterContent() {
  const [location] = useLocation();
  const { isDark } = useTeamsTheme();
  const isOrders = location === "/teams-tab/orders";

  return (
    <div
      className={`flex flex-col min-h-screen ${isDark ? "dark bg-gray-900" : "bg-white"}`}
      style={{
        paddingTop: "env(safe-area-inset-top)",
        paddingBottom: "calc(3.5rem + env(safe-area-inset-bottom))",
      }}
    >
      {isOrders ? (
        <div className="flex flex-col flex-1 min-h-0 overflow-hidden">
          <OrdersTab />
        </div>
      ) : (
        <SubmitTab />
      )}
      <TeamsBottomNav />
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

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

const TEAMS_PATHS = ["/teams-submit-cg7k2x9m", "/teams-tab", "/teams-tab/orders"];

// Detect an MSAL redirect callback landing on "/" so we can restore the Teams UI
function getMsalRedirectTarget(): string | null {
  const hasMsalParams =
    window.location.hash.includes("code=") ||
    window.location.hash.includes("error=") ||
    window.location.search.includes("code=") ||
    window.location.search.includes("error=");
  if (!hasMsalParams) return null;
  return sessionStorage.getItem("teams-auth-redirect-target");
}

function TeamsBottomNav() {
  const [location] = useLocation();
  const isOrders = location === "/teams-tab/orders";

  return (
    <nav className="fixed bottom-0 left-0 right-0 bg-white border-t border-gray-200 flex z-50 safe-bottom">
      <Link
        href="/teams-tab"
        className={`flex-1 flex flex-col items-center justify-center py-2 gap-0.5 text-xs font-medium transition-colors ${
          !isOrders ? "text-blue-600" : "text-gray-400 hover:text-gray-600"
        }`}
      >
        <Send className="h-5 w-5" />
        <span>Submit</span>
      </Link>
      <Link
        href="/teams-tab/orders"
        className={`flex-1 flex flex-col items-center justify-center py-2 gap-0.5 text-xs font-medium transition-colors ${
          isOrders ? "text-purple-600" : "text-gray-400 hover:text-gray-600"
        }`}
      >
        <ShoppingCart className="h-5 w-5" />
        <span>Orders</span>
      </Link>
    </nav>
  );
}

function TeamsRouter() {
  const [location] = useLocation();
  const isOrders = location === "/teams-tab/orders";

  return (
    <div className="flex flex-col min-h-screen pb-14">
      {isOrders ? (
        <div className="flex flex-col flex-1 min-h-0">
          <div className="bg-gradient-to-r from-purple-700 to-purple-500 px-5 pt-6 pb-7 text-white">
            <div className="max-w-lg mx-auto flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-white/15 backdrop-blur flex items-center justify-center">
                <ShoppingCart className="h-5 w-5" />
              </div>
              <div>
                <h1 className="text-base font-bold leading-tight">Ordering List</h1>
                <p className="text-xs text-purple-100/90">Cranfield Glass Christchurch</p>
              </div>
            </div>
          </div>
          <div className="flex-1 overflow-hidden -mt-4">
            <div className="bg-gray-50 rounded-t-2xl h-full pt-4">
              <OrdersTab />
            </div>
          </div>
        </div>
      ) : (
        <SubmitTab />
      )}
      <TeamsBottomNav />
    </div>
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
  const [location, navigate] = useLocation();

  // Handle MSAL redirect callback landing on "/" — restore the correct Teams route
  const msalTarget = getMsalRedirectTarget();
  if (msalTarget && location === "/") {
    sessionStorage.removeItem("teams-auth-redirect-target");
    // Navigate to the stored Teams path so TeamsRouter renders and OrdersTab
    // can call handleRedirectPromise() to consume the auth tokens.
    navigate(msalTarget, { replace: true });
    return null;
  }

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

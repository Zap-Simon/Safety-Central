import { Switch, Route } from "wouter";
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

import NotFound from "@/pages/not-found";

function Router() {
  return (
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
  );
}

function App() {
  return (
    <QueryClientProvider client={queryClient}>
      <TooltipProvider>
        <Toaster />
        <AuthWrapper>
          <Router />
        </AuthWrapper>
      </TooltipProvider>
    </QueryClientProvider>
  );
}

export default App;

import Overview from "@/components/sections/overview";
import VisualOverview from "@/components/sections/visual-overview";
import WorkflowDiagram from "@/components/sections/workflow-diagram";
import ListColumns from "@/components/sections/list-columns";
import WorkflowSteps from "@/components/sections/workflow-steps";
import Benefits from "@/components/sections/benefits";
import Sidebar from "@/components/sidebar";

export default function DocumentationTab() {
  return (
    <div className="lg:grid lg:grid-cols-4 lg:gap-6 xl:gap-8">
      <div className="hidden lg:block">
        <Sidebar />
      </div>
      
      <main className="lg:col-span-4 lg:ml-72 space-y-4 sm:space-y-6 lg:space-y-8 p-4">
        <Overview />
        <VisualOverview />
        <WorkflowDiagram />
        <WorkflowSteps />
        <ListColumns />
        <Benefits />
      </main>
    </div>
  );
}
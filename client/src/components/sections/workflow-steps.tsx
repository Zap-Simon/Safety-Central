import { useState } from "react";

export default function WorkflowSteps() {
  const [expandedSteps, setExpandedSteps] = useState<Record<string, boolean>>({
    "step1": false,
    "step2": false,
    "step3": false,
    "step4": false
  });

  const toggleStep = (stepId: string) => {
    setExpandedSteps(prev => ({
      ...prev,
      [stepId]: !prev[stepId]
    }));
  };

  const steps = [
    {
      id: "step1",
      title: "1. Form Submission & Processing",
      bgColor: "ms-blue",
      hoverColor: "hover:ms-blue-dark",
      iconColor: "bg-blue-400",
      items: [
        {
          title: "Access forms via Teams Viva card",
          description: "Staff click the Health & Safety Viva card in Teams navigation"
        },
        {
          title: "User submits Microsoft Form",
          description: "Safety, Business Ideas, or Near Miss forms submitted"
        },
        {
          title: "Power Automate triggered",
          description: "Automated workflow processes the submission"
        },
        {
          title: "Added to relevant SharePoint List",
          description: "Item created with \"Submitted\" status"
        },
        {
          title: "Teams notification sent",
          description: "Team receives instant notification of new submission"
        }
      ]
    },
    {
      id: "step2",
      title: "2. Microsoft Teams Collaboration",
      bgColor: "bg-purple-600",
      hoverColor: "hover:bg-purple-700",
      iconColor: "bg-purple-500",
      items: [
        {
          title: "Real-time notifications",
          description: "Teams receives instant alerts when forms are submitted"
        },
        {
          title: "Two-week preparation period",
          description: "Team reviews submissions and prepares for meeting discussions"
        },
        {
          title: "Meeting coordination",
          description: "Teams facilitates fortnightly meeting scheduling and agenda sharing"
        }
      ]
    },
    {
      id: "step3",
      title: "3. Fortnightly Meeting Decisions",
      bgColor: "ms-amber",
      hoverColor: "hover:bg-yellow-500",
      iconColor: "bg-yellow-500",
      items: [
        {
          title: "Review submitted items",
          description: "Team examines new submissions from all three lists"
        },
        {
          title: "Review in-discussion progress",
          description: "Check status of items being actively discussed"
        },
        {
          title: "Decision point: New items",
          description: "Mark as \"In Discussion\" for further review or \"Actioned\" for immediate implementation"
        },
        {
          title: "Decision point: Ongoing items",
          description: "Continue in discussion or mark completed items as \"Actioned\""
        }
      ]
    },
    {
      id: "step4",
      title: "4. Action List Management",
      bgColor: "bg-orange-600",
      hoverColor: "hover:bg-orange-700",
      iconColor: "bg-orange-500",
      items: [
        {
          title: "Auto-copy to Action List",
          description: "Power Automate copies \"Actioned\" items to central tracking list"
        },
        {
          title: "Initial status: Not Started",
          description: "All new action items begin with \"Not Started\" status"
        },
        {
          title: "Admin assignment & tracking",
          description: "Responsible person assigned, progress tracked through statuses"
        },
        {
          title: "Status progression",
          description: "In Progress → Completed → Original item auto-updated to \"Closed\""
        },
        {
          title: "Board removal",
          description: "Closed items removed from active board view"
        }
      ]
    }
  ];

  return (
    <section id="workflow-steps" className="bg-gradient-to-br from-white to-gray-50 rounded-xl shadow-lg p-4 sm:p-6 lg:p-8 mb-6 sm:mb-8 border border-gray-200 touch-manipulation" style={{ touchAction: 'pan-y' }}>
      <div className="flex items-center space-x-3 mb-6">
        <div className="w-10 h-10 bg-gradient-to-r from-blue-600 to-blue-700 rounded-xl flex items-center justify-center shadow-lg flex-shrink-0">
          <i className="fas fa-list-ol text-white text-lg"></i>
        </div>
        <div className="min-w-0 flex-1">
          <h2 className="text-xl sm:text-2xl font-bold ms-gray-900">Process Walkthrough</h2>
          <p className="text-xs sm:text-sm text-gray-600">Step-by-step breakdown connecting Data Flow Architecture and Process Workflow</p>
        </div>
      </div>
      
      {steps.map((step) => (
        <div key={step.id} className="border border-gray-100 rounded-lg mb-3 sm:mb-4 touch-manipulation" style={{ touchAction: 'pan-y' }}>
          <button 
            onClick={() => toggleStep(step.id)}
            className={`w-full px-4 sm:px-6 py-3 sm:py-4 text-left ${step.bgColor} text-white rounded-t-lg ${step.hoverColor} active:opacity-90 transition-all touch-manipulation flex items-center justify-between`}
          >
            <span className="font-medium sm:font-semibold text-sm sm:text-base whitespace-pre-line">{step.title}</span>
            <i className={`fas fa-chevron-down transform transition-transform ${expandedSteps[step.id] ? 'rotate-180' : ''}`}></i>
          </button>
          <div className={`collapsible-content px-4 sm:px-6 py-3 sm:py-4 touch-manipulation ${expandedSteps[step.id] ? '' : 'hidden'}`} style={{ touchAction: 'pan-y' }}>
            <div className="space-y-2">
              {step.items.map((item, index) => (
                <div key={index} className="border-l-2 border-gray-200 pl-3 py-1 touch-manipulation" style={{ touchAction: 'pan-y' }}>
                  <h4 className="font-medium text-gray-900 text-sm">{item.title}</h4>
                  <p className="text-xs text-gray-600">{item.description}</p>
                </div>
              ))}
            </div>
          </div>
        </div>
      ))}
    </section>
  );
}

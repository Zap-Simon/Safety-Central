import { useEffect } from "react";
import { Link } from "wouter";
import { ChevronLeft, Users, Calendar, Bell, FileText, MessageSquare, Clock } from "lucide-react";

export default function TeamsIntegrationPlan() {
  useEffect(() => {
    window.scrollTo(0, 0);
  }, []);
  const phases = [
    {
      phase: "Phase 1",
      title: "Pre-Meeting Automation Setup",
      duration: "Days 1-2",
      icon: <Calendar className="w-6 h-6" />,
      status: "Foundation Ready",
      tasks: [
        "✓ SharePoint credentials already configured in Power Automate",
        "✓ Meeting History page structure complete for data integration",
        "→ Create Power Automate flow triggered 2 days before meetings",
        "→ Build agenda extraction logic from SharePoint lists",
        "→ Filter items submitted since last meeting date"
      ]
    },
    {
      phase: "Phase 2", 
      title: "Agenda Generation & Teams Distribution",
      duration: "Days 3-4",
      icon: <FileText className="w-6 h-6" />,
      status: "PDF System Ready",
      tasks: [
        "✓ Professional PDF generation system already implemented",
        "✓ Company branding and formatting templates complete",
        "→ Adapt existing PDF system for agenda format",
        "→ Include action items from previous meetings",
        "→ Configure Teams channel posting automation"
      ]
    },
    {
      phase: "Phase 3",
      title: "Teams Notifications & Post-Meeting",
      duration: "Days 5-6",
      icon: <Bell className="w-6 h-6" />,
      status: "Integration Phase",
      tasks: [
        "→ Implement @mention notifications for all attendees",
        "→ Post-meeting minutes distribution to Teams channel",
        "→ Teams chat alerts for new action items",
        "→ Connect with existing Viva card workflow",
        "→ End-to-end testing and deployment"
      ]
    }
  ];

  const features = [
    {
      title: "Automated Agenda Creation",
      description: "Extract new submissions and generate professional agendas",
      icon: <Calendar className="w-5 h-5" />
    },
    {
      title: "Teams Channel Integration", 
      description: "Post agendas and minutes directly to Teams channels",
      icon: <MessageSquare className="w-5 h-5" />
    },
    {
      title: "Smart Notifications",
      description: "@mention attendees with agenda reminders and action items",
      icon: <Bell className="w-5 h-5" />
    },
    {
      title: "Workflow Continuity",
      description: "Connect pre-meeting, meeting, and post-meeting processes",
      icon: <Users className="w-5 h-5" />
    }
  ];

  const getStatusColor = (status: string) => {
    switch (status) {
      case "Foundation Ready":
        return "bg-emerald-50 text-emerald-700 border-emerald-200";
      case "PDF System Ready":
        return "bg-blue-50 text-blue-700 border-blue-200";
      case "Integration Phase":
        return "bg-purple-50 text-purple-700 border-purple-200";
      default:
        return "bg-gray-50 text-gray-700 border-gray-200";
    }
  };

  return (
    <div className="min-h-screen bg-gray-50">
      <div className="max-w-6xl mx-auto px-4 sm:px-6 lg:px-8 py-4 sm:py-6">
        {/* Back Navigation */}
        <div className="mb-6">
          <Link href="/">
            <button className="inline-flex items-center px-4 py-2 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded-lg hover:bg-gray-50 transition-colors duration-200">
              <ChevronLeft className="w-4 h-4 mr-2" />
              Back to Documentation
            </button>
          </Link>
        </div>

        {/* Header */}
        <div className="bg-white rounded-lg border border-gray-200 p-4 sm:p-6 mb-6">
          <div className="flex flex-col sm:flex-row sm:items-center space-y-3 sm:space-y-0 sm:space-x-3 mb-4">
            <div className="w-12 h-12 bg-gradient-to-r from-blue-600 to-blue-700 rounded-lg flex items-center justify-center flex-shrink-0">
              <Users className="w-6 h-6 text-white" />
            </div>
            <div className="min-w-0">
              <h1 className="text-xl sm:text-2xl font-bold text-gray-900">Enhanced Microsoft Teams Integration</h1>
              <p className="text-sm sm:text-base text-gray-600">6-day deployment for complete Teams workflow automation</p>
            </div>
          </div>
          
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 sm:gap-4 mt-4 sm:mt-6">
            <div className="bg-gray-50 rounded-lg p-3 sm:p-4">
              <div className="flex items-center space-x-2">
                <Clock className="w-4 h-4 sm:w-5 sm:h-5 text-gray-600 flex-shrink-0" />
                <span className="font-medium text-gray-900 text-sm sm:text-base">Timeline</span>
              </div>
              <p className="text-xs sm:text-sm text-gray-600 mt-1">6 days</p>
            </div>
            <div className="bg-gray-50 rounded-lg p-3 sm:p-4">
              <div className="flex items-center space-x-2">
                <Users className="w-4 h-4 sm:w-5 sm:h-5 text-gray-600 flex-shrink-0" />
                <span className="font-medium text-gray-900 text-sm sm:text-base">Dependencies</span>
              </div>
              <p className="text-xs sm:text-sm text-gray-600 mt-1">Meeting History system</p>
            </div>
            <div className="bg-gray-50 rounded-lg p-3 sm:p-4">
              <div className="flex items-center space-x-2">
                <MessageSquare className="w-4 h-4 sm:w-5 sm:h-5 text-gray-600 flex-shrink-0" />
                <span className="font-medium text-gray-900 text-sm sm:text-base">Integration</span>
              </div>
              <p className="text-xs sm:text-sm text-gray-600 mt-1">Microsoft 365 ecosystem</p>
            </div>
          </div>
        </div>

        {/* Key Features */}
        <div className="bg-white rounded-lg border border-gray-200 p-4 sm:p-6 mb-6">
          <h2 className="text-lg font-semibold text-gray-900 mb-4">Key Features</h2>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            {features.map((feature, index) => (
              <div key={index} className="flex items-start space-x-3 p-3 bg-gray-50 rounded-lg">
                <div className="w-8 h-8 bg-blue-100 rounded-lg flex items-center justify-center flex-shrink-0">
                  {feature.icon}
                </div>
                <div>
                  <h3 className="font-medium text-gray-900 text-sm">{feature.title}</h3>
                  <p className="text-xs text-gray-600 mt-1">{feature.description}</p>
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* Implementation Phases */}
        <div className="space-y-4 sm:space-y-6 mb-6">
          {phases.map((phase, index) => (
            <div key={index} className="bg-white rounded-lg border border-gray-200 p-4 sm:p-6 hover:shadow-md transition-shadow duration-200">
              <div className="flex flex-col sm:flex-row sm:items-start space-y-3 sm:space-y-0 sm:space-x-4 mb-4">
                <div className="w-10 h-10 sm:w-12 sm:h-12 bg-gradient-to-r from-blue-600 to-blue-700 rounded-lg flex items-center justify-center flex-shrink-0">
                  {phase.icon}
                </div>
                <div className="flex-1 min-w-0">
                  <div className="flex flex-col space-y-2 mb-2">
                    <h3 className="text-base sm:text-lg font-semibold text-gray-900">{phase.phase}: {phase.title}</h3>
                    <p className="text-xs sm:text-sm text-gray-600">{phase.duration}</p>
                    <span className={`px-2 py-1 rounded-full text-xs font-medium border ${getStatusColor(phase.status)} self-start`}>
                      {phase.status}
                    </span>
                  </div>
                </div>
              </div>
              
              <div className="space-y-2">
                {phase.tasks.map((task, taskIndex) => (
                  <div key={taskIndex} className="text-xs sm:text-sm text-gray-700">
                    <span className="text-blue-600 font-medium">• </span>
                    <span>{task}</span>
                  </div>
                ))}
              </div>
            </div>
          ))}
        </div>

        {/* Benefits */}
        <div className="bg-white rounded-lg border border-gray-200 p-4 sm:p-6">
          <h2 className="text-lg font-semibold text-gray-900 mb-4">Expected Outcomes</h2>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-sm text-gray-700">
            <div>
              <h3 className="font-medium text-gray-900 mb-2">Efficiency Gains</h3>
              <ul className="space-y-1">
                <li>• Automated agenda preparation and distribution</li>
                <li>• Reduced meeting preparation time by 75%</li>
                <li>• Consistent communication workflow</li>
              </ul>
            </div>
            <div>
              <h3 className="font-medium text-gray-900 mb-2">Team Collaboration</h3>
              <ul className="space-y-1">
                <li>• Centralised Teams-based workflow</li>
                <li>• Real-time notifications and updates</li>
                <li>• Seamless integration with existing tools</li>
              </ul>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
import { useState } from "react";

export default function ListColumns() {
  const [expandedSections, setExpandedSections] = useState<Record<string, boolean>>({
    "business": false,
    "safety": false,
    "nearmiss": false,
    "action": false
  });

  const toggleSection = (sectionId: string) => {
    setExpandedSections(prev => ({
      ...prev,
      [sectionId]: !prev[sectionId]
    }));
  };

  const renderColumnValues = (values: string) => {
    // Only show values section if there are actual choice options (not just "-")
    if (values === "-") return null;
    return (
      <div className="text-xs sm:text-sm text-gray-500 mt-1">
        <strong>Values:</strong> {values}
      </div>
    );
  };
  const businessIdeasColumns = [
    { name: "Title", type: "Single line of text", description: "Short summary of the business idea", values: "-" },
    { name: "Submitted By", type: "Person or Group", description: "Person who submitted the form", values: "-" },
    { name: "Status", type: "Choice", description: "Current workflow status", values: "Unassigned, Submitted, In Discussion, Actioned" },
    { name: "Idea Date", type: "Date and Time", description: "When the idea was submitted", values: "-" },
    { name: "Idea Type", type: "Choice", description: "Category of business idea", values: "Value Add Idea, Waste Reduce, Feedback" },
    { name: "Business Idea", type: "Multiple lines of text", description: "Detailed description of the business idea", values: "-" },
    { name: "Meeting Date", type: "Date and Time", description: "Date when discussed in meeting", values: "-" },
    { name: "Meeting Notes", type: "Multiple lines of text", description: "Notes and outcomes from meeting discussion", values: "-" }
  ];

  const safetyIdeasColumns = [
    { name: "Title", type: "Single line of text", description: "Short summary of the safety idea", values: "-" },
    { name: "Submitted By", type: "Person or Group", description: "Person who submitted the form", values: "-" },
    { name: "Status", type: "Choice", description: "Current workflow status", values: "Unassigned, Submitted, In Discussion, Actioned" },
    { name: "Idea Date", type: "Date and Time", description: "When the idea was submitted", values: "-" },
    { name: "Idea Type", type: "Choice", description: "Category of safety idea", values: "Safety Update, Order, Hazard Report" },
    { name: "Safety Idea", type: "Multiple lines of text", description: "Detailed description of the safety improvement", values: "-" },
    { name: "Meeting Date", type: "Date and Time", description: "Date when discussed in meeting", values: "-" },
    { name: "Meeting Notes", type: "Multiple lines of text", description: "Notes and outcomes from meeting discussion", values: "-" }
  ];

  const nearMissColumns = [
    { name: "Title", type: "Single line of text", description: "Short summary of the incident", values: "-" },
    { name: "Name", type: "Person or Group", description: "Person reporting the incident", values: "-" },
    { name: "Event Date", type: "Date and Time", description: "When the incident occurred", values: "-" },
    { name: "Event Type", type: "Choice", description: "Type of safety event", values: "Near Miss, Harm Accident, Serious Harm, Property Damage" },
    { name: "What happened?", type: "Multiple lines of text", description: "Description of what occurred", values: "-" },
    { name: "How It happened?", type: "Multiple lines of text", description: "Explanation of how the incident occurred", values: "-" },
    { name: "Investigation?", type: "Yes/No", description: "Whether formal investigation is required", values: "Yes, No" },
    { name: "Report Link", type: "Hyperlink or Picture", description: "Link to formal incident report if available", values: "-" },
    { name: "Status", type: "Choice", description: "Current workflow status", values: "Unassigned, Submitted, In Discussion, Actioned" },
    { name: "Meeting Date", type: "Date and Time", description: "Date when discussed in meeting", values: "-" },
    { name: "Meeting Notes", type: "Multiple lines of text", description: "Notes and outcomes from meeting discussion", values: "-" }
  ];

  const actionListColumns = [
    { name: "Idea / Event", type: "Multiple lines of text", description: "Content copied from original item", values: "-" },
    { name: "Title", type: "Single line of text", description: "Title copied from source item", values: "-" },
    { name: "Assigned to", type: "Person or Group", description: "Person responsible for the action", values: "-" },
    { name: "Status", type: "Choice", description: "Action tracking status", values: "Not Started, In Progress / Assigned, On Hold, Awaiting Discussion, Completed" },
    { name: "List Origin", type: "Choice", description: "Source list category", values: "Business Idea, Safety Idea, Near Miss / Accident, Meeting Minutes, Compliance, Audit" },
    { name: "Priority", type: "Choice", description: "Action priority level", values: "Low, Medium, High, Long Term" },
    { name: "Action Start Date", type: "Date and Time", description: "When action work began", values: "-" },
    { name: "Action Completion Date", type: "Date and Time", description: "When action was completed", values: "-" },
    { name: "Meeting Notes", type: "Multiple lines of text", description: "Notes from original meeting discussion", values: "-" },
    { name: "Outcome Result", type: "Multiple lines of text", description: "Final outcome and results achieved", values: "-" },
    { name: "Submitted By", type: "Person or Group", description: "Original submitter from source list", values: "-" },
    { name: "Idea / Event Type", type: "Choice", description: "Type from original submission", values: "Value Add Idea, Waste Reduce, Feedback, Safety Update, Order, Hazard Report, Near Miss, Harm Accident, Serious Harm, Property Damage" },
    { name: "Idea Date", type: "Date and Time", description: "Original submission date", values: "-" },
    { name: "Link", type: "Hyperlink or Picture", description: "Link back to original list item", values: "-" }
  ];

  return (
    <section id="list-columns" className="bg-gradient-to-br from-white to-gray-50 rounded-xl shadow-lg p-4 sm:p-6 lg:p-8 mb-6 sm:mb-8 border border-gray-200">
      <div className="flex items-center space-x-3 mb-6 sm:mb-8">
        <div className="w-10 h-10 bg-gradient-to-r from-blue-600 to-blue-700 rounded-xl flex items-center justify-center shadow-lg flex-shrink-0">
          <i className="fas fa-table text-white text-lg"></i>
        </div>
        <div className="min-w-0 flex-1">
          <h2 className="text-lg sm:text-xl font-bold ms-gray-900">Data Structures</h2>
          <p className="text-xs sm:text-sm text-gray-600">SharePoint list column configurations</p>
        </div>
      </div>
      
      {/* Business Ideas Table */}
      <div className="mb-6 sm:mb-8">
        <button
          onClick={() => toggleSection("business")}
          className="w-full bg-gradient-to-r from-blue-600 to-blue-700 rounded-lg p-4 sm:p-5 lg:p-6 text-left hover:from-blue-700 hover:to-blue-800 transition-all duration-200 touch-manipulation"
        >
          <div className="flex items-center justify-between">
            <div className="flex items-center">
              <i className="fas fa-lightbulb mr-3 sm:mr-4 text-white text-lg sm:text-xl flex-shrink-0"></i>
              <div>
                <h3 className="text-base sm:text-lg font-semibold text-white">Business Ideas List</h3>
                <p className="text-blue-100 text-xs sm:text-sm mt-1">Value-add improvements and waste reduction suggestions</p>
              </div>
            </div>
            <i className={`fas fa-chevron-${expandedSections.business ? 'up' : 'down'} text-white text-base sm:text-lg transition-transform duration-200 flex-shrink-0`}></i>
          </div>
        </button>
        {expandedSections.business && (
          <div className="bg-white rounded-b-lg border border-gray-200">
            <div className="p-4 sm:p-6 lg:p-8 space-y-4">
              {businessIdeasColumns.map((column: any, index: number) => (
                <div key={index} className="border-l-4 border-blue-200 pl-4 py-3">
                  <div className="mb-1">
                    <h4 className="font-medium text-gray-900 text-xs sm:text-sm">{column.name}</h4>
                    <span className="text-xs text-blue-600">{column.type}</span>
                  </div>
                  <p className="text-xs text-gray-600">{column.description}</p>
                  {renderColumnValues(column.values)}
                </div>
              ))}
            </div>
          </div>
        )}
      </div>

      {/* Safety Ideas Table */}
      <div className="mb-6 sm:mb-8">
        <button
          onClick={() => toggleSection("safety")}
          className="w-full bg-gradient-to-r from-green-600 to-green-700 rounded-lg p-4 sm:p-5 lg:p-6 text-left hover:from-green-700 hover:to-green-800 transition-all duration-200 touch-manipulation"
        >
          <div className="flex items-center justify-between">
            <div className="flex items-center">
              <i className="fas fa-shield-alt mr-3 sm:mr-4 text-white text-lg sm:text-xl flex-shrink-0"></i>
              <div>
                <h3 className="text-base sm:text-lg font-semibold text-white">Safety Ideas List</h3>
                <p className="text-green-100 text-xs sm:text-sm mt-1">Safety improvements, hazard reports, and protective measures</p>
              </div>
            </div>
            <i className={`fas fa-chevron-${expandedSections.safety ? 'up' : 'down'} text-white text-base sm:text-lg transition-transform duration-200 flex-shrink-0`}></i>
          </div>
        </button>
        {expandedSections.safety && (
          <div className="bg-white rounded-b-lg border border-gray-200">
            <div className="p-4 sm:p-6 lg:p-8 space-y-4">
              {safetyIdeasColumns.map((column: any, index: number) => (
                <div key={index} className="border-l-4 border-green-200 pl-4 py-3">
                  <div className="mb-1">
                    <h4 className="font-medium text-gray-900 text-xs sm:text-sm">{column.name}</h4>
                    <span className="text-xs text-green-600">{column.type}</span>
                  </div>
                  <p className="text-xs text-gray-600">{column.description}</p>
                  {renderColumnValues(column.values)}
                </div>
              ))}
            </div>
          </div>
        )}
      </div>

      {/* Near Miss Table */}
      <div className="mb-6 sm:mb-8">
        <button
          onClick={() => toggleSection("nearmiss")}
          className="w-full bg-gradient-to-r from-red-600 to-red-700 rounded-lg p-4 sm:p-5 lg:p-6 text-left hover:from-red-700 hover:to-red-800 transition-all duration-200 touch-manipulation"
        >
          <div className="flex items-center justify-between">
            <div className="flex items-center">
              <i className="fas fa-exclamation-triangle mr-3 sm:mr-4 text-white text-lg sm:text-xl flex-shrink-0"></i>
              <div>
                <h3 className="text-base sm:text-lg font-semibold text-white">Accident Safety Register</h3>
                <p className="text-red-100 text-xs sm:text-sm mt-1">Incident reporting and property damage documentation</p>
              </div>
            </div>
            <i className={`fas fa-chevron-${expandedSections.nearmiss ? 'up' : 'down'} text-white text-base sm:text-lg transition-transform duration-200 flex-shrink-0`}></i>
          </div>
        </button>
        {expandedSections.nearmiss && (
          <div className="bg-white rounded-b-lg border border-gray-200">
            <div className="p-4 sm:p-6 lg:p-8 space-y-4">
              {nearMissColumns.map((column: any, index: number) => (
                <div key={index} className="border-l-4 border-red-200 pl-4 py-3">
                  <div className="mb-1">
                    <h4 className="font-medium text-gray-900 text-xs sm:text-sm">{column.name}</h4>
                    <span className="text-xs text-red-600">{column.type}</span>
                  </div>
                  <p className="text-xs text-gray-600">{column.description}</p>
                  {renderColumnValues(column.values)}
                </div>
              ))}
            </div>
          </div>
        )}
      </div>

      {/* Action List Table */}
      <div>
        <button
          onClick={() => toggleSection("action")}
          className="w-full bg-gradient-to-r from-orange-600 to-orange-700 rounded-lg p-4 sm:p-5 lg:p-6 text-left hover:from-orange-700 hover:to-orange-800 transition-all duration-200 touch-manipulation"
        >
          <div className="flex items-center justify-between">
            <div className="flex items-center">
              <i className="fas fa-tasks mr-3 sm:mr-4 text-white text-lg sm:text-xl flex-shrink-0"></i>
              <div>
                <h3 className="text-base sm:text-lg font-semibold text-white">Central Action List</h3>
                <p className="text-orange-100 text-xs sm:text-sm mt-1">Central tracking for all approved actions and their completion status</p>
              </div>
            </div>
            <i className={`fas fa-chevron-${expandedSections.action ? 'up' : 'down'} text-white text-base sm:text-lg transition-transform duration-200 flex-shrink-0`}></i>
          </div>
        </button>
        {expandedSections.action && (
          <div className="bg-white rounded-b-lg border border-gray-200">
            <div className="p-4 sm:p-6 lg:p-8 space-y-4">
              {actionListColumns.map((column: any, index: number) => (
                <div key={index} className="border-l-4 border-orange-200 pl-4 py-3">
                  <div className="mb-1">
                    <h4 className="font-medium text-gray-900 text-xs sm:text-sm">{column.name}</h4>
                    <span className="text-xs text-orange-600">{column.type}</span>
                  </div>
                  <p className="text-xs text-gray-600">{column.description}</p>
                  {renderColumnValues(column.values)}
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    </section>
  );
}

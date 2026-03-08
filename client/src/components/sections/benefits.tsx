export default function Benefits() {
  const keyBenefits = [
    "Centralised Health & Safety tracking",
    "Automated compliance record-keeping", 
    "Real-time Teams notifications",
    "Streamlined meeting processes"
  ];

  const implementedFeatures = [
    {
      title: "AI-Powered Title Generation",
      icon: "fas fa-robot",
      description: "Smart title generation for submissions using OpenAI GPT-4 for professional documentation",
      details: [
        "Automatically generates descriptive titles for blank submissions",
        "Analyses content to create meaningful, searchable titles",
        "Bulk processing for multiple items at once",
        "Fallback system when AI is unavailable"
      ],
      status: "Implemented"
    },
    {
      title: "SharePoint Integration",
      icon: "fas fa-cloud",
      description: "Real-time data synchronisation with Microsoft SharePoint across multiple lists",
      details: [
        "Live data from Business Ideas, Safety Ideas, and Near Miss lists",
        "Proper authentication with Microsoft 365",
        "Clean data processing with HTML content stripping",
        "Person field resolution for staff names"
      ],
      status: "Implemented"
    },
    {
      title: "Professional Document Export",
      icon: "fas fa-file-export",
      description: "Enterprise-quality meeting minutes in HTML and CSV formats for maximum compatibility",
      details: [
        "Professional HTML documents with Cranfield Glass branding and A4 print formatting",
        "Structured CSV exports perfect for Excel analysis and reporting",
        "Dynamic attendance sections that sync with UI selections",
        "Clean white styling and professional typography for compliance documents"
      ],
      status: "Implemented"
    },
    {
      title: "Secure Document Access",
      icon: "fas fa-shield-alt",
      description: "Protected access to Health & Safety policy with clean URLs",
      details: [
        "Microsoft 365 authentication required",
        "Token validation for security",
        "Clean domain URLs hiding SharePoint complexity",
        "Professional error handling"
      ],
      status: "Implemented"
    }
  ];

  return (
    <section id="benefits" className="bg-gradient-to-br from-white to-gray-50 rounded-xl shadow-lg p-4 sm:p-6 lg:p-8 mb-8 border border-gray-200">
      <div className="flex items-center space-x-3 mb-6">
        <div className="w-10 h-10 bg-gradient-to-r from-blue-600 to-blue-700 rounded-xl flex items-center justify-center shadow-lg flex-shrink-0">
          <i className="fas fa-star text-white text-lg"></i>
        </div>
        <div className="min-w-0 flex-1">
          <h2 className="text-xl sm:text-2xl font-bold ms-gray-900">Key Benefits</h2>
          <p className="text-xs sm:text-sm text-gray-600">System advantages and improvements</p>
        </div>
      </div>
      
      {/* Core Benefits */}
      <div className="bg-white border border-gray-200 rounded-xl p-4 sm:p-6 shadow-sm mb-6">
        <h3 className="font-semibold text-gray-900 mb-4 text-sm">Core System Benefits</h3>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          {keyBenefits.map((benefit, index) => (
            <div key={index} className="flex items-center space-x-3">
              <i className="fas fa-check-circle text-green-600 text-sm flex-shrink-0"></i>
              <span className="text-sm text-gray-700">{benefit}</span>
            </div>
          ))}
        </div>
      </div>

      {/* Implementation Features */}
      <div className="bg-white border border-gray-200 rounded-xl p-4 sm:p-6 shadow-sm">
        <h3 className="font-semibold text-gray-900 mb-4 text-sm">Implemented System Features</h3>
        
        {/* Summary Cards */}
        <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
          {implementedFeatures.map((feature, index) => (
            <div key={index} className="bg-gray-50 rounded-lg border border-gray-200 p-4 text-center hover:shadow-md transition-shadow duration-200">
              <div className="w-8 h-8 bg-gradient-to-r from-emerald-500 to-emerald-600 rounded-lg flex items-center justify-center mx-auto mb-2">
                <i className={`${feature.icon} text-white text-sm`}></i>
              </div>
              <h4 className="font-semibold text-gray-900 text-xs sm:text-sm mb-1">{feature.title}</h4>
              <span className="px-2 py-1 rounded-full text-xs font-medium bg-emerald-50 text-emerald-700 border border-emerald-200">
                {feature.status}
              </span>
            </div>
          ))}
        </div>

        {/* Detailed View - Collapsible */}
        <details className="bg-gray-50 rounded-lg border border-gray-200 overflow-hidden">
          <summary className="cursor-pointer p-4 bg-gray-100 hover:bg-gray-200 transition-colors duration-200 flex items-center justify-between">
            <div className="flex items-center space-x-2">
              <i className="fas fa-list-ul text-gray-600"></i>
              <span className="font-medium text-gray-900 text-sm">View Detailed Features</span>
            </div>
            <i className="fas fa-chevron-down text-gray-400 text-sm"></i>
          </summary>
          
          <div className="p-4 space-y-4">
            {implementedFeatures.map((feature, index) => (
              <div key={index} className="border-b border-gray-100 last:border-b-0 pb-4 last:pb-0">
                <div className="flex items-start space-x-3 mb-2">
                  <div className="w-6 h-6 bg-gradient-to-r from-emerald-500 to-emerald-600 rounded flex items-center justify-center flex-shrink-0">
                    <i className={`${feature.icon} text-white text-xs`}></i>
                  </div>
                  <div className="flex-1">
                    <h4 className="font-semibold text-gray-900 text-sm">{feature.title}</h4>
                    <p className="text-gray-600 text-xs mb-2">{feature.description}</p>
                    <div className="space-y-1">
                      {feature.details.map((detail, idx) => (
                        <div key={idx} className="text-xs text-gray-700">
                          <span className="text-emerald-600 font-medium">✓ </span>
                          <span>{detail}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              </div>
            ))}
          </div>
        </details>
      </div>
    </section>
  );
}
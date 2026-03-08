import { MicrosoftLogo } from "../ui/microsoft-logo";

export default function VisualOverview() {
  return (
    <section id="visual-overview" className="bg-gradient-to-br from-white to-gray-50 rounded-xl shadow-lg p-4 sm:p-6 lg:p-8 mb-8 border border-gray-200">
      <div className="flex items-center space-x-3 mb-6">
        <div className="w-10 h-10 bg-gradient-to-r from-blue-600 to-blue-700 rounded-xl flex items-center justify-center shadow-lg flex-shrink-0">
          <i className="fas fa-sitemap text-white text-lg"></i>
        </div>
        <div className="min-w-0 flex-1">
          <h2 className="text-xl sm:text-2xl font-bold ms-gray-900">Teams Viva & Data Flow</h2>
          <p className="text-xs sm:text-sm text-gray-600">Simple workflow from forms to actions</p>
        </div>
      </div>

      {/* Microsoft Teams Viva Card Access Point */}
      <div className="bg-gradient-to-r from-purple-50 to-indigo-50 rounded-xl p-4 sm:p-6 mb-6 border border-purple-200">
        <div className="flex items-center space-x-3 mb-4">
          <div className="w-8 h-8 bg-purple-600 rounded-lg flex items-center justify-center flex-shrink-0">
            <MicrosoftLogo size={16} variant="white" />
          </div>
          <div className="min-w-0 flex-1">
            <h3 className="text-base font-semibold text-purple-900">Teams Viva Card Access</h3>
            <p className="text-sm text-purple-700">Staff click Viva card in Teams for all forms</p>
          </div>
        </div>
        
        <div className="bg-white rounded-lg p-4 border border-purple-200">
          <p className="text-sm text-gray-700 mb-3">
            Staff access all Health & Safety forms directly from Microsoft Teams through a dedicated Viva card.
          </p>
          <div className="space-y-2">
            <div className="flex items-center space-x-2">
              <i className="fas fa-lightbulb text-purple-600 text-sm flex-shrink-0"></i>
              <span className="text-sm text-gray-900">Business Ideas</span>
            </div>
            <div className="flex items-center space-x-2">
              <i className="fas fa-shield-alt text-purple-600 text-sm flex-shrink-0"></i>
              <span className="text-sm text-gray-900">Safety Ideas</span>
            </div>
            <div className="flex items-center space-x-2">
              <i className="fas fa-exclamation-triangle text-purple-600 text-sm flex-shrink-0"></i>
              <span className="text-sm text-gray-900">Near Miss Reports</span>
            </div>
          </div>
        </div>
      </div>

      {/* Simplified Data Flow */}
      <div className="bg-gradient-to-r from-blue-50 to-indigo-50 rounded-xl p-4 sm:p-6 border border-blue-200">
        <div className="flex items-center space-x-3 mb-4">
          <div className="w-8 h-8 bg-blue-600 rounded-lg flex items-center justify-center flex-shrink-0">
            <i className="fas fa-stream text-white"></i>
          </div>
          <div className="min-w-0 flex-1">
            <h3 className="text-base font-semibold text-blue-900">Four-Step Process</h3>
            <p className="text-sm text-blue-700">From submission to completion</p>
          </div>
        </div>
        
        <div className="space-y-3">
          {/* Step 1 */}
          <div className="bg-white rounded-lg p-3 border border-blue-200">
            <div className="flex items-center space-x-3">
              <div className="w-6 h-6 bg-blue-600 rounded-full flex items-center justify-center flex-shrink-0">
                <span className="text-white text-xs font-bold">1</span>
              </div>
              <div className="min-w-0 flex-1">
                <h4 className="font-medium text-gray-900">Forms</h4>
                <p className="text-xs text-gray-600">Submit via Teams Viva</p>
              </div>
            </div>
          </div>
          
          {/* Arrow */}
          <div className="flex justify-center">
            <i className="fas fa-arrow-down text-blue-500"></i>
          </div>
          
          {/* Step 2 */}
          <div className="bg-white rounded-lg p-3 border border-green-200">
            <div className="flex items-center space-x-3">
              <div className="w-6 h-6 bg-green-600 rounded-full flex items-center justify-center flex-shrink-0">
                <span className="text-white text-xs font-bold">2</span>
              </div>
              <div className="min-w-0 flex-1">
                <h4 className="font-medium text-gray-900">Lists</h4>
                <p className="text-xs text-gray-600">Auto-stored in SharePoint</p>
              </div>
            </div>
          </div>
          
          {/* Arrow */}
          <div className="flex justify-center">
            <i className="fas fa-arrow-down text-green-500"></i>
          </div>
          
          {/* Step 3 */}
          <div className="bg-white rounded-lg p-3 border border-purple-200">
            <div className="flex items-center space-x-3">
              <div className="w-6 h-6 bg-purple-600 rounded-full flex items-center justify-center flex-shrink-0">
                <span className="text-white text-xs font-bold">3</span>
              </div>
              <div className="min-w-0 flex-1">
                <h4 className="font-medium text-gray-900">Teams</h4>
                <p className="text-xs text-gray-600">Meeting decisions</p>
              </div>
            </div>
          </div>
          
          {/* Arrow */}
          <div className="flex justify-center">
            <i className="fas fa-arrow-down text-purple-500"></i>
          </div>
          
          {/* Step 4 */}
          <div className="bg-white rounded-lg p-3 border border-orange-200">
            <div className="flex items-center space-x-3">
              <div className="w-6 h-6 bg-orange-600 rounded-full flex items-center justify-center flex-shrink-0">
                <span className="text-white text-xs font-bold">4</span>
              </div>
              <div className="min-w-0 flex-1">
                <h4 className="font-medium text-gray-900">Actions</h4>
                <p className="text-xs text-gray-600">Track completion</p>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Power Automate Integration */}
      <div className="bg-white rounded-lg border border-gray-200 p-4 sm:p-6 mt-6">
        <div className="flex items-center space-x-3 mb-4">
          <div className="w-8 h-8 bg-blue-600 rounded-lg flex items-center justify-center flex-shrink-0">
            <i className="fas fa-cogs text-white"></i>
          </div>
          <div className="min-w-0 flex-1">
            <h3 className="text-base font-semibold text-gray-900">Power Automate</h3>
            <p className="text-sm text-gray-600">Automated workflows handle the connections</p>
          </div>
        </div>
        
        <div className="space-y-2">
          <div className="flex items-start space-x-2">
            <i className="fas fa-check text-green-600 text-xs mt-1 flex-shrink-0"></i>
            <span className="text-sm text-gray-700">Form submissions trigger automatic list creation</span>
          </div>
          <div className="flex items-start space-x-2">
            <i className="fas fa-check text-green-600 text-xs mt-1 flex-shrink-0"></i>
            <span className="text-sm text-gray-700">Teams notifications sent instantly</span>
          </div>
          <div className="flex items-start space-x-2">
            <i className="fas fa-check text-green-600 text-xs mt-1 flex-shrink-0"></i>
            <span className="text-sm text-gray-700">Action items created from approved submissions</span>
          </div>
        </div>
      </div>
    </section>
  );
}
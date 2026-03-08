export default function Tools() {
  return (
    <section id="tools" className="bg-gradient-to-br from-white to-gray-50 rounded-xl shadow-lg p-8 mb-8 border border-gray-200">
      <div className="flex items-center space-x-3 mb-6">
        <div className="w-10 h-10 bg-gradient-to-r from-blue-600 to-blue-700 rounded-xl flex items-center justify-center shadow-lg">
          <i className="fas fa-tools text-white text-lg"></i>
        </div>
        <div>
          <h2 className="text-2xl font-bold ms-gray-900">Core Technologies</h2>
          <p className="text-sm text-gray-600">Microsoft 365 platform components</p>
        </div>
      </div>
      
      <div className="grid md:grid-cols-1 gap-6">
        <div className="bg-white border border-gray-200 rounded-xl p-6 shadow-sm">
          <div className="flex items-start mb-4">
            <div className="w-10 h-10 bg-blue-500 rounded-xl flex items-center justify-center mr-4">
              <i className="fas fa-wpforms text-white text-sm"></i>
            </div>
            <div className="flex-1">
              <h3 className="font-semibold ms-gray-900 text-lg">Microsoft Forms & Lists</h3>
              <p className="text-sm text-gray-600">Data collection and storage</p>
            </div>
          </div>
          <div className="ml-14">
            <p className="text-sm ms-gray-700">Three submission forms feed into SharePoint Lists for structured data management and tracking.</p>
          </div>
        </div>

        <div className="bg-white border border-gray-200 rounded-xl p-6 shadow-sm">
          <div className="flex items-start mb-4">
            <div className="w-10 h-10 bg-blue-500 rounded-xl flex items-center justify-center mr-4">
              <i className="fas fa-bolt text-white text-sm"></i>
            </div>
            <div className="flex-1">
              <h3 className="font-semibold ms-gray-900 text-lg">Power Automate & Teams</h3>
              <p className="text-sm text-gray-600">Automation and communication</p>
            </div>
          </div>
          <div className="ml-14">
            <p className="text-sm ms-gray-700">Automated workflows handle data processing and notifications through Microsoft Teams integration.</p>
          </div>
        </div>
      </div>
    </section>
  );
}

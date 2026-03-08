export default function Automation() {
  const flows = [
    {
      title: "Form to List Connection",
      icon: "fas fa-play",
      bgColor: "bg-blue-400",
      description: "Triggered when new response submitted to Microsoft Form",
      items: [
        "Gets response details",
        "Gets user profile", 
        "Creates list item",
        "Posts Teams notification",
        "Calculates next meeting date"
      ]
    },
    {
      title: "Create Action Item",
      icon: "fas fa-copy",
      bgColor: "ms-green",
      description: "Triggered when item status changed to \"Actioned\"",
      items: [
        "Monitors list modifications",
        "Checks for \"Actioned\" status",
        "Cleans HTML content",
        "Creates action list item",
        "Links to original item"
      ]
    },
    {
      title: "Update Related Items",
      icon: "fas fa-sync",
      bgColor: "bg-yellow-500",
      description: "Triggered when action marked \"Completed\"",
      items: [
        "Monitors action list changes",
        "Filters for \"Completed\" status",
        "Updates original item",
        "Sets status to \"Closed\"",
        "Removes from board view"
      ]
    }
  ];

  return (
    <section id="automation" className="bg-gradient-to-br from-white to-gray-50 rounded-xl shadow-lg p-8 mb-8 border border-gray-200">
      <div className="flex items-center space-x-3 mb-6">
        <div className="w-10 h-10 bg-gradient-to-r from-blue-600 to-blue-700 rounded-xl flex items-center justify-center shadow-lg">
          <i className="fas fa-bolt text-white text-lg"></i>
        </div>
        <div>
          <h2 className="text-2xl font-bold ms-gray-900">Power Automate Flows</h2>
          <p className="text-sm text-gray-600">Automated workflow processes</p>
        </div>
      </div>
      
      <div className="grid md:grid-cols-1 gap-6">
        {flows.map((flow, index) => (
          <div key={index} className="bg-white border border-gray-200 rounded-xl p-6 shadow-sm">
            <div className="flex items-start mb-4">
              <div className="w-10 h-10 bg-blue-500 rounded-xl flex items-center justify-center mr-4">
                <i className={`${flow.icon} text-white text-sm`}></i>
              </div>
              <div className="flex-1">
                <h3 className="font-semibold text-gray-900 text-lg">{flow.title}</h3>
                <p className="text-sm text-gray-600">{flow.description}</p>
              </div>
            </div>
            
            <div className="ml-14">
              <ul className="space-y-2">
                {flow.items.map((item, itemIndex) => (
                  <li key={itemIndex} className="flex items-start space-x-2 text-sm text-gray-700">
                    <i className="fas fa-arrow-right text-blue-500 text-xs mt-1"></i>
                    <span>{item}</span>
                  </li>
                ))}
              </ul>
            </div>
          </div>
        ))}
      </div>
    </section>
  );
}

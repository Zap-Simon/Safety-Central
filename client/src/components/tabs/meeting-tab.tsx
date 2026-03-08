import { Link } from "wouter";
import { Calendar, FileText, Users, Clock, CheckCircle } from "lucide-react";

export default function MeetingTab() {
  return (
    <div className="p-6">
      <div className="max-w-4xl mx-auto">
        {/* Header */}
        <div className="text-center mb-8">
          <div className="w-16 h-16 bg-blue-100 rounded-full flex items-center justify-center mx-auto mb-4">
            <Calendar className="text-blue-600" size={32} />
          </div>
          <h1 className="text-2xl font-bold text-gray-900 mb-2">Meeting Minutes System</h1>
          <p className="text-gray-600">
            Manage Health & Safety meetings, capture ideas, and generate professional minutes
          </p>
        </div>

        {/* Quick Access */}
        <div className="bg-gradient-to-r from-blue-50 to-indigo-50 rounded-xl p-6 mb-8">
          <div className="text-center">
            <h2 className="text-lg font-semibold text-gray-900 mb-4">Ready to start a meeting?</h2>
            <Link href="/meeting-history">
              <button className="bg-blue-600 hover:bg-blue-700 text-white px-8 py-3 rounded-lg font-medium flex items-center space-x-2 mx-auto transition-colors">
                <Calendar size={20} />
                <span>Access Meeting System</span>
              </button>
            </Link>
          </div>
        </div>

        {/* Features */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-8">
          <div className="bg-white border border-gray-200 rounded-xl p-6">
            <div className="w-10 h-10 bg-green-100 rounded-lg flex items-center justify-center mb-4">
              <FileText className="text-green-600" size={20} />
            </div>
            <h3 className="font-semibold text-gray-900 mb-2">Professional Minutes</h3>
            <p className="text-gray-600 text-sm">
              Generate clean, branded meeting minutes in HTML and CSV formats with proper formatting and structure.
            </p>
          </div>

          <div className="bg-white border border-gray-200 rounded-xl p-6">
            <div className="w-10 h-10 bg-blue-100 rounded-lg flex items-center justify-center mb-4">
              <Users className="text-blue-600" size={20} />
            </div>
            <h3 className="font-semibold text-gray-900 mb-2">Attendance Tracking</h3>
            <p className="text-gray-600 text-sm">
              Track who attended meetings, manage attendance lists, and keep records for compliance purposes.
            </p>
          </div>

          <div className="bg-white border border-gray-200 rounded-xl p-6">
            <div className="w-10 h-10 bg-orange-100 rounded-lg flex items-center justify-center mb-4">
              <Clock className="text-orange-600" size={20} />
            </div>
            <h3 className="font-semibold text-gray-900 mb-2">Action Items</h3>
            <p className="text-gray-600 text-sm">
              Track action items from meetings, assign responsibilities, and follow up on completion status.
            </p>
          </div>

          <div className="bg-white border border-gray-200 rounded-xl p-6">
            <div className="w-10 h-10 bg-purple-100 rounded-lg flex items-center justify-center mb-4">
              <CheckCircle className="text-purple-600" size={20} />
            </div>
            <h3 className="font-semibold text-gray-900 mb-2">Ideas Management</h3>
            <p className="text-gray-600 text-sm">
              Capture business ideas, safety improvements, and near-miss reports directly in meetings.
            </p>
          </div>
        </div>

        {/* How it Works */}
        <div className="bg-white border border-gray-200 rounded-xl p-6">
          <h3 className="font-semibold text-gray-900 mb-4">How it works</h3>
          <div className="space-y-4">
            <div className="flex items-start space-x-3">
              <div className="w-6 h-6 bg-blue-100 rounded-full flex items-center justify-center flex-shrink-0 mt-0.5">
                <span className="text-blue-600 text-xs font-medium">1</span>
              </div>
              <div>
                <h4 className="font-medium text-gray-900">Start a meeting</h4>
                <p className="text-gray-600 text-sm">Choose a meeting date and select attendees from your team</p>
              </div>
            </div>
            <div className="flex items-start space-x-3">
              <div className="w-6 h-6 bg-blue-100 rounded-full flex items-center justify-center flex-shrink-0 mt-0.5">
                <span className="text-blue-600 text-xs font-medium">2</span>
              </div>
              <div>
                <h4 className="font-medium text-gray-900">Review items</h4>
                <p className="text-gray-600 text-sm">Go through business ideas, safety suggestions, and near-miss reports</p>
              </div>
            </div>
            <div className="flex items-start space-x-3">
              <div className="w-6 h-6 bg-blue-100 rounded-full flex items-center justify-center flex-shrink-0 mt-0.5">
                <span className="text-blue-600 text-xs font-medium">3</span>
              </div>
              <div>
                <h4 className="font-medium text-gray-900">Generate minutes</h4>
                <p className="text-gray-600 text-sm">Export professional meeting minutes with all discussions and action items</p>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
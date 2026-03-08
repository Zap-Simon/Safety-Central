import { Link } from "wouter";
import { BookOpen, Play, Users, Award } from "lucide-react";

export default function TrainingTab() {
  return (
    <div className="p-6">
      <div className="max-w-4xl mx-auto">
        {/* Header */}
        <div className="text-center mb-8">
          <div className="w-16 h-16 bg-green-100 rounded-full flex items-center justify-center mx-auto mb-4">
            <BookOpen className="text-green-600" size={32} />
          </div>
          <h1 className="text-2xl font-bold text-gray-900 mb-2">Staff Training Portal</h1>
          <p className="text-gray-600">
            AroFlo setup, workflows, and system training resources
          </p>
        </div>

        {/* Quick Access */}
        <div className="bg-gradient-to-r from-green-50 to-emerald-50 rounded-xl p-6 mb-8">
          <div className="text-center">
            <h2 className="text-lg font-semibold text-gray-900 mb-4">Access Training Materials</h2>
            <Link href="/staff-training">
              <button className="bg-green-600 hover:bg-green-700 text-white px-8 py-3 rounded-lg font-medium flex items-center space-x-2 mx-auto transition-colors">
                <Play size={20} />
                <span>Start Training</span>
              </button>
            </Link>
          </div>
        </div>

        {/* Training Areas */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-8">
          <div className="bg-white border border-gray-200 rounded-xl p-6">
            <div className="w-10 h-10 bg-blue-100 rounded-lg flex items-center justify-center mb-4">
              <BookOpen className="text-blue-600" size={20} />
            </div>
            <h3 className="font-semibold text-gray-900 mb-2">AroFlo Basics</h3>
            <p className="text-gray-600 text-sm">
              Learn the fundamentals of AroFlo system - from navigation to basic operations and daily workflows.
            </p>
          </div>

          <div className="bg-white border border-gray-200 rounded-xl p-6">
            <div className="w-10 h-10 bg-purple-100 rounded-lg flex items-center justify-center mb-4">
              <Users className="text-purple-600" size={20} />
            </div>
            <h3 className="font-semibold text-gray-900 mb-2">Team Workflows</h3>
            <p className="text-gray-600 text-sm">
              Understand how different team roles interact with AroFlo and coordinate work effectively.
            </p>
          </div>

          <div className="bg-white border border-gray-200 rounded-xl p-6">
            <div className="w-10 h-10 bg-orange-100 rounded-lg flex items-center justify-center mb-4">
              <Play className="text-orange-600" size={20} />
            </div>
            <h3 className="font-semibold text-gray-900 mb-2">Video Tutorials</h3>
            <p className="text-gray-600 text-sm">
              Step-by-step video guides for common tasks and advanced features in AroFlo.
            </p>
          </div>

          <div className="bg-white border border-gray-200 rounded-xl p-6">
            <div className="w-10 h-10 bg-green-100 rounded-lg flex items-center justify-center mb-4">
              <Award className="text-green-600" size={20} />
            </div>
            <h3 className="font-semibold text-gray-900 mb-2">Best Practices</h3>
            <p className="text-gray-600 text-sm">
              Learn proven methods and tips for getting the most out of AroFlo in your daily work.
            </p>
          </div>
        </div>

        {/* Features */}
        <div className="bg-white border border-gray-200 rounded-xl p-6">
          <h3 className="font-semibold text-gray-900 mb-4">What you'll learn</h3>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div className="flex items-center space-x-3">
              <div className="w-2 h-2 bg-green-500 rounded-full"></div>
              <span className="text-gray-700 text-sm">AroFlo system navigation</span>
            </div>
            <div className="flex items-center space-x-3">
              <div className="w-2 h-2 bg-green-500 rounded-full"></div>
              <span className="text-gray-700 text-sm">Job management workflows</span>
            </div>
            <div className="flex items-center space-x-3">
              <div className="w-2 h-2 bg-green-500 rounded-full"></div>
              <span className="text-gray-700 text-sm">Customer communication</span>
            </div>
            <div className="flex items-center space-x-3">
              <div className="w-2 h-2 bg-green-500 rounded-full"></div>
              <span className="text-gray-700 text-sm">Time tracking and reporting</span>
            </div>
            <div className="flex items-center space-x-3">
              <div className="w-2 h-2 bg-green-500 rounded-full"></div>
              <span className="text-gray-700 text-sm">Quality control processes</span>
            </div>
            <div className="flex items-center space-x-3">
              <div className="w-2 h-2 bg-green-500 rounded-full"></div>
              <span className="text-gray-700 text-sm">Troubleshooting common issues</span>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
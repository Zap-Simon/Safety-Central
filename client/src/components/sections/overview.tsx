import { Users, Info, Target, Heart, Settings, Quote, TrendingUp, Shield, Star, RotateCcw, FileText, List, Zap, UserCheck, Lightbulb, CheckCircle, Workflow, Calendar } from 'lucide-react';
import { Link } from "wouter";

export default function Overview() {
  return (
    <section id="overview" className="space-y-6">
      {/* Staff Engagement System Card */}
      <div className="bg-gradient-to-br from-white to-gray-50 rounded-xl shadow-lg p-4 sm:p-6 lg:p-8 border border-gray-200">
        <div className="flex items-center space-x-3">
          <div className="w-10 h-10 bg-gradient-to-r from-blue-600 to-blue-700 rounded-xl flex items-center justify-center shadow-lg flex-shrink-0">
            <Users className="text-white" size={20} />
          </div>
          <div className="min-w-0 flex-1">
            <h2 className="text-xl sm:text-2xl font-bold ms-gray-900">Staff Engagement System</h2>
            <p className="text-xs sm:text-sm text-gray-600">Policy framework for continuous improvement</p>
          </div>
        </div>
      </div>
      
      {/* Meeting System Overview - Outer Container */}
      <div className="bg-gradient-to-br from-white to-gray-50 rounded-xl shadow-lg p-4 sm:p-6 lg:p-8 border border-gray-200 space-y-6">
        
        {/* Meeting Minutes & Idea Management */}
        <div className="bg-white border border-gray-200 rounded-xl p-4 sm:p-6 shadow-sm">
          <div className="flex items-start mb-4">
            <div className="w-10 h-10 bg-gradient-to-r from-blue-600 to-blue-700 rounded-xl flex items-center justify-center mr-3 sm:mr-4 flex-shrink-0">
              <Calendar className="text-white" size={20} />
            </div>
            <div className="min-w-0 flex-1">
              <h3 className="font-semibold text-gray-900 text-base sm:text-lg">Meeting Minutes & Idea Management</h3>
              <p className="text-xs sm:text-sm text-gray-600">Staff engagement meeting solution</p>
            </div>
          </div>
          <p className="text-sm text-gray-700">
            This engagement system captures staff business ideas, safety improvements, and near-miss incidents through structured meetings. 
            Generate professional meeting minutes, track attendance, and follow up on action items.
          </p>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4 sm:gap-6">
          <div className="bg-white border border-gray-200 rounded-xl p-4 sm:p-6 shadow-sm">
            <div className="flex items-start mb-4">
              <div className="w-10 h-10 bg-blue-500 rounded-xl flex items-center justify-center mr-3 sm:mr-4 flex-shrink-0">
                <Target className="text-white" size={20} />
              </div>
              <div className="min-w-0 flex-1">
                <h3 className="font-semibold text-gray-900 text-base sm:text-lg">Policy Framework</h3>
                <p className="text-xs sm:text-sm text-gray-600">Health & Safety Policy implementation</p>
              </div>
            </div>
            <p className="text-sm text-gray-700 mb-3">
              Implements Section 4.4 Staff Engagement of our Health & Safety Policy for structured idea management.
            </p>
            <div className="bg-blue-50 rounded-lg p-3 border border-blue-200">
              <p className="text-xs text-blue-700 italic">
                "Staff actively contribute to Health, Safety, and business improvement ideas as a key part of everyone's role."
              </p>
            </div>
          </div>

          <div className="bg-white border border-gray-200 rounded-xl p-4 sm:p-6 shadow-sm">
            <div className="flex items-start mb-4">
              <div className="w-10 h-10 bg-green-500 rounded-xl flex items-center justify-center mr-3 sm:mr-4 flex-shrink-0">
                <Heart className="text-white" size={20} />
              </div>
              <div className="min-w-0 flex-1">
                <h3 className="font-semibold text-gray-900 text-base sm:text-lg">Core Values</h3>
                <p className="text-xs sm:text-sm text-gray-600">Cultural principles and mindset</p>
              </div>
            </div>
            <div className="space-y-3 text-sm text-gray-700">
              <div className="flex items-center space-x-2">
                <i className="fas fa-arrow-up text-blue-500 text-xs"></i>
                <span>Continuous improvement culture</span>
              </div>
              <div className="flex items-center space-x-2">
                <i className="fas fa-shield-alt text-green-500 text-xs"></i>
                <span>Protect our people priority</span>
              </div>
              <div className="flex items-center space-x-2">
                <i className="fas fa-star text-yellow-500 text-xs"></i>
                <span>Go the extra mile mindset</span>
              </div>
            </div>
          </div>
        </div>

        <div className="bg-white border border-gray-200 rounded-xl p-4 sm:p-6 shadow-sm">
          <div className="flex items-start mb-4">
            <div className="w-10 h-10 bg-blue-500 rounded-xl flex items-center justify-center mr-3 sm:mr-4 flex-shrink-0">
              <Settings className="text-white" size={20} />
            </div>
            <div className="min-w-0 flex-1">
              <h3 className="font-semibold text-gray-900 text-base sm:text-lg">Digital Transformation</h3>
              <p className="text-xs sm:text-sm text-gray-600">Microsoft 365-based system implementation</p>
            </div>
          </div>
          
          <p className="text-sm text-gray-700">
            This Microsoft 365-based system replaces our previous Trello-based approach while maintaining core principles. 
            The system integrates Forms, Lists, Teams, and SharePoint through automated workflows.
          </p>
        </div>
        
      </div>
    </section>
  );
}

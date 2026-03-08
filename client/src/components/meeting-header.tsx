import { Link } from "wouter";
import { Home, ArrowLeft } from "lucide-react";
import UserAuth from "@/components/user-auth";

export default function MeetingHeader() {
  return (
    <header className="bg-gradient-to-r from-blue-600 to-blue-700 text-white shadow-lg fixed top-0 left-0 right-0 z-30">
      <div className="container mx-auto px-4">
        <div className="flex items-center justify-between h-16">
          {/* Logo and Title */}
          <div className="flex items-center space-x-2 sm:space-x-3">
            <div 
              className="w-8 h-8 sm:w-10 sm:h-10 bg-white bg-opacity-20 rounded-lg flex items-center justify-center flex-shrink-0"
              title="Cranfield Glass"
            >
              <Home className="w-4 h-4 sm:w-5 sm:h-5 text-white" />
            </div>
            <div className="min-w-0 flex-1">
              <h1 className="text-base sm:text-xl font-bold text-white truncate">Cranfield Glass Christchurch</h1>
              <p className="text-blue-100 text-xs sm:text-sm truncate">Meeting Minutes</p>
            </div>
          </div>

          {/* Back Button and User Authentication */}
          <div className="flex items-center space-x-3">
            <Link href="/">
              <button 
                className="bg-white bg-opacity-20 rounded-lg flex items-center space-x-2 px-3 py-2 hover:bg-opacity-30 transition-all"
                title="Back to Home"
              >
                <ArrowLeft className="w-4 h-4 text-white" />
                <span className="text-white text-sm font-medium hidden sm:inline">Back</span>
              </button>
            </Link>
            <UserAuth />
          </div>

        </div>
      </div>
    </header>
  );
}
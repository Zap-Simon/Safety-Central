import { Link } from "wouter";
import { Home, ArrowLeft } from "lucide-react";
import UserAuth from "@/components/user-auth";

export default function PolicyHeader() {
  return (
    <header className="bg-white border-b border-gray-200 shadow-sm fixed top-0 left-0 right-0 z-30">
      <div className="container mx-auto px-4">
        <div className="flex items-center justify-between h-16">
          {/* Logo and Title */}
          <div className="flex items-center space-x-2 sm:space-x-3">
            <div 
              className="w-8 h-8 sm:w-10 sm:h-10 bg-red-100 rounded-lg flex items-center justify-center flex-shrink-0"
              title="Cranfield Glass"
            >
              <Home className="w-4 h-4 sm:w-5 sm:h-5 text-red-600" />
            </div>
            <div className="min-w-0 flex-1">
              <h1 className="text-base sm:text-xl font-bold text-gray-900 truncate">Cranfield Glass Christchurch</h1>
              <p className="text-gray-600 text-xs sm:text-sm truncate">Health & Safety Policy</p>
            </div>
          </div>

          {/* Back Button and User Authentication */}
          <div className="flex items-center space-x-3">
            <Link href="/">
              <button 
                className="bg-gray-100 hover:bg-gray-200 rounded-lg flex items-center space-x-2 px-3 py-2 transition-all"
                title="Back to Home"
              >
                <ArrowLeft className="w-4 h-4 text-gray-700" />
                <span className="text-gray-700 text-sm font-medium hidden sm:inline">Back</span>
              </button>
            </Link>
            <UserAuth />
          </div>

        </div>
      </div>
    </header>
  );
}
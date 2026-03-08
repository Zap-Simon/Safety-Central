import { Link, useLocation } from "wouter";
import { useState, useEffect } from "react";
import { authService } from "@/auth/authService";
import { Shield, Globe, User, LogOut, ChevronDown, Users, BookOpen, FileText, Menu, X } from "lucide-react";
import Sidebar from "@/components/sidebar";
import UserAuth from "@/components/user-auth";

export default function Header() {
  const [location] = useLocation();
  const [isAuthenticated, setIsAuthenticated] = useState<boolean>(false);
  const [isSigningIn, setIsSigningIn] = useState<boolean>(false);
  const [userInfo, setUserInfo] = useState<any>(null);
  const [isUserDropdownOpen, setIsUserDropdownOpen] = useState<boolean>(false);
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState<boolean>(false);
  
  const isOnMeetingHistory = location.includes('/meeting-history');

  useEffect(() => {
    const checkAuth = async () => {
      try {
        await authService.initialize();
        setIsAuthenticated(authService.isAuthenticated());
        if (authService.isAuthenticated()) {
          setUserInfo(authService.getCurrentUser());
        }
      } catch (error) {
        console.error('Auth initialization failed:', error);
        setIsAuthenticated(false);
      }
    };
    checkAuth();
  }, []);

  // Prevent body scrolling when mobile menu is open
  useEffect(() => {
    if (isMobileMenuOpen) {
      document.body.style.overflow = 'hidden';
    } else {
      document.body.style.overflow = 'unset';
    }
    
    // Cleanup on unmount
    return () => {
      document.body.style.overflow = 'unset';
    };
  }, [isMobileMenuOpen]);

  // Only check auth status when window regains focus (no polling needed)
  useEffect(() => {
    const checkAuthStatus = () => {
      const currentAuthState = authService.isAuthenticated();
      if (currentAuthState !== isAuthenticated) {
        setIsAuthenticated(currentAuthState);
        if (currentAuthState) {
          setUserInfo(authService.getCurrentUser());
        } else {
          setUserInfo(null);
        }
      }
    };

    const handleFocus = () => checkAuthStatus();
    window.addEventListener('focus', handleFocus);
    
    return () => {
      window.removeEventListener('focus', handleFocus);
    };
  }, [isAuthenticated]);

  const handleSignIn = async () => {
    if (isSigningIn) return;
    
    setIsSigningIn(true);
    try {
      await authService.signIn();
      setIsAuthenticated(authService.isAuthenticated());
      if (authService.isAuthenticated()) {
        setUserInfo(authService.getCurrentUser());
      }
    } catch (error) {
      console.error('Sign in failed:', error);
    } finally {
      setIsSigningIn(false);
    }
  };

  const handleSignOut = async () => {
    try {
      await authService.signOut();
      setIsAuthenticated(false);
      setUserInfo(null);
    } catch (error) {
      console.error('Sign out failed:', error);
    }
  };



  return (
    <header className="bg-gradient-to-r from-blue-600 to-blue-700 text-white shadow-lg fixed top-0 left-0 right-0 z-30">
      <div className="container mx-auto px-4">
        <div className="flex items-center justify-between h-16">
          {/* Logo - Home Link */}
          <Link href="/">
            <div className="flex items-center space-x-2 sm:space-x-3 cursor-pointer hover:opacity-90 transition-opacity">
              <div className="w-8 h-8 sm:w-10 sm:h-10 bg-white bg-opacity-20 rounded-lg flex items-center justify-center flex-shrink-0 hover:bg-opacity-30 transition-all">
                <Shield className="w-4 h-4 sm:w-5 sm:h-5 text-white" />
              </div>
              <div className="min-w-0 flex-1 pr-2">
                <h1 className="text-base sm:text-xl font-bold text-white truncate">Cranfield Glass Christchurch</h1>
                <p className="text-blue-100 text-xs sm:text-sm truncate">HSEQ - Standards & Compliance</p>
              </div>
            </div>
          </Link>

          {/* Desktop Navigation */}
          <div className="hidden md:flex items-center space-x-2 lg:space-x-4">
            <Link href="/meeting-history">
              <button className="bg-white bg-opacity-20 hover:bg-opacity-30 text-white px-2 lg:px-4 py-2 rounded-lg text-sm font-medium transition-all duration-200 flex items-center space-x-1 lg:space-x-2">
                <FileText className="w-4 h-4 flex-shrink-0" />
                <span className="hidden lg:inline">Meeting History</span>
                <span className="lg:hidden">Meeting History</span>
              </button>
            </Link>



            <UserAuth />
          </div>

          {/* Tablet Navigation (md to lg) */}
          <div className="hidden sm:flex md:hidden items-center space-x-2">
            <Link href="/meeting-history">
              <button className="bg-white bg-opacity-20 hover:bg-opacity-30 text-white w-9 h-9 rounded-lg transition-all duration-200 flex items-center justify-center" title="Meeting History">
                <FileText className="w-4 h-4" />
              </button>
            </Link>

            <UserAuth />
          </div>

          {/* Mobile Hamburger Menu */}
          <button 
            onClick={() => setIsMobileMenuOpen(!isMobileMenuOpen)}
            className="sm:hidden w-9 h-9 bg-white bg-opacity-20 hover:bg-opacity-30 text-white rounded-lg flex items-center justify-center transition-all duration-200"
            aria-label="Open menu"
          >
            <Menu className="w-4 h-4" />
          </button>
        </div>
      </div>

      {/* Mobile Menu Overlay */}
      {isMobileMenuOpen && (
        <div 
          className="fixed inset-0 bg-black bg-opacity-50 z-50 overflow-hidden" 
          onClick={() => setIsMobileMenuOpen(false)}
        >
          <div 
            className="bg-white w-72 h-full shadow-xl ml-auto transform transition-transform duration-300 ease-in-out" 
            onClick={e => e.stopPropagation()}
          >
            {/* Mobile Menu Header */}
            <div className="h-16 border-b border-gray-200 flex items-center justify-between px-4">
              <h3 className="text-lg font-semibold text-gray-900">Menu</h3>
              <button
                onClick={() => setIsMobileMenuOpen(false)}
                className="p-2 rounded-lg hover:bg-gray-100 transition-colors"
                aria-label="Close menu"
              >
                <X className="w-5 h-5 text-gray-600" />
              </button>
            </div>
            
            {/* Mobile Menu Content */}
            <div className="flex flex-col h-full">
              {/* Action Buttons - Fixed at top */}
              <div className="p-4 border-b border-gray-100 flex-shrink-0">
                <div className="space-y-3">
                  {!isOnMeetingHistory && (
                    <Link href="/meeting-history">
                      <button 
                        className="w-full bg-blue-600 hover:bg-blue-700 text-white px-4 py-3 rounded-lg transition-all duration-200 flex items-center space-x-3"
                        onClick={() => setIsMobileMenuOpen(false)}
                      >
                        <FileText className="w-5 h-5" />
                        <span className="text-sm font-medium">Meeting History</span>
                      </button>
                    </Link>
                  )}

                  <div className="flex justify-center">
                    <UserAuth />
                  </div>
                </div>
              </div>

              {/* Navigation - Scrollable area */}
              {!isOnMeetingHistory && (
                <div className="flex-1 overflow-y-auto min-h-0">
                  <div className="p-4">
                    <div className="text-xs font-medium text-gray-500 mb-3 px-1">PAGE NAVIGATION</div>
                    <div onClick={() => setIsMobileMenuOpen(false)} className="mobile-sidebar">
                      <Sidebar />
                    </div>
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

    </header>
  );
}
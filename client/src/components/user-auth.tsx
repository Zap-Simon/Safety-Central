import { useState, useEffect } from "react";
import { authService } from "@/auth/authService";
import { User, LogOut, ChevronDown } from "lucide-react";

export default function UserAuth() {
  const [isAuthenticated, setIsAuthenticated] = useState<boolean>(false);
  const [isSigningIn, setIsSigningIn] = useState<boolean>(false);
  const [userInfo, setUserInfo] = useState<any>(null);
  const [isUserDropdownOpen, setIsUserDropdownOpen] = useState<boolean>(false);

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
      setIsUserDropdownOpen(false);
    } catch (error) {
      console.error('Sign out failed:', error);
    }
  };

  return (
    <div className="flex items-center">
      {!isAuthenticated ? (
        <button 
          onClick={handleSignIn}
          disabled={isSigningIn}
          className="w-9 h-9 bg-white bg-opacity-20 hover:bg-opacity-30 text-white rounded-lg transition-all duration-200 flex items-center justify-center disabled:opacity-50"
          title={isSigningIn ? 'Signing in...' : 'Sign In with Microsoft'}
        >
          <User className="w-4 h-4" />
        </button>
      ) : (
        <div className="relative">
          <button 
            onClick={() => setIsUserDropdownOpen(!isUserDropdownOpen)}
            className="w-9 h-9 bg-white bg-opacity-20 hover:bg-opacity-30 text-white rounded-lg transition-all duration-200 flex items-center justify-center"
            title={userInfo?.name || 'User menu'}
          >
            <User className="w-4 h-4" />
          </button>
          
          {isUserDropdownOpen && (
            <div className="absolute top-full right-0 mt-2 w-48 bg-white rounded-lg shadow-xl border z-[60]">
              <div className="p-2">
                <div className="px-3 py-2 border-b border-gray-100 mb-1">
                  <div className="text-xs text-gray-500">Signed in as:</div>
                  <div className="text-sm font-medium text-gray-900">{userInfo?.name || 'User'}</div>
                </div>
                <button 
                  onClick={handleSignOut}
                  className="w-full text-left px-3 py-2 text-sm text-gray-700 hover:bg-gray-100 rounded-lg flex items-center space-x-2"
                >
                  <LogOut className="w-4 h-4" />
                  <span>Sign Out</span>
                </button>
              </div>
            </div>
          )}
        </div>
      )}

      {/* Backdrop to close dropdown */}
      {isUserDropdownOpen && (
        <div 
          className="fixed inset-0 z-[50]" 
          onClick={() => setIsUserDropdownOpen(false)}
        />
      )}
    </div>
  );
}
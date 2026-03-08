import { useState, useEffect } from "react";
import { authService } from "./authService";
import { User, LogOut, ChevronDown } from "lucide-react";

export default function LoginComponent() {
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

  // Listen for authentication state changes
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

    // Check auth status periodically
    const authCheckInterval = setInterval(checkAuthStatus, 1000);
    
    // Also check when window regains focus
    const handleFocus = () => checkAuthStatus();
    window.addEventListener('focus', handleFocus);
    
    return () => {
      clearInterval(authCheckInterval);
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
    <div className="flex items-center space-x-4">
      {!isAuthenticated ? (
        <button 
          onClick={handleSignIn}
          disabled={isSigningIn}
          className="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-lg text-sm font-medium transition-all duration-200 flex items-center space-x-2 disabled:opacity-50"
          title="Sign In with Microsoft"
        >
          <User className="w-4 h-4" />
          <span>{isSigningIn ? 'Signing in...' : 'Sign In'}</span>
        </button>
      ) : (
        <div className="relative">
          <button
            onClick={() => setIsUserDropdownOpen(!isUserDropdownOpen)}
            className="bg-white bg-opacity-20 hover:bg-opacity-30 text-white px-3 py-2 rounded-lg text-sm font-medium transition-all duration-200 flex items-center space-x-2"
            title={`Signed in as ${userInfo?.name || userInfo?.username || 'User'}`}
          >
            <User className="w-4 h-4" />
            <span className="max-w-32 truncate">
              {userInfo?.name || userInfo?.username || 'User'}
            </span>
            <ChevronDown className={`w-3 h-3 transition-transform ${isUserDropdownOpen ? 'rotate-180' : ''}`} />
          </button>
          
          {isUserDropdownOpen && (
            <div className="absolute right-0 top-full mt-2 w-56 bg-white rounded-lg shadow-lg border border-gray-200 py-2 z-50">
              <div className="px-4 py-2 border-b border-gray-100">
                <p className="text-sm font-medium text-gray-900">
                  {userInfo?.name || userInfo?.username || 'User'}
                </p>
                <p className="text-xs text-gray-500 truncate">
                  {userInfo?.username || userInfo?.email || ''}
                </p>
              </div>
              
              <button
                onClick={handleSignOut}
                className="w-full px-4 py-2 text-left text-sm text-gray-700 hover:bg-gray-50 flex items-center space-x-2"
              >
                <LogOut className="w-4 h-4" />
                <span>Sign Out</span>
              </button>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
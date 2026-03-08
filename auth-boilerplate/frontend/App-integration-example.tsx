import { useEffect } from 'react';
import { Route, Switch } from 'wouter';
import { authService } from './auth/authService';
import LoginComponent from './auth/LoginComponent';

// Example protected component
function ProtectedDashboard() {
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    const checkAuth = () => {
      setIsAuthenticated(authService.isAuthenticated());
      setLoading(false);
    };
    
    checkAuth();
  }, []);

  if (loading) {
    return <div>Loading...</div>;
  }

  if (!isAuthenticated) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-gray-50">
        <div className="max-w-md w-full space-y-8 p-8">
          <div className="text-center">
            <h2 className="text-3xl font-bold text-gray-900">Sign In Required</h2>
            <p className="mt-2 text-gray-600">Please sign in to access the dashboard</p>
          </div>
          <LoginComponent />
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gray-50">
      <header className="bg-white shadow">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4 flex justify-between items-center">
          <h1 className="text-2xl font-bold text-gray-900">Dashboard</h1>
          <LoginComponent />
        </div>
      </header>
      <main className="max-w-7xl mx-auto py-6 sm:px-6 lg:px-8">
        <div className="px-4 py-6 sm:px-0">
          <div className="border-4 border-dashed border-gray-200 rounded-lg h-96 p-8">
            <h2 className="text-xl font-semibold mb-4">Protected Content</h2>
            <p>This content is only visible to authenticated users.</p>
          </div>
        </div>
      </main>
    </div>
  );
}

// Main App component
function App() {
  useEffect(() => {
    // Initialize authentication on app startup
    authService.initialize().catch(console.error);
  }, []);

  return (
    <div className="App">
      <Switch>
        {/* Public route */}
        <Route path="/">
          <div className="min-h-screen bg-gray-50">
            <header className="bg-white shadow">
              <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4 flex justify-between items-center">
                <h1 className="text-2xl font-bold text-gray-900">My App</h1>
                <LoginComponent />
              </div>
            </header>
            <main className="max-w-7xl mx-auto py-6 sm:px-6 lg:px-8">
              <div className="px-4 py-6 sm:px-0">
                <h2 className="text-3xl font-bold text-center text-gray-900">Welcome</h2>
                <p className="mt-4 text-center text-gray-600">
                  Sign in to access protected features.
                </p>
              </div>
            </main>
          </div>
        </Route>

        {/* Protected route */}
        <Route path="/dashboard">
          <ProtectedDashboard />
        </Route>

        {/* 404 route */}
        <Route>
          <div className="min-h-screen flex items-center justify-center">
            <div className="text-center">
              <h1 className="text-4xl font-bold text-gray-900">404</h1>
              <p className="text-gray-600">Page not found</p>
            </div>
          </div>
        </Route>
      </Switch>
    </div>
  );
}

export default App;
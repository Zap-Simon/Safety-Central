import { useState, useEffect } from "react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Alert, AlertDescription } from "@/components/ui/alert";
import { Skeleton } from "@/components/ui/skeleton";
import { ExternalLink, Shield, Building2, FileText, Users } from "lucide-react";
import Header from "@/components/header";
import Footer from "@/components/footer";

function BrightHRAccess() {
  const [connectionStatus, setConnectionStatus] = useState<'checking' | 'available' | 'error'>('checking');

  // Test connection to BrightHR (without embedding for security)
  const testConnection = async () => {
    try {
      const response = await fetch('/api/brighthr-proxy/dashboard', { 
        method: 'HEAD',
        signal: AbortSignal.timeout(10000) // 10 second timeout
      });
      setConnectionStatus(response.ok ? 'available' : 'error');
    } catch {
      setConnectionStatus('error');
    }
  };

  useEffect(() => {
    testConnection();
  }, []);

  const openBrightHR = () => {
    // Open BrightHR in a new tab for security
    window.open('https://app.brighthr.com.au/dashboard', '_blank', 'noopener,noreferrer');
  };

  const openProxiedAccess = () => {
    // Open proxied access in new tab for security isolation
    window.open('/api/brighthr-proxy/dashboard', '_blank', 'noopener,noreferrer');
  };

  return (
    <div className="w-full border rounded-lg overflow-hidden bg-white dark:bg-gray-900 p-8">
      <div className="text-center space-y-6">
        <div className="mx-auto w-16 h-16 bg-blue-100 dark:bg-blue-900 rounded-full flex items-center justify-center">
          <Building2 className="h-8 w-8 text-blue-600 dark:text-blue-400" />
        </div>
        
        <div>
          <h3 className="text-xl font-semibold text-gray-900 dark:text-gray-100 mb-2">
            Access BrightHR Platform
          </h3>
          <p className="text-gray-600 dark:text-gray-400 max-w-md mx-auto">
            For security reasons, BrightHR opens in a separate tab. Choose your preferred access method below.
          </p>
        </div>

        <div className="flex flex-col sm:flex-row gap-4 justify-center">
          <Button 
            onClick={openBrightHR}
            className="flex items-center gap-2"
            data-testid="button-open-brighthr-direct"
          >
            <ExternalLink className="h-4 w-4" />
            Open BrightHR Directly
          </Button>
          
          <Button 
            variant="outline"
            onClick={openProxiedAccess}
            disabled={connectionStatus === 'error'}
            className="flex items-center gap-2"
            data-testid="button-open-brighthr-proxy"
          >
            <Shield className="h-4 w-4" />
            Open via Proxy
            {connectionStatus === 'checking' && <span className="ml-2 text-xs">(Testing...)</span>}
            {connectionStatus === 'error' && <span className="ml-2 text-xs text-red-500">(Unavailable)</span>}
          </Button>
        </div>

        {connectionStatus === 'error' && (
          <div className="bg-amber-50 dark:bg-amber-950 border border-amber-200 dark:border-amber-800 rounded-lg p-4">
            <div className="flex items-start gap-3">
              <Shield className="h-5 w-5 text-amber-600 mt-0.5 flex-shrink-0" />
              <div className="text-left">
                <h4 className="font-medium text-amber-800 dark:text-amber-200 mb-1">
                  Proxy Connection Issue
                </h4>
                <p className="text-sm text-amber-700 dark:text-amber-300">
                  The proxy connection to BrightHR is currently unavailable. Please use the direct access option.
                </p>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

export default function HRCompliance() {

  return (
    <div className="min-h-screen bg-gray-50 dark:bg-gray-900">
      <Header />
      
      <main className="max-w-7xl mx-auto py-6 px-4 sm:px-6 lg:px-8">
        {/* Page Header */}
        <div className="mb-8">
          <div className="flex items-center gap-3 mb-4">
            <div className="p-2 bg-blue-100 dark:bg-blue-900 rounded-lg">
              <Building2 className="h-6 w-6 text-blue-600 dark:text-blue-400" />
            </div>
            <div>
              <h1 className="text-3xl font-bold text-gray-900 dark:text-gray-100" data-testid="heading-hr-compliance">
                HR Compliance Dashboard
              </h1>
              <p className="text-gray-600 dark:text-gray-400">
                Access BrightHR platform for comprehensive HR management
              </p>
            </div>
          </div>
        </div>


        {/* Quick Access Cards */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
          <Card data-testid="card-employee-management">
            <CardHeader className="flex flex-row items-center space-y-0 pb-2">
              <CardTitle className="text-sm font-medium">Employee Management</CardTitle>
              <Users className="h-4 w-4 ml-auto text-blue-500" />
            </CardHeader>
            <CardContent>
              <CardDescription>
                Manage employee records, contracts, and personal information
              </CardDescription>
            </CardContent>
          </Card>

          <Card data-testid="card-compliance-tracking">
            <CardHeader className="flex flex-row items-center space-y-0 pb-2">
              <CardTitle className="text-sm font-medium">Compliance Tracking</CardTitle>
              <Shield className="h-4 w-4 ml-auto text-green-500" />
            </CardHeader>
            <CardContent>
              <CardDescription>
                Monitor compliance requirements and regulatory obligations
              </CardDescription>
            </CardContent>
          </Card>

          <Card data-testid="card-document-management">
            <CardHeader className="flex flex-row items-center space-y-0 pb-2">
              <CardTitle className="text-sm font-medium">Document Management</CardTitle>
              <FileText className="h-4 w-4 ml-auto text-purple-500" />
            </CardHeader>
            <CardContent>
              <CardDescription>
                Store and manage HR documents, policies, and procedures
              </CardDescription>
            </CardContent>
          </Card>
        </div>

        {/* BrightHR Integration */}
        <Card>
          <CardHeader>
            <div className="flex items-center justify-between">
              <div>
                <CardTitle className="flex items-center gap-2">
                  <Shield className="h-5 w-5 text-blue-500" />
                  BrightHR Platform
                </CardTitle>
                <CardDescription>
                  Integrated HR management system for Cranfield Glass
                </CardDescription>
              </div>
              <div className="flex gap-2">
                <Button 
                  variant="outline" 
                  size="sm"
                  onClick={() => window.open('https://app.brighthr.com.au/dashboard', '_blank')}
                  data-testid="button-open-new-tab"
                >
                  <ExternalLink className="h-4 w-4 mr-2" />
                  Open in New Tab
                </Button>
              </div>
            </div>
          </CardHeader>
          <CardContent>
            <BrightHRAccess />
          </CardContent>
        </Card>

        {/* Additional Information */}
        <div className="mt-8 p-4 bg-blue-50 dark:bg-blue-950 rounded-lg border border-blue-200 dark:border-blue-800">
          <div className="flex items-start gap-3">
            <Shield className="h-5 w-5 text-blue-600 dark:text-blue-400 mt-0.5 flex-shrink-0" />
            <div>
              <h3 className="font-semibold text-blue-900 dark:text-blue-100 mb-2">
                About This Integration
              </h3>
              <p className="text-blue-800 dark:text-blue-200 text-sm leading-relaxed">
                This proxy integration allows you to access BrightHR's HR compliance tools directly within 
                our safety management platform. All data remains secure and follows the same authentication 
                and privacy standards as your main BrightHR account.
              </p>
            </div>
          </div>
        </div>
      </main>

      <Footer />
    </div>
  );
}
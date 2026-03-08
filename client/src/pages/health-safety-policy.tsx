import { useEffect } from "react";
import { useQuery } from "@tanstack/react-query";
import { Shield, Download, FileText, File, ExternalLink } from "lucide-react";
import PolicyHeader from "@/components/policy-header";
import Footer from "@/components/footer";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { authService } from "@/auth/authService";

interface PolicyDocument {
  name: string;
  url: string;
  type: 'word' | 'pdf';
  size: string;
  modified: string;
}

export default function HealthSafetyPolicy() {
  useEffect(() => {
    // Scroll to top when page loads
    window.scrollTo(0, 0);
  }, []);

  // Fetch policy documents from the folder with authentication
  const { data: policyDocuments, isLoading, error } = useQuery({
    queryKey: ['/api/policy-documents'],
    queryFn: async () => {
      try {
        const token = await authService.getAccessToken();
        const response = await fetch('/api/policy-documents', {
          headers: {
            'Authorization': `Bearer ${token}`,
            'Accept': 'application/json'
          },
          credentials: 'include'
        });
        
        if (!response.ok) {
          const errorText = await response.text();
          console.error('Policy documents API error:', response.status, errorText);
          throw new Error(`Failed to fetch policy documents: ${response.status} - ${errorText}`);
        }
        
        return await response.json();
      } catch (error) {
        console.error('Error fetching policy documents:', error);
        throw error;
      }
    },
    enabled: true,
    retry: false
  });

  const handleDocumentPreview = (policyDoc: PolicyDocument) => {
    // Open directly in SharePoint - cleaner and more secure
    window.open(policyDoc.url, '_blank');
  };

  const handleDocumentDownload = async (policyDoc: PolicyDocument) => {
    try {
      const token = await authService.getAccessToken();
      const downloadUrl = `/api/policy-document-proxy?url=${encodeURIComponent(policyDoc.url)}&access_token=${encodeURIComponent(token)}&download=true`;
      
      // Create a temporary link to trigger download
      const link = document.createElement('a');
      link.href = downloadUrl;
      link.download = policyDoc.name;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } catch (error) {
      console.error('Failed to download document:', error);
    }
  };

  return (
    <div className="font-inter bg-gradient-to-br from-gray-50 to-gray-100 text-ms-gray-900 min-h-screen flex flex-col">
      <PolicyHeader />
      
      {/* Main Content with Header Offset */}
      <div className="pt-16 flex-1">
        <div className="max-w-4xl mx-auto px-3 sm:px-6 py-4 sm:py-8">
          
          {/* Page Header */}
          <div className="bg-white rounded-lg shadow-sm p-4 mb-6 border border-gray-200">
            <div className="flex items-center space-x-2 sm:space-x-3">
              <div className="w-8 h-8 sm:w-10 sm:h-10 bg-gradient-to-r from-red-600 to-red-700 rounded-lg sm:rounded-xl flex items-center justify-center shadow-lg flex-shrink-0">
                <Shield className="text-white" size={16} />
              </div>
              <div className="min-w-0 flex-1">
                <h1 className="text-lg sm:text-2xl font-bold text-gray-900 leading-tight">Health & Safety Policy</h1>
                <p className="text-xs text-gray-600 hidden sm:block">Official company policy documents and procedures</p>
              </div>
            </div>
          </div>

          {/* Policy Documents */}
          <div className="grid gap-6">
            
            {/* Current Policy Documents */}
            <Card>
              <CardHeader>
                <CardTitle>
                  <h2 className="text-lg font-semibold text-gray-900">Current Policy Documents</h2>
                </CardTitle>
              </CardHeader>
              <CardContent>
                {isLoading ? (
                  <div className="flex items-center justify-center py-8">
                    <div className="text-gray-500">Loading policy documents...</div>
                  </div>
                ) : policyDocuments && (policyDocuments as any)?.documents?.length > 0 ? (
                  <div className="grid gap-4 md:grid-cols-2">
                    {(policyDocuments as any).documents.map((document: PolicyDocument, index: number) => (
                      <div
                        key={index}
                        className="border border-gray-200 rounded-lg p-4 hover:border-gray-300 transition-colors flex flex-col"
                      >
                        <div className="flex items-start justify-between mb-3">
                          <div className="flex items-center space-x-3 flex-1 min-w-0">
                            {document.type === 'word' ? (
                              <FileText className="text-blue-600 flex-shrink-0" size={20} />
                            ) : (
                              <File className="text-red-600 flex-shrink-0" size={20} />
                            )}
                            <div className="min-w-0 flex-1">
                              <h3 className="font-medium text-gray-900 truncate">{document.name}</h3>
                              <p className="text-sm text-gray-500 truncate">
                                {document.size}
                              </p>
                            </div>
                          </div>
                          <Badge variant={document.type === 'word' ? 'default' : 'secondary'} className="flex-shrink-0">
                            {document.type.toUpperCase()}
                          </Badge>
                        </div>
                        
                        <div className="flex gap-2 mt-auto">
                          <Button
                            onClick={() => handleDocumentPreview(document)}
                            className="flex-1"
                            variant="default"
                          >
                            <ExternalLink size={16} className="mr-2" />
                            Open Preview
                          </Button>
                          <Button
                            onClick={() => handleDocumentDownload(document)}
                            className="flex-1"
                            variant="outline"
                          >
                            <Download size={16} className="mr-2" />
                            Download
                          </Button>
                        </div>
                      </div>
                    ))}
                  </div>
                ) : (
                  <div className="text-center py-8">
                    <Shield className="mx-auto h-12 w-12 text-gray-400 mb-4" />
                    <p className="text-gray-500 mb-4">
                      Policy documents will appear here when available.
                    </p>
                    <p className="text-sm text-gray-400">
                      Check back later for updated policy documents.
                    </p>
                  </div>
                )}
              </CardContent>
            </Card>


          </div>
        </div>
      </div>
      
      <Footer />
    </div>
  );
}
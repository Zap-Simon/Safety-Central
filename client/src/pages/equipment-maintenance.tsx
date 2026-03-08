import { useEffect, useState } from "react";
import { useQuery } from "@tanstack/react-query";
import EquipmentHeader from "@/components/equipment-header";
import Footer from "@/components/footer";
import { Tabs, TabsList, TabsTrigger, TabsContent } from "@/components/ui/tabs";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import { Settings, Zap, Shield, FileText, File, ExternalLink, Download, FileSpreadsheet } from "lucide-react";
import { authService } from "@/auth/authService";
import { ExcelTable } from "@/components/ExcelTable";

interface TestTagDocument {
  name: string;
  url: string;
  type: 'word' | 'pdf';
  size: string;
  modified: string;
}

export default function EquipmentMaintenance() {
  useEffect(() => {
    // Scroll to top when page loads
    window.scrollTo(0, 0);
  }, []);

  const [activeTab, setActiveTab] = useState('overview');

  // Fetch test & tag documents only when the tab is selected to reduce initial load
  const { data: testTagDocuments, isLoading, error } = useQuery({
    queryKey: ['/api/equipment-test-tag-documents'],
    queryFn: async () => {
      try {
        const token = await authService.getAccessToken();
        const response = await fetch('/api/equipment-test-tag-documents', {
          headers: {
            'Authorization': `Bearer ${token}`,
            'Accept': 'application/json'
          },
          credentials: 'include'
        });
        
        if (!response.ok) {
          const errorText = await response.text();
          console.error('Test & tag documents API error:', response.status, errorText);
          throw new Error(`Failed to fetch test & tag documents: ${response.status} - ${errorText}`);
        }
        
        return await response.json();
      } catch (error) {
        console.error('Error fetching test & tag documents:', error);
        throw error;
      }
    },
    enabled: activeTab === 'test-tag',
    retry: false
  });

  const handleDocumentPreview = (document: TestTagDocument) => {
    // Open directly in SharePoint - cleaner and more secure
    window.open(document.url, '_blank');
  };

  const handleDocumentDownload = async (testTagDoc: TestTagDocument) => {
    try {
      const token = await authService.getAccessToken();
      const downloadUrl = `/api/policy-document-proxy?url=${encodeURIComponent(testTagDoc.url)}&access_token=${encodeURIComponent(token)}&download=true`;
      
      // Create a temporary link to trigger download
      const link = document.createElement('a');
      link.href = downloadUrl;
      link.download = testTagDoc.name;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } catch (error) {
      console.error('Failed to download document:', error);
    }
  };

  return (
    <div className="font-inter bg-gradient-to-br from-gray-50 to-gray-100 text-ms-gray-900 min-h-screen flex flex-col">
      <EquipmentHeader />
      
      {/* Main Content with Header Offset */}
      <div className="pt-16 flex-1">
        <div className="max-w-4xl mx-auto px-3 sm:px-6 py-4 sm:py-8">
          
          {/* Page Header */}
          <div className="bg-white rounded-lg shadow-sm p-4 mb-6 border border-gray-200">
            <div className="flex items-center space-x-2 sm:space-x-3">
              <div className="w-8 h-8 sm:w-10 sm:h-10 bg-gradient-to-r from-green-600 to-green-700 rounded-lg sm:rounded-xl flex items-center justify-center shadow-lg flex-shrink-0">
                <Settings className="text-white" size={16} />
              </div>
              <div className="min-w-0 flex-1">
                <h1 className="text-lg sm:text-2xl font-bold text-gray-900 leading-tight">Equipment & Maintenance</h1>
                <p className="text-xs text-gray-600 hidden sm:block">Plant & equipment safety records and maintenance schedules</p>
              </div>
            </div>
          </div>

          <div className="bg-white rounded-lg shadow-sm p-4 mb-6 border border-gray-200">
            <Tabs value={activeTab} onValueChange={setActiveTab} className="w-full">
              <TabsList className="grid w-full grid-cols-3 mb-6">
                <TabsTrigger value="overview" className="flex items-center space-x-2">
                  <Settings size={16} />
                  <span>Equipment Overview</span>
                </TabsTrigger>
                <TabsTrigger value="test-tag" className="flex items-center space-x-2">
                  <Zap size={16} />
                  <span>Test & Tag Records</span>
                </TabsTrigger>
                <TabsTrigger value="maintenance" className="flex items-center space-x-2">
                  <Shield size={16} />
                  <span>Inspection Records</span>
                </TabsTrigger>
              </TabsList>
              
              <TabsContent value="overview">
                <ExcelTable
                  siteName="EquipmentMaintenance"
                  folderPath="/Equipment Test - Tag"
                  title="Equipment Inventory"
                  description="Current tools and equipment status"
                  pageSize={25}
                />
              </TabsContent>
              
              <TabsContent value="test-tag">
                <Card>
                  <CardHeader>
                    <CardTitle>
                      <h2 className="text-lg font-semibold text-gray-900">Test & Tag Records</h2>
                    </CardTitle>
                  </CardHeader>
                  <CardContent>
                    {isLoading ? (
                      <div className="flex items-center justify-center py-8">
                        <div className="text-gray-500">Loading test & tag documents...</div>
                      </div>
                    ) : testTagDocuments && (testTagDocuments as any)?.documents?.length > 0 ? (
                      <div className="grid gap-4 md:grid-cols-2">
                        {(testTagDocuments as any).documents.map((document: TestTagDocument, index: number) => (
                          <div
                            key={index}
                            className="border border-gray-200 rounded-lg p-4 hover:border-gray-300 transition-colors flex flex-col"
                            data-testid={`card-testag-document-${index}`}
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
                                  <div className="flex items-center space-x-2 mt-1">
                                    <Badge variant="secondary" className="text-xs">
                                      {document.type.toUpperCase()}
                                    </Badge>
                                    <span className="text-xs text-gray-500">{document.size}</span>
                                  </div>
                                </div>
                              </div>
                            </div>
                            
                            <div className="text-xs text-gray-500 mb-3">
                              Last modified: {document.modified}
                            </div>
                            
                            <div className="flex space-x-2 mt-auto">
                              <Button 
                                variant="outline" 
                                size="sm" 
                                onClick={() => handleDocumentPreview(document)}
                                className="flex-1"
                                data-testid={`button-preview-${index}`}
                              >
                                <ExternalLink size={14} className="mr-1" />
                                Preview
                              </Button>
                              <Button 
                                variant="outline" 
                                size="sm" 
                                onClick={() => handleDocumentDownload(document)}
                                className="flex-1"
                                data-testid={`button-download-${index}`}
                              >
                                <Download size={14} className="mr-1" />
                                Download
                              </Button>
                            </div>
                          </div>
                        ))}
                      </div>
                    ) : error ? (
                      <div className="text-center py-8">
                        <div className="text-red-600 mb-2">Failed to load test & tag documents</div>
                        <div className="text-sm text-gray-500">Please check your SharePoint access or try again later</div>
                      </div>
                    ) : (
                      <div className="text-center py-8">
                        <div className="text-gray-500 mb-2">No test & tag documents found</div>
                        <div className="text-sm text-gray-400">Documents will appear here once available in SharePoint</div>
                      </div>
                    )}
                  </CardContent>
                </Card>
              </TabsContent>
              
              <TabsContent value="maintenance">
                {/* Maintenance content will be added here when needed */}
              </TabsContent>
            </Tabs>
          </div>
          
        </div>
      </div>
      
      <Footer />
    </div>
  );
}
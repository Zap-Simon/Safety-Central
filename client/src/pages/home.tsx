import { useEffect, useState } from "react";
import { useLocation } from "wouter";
import Header from "@/components/header";
import Footer from "@/components/footer";
import { Tabs, TabsList, TabsTrigger, TabsContent } from "@/components/ui/tabs";
import OverviewTab from "@/components/tabs/overview-tab";
import DocumentationTab from "@/components/tabs/documentation-tab";
import EnvironmentQualityTab from "@/components/tabs/environment-quality-tab";
import ComplianceTab from "@/components/tabs/compliance-tab";
import MeetingTab from "@/components/tabs/meeting-tab";
import TrainingTab from "@/components/tabs/training-tab";
import { Home as HomeIcon, FileText, Calendar, BookOpen, Leaf, Shield } from "lucide-react";
import { initializeMermaid } from "@/lib/mermaid";

export default function Home() {
  const [location] = useLocation();

  // Get tab from URL search params
  const getActiveTab = () => {
    const urlParams = new URLSearchParams(window.location.search);
    const tab = urlParams.get('tab');
    return tab || 'home';
  };

  const [activeTab, setActiveTab] = useState(() => getActiveTab());

  // Update URL when tab changes
  const handleTabChange = (value: string) => {
    setActiveTab(value);
    const url = new URL(window.location.href);
    url.searchParams.set('tab', value);
    window.history.replaceState({}, '', url.toString());
  };

  useEffect(() => {
    initializeMermaid();
    
    // Handle print event - expand all collapsible sections
    const handleBeforePrint = () => {
      const collapsibleElements = document.querySelectorAll('.collapsible-content');
      collapsibleElements.forEach(el => {
        el.classList.remove('hidden');
      });
    };

    // Handle browser back/forward buttons
    const handlePopState = () => {
      setActiveTab(getActiveTab());
    };

    window.addEventListener('beforeprint', handleBeforePrint);
    window.addEventListener('popstate', handlePopState);
    
    return () => {
      window.removeEventListener('beforeprint', handleBeforePrint);
      window.removeEventListener('popstate', handlePopState);
    };
  }, []);

  return (
    <div className="font-inter bg-gray-50 text-ms-gray-900 min-h-screen flex flex-col">
      <Header />
      
      {/* Main Content with Header Offset */}
      <div className="pt-16 flex-1">
        <div className="max-w-7xl mx-auto px-3 sm:px-4 lg:px-8 py-3 sm:py-4 lg:py-6">
          
          <Tabs value={activeTab} onValueChange={handleTabChange} className="w-full">
            <TabsList className="grid w-full grid-cols-4 mb-6">
              <TabsTrigger value="home" className="flex items-center space-x-2">
                <HomeIcon size={16} />
                <span>Health & Safety</span>
              </TabsTrigger>
              <TabsTrigger value="environment-quality" className="flex items-center space-x-2">
                <Leaf size={16} />
                <span>Environment & Quality</span>
              </TabsTrigger>
              <TabsTrigger value="compliance" className="flex items-center space-x-2">
                <Shield size={16} />
                <span>Standards & Compliance</span>
              </TabsTrigger>
              <TabsTrigger value="documentation" className="flex items-center space-x-2">
                <FileText size={16} />
                <span>Engagement</span>
              </TabsTrigger>
            </TabsList>
            
            <TabsContent value="home">
              <OverviewTab />
            </TabsContent>
            
            <TabsContent value="environment-quality">
              <EnvironmentQualityTab />
            </TabsContent>
            
            <TabsContent value="compliance">
              <ComplianceTab />
            </TabsContent>
            
            <TabsContent value="documentation">
              <DocumentationTab />
            </TabsContent>
          </Tabs>
          
        </div>
      </div>
      
      <Footer />
    </div>
  );
}

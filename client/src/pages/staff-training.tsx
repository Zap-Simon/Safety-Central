import { useEffect } from "react";
import TrainingHeader from "@/components/training-header";
import Footer from "@/components/footer";
import { Tabs, TabsList, TabsTrigger, TabsContent } from "@/components/ui/tabs";
import FieldTrainingTab from "@/components/tabs/field-training-tab";
import OfficeTrainingTab from "@/components/tabs/office-training-tab";
import SkillsMatrixTab from "@/components/tabs/skills-matrix-tab";
import { HardHat, Building2, Award } from "lucide-react";

export default function StaffTraining() {
  useEffect(() => {
    // Scroll to top when page loads
    window.scrollTo(0, 0);
  }, []);

  return (
    <div className="font-inter bg-gradient-to-br from-gray-50 to-gray-100 text-ms-gray-900 min-h-screen flex flex-col">
      <TrainingHeader />
      
      {/* Main Content with Header Offset */}
      <div className="pt-16 flex-1">
        <div className="max-w-4xl mx-auto px-3 sm:px-6 py-4 sm:py-8">
          
          {/* Page Header */}
          <div className="bg-white rounded-lg shadow-sm p-4 mb-6 border border-gray-200">
            <div className="flex items-center space-x-2 sm:space-x-3">
              <div className="w-8 h-8 sm:w-10 sm:h-10 bg-gradient-to-r from-orange-600 to-orange-700 rounded-lg sm:rounded-xl flex items-center justify-center shadow-lg flex-shrink-0">
                <HardHat className="text-white" size={16} />
              </div>
              <div className="min-w-0 flex-1">
                <h1 className="text-lg sm:text-2xl font-bold text-gray-900 leading-tight">Staff Training</h1>
                <p className="text-xs text-gray-600 hidden sm:block">Job workflows & training materials for field and office staff</p>
              </div>
            </div>
          </div>

          <div className="bg-white rounded-lg shadow-sm p-4 mb-6 border border-gray-200">
            <Tabs defaultValue="skills-matrix" className="w-full">
              <TabsList className="grid w-full grid-cols-3 mb-6 h-auto">
                <TabsTrigger value="skills-matrix" className="flex items-center justify-center space-x-1 sm:space-x-2 px-2 sm:px-3 py-2 sm:py-1.5">
                  <Award className="w-4 h-4 flex-shrink-0" />
                  <span className="text-xs sm:text-sm font-medium truncate">Skills Matrix</span>
                </TabsTrigger>
                <TabsTrigger value="field" className="flex items-center justify-center space-x-1 sm:space-x-2 px-2 sm:px-3 py-2 sm:py-1.5">
                  <HardHat className="w-4 h-4 flex-shrink-0" />
                  <span className="text-xs sm:text-sm font-medium truncate">Able To Use</span>
                </TabsTrigger>
                <TabsTrigger value="office" className="flex items-center justify-center space-x-1 sm:space-x-2 px-2 sm:px-3 py-2 sm:py-1.5">
                  <Building2 className="w-4 h-4 flex-shrink-0" />
                  <span className="text-xs sm:text-sm font-medium truncate">Administration</span>
                </TabsTrigger>
              </TabsList>
              
              <TabsContent value="skills-matrix">
                <SkillsMatrixTab />
              </TabsContent>
              
              <TabsContent value="field">
                <FieldTrainingTab />
              </TabsContent>
              
              <TabsContent value="office">
                <OfficeTrainingTab />
              </TabsContent>
            </Tabs>
          </div>

        </div>
      </div>
      
      <Footer />
    </div>
  );
}
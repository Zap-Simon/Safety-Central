import { useQuery } from "@tanstack/react-query";
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { Accordion, AccordionContent, AccordionItem, AccordionTrigger } from "@/components/ui/accordion";
import { Badge } from "@/components/ui/badge";
import { Skeleton } from "@/components/ui/skeleton";
import { CSCWorkflowPopup } from "@/components/csc-workflow-popup";
import { CSC_WORKFLOW_MODULES } from "@shared/schema";
import { 
  Building2, 
  FileText, 
  Users, 
  Shield, 
  DollarSign, 
  Phone,
  Settings,
  Clock,
  CheckCircle,
  AlertTriangle,
  BookOpen
} from "lucide-react";

// Define types for training data
type TrainingClassification = {
  id: number;
  key: string;
  name: string;
  description: string | null;
  audience: string | null;
  displayOrder: number;
  isActive: boolean;
};

type TrainingModule = {
  id: number;
  classificationId: number;
  code: string;
  name: string;
  description: string | null;
  validityPeriod: number | null;
  displayOrder: number;
  isActive: boolean;
  isSafetyCritical?: boolean;
  requiresCertification?: boolean;
};

export default function OfficeTrainingTab() {
  // Fetch real data from API - only administration audience classifications
  const { data: classificationsData, isLoading: classificationsLoading } = useQuery({
    queryKey: ['/api/training-classifications?audience=administration'],
  });

  const { data: modulesData, isLoading: modulesLoading } = useQuery({
    queryKey: ['/api/training-modules'],
  });

  if (classificationsLoading || modulesLoading) {
    return (
      <div className="space-y-6">
        <div className="mb-6">
          <Skeleton className="h-8 w-96 mb-2" />
          <Skeleton className="h-4 w-full max-w-2xl" />
        </div>
        {Array.from({ length: 6 }).map((_, i) => (
          <Skeleton key={i} className="h-64 w-full" />
        ))}
      </div>
    );
  }

  const classifications: TrainingClassification[] = (classificationsData as any)?.data || [];
  const modules: TrainingModule[] = (modulesData as any)?.data || [];

  // Group modules by classification
  const modulesByClassification = modules.reduce((acc, module) => {
    if (!acc[module.classificationId]) {
      acc[module.classificationId] = [];
    }
    acc[module.classificationId].push(module);
    return acc;
  }, {} as Record<number, TrainingModule[]>);

  const getClassificationIcon = (key: string) => {
    switch (key) {
      case 'office-procedures': return Building2;
      case 'software-systems': return Settings;
      case 'compliance-governance': return Shield;
      case 'hr-onboarding': return Users;
      case 'finance-invoicing': return DollarSign;
      case 'customer-service': return Phone;
      default: return FileText;
    }
  };

  const getCategoryFromClassification = (classificationId: number, classificationModules: TrainingModule[]) => {
    const hasSafetyCritical = classificationModules.some(m => m.isSafetyCritical);
    const hasCertificationRequired = classificationModules.some(m => m.requiresCertification);
    
    if (hasSafetyCritical) return 'Safety Critical';
    if (hasCertificationRequired) return 'Equipment Level';
    return 'Foundation Level';
  };

  // Filter for administration classifications only, prioritizing safety-critical classifications at top
  const adminClassifications = classifications
    .filter(c => c.isActive && c.audience === 'administration')
    .map(classification => {
      const classificationModules = modulesByClassification[classification.id] || [];
      const category = getCategoryFromClassification(classification.id, classificationModules);
      
      return {
        ...classification,
        category: category,
        hasSafetyCritical: classificationModules.some(m => m.isSafetyCritical),
        safetyCriticalCount: classificationModules.filter(m => m.isSafetyCritical).length
      };
    })
    .sort((a, b) => {
      // Sort classifications: Safety Critical first, then Equipment Level, then Foundation Level
      const categoryOrder = { 'Safety Critical': 0, 'Equipment Level': 1, 'Foundation Level': 2 };
      const aOrder = categoryOrder[a.category as keyof typeof categoryOrder] ?? 3;
      const bOrder = categoryOrder[b.category as keyof typeof categoryOrder] ?? 3;
      
      if (aOrder !== bOrder) return aOrder - bOrder;
      // Within same category, sort by original display order
      return a.displayOrder - b.displayOrder;
    });

  const getClassificationColor = (category: string) => {
    switch (category) {
      case 'Foundation Level': return 'from-purple-600 to-purple-700';
      case 'Equipment Level': return 'from-orange-600 to-orange-700';
      case 'Safety Critical': return 'from-red-600 to-red-700';
      default: return 'from-gray-600 to-gray-700';
    }
  };

  const getCategoryBadgeColor = (category: string) => {
    switch (category) {
      case 'Foundation Level': return 'bg-purple-100 text-purple-700 border-purple-200';
      case 'Equipment Level': return 'bg-orange-100 text-orange-700 border-orange-200';
      case 'Safety Critical': return 'bg-red-100 text-red-700 border-red-200';
      default: return 'bg-gray-100 text-gray-700 border-gray-200';
    }
  };

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="mb-6">
        <h2 className="text-2xl font-bold text-gray-900 mb-2">Administration Training Modules</h2>
        <p className="text-gray-600">
          Essential office procedures, software systems, and administrative skills training for support staff. 
          Staff must complete relevant modules to ensure efficient business operations and compliance.
        </p>
      </div>

      {/* Administration Training Classifications */}
      <div className="grid gap-6">
        {adminClassifications.map((classification) => {
          const IconComponent = getClassificationIcon(classification.key);
          const classificationModules = modulesByClassification[classification.id] || [];
          
          return (
            <Card key={classification.id} className="overflow-hidden">
              <CardHeader className="pb-4">
                <div className="flex items-start justify-between">
                  <div className="flex items-center space-x-3">
                    <div className={`w-12 h-12 bg-gradient-to-r ${getClassificationColor(classification.category)} rounded-lg flex items-center justify-center shadow-lg`}>
                      <IconComponent className="text-white" size={24} />
                    </div>
                    <div>
                      <CardTitle className="text-xl text-gray-900">{classification.name}</CardTitle>
                      <CardDescription className="text-sm mt-1">{classification.description}</CardDescription>
                    </div>
                  </div>
                  <div className="flex items-center space-x-2">
                    {classification.category !== 'Foundation Level' && (
                      <Badge variant="outline" className={`border ${getCategoryBadgeColor(classification.category)} font-semibold`} data-testid={`badge-category-${classification.id}`}>
                        {classification.category}
                      </Badge>
                    )}
                    <Badge variant="secondary" className="bg-gray-100 text-gray-800" data-testid={`badge-modules-${classification.id}`}>
                      {classificationModules.length + (classification.key === 'office-procedures' ? Object.keys(CSC_WORKFLOW_MODULES).length : 0)} modules
                    </Badge>
                    {classification.hasSafetyCritical && (
                      <Badge variant="destructive" className="bg-red-600 text-white" data-testid={`badge-safety-critical-${classification.id}`}>
                        <AlertTriangle className="w-3 h-3 mr-1" />
                        {classification.safetyCriticalCount} Critical
                      </Badge>
                    )}
                  </div>
                </div>
              </CardHeader>
              
              <CardContent>
                <Accordion type="single" collapsible>
                  <AccordionItem value="training-modules">
                    <AccordionTrigger className="text-sm font-medium">
                      Training Modules
                    </AccordionTrigger>
                    <AccordionContent>
                      {classificationModules.length > 0 ? (
                        <div className="space-y-3">
                          {classificationModules
                            .sort((a, b) => {
                              // Safety-critical modules first, then by display order
                              if (a.isSafetyCritical && !b.isSafetyCritical) return -1;
                              if (!a.isSafetyCritical && b.isSafetyCritical) return 1;
                              return a.displayOrder - b.displayOrder;
                            })
                            .map((module) => (
                              <div 
                                key={module.id}
                                className={`p-3 bg-gray-50 rounded-md border-l-4 ${module.isSafetyCritical ? 'border-red-500 bg-red-50' : module.requiresCertification ? 'border-orange-500 bg-orange-50' : 'border-purple-500'}`}
                                data-testid={`training-module-${module.code}`}
                              >
                                <div className="flex items-start justify-between">
                                  <div className="flex-1">
                                    <div className="flex items-center space-x-2 mb-1">
                                      <span className={`text-sm font-mono font-medium ${module.isSafetyCritical ? 'text-red-600' : module.requiresCertification ? 'text-orange-600' : 'text-purple-600'}`}>{module.code}</span>
                                      <span className="text-sm font-medium text-gray-900">{module.name}</span>
                                      {module.isSafetyCritical && (
                                        <Badge variant="destructive" className="ml-2 bg-red-600 text-white text-xs">
                                          <AlertTriangle className="w-3 h-3 mr-1" />
                                          Safety Critical
                                        </Badge>
                                      )}
                                      {module.requiresCertification && !module.isSafetyCritical && (
                                        <Badge variant="secondary" className="ml-2 bg-orange-100 text-orange-700 text-xs">
                                          Certification Required
                                        </Badge>
                                      )}
                                    </div>
                                    {module.description && (
                                      <p className="text-xs text-gray-600 leading-relaxed">{module.description}</p>
                                    )}
                                  </div>
                                  <div className="ml-3">
                                    <Badge variant="outline" className="inline-flex items-center gap-1 px-2 py-1 text-xs font-medium text-gray-600 bg-gray-50 hover:bg-gray-100 border-gray-300">
                                      <Clock className="h-3 w-3" />
                                      Coming Soon
                                    </Badge>
                                  </div>
                                </div>
                              </div>
                            ))
                          }
                          
                          {/* Add CSC Workflow Modules if this is office-procedures classification */}
                          {classification.key === 'office-procedures' && (
                            <>
                              {Object.entries(CSC_WORKFLOW_MODULES).map(([key, workflowModule]) => (
                                <div 
                                  key={key}
                                  className="p-3 bg-gray-50 rounded-md border-l-4 border-purple-500"
                                  data-testid={`training-module-${workflowModule.code}`}
                                >
                                  <div className="flex items-start justify-between">
                                    <div className="flex-1">
                                      <div className="flex items-center space-x-2 mb-1">
                                        <span className="text-sm font-mono font-medium text-purple-600">{workflowModule.code}</span>
                                        <span className="text-sm font-medium text-gray-900">{workflowModule.name}</span>
                                      </div>
                                      <p className="text-xs text-gray-600 leading-relaxed">{workflowModule.description}</p>
                                    </div>
                                    <div className="ml-3">
                                      <CSCWorkflowPopup 
                                        workflowCode={key as keyof typeof CSC_WORKFLOW_MODULES}
                                        triggerButton={
                                          <button className="inline-flex items-center gap-1 px-2 py-1 text-xs font-medium text-purple-700 bg-purple-100 hover:bg-purple-200 rounded-md border border-purple-300 transition-colors">
                                            <FileText className="h-3 w-3" />
                                            View SOP
                                          </button>
                                        }
                                      />
                                    </div>
                                  </div>
                                </div>
                              ))}
                            </>
                          )}
                        </div>
                      ) : (
                        <div className="text-sm text-gray-500 italic p-3 bg-gray-50 rounded-md">
                          No training modules defined for this category yet.
                        </div>
                      )}
                    </AccordionContent>
                  </AccordionItem>
                </Accordion>
              </CardContent>
            </Card>
          );
        })}
      </div>

      {/* Training Summary */}
      <Card className="bg-purple-50 border-purple-200">
        <CardContent className="p-6">
          <div className="flex items-start space-x-3">
            <Building2 className="w-6 h-6 text-purple-600 mt-1" />
            <div>
              <h3 className="font-semibold text-purple-900 mb-2">Training Management</h3>
              <p className="text-purple-800 text-sm mb-4">
                All Administration training records are tracked in the Skills Matrix. 
                Staff competency levels are updated as training progresses and assessments are completed.
              </p>
              <div className="flex justify-center text-sm">
                <div className="flex items-center space-x-2">
                  <CheckCircle className="w-4 h-4 text-green-600" />
                  <span className="text-purple-800">Competency tracking</span>
                </div>
              </div>
            </div>
          </div>
        </CardContent>
      </Card>
    </div>
  );
}
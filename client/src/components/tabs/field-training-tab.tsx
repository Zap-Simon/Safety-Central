import { useState } from "react";
import { useQuery } from "@tanstack/react-query";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Accordion, AccordionContent, AccordionItem, AccordionTrigger } from "@/components/ui/accordion";
import { Badge } from "@/components/ui/badge";
import { Skeleton } from "@/components/ui/skeleton";
import { HardHat, Shield, Wrench, Settings, AlertTriangle, CheckCircle, Clock, User, Car, Hammer, XCircle, Crown, BookOpen, Users } from "lucide-react";

type TrainingClassification = {
  id: number;
  key: string;
  name: string;
  description: string | null;
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
  isSafetyCritical: boolean;
  requiresCertification: boolean;
  sopUrl: string | null;
  trainingVideoUrl: string | null;
};

export default function FieldTrainingTab() {
  // Fetch real data from API - only field audience classifications
  const { data: classificationsData, isLoading: classificationsLoading } = useQuery({
    queryKey: ['/api/training-classifications?audience=field'],
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
      case 'basic-hand-tools': return Hammer;
      case 'power-tools-equipment': return Wrench;
      case 'lifting-handling-equipment': return HardHat;
      case 'safety-equipment-ppe': return Shield;
      case 'vehicles-transport': return Car;
      case 'working-at-heights': return Settings;
      default: return Settings;
    }
  };

  const getCategoryFromClassification = (key: string) => {
    // Safety Critical: Lifting & Handling Equipment, Safety Equipment & PPE, Working at Heights
    if (['lifting-handling-equipment', 'safety-equipment-ppe', 'working-at-heights'].includes(key)) {
      return 'Safety Critical';
    }
    // Equipment Level: Power Tools & Equipment, Vehicles & Transport
    if (['power-tools-equipment', 'vehicles-transport'].includes(key)) {
      return 'Equipment Level';
    }
    // Foundation Level: Basic Hand Tools, Chemicals (and all others)
    return 'Foundation Level';
  };

  const getClassificationColor = (category: string) => {
    switch (category) {
      case 'Foundation Level': return 'from-blue-600 to-blue-700';
      case 'Equipment Level': return 'from-orange-600 to-orange-700';
      case 'Safety Critical': return 'from-red-600 to-red-700';
      default: return 'from-gray-600 to-gray-700';
    }
  };

  // Use real data from API instead of hardcoded data, prioritizing safety-critical classifications at top
  const ableToUseModules = classifications
    .filter(c => c.isActive)
    .map(classification => {
      const classificationModules = modulesByClassification[classification.id] || [];
      const IconComponent = getClassificationIcon(classification.key);
      const category = getCategoryFromClassification(classification.key);
      
      return {
        id: classification.key,
        title: classification.name,
        category: category,
        description: classification.description || `Training modules for ${classification.name.toLowerCase()}`,
        icon: IconComponent,
        color: getClassificationColor(category),
        competencyLevels: ['Not Trained', 'In Training (Supervised)', 'Competent – Supervised', 'Competent – SOP/Module', 'Expert'],
        modules: classificationModules
          .sort((a, b) => {
            // Safety-critical modules first, then by display order
            if (a.isSafetyCritical && !b.isSafetyCritical) return -1;
            if (!a.isSafetyCritical && b.isSafetyCritical) return 1;
            return a.displayOrder - b.displayOrder;
          }),
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
      return classifications.find(c => c.key === a.id)?.displayOrder ?? 0 - 
             (classifications.find(c => c.key === b.id)?.displayOrder ?? 0);
    });


  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="mb-4 sm:mb-6">
        <h2 className="text-lg sm:text-2xl font-bold text-gray-900 mb-2">Field Training - "Able to Use" Modules</h2>
        <p className="text-sm sm:text-base text-gray-600">
          Essential equipment and tool competency training for glazing field staff. 
          Staff must demonstrate competency before being authorized to use equipment.
        </p>
      </div>

      {/* Training Modules */}
      <div className="grid gap-6">
        {ableToUseModules.map((module) => {
          const IconComponent = module.icon;
          
          return (
            <Card key={module.id} className="overflow-hidden">
              <CardHeader className="pb-3 sm:pb-4">
                <div className="flex flex-col sm:flex-row sm:items-start sm:justify-between space-y-3 sm:space-y-0">
                  <div className="flex items-center space-x-3 min-w-0 flex-1">
                    <div className={`w-10 h-10 sm:w-12 sm:h-12 bg-gradient-to-r ${module.color} rounded-lg flex items-center justify-center shadow-lg flex-shrink-0`}>
                      <IconComponent className="text-white" size={20} />
                    </div>
                    <div className="min-w-0 flex-1">
                      <CardTitle className="text-base sm:text-xl text-gray-900 truncate">{module.title}</CardTitle>
                      <CardDescription className="text-xs sm:text-sm mt-1 line-clamp-2">{module.description}</CardDescription>
                    </div>
                  </div>
                  <div className="flex items-center flex-wrap gap-2 sm:space-x-2 sm:flex-nowrap">
                    <Badge 
                      variant={module.category === 'Safety Critical' ? 'destructive' : module.category === 'Equipment Level' ? 'default' : 'secondary'} 
                      className={`whitespace-nowrap ${
                        module.category === 'Safety Critical' 
                          ? 'bg-red-100 text-red-800 border-red-200' 
                          : module.category === 'Equipment Level' 
                          ? 'bg-orange-100 text-orange-800 border-orange-200' 
                          : 'bg-blue-100 text-blue-800 border-blue-200'
                      }`}
                      data-testid={`badge-category-${module.id}`}
                    >
                      {module.category}
                    </Badge>
                    <Badge variant="secondary" className="whitespace-nowrap bg-gray-100 text-gray-800" data-testid={`badge-modules-${module.id}`}>
                      {module.modules?.length || 0} modules
                    </Badge>
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
                      <div className="space-y-3">
                        {module.modules && module.modules.length > 0 ? (
                          module.modules.map((trainingModule) => (
                            <div 
                              key={trainingModule.id} 
                              className="p-3 bg-gray-50 rounded-md border-l-4 border-blue-500"
                              data-testid={`training-module-${trainingModule.code}`}
                            >
                              <div className="flex items-start justify-between">
                                <div className="flex-1">
                                  <div className="flex items-center space-x-2 mb-1">
                                    <span className="text-sm font-mono font-medium text-blue-600">{trainingModule.code}</span>
                                    <span className="text-sm font-medium text-gray-900">{trainingModule.name}</span>
                                  </div>
                                  {trainingModule.description && (
                                    <p className="text-xs text-gray-600 leading-relaxed line-clamp-2">{trainingModule.description}</p>
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
                        ) : (
                          <div className="text-sm text-gray-500 italic p-3 bg-gray-50 rounded-md">
                            No training modules defined for this category yet.
                          </div>
                        )}
                      </div>
                    </AccordionContent>
                  </AccordionItem>
                  
                  <AccordionItem value="competency-levels">
                    <AccordionTrigger className="text-sm font-medium">
                      Competency Levels
                    </AccordionTrigger>
                    <AccordionContent>
                      <div className="space-y-4">
                        {module.category === 'Safety Critical' ? (
                          <>
                            <div className="flex items-start space-x-3">
                              <XCircle className="w-5 h-5 text-gray-400 mt-0.5" />
                              <div>
                                <h4 className="font-medium text-sm">Not Trained</h4>
                                <p className="text-sm text-gray-600">
                                  No formal training completed; cannot use equipment or perform tasks.
                                </p>
                              </div>
                            </div>
                            
                            <div className="flex items-start space-x-3">
                              <BookOpen className="w-5 h-5 text-yellow-500 mt-0.5" />
                              <div>
                                <h4 className="font-medium text-sm">In Training (Supervised)</h4>
                                <p className="text-sm text-gray-600">
                                  Initial training and supervised practice underway; requires constant supervision.
                                </p>
                              </div>
                            </div>
                            
                            <div className="flex items-start space-x-3">
                              <Users className="w-5 h-5 text-green-400 mt-0.5" />
                              <div>
                                <h4 className="font-medium text-sm">Competent – Supervised</h4>
                                <p className="text-sm text-gray-600">
                                  Passed basic competency checks; can operate under supervision only.
                                </p>
                              </div>
                            </div>
                            
                            <div className="flex items-start space-x-3">
                              <CheckCircle className="w-5 h-5 text-green-600 mt-0.5" />
                              <div>
                                <h4 className="font-medium text-sm">Competent – SOP/Module</h4>
                                <p className="text-sm text-gray-600">
                                  Signed off against SOP/module requirements; authorized to work independently.
                                </p>
                              </div>
                            </div>
                            
                            <div className="flex items-start space-x-3">
                              <Crown className="w-5 h-5 text-blue-600 mt-0.5" />
                              <div>
                                <h4 className="font-medium text-sm">Expert</h4>
                                <p className="text-sm text-gray-600">
                                  Demonstrates advanced proficiency; may mentor others and conduct assessments.
                                </p>
                              </div>
                            </div>
                            
                            <div className="flex items-start space-x-3">
                              <AlertTriangle className="w-5 h-5 text-red-500 mt-0.5" />
                              <div>
                                <h4 className="font-medium text-sm">Ongoing Requirements</h4>
                                <p className="text-sm text-gray-600">
                                  Reassessment required as per the SOP for safety-critical equipment. 
                                  Additional licensing may be required (e.g., forklift operation).
                                </p>
                              </div>
                            </div>
                            
                            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mt-4 pt-4 border-t">
                              <div className="flex items-start space-x-3">
                                <Clock className="w-5 h-5 text-orange-500 mt-0.5" />
                                <div>
                                  <h4 className="font-medium text-sm">Expiry Date</h4>
                                  <p className="text-sm text-gray-600">
                                    Training expires as per SOP requirements and needs renewal.
                                  </p>
                                </div>
                              </div>
                              
                              <div className="flex items-start space-x-3">
                                <CheckCircle className="w-5 h-5 text-blue-500 mt-0.5" />
                                <div>
                                  <h4 className="font-medium text-sm">Applied Date</h4>
                                  <p className="text-sm text-gray-600">
                                    When this training was first applied/used in practice.
                                  </p>
                                </div>
                              </div>
                            </div>
                          </>
                        ) : (
                          <>
                            <div className="flex flex-wrap gap-2 mb-4">
                              {module.competencyLevels.map((level, index) => {
                                const getBadgeStyle = (level: string) => {
                                  if (level.includes('Not Trained')) return 'bg-gray-100 text-gray-800 border-gray-200';
                                  if (level.includes('In Training')) return 'bg-yellow-100 text-yellow-800 border-yellow-200';
                                  if (level.includes('Competent')) return 'bg-green-100 text-green-800 border-green-200';
                                  if (level.includes('Expert')) return 'bg-blue-100 text-blue-800 border-blue-200';
                                  return 'bg-gray-100 text-gray-800 border-gray-200';
                                };
                                
                                return (
                                  <Badge 
                                    key={index} 
                                    variant="secondary" 
                                    className={`${getBadgeStyle(level)} whitespace-nowrap`}
                                    data-testid={`badge-competency-${level.toLowerCase().replace(/[^a-z0-9]/g, '-')}`}
                                  >
                                    {level}
                                  </Badge>
                                );
                              })}
                            </div>
                            
                            <p className="text-sm text-gray-600">
                              Competency maintained through regular practice and application. 
                              No Certificate Number or Photo Evidence required for non-critical skills.
                            </p>
                          </>
                        )}
                      </div>
                    </AccordionContent>
                  </AccordionItem>
                </Accordion>
              </CardContent>
            </Card>
          );
        })}
      </div>

      {/* Training Summary */}
      <Card className="bg-blue-50 border-blue-200">
        <CardContent className="p-6">
          <div className="flex items-start space-x-3">
            <HardHat className="w-6 h-6 text-blue-600 mt-1" />
            <div>
              <h3 className="font-semibold text-blue-900 mb-2">Training Management</h3>
              <p className="text-blue-800 text-sm mb-4">
                All "Able to Use" training records are tracked in the Skills Matrix. 
                Staff competency levels are updated as training progresses and assessments are completed.
              </p>
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-sm">
                <div className="flex items-center space-x-2">
                  <CheckCircle className="w-4 h-4 text-green-600" />
                  <span className="text-blue-800">Competency tracking</span>
                </div>
                <div className="flex items-center space-x-2">
                  <Clock className="w-4 h-4 text-orange-600" />
                  <span className="text-blue-800">Expiry monitoring</span>
                </div>
                <div className="flex items-center space-x-2">
                  <AlertTriangle className="w-4 h-4 text-red-600" />
                  <span className="text-blue-800">Safety compliance</span>
                </div>
              </div>
            </div>
          </div>
        </CardContent>
      </Card>
    </div>
  );
}
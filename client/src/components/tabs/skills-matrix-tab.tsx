import { useState, useMemo, useEffect } from "react";
import { useQuery, useMutation } from "@tanstack/react-query";
import { useForm } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import { z } from "zod";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Accordion, AccordionContent, AccordionItem, AccordionTrigger } from "@/components/ui/accordion";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogTrigger, DialogFooter, DialogDescription } from "@/components/ui/dialog";
import { DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuTrigger } from "@/components/ui/dropdown-menu";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Input } from "@/components/ui/input";
import { Skeleton } from "@/components/ui/skeleton";
import { Form, FormControl, FormField, FormItem, FormLabel, FormMessage } from "@/components/ui/form";
import { Textarea } from "@/components/ui/textarea";
import { Checkbox } from "@/components/ui/checkbox";
import { Label } from "@/components/ui/label";
import { useToast } from "@/hooks/use-toast";
import { apiRequest, queryClient } from "@/lib/queryClient";
import { authService } from "@/auth/authService";
import { COMPETENCY_LEVELS, FIELD_COMPETENCY_LEVELS, ADMIN_COMPETENCY_LEVELS, getCompetencyLevels } from "@shared/schema";

import { 
  User, 
  Award, 
  Calendar, 
  AlertTriangle, 
  CheckCircle, 
  Plus, 
  Download,
  Camera,
  RotateCcw,
  Filter,
  Upload,
  FileText,
  Save,
  X,
  ExternalLink,
  Clock,
  Shield
} from "lucide-react";

type Staff = {
  id: number;
  name: string;
  email: string | null;
  jobTitle: string | null;
  startDate: string | null;
  isActive: boolean;
  isFieldStaff: boolean;
  isAdministrationStaff: boolean;
};

type TrainingClassification = {
  id: number;
  key: string;
  name: string;
  description: string | null;
  audience: string;
  displayOrder: number;
  color: string;
  isActive: boolean;
};

type TrainingModule = {
  id: number;
  classificationId: number;
  code: string;
  name: string;
  description: string | null;
  audience: string;
  sopUrl: string | null;
  trainingVideoUrl: string | null;
  validityPeriod: number | null;
  displayOrder: number;
  isActive: boolean;
  createdAt: string;
  updatedAt: string;
};

type StaffModuleProgress = {
  id: number;
  staffId: number;
  moduleId: number;
  status: string;
  competencyLevel: string | null;
  completedDate: string | null;
  expiryDate: string | null;
  assessorName: string | null;
  certificateNumber: string | null;
  evidenceUrl: string | null;
  evidenceFilename: string | null;
  notes: string | null;
  createdAt: string;
  updatedAt: string;
};

type TrainingModuleMatrixRow = {
  id?: number;
  staffId: number;
  staffName: string;
  moduleId: number;
  moduleName: string;
  moduleCode: string;
  classificationName: string;
  status: string;
  competencyLevel: string | null;
  completedDate: Date | null;
  expiryDate: Date | null;
  ableToUse?: boolean;
  achievedDate?: string;
  assessorName?: string | null;
  trainingProvider?: string | null;
  certificateNumber?: string | null;
  notes?: string | null;
  appliedDate?: string | null;
  photoEvidenceUrl?: string | null;
  photoEvidenceFilename?: string | null;
};

type PhotoAsset = {
  id: number;
  staffId: number;
  fileName: string;
  sharePointFileId: string;
  sharePointWebUrl: string | null;
  thumbnailUrl: string | null;
  fileSize: number;
  mimeType: string;
  uploadedBy: string;
  uploadedAt: string;
  isActive: boolean;
};

// Competency level colors and badges - Updated to 5-level system
const competencyConfig = {
  'Not Trained': { color: 'bg-gray-100 text-gray-700 border-gray-300', priority: 0 },
  'In Training (Supervised)': { color: 'bg-yellow-100 text-yellow-700 border-yellow-300', priority: 1 },
  'Competent – Supervised': { color: 'bg-orange-100 text-orange-700 border-orange-300', priority: 2 },
  'Competent – SOP/Module': { color: 'bg-green-100 text-green-700 border-green-300', priority: 3 },
  'Expert': { color: 'bg-blue-100 text-blue-700 border-blue-300', priority: 4 }
};

function AbilityBadge({ 
  ableToUse, 
  expiryDate, 
  competencyLevel,
  onClick 
}: { 
  ableToUse?: boolean; 
  expiryDate: string | null; 
  competencyLevel?: string;
  onClick?: () => void;
}) {
  const isExpired = expiryDate && new Date(expiryDate) < new Date();
  const isExpiringSoon = expiryDate && !isExpired && new Date(expiryDate) < new Date(Date.now() + 90 * 24 * 60 * 60 * 1000); // 3 months = 90 days

  // H&S Policy Color Coding: RED for expired, YELLOW for expiring within 3 months
  let badgeClass = '';
  let displayText = '';
  
  if (isExpired) {
    badgeClass = 'bg-red-500 text-white border-red-500'; // RED for expired (as per H&S policy)
    displayText = 'EXPIRED';
  } else if (isExpiringSoon) {
    badgeClass = 'bg-yellow-500 text-white border-yellow-500'; // YELLOW for expiring soon (as per H&S policy)
    displayText = ableToUse ? 'Able to Use' : 'Not Able';
  } else if (ableToUse) {
    badgeClass = 'bg-green-500 text-white border-green-500'; // GREEN for able to use
    displayText = 'Able to Use';
  } else {
    badgeClass = 'bg-gray-400 text-white border-gray-400'; // GRAY for not able to use
    displayText = 'Not Able';
  }

  return (
    <div 
      className="flex items-center justify-center cursor-pointer hover:opacity-80" 
      onClick={onClick}
      title={`${displayText}${competencyLevel ? ` (${competencyLevel})` : ''}${expiryDate ? ` - Expires: ${new Date(expiryDate).toLocaleDateString('en-GB')}` : ''}`}
    >
      <Badge 
        className={`text-xs font-medium ${badgeClass} ${isExpired ? 'line-through' : ''}`}
        data-testid={`badge-able-to-use-${ableToUse ? 'yes' : 'no'}`}
      >
        {displayText}
      </Badge>
      {isExpired && (
        <AlertTriangle className="w-3 h-3 text-red-600 ml-1" data-testid="icon-expired" />
      )}
      {isExpiringSoon && !isExpired && (
        <Calendar className="w-3 h-3 text-yellow-600 ml-1" data-testid="icon-expiring-soon" />
      )}
    </div>
  );
}

function StaffPhotoCell({ staffId, name, jobTitle }: { staffId: number; name: string; jobTitle?: string | null }) {
  const { data: photo } = useQuery<{ success: boolean; data: PhotoAsset | null }>({
    queryKey: [`/api/photos/staff/${staffId}`],
    enabled: !!staffId
  });

  return (
    <div className="flex items-center space-x-3" data-testid={`staff-row-${staffId}`}>
      <div className="w-10 h-10 rounded-full bg-gray-200 flex items-center justify-center overflow-hidden">
        {photo?.data?.thumbnailUrl ? (
          <img 
            src={photo.data.thumbnailUrl} 
            alt={name}
            className="w-full h-full object-cover"
            data-testid={`photo-${staffId}`}
          />
        ) : (
          <User className="w-5 h-5 text-gray-500" data-testid={`photo-placeholder-${staffId}`} />
        )}
      </div>
      <div>
        <div className="font-medium text-sm" data-testid={`text-staff-name-${staffId}`}>{name}</div>
        {jobTitle && (
          <div className="text-xs text-gray-500 mt-1" data-testid={`text-staff-title-${staffId}`}>{jobTitle}</div>
        )}
      </div>
    </div>
  );
}

export default function SkillsMatrixTab() {
  const { toast } = useToast();
  const [selectedStaffCategory, setSelectedStaffCategory] = useState<string>('field');
  
  // Fetch data
  const { data: staffData, isLoading: staffLoading } = useQuery<{ success: boolean; data: Staff[] }>({
    queryKey: ['/api/staff']
  });

  const { data: classificationsData, isLoading: classificationsLoading } = useQuery<{ success: boolean; data: TrainingClassification[] }>({
    queryKey: ['/api/training-classifications']
  });

  const { data: modulesData, isLoading: modulesLoading } = useQuery<{ success: boolean; data: TrainingModule[] }>({
    queryKey: ['/api/training-modules']
  });

  const { data: matrixApiData, isLoading: matrixLoading } = useQuery<{ success: boolean; data: TrainingModuleMatrixRow[] }>({
    queryKey: ['/api/training-module-matrix']
  });

  // Get current user's role for filtering
  const { data: currentUserData, isLoading: userLoading } = useQuery<{ 
    success: boolean; 
    user: { 
      email: string; 
      name: string; 
      jobTitle: string; 
      role: string; 
      roleRank: number; 
    } 
  }>({
    queryKey: ['/api/current-user/role'],
    queryFn: async () => {
      const token = await authService.getAccessToken();
      const response = await fetch('/api/current-user/role', {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Accept': 'application/json'
        },
        credentials: 'include'
      });
      
      if (!response.ok) {
        throw new Error('Failed to fetch user role');
      }
      
      return await response.json();
    }
  });

  // SharePoint sync mutation
  const syncMutation = useMutation({
    mutationFn: async () => {
      const token = await authService.getAccessToken();
      const response = await fetch('/api/staff/sync-sharepoint', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Accept': 'application/json'
        },
        credentials: 'include'
      });
      
      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Failed to sync staff from SharePoint: ${response.status} - ${errorText}`);
      }
      
      return await response.json();
    },
    onSuccess: (data) => {
      toast({
        title: "SharePoint Sync Successful",
        description: data.message,
      });
      // Invalidate staff data to refresh the matrix
      // Invalidate all related queries to force refresh
      queryClient.invalidateQueries({ queryKey: ['/api/staff'] });
      queryClient.invalidateQueries({ queryKey: ['/api/training-classifications'] });
      queryClient.invalidateQueries({ queryKey: ['/api/training-modules'] });
      queryClient.invalidateQueries({ queryKey: ['/api/training-module-matrix'] });
      
      // Force refetch to bypass cache
      queryClient.refetchQueries({ queryKey: ['/api/staff'] });
    },
    onError: (error) => {
      toast({
        title: "Sync Failed",
        description: error instanceof Error ? error.message : 'Failed to sync staff from SharePoint',
        variant: "destructive",
      });
    }
  });

  // Handle SharePoint sync button click
  const handleSharePointSync = () => {
    syncMutation.mutate();
  };

  // Bulk compliance update mutation
  const bulkUpdateMutation = useMutation({
    mutationFn: async ({ expiryDate }: { expiryDate: string }) => {
      const response = await apiRequest('POST', '/api/training-records/bulk-compliance-update', { expiryDate });
      return await response.json();
    },
    onSuccess: (data) => {
      const expiryDateFormatted = new Date(data.data.expiryDate).toLocaleDateString('en-NZ', { 
        year: 'numeric', 
        month: 'long', 
        day: 'numeric' 
      });
      toast({
        title: "Compliance Update Successful",
        description: `Successfully processed ${data.data.recordsProcessed} training records for ${data.data.totalStaff} field staff across ${data.data.totalModules} modules. All set to "Able to Use" with expiry date ${expiryDateFormatted}.`,
      });
      // Invalidate all related queries to refresh the matrix
      queryClient.invalidateQueries({ queryKey: ['/api/training-module-matrix'] });
      queryClient.invalidateQueries({ queryKey: ['/api/training-records'] });
      queryClient.refetchQueries({ queryKey: ['/api/training-module-matrix'] });
    },
    onError: (error) => {
      toast({
        title: "Bulk Update Failed",
        description: error instanceof Error ? error.message : 'Failed to perform bulk compliance update',
        variant: "destructive",
      });
    }
  });

  // Bulk compliance update dialog state
  const [showComplianceDialog, setShowComplianceDialog] = useState(false);
  const [complianceExpiryDate, setComplianceExpiryDate] = useState('2026-02-28');

  // Handle bulk compliance update button click
  const handleBulkComplianceUpdate = () => {
    setShowComplianceDialog(false);
    bulkUpdateMutation.mutate({ expiryDate: complianceExpiryDate });
  };

  // Export to Enhanced CSV function - Visual and Professional
  const exportToCSV = async () => {
    try {
      // Show loading state
      const button = document.querySelector('.export-csv-btn') as HTMLButtonElement;
      if (button) {
        button.disabled = true;
        button.innerHTML = '<i class="fas fa-spinner fa-spin text-xs"></i> <span>Generating Professional CSV...</span>';
      }

      // Process data to include visual indicators and risk levels
      const enhancedData = staff.map(staffMember => {
        const row: any = {
          'Staff Member': staffMember.name,
          'Job Title': staffMember.jobTitle,
          'Start Date': staffMember.startDate ? new Date(staffMember.startDate).toLocaleDateString('en-GB') : 'Not Set',
        };

        // Group modules by classification for better organization
        const classificationGroups = classifications.reduce((acc, classification) => {
          const classificationModules = modules.filter(m => m.classificationId === classification.id);
          if (classificationModules.length > 0) {
            const riskLevel = classification.key === 'safety-critical' ? 'SAFETY CRITICAL' :
                            classification.key === 'equipment-training' ? 'EQUIPMENT LEVEL' : 'FOUNDATION LEVEL';
            
            acc[`${classification.name} (${riskLevel})`] = classificationModules;
          }
          return acc;
        }, {} as Record<string, any[]>);

        // Add training status for each module with visual indicators
        Object.entries(classificationGroups).forEach(([groupName, groupModules]) => {
          groupModules.forEach(module => {
            const trainingRecord = processedMatrixData[staffMember.id]?.[module.id];
            const columnName = `${module.name} (${module.code})`;
            
            if (trainingRecord) {
              const isExpired = trainingRecord.expiryDate && new Date(trainingRecord.expiryDate) < new Date();
              const isExpiringSoon = trainingRecord.expiryDate && !isExpired && 
                new Date(trainingRecord.expiryDate) < new Date(Date.now() + 90 * 24 * 60 * 60 * 1000);

              let status = '';
              if (isExpired) {
                status = '🔴 EXPIRED';
              } else if (trainingRecord.ableToUse) {
                status = isExpiringSoon ? '🟡 ABLE TO USE (Expires Soon)' : '🟢 ABLE TO USE';
              } else {
                status = '🔴 NOT ABLE TO USE';
              }

              const competencyLevel = trainingRecord.competencyLevel || 'Not Set';
              const expiryInfo = trainingRecord.expiryDate 
                ? ` | Expires: ${new Date(trainingRecord.expiryDate).toLocaleDateString('en-GB')}` 
                : ' | No Expiry Set';

              row[columnName] = `${status} | ${competencyLevel}${expiryInfo}`;
            } else {
              row[columnName] = '⚪ NOT TRAINED';
            }
          });
        });

        return row;
      });

      // Create CSV content manually with enhanced formatting
      const csvContent = convertToEnhancedCSV(enhancedData);
      
      // Download the enhanced CSV
      const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.style.display = 'none';
      a.href = url;
      a.download = `Cranfield-Skills-Matrix-Professional-${new Date().toISOString().split('T')[0]}.csv`;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);

      toast({
        title: "Professional Export Successful",
        description: "Enhanced Skills Matrix CSV with visual indicators downloaded successfully",
      });

    } catch (error) {
      console.error('Error generating enhanced CSV:', error);
      toast({
        title: "Export Failed", 
        description: "Failed to generate professional CSV. Please try again.",
        variant: "destructive",
      });
    } finally {
      // Reset button state
      const button = document.querySelector('.export-csv-btn') as HTMLButtonElement;
      if (button) {
        button.disabled = false;
        button.innerHTML = '<i class="fas fa-file-csv text-xs"></i> <span>Export Professional CSV</span>';
      }
    }
  };

  // Helper function to convert data to enhanced CSV format
  const convertToEnhancedCSV = (data: any[]): string => {
    if (data.length === 0) return '';
    
    const headers = Object.keys(data[0]);
    const csvRows = [headers.join(',')];
    
    // Add header row with classification info
    csvRows.unshift(`"CRANFIELD GLASS CHRISTCHURCH - STAFF TRAINING MATRIX - PROFESSIONAL EXPORT"`);
    csvRows.unshift(`"Generated: ${new Date().toLocaleString('en-GB')}"`);
    csvRows.unshift(`"Legend: 🟢=Able to Use | 🟡=Expiring Soon | 🔴=Expired/Not Able | ⚪=Not Trained"`);
    csvRows.unshift('');
    
    data.forEach(row => {
      const values = headers.map(header => {
        const value = row[header];
        // Escape quotes and wrap in quotes if contains comma or quotes
        const escapedValue = String(value).replace(/"/g, '""');
        return `"${escapedValue}"`;
      });
      csvRows.push(values.join(','));
    });
    
    return csvRows.join('\n');
  };

  // Training Record Form Schema (H&S Policy requirements)
  const trainingRecordSchema = z.object({
    competencyLevel: z.string().min(1, "Competency level is required"),
    achievedDate: z.string().optional().refine((val) => {
      // Convert empty string to undefined for optional dates
      return val === "" || val === undefined || !isNaN(Date.parse(val));
    }, "Invalid date format"),
    expiryDate: z.string().optional().refine((val) => {
      return val === "" || val === undefined || !isNaN(Date.parse(val));
    }, "Invalid date format"),
    assessorName: z.string().optional(),
    trainingProvider: z.string().optional(),
    certificateNumber: z.string().optional(),
    notes: z.string().optional(),
    ableToUse: z.boolean().default(false),
    appliedDate: z.string().optional().refine((val) => {
      return val === "" || val === undefined || !isNaN(Date.parse(val));
    }, "Invalid date format"),
    photoEvidenceUrl: z.string().optional(),
    photoEvidenceFilename: z.string().optional(),
    // These fields will be added programmatically in the mutation
    staffId: z.number().optional(),
    skillId: z.number().optional(),
    status: z.string().optional(),
  });

  // Modal state for training record editing
  const [editModalOpen, setEditModalOpen] = useState(false);
  const [selectedRecord, setSelectedRecord] = useState<{
    staff: Staff;
    skill: TrainingModule;
    record: TrainingModuleMatrixRow | null;
  } | null>(null);
  

  // Enhanced date validation schema with object-level validation
  const enhancedTrainingRecordSchema = trainingRecordSchema.extend({}).superRefine((data, ctx) => {
    // Achieved date validation
    if (data.achievedDate) {
      const achievedDate = new Date(data.achievedDate);
      const today = new Date();
      today.setHours(23, 59, 59, 999); // Allow today
      if (achievedDate > today) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message: "Achieved date cannot be in the future",
          path: ["achievedDate"]
        });
      }
    }

    // Expiry date validation
    if (data.expiryDate && data.achievedDate) {
      const expiryDate = new Date(data.expiryDate);
      const achievedDate = new Date(data.achievedDate);
      if (expiryDate <= achievedDate) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message: "Expiry date must be after achieved date",
          path: ["expiryDate"]
        });
      }
    }

    // Applied date validation
    if (data.appliedDate && data.achievedDate) {
      const appliedDate = new Date(data.appliedDate);
      const achievedDate = new Date(data.achievedDate);
      if (appliedDate < achievedDate) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message: "Applied date must be on or after achieved date",
          path: ["appliedDate"]
        });
      }
    }
  });

  // Form for training record editing
  const form = useForm<z.infer<typeof trainingRecordSchema>>({
    resolver: zodResolver(enhancedTrainingRecordSchema),
    defaultValues: {
      competencyLevel: "Not Trained",
      achievedDate: "",
      expiryDate: "",
      assessorName: "",
      trainingProvider: "",
      certificateNumber: "",
      notes: "",
      ableToUse: false,
      appliedDate: "",
      photoEvidenceUrl: "",
      photoEvidenceFilename: "",
    },
  });

  // Handle cell click for editing training records (H&S Policy requirement)
  const handleCellClick = (staffMember: Staff, skill: TrainingModule, trainingRecord: TrainingModuleMatrixRow | null) => {
    setSelectedRecord({ staff: staffMember, skill, record: trainingRecord });
    
    // Pre-populate form with existing data
    if (trainingRecord) {
      form.reset({
        competencyLevel: trainingRecord.competencyLevel || "Not Trained",
        achievedDate: trainingRecord.achievedDate ? trainingRecord.achievedDate.split('T')[0] : "", // Convert to date format
        expiryDate: trainingRecord.expiryDate ? String(trainingRecord.expiryDate).split('T')[0] : "",
        assessorName: trainingRecord.assessorName || "",
        trainingProvider: trainingRecord.trainingProvider || "",
        certificateNumber: trainingRecord.certificateNumber || "",
        notes: trainingRecord.notes || "",
        ableToUse: trainingRecord.status === 'competent' || false,
        appliedDate: trainingRecord.appliedDate ? trainingRecord.appliedDate.split('T')[0] : "",
        photoEvidenceUrl: trainingRecord.photoEvidenceUrl || "",
        photoEvidenceFilename: trainingRecord.photoEvidenceFilename || "",
      });
    } else {
      form.reset();
    }
    
    setEditModalOpen(true);
  };

  // Mutation for saving training records
  const saveTrainingMutation = useMutation({
    mutationFn: async (data: z.infer<typeof trainingRecordSchema>) => {
      if (!selectedRecord) throw new Error("No record selected");
      
      const payload: any = {
        ...data,
        staffId: selectedRecord.staff.id,
        skillId: selectedRecord.skill.id,
        status: 'Active',
        // Convert empty strings to null for date fields to match backend schema
        achievedDate: data.achievedDate && data.achievedDate.trim() !== "" ? data.achievedDate : null,
        appliedDate: data.appliedDate && data.appliedDate.trim() !== "" ? data.appliedDate : null,
        expiryDate: data.expiryDate && data.expiryDate.trim() !== "" ? data.expiryDate : null,
      };

      if (selectedRecord.record) {
        // Versioning system: Always create new record, don't overwrite existing
        payload.previousRecordId = selectedRecord.record.id; // Track previous version
      }
      
      return apiRequest('POST', '/api/training-records', payload);
    },
    onSuccess: () => {
      toast({
        title: "Success",
        description: "Training record saved successfully",
      });
      
      // Invalidate and refetch data - fix persistence across users
      queryClient.invalidateQueries({ queryKey: ['/api/training-records'] });
      queryClient.invalidateQueries({ queryKey: ['/api/training-module-matrix'] }); // Main matrix data
      
      // Close modal and reset form
      setEditModalOpen(false);
      setSelectedRecord(null);
      form.reset();
    },
    onError: (error: any) => {
      toast({
        title: "Error",
        description: error.message || "Failed to save training record",
        variant: "destructive",
      });
    }
  });

  // Handle form submission
  const onSubmit = (data: z.infer<typeof trainingRecordSchema>) => {
    saveTrainingMutation.mutate(data);
  };

  // Export to HTML function
  const exportToHTML = async () => {
    try {
      // Show loading state
      const button = document.querySelector('.export-html-btn') as HTMLButtonElement;
      if (button) {
        button.disabled = true;
        button.innerHTML = '<i class="fas fa-spinner fa-spin text-xs"></i> <span>Generating HTML...</span>';
      }

      const response = await fetch('/api/export-skills-matrix-html', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          staffFilter: selectedStaffCategory,
          skillCategory: selectedStaffCategory || 'all'
        })
      });

      if (!response.ok) {
        throw new Error('Failed to generate HTML');
      }

      const result = await response.json();
      
      // Download the HTML file
      const blob = new Blob([result.htmlContent], { type: 'text/html' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.style.display = 'none';
      a.href = url;
      a.download = result.filename || 'skills-matrix.html';
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);

      toast({
        title: "Export Successful",
        description: "Skills Matrix HTML downloaded successfully",
      });

    } catch (error) {
      console.error('Error generating HTML:', error);
      toast({
        title: "Export Failed",
        description: "Failed to generate HTML. Please try again.",
        variant: "destructive",
      });
    } finally {
      // Reset button state
      const button = document.querySelector('.export-html-btn') as HTMLButtonElement;
      if (button) {
        button.disabled = false;
        button.innerHTML = '<i class="fas fa-file-code text-xs"></i> <span>Export HTML</span>';
      }
    }
  };

  // Filter and organize data
  const staff = useMemo(() => {
    if (!staffData?.data) return [];
    
    // Define job title hierarchy for sorting
    const titleHierarchy: Record<string, number> = {
      'Managing Director': 1,
      'Office Manager': 2,
      'Accounts Manager': 3,
      'Field Supervisor': 4,
      'Senior Glazier': 5,
      'Glazier': 6,
      'Apprentice Glazier': 7,
      'Apprentice': 8,
    };
    
    return staffData.data
      .filter(s => 
        // Only show active staff
        s.isActive &&
        // Only show staff with proper job titles (exclude system accounts)
        s.jobTitle && s.jobTitle.trim() !== '' &&
        // Filter by staff category
        ((selectedStaffCategory === 'field' && s.isFieldStaff) ||
         (selectedStaffCategory === 'administration' && s.isAdministrationStaff))
      )
      .sort((a, b) => {
        // Sort by job title hierarchy, then by name
        const aOrder = titleHierarchy[a.jobTitle || ''] || 999;
        const bOrder = titleHierarchy[b.jobTitle || ''] || 999;
        
        if (aOrder !== bOrder) {
          return aOrder - bOrder;
        }
        
        // If same hierarchy level, sort by name
        return (a.name || '').localeCompare(b.name || '');
      });
  }, [staffData, selectedStaffCategory]);

  // Helper to determine allowed audiences based on staff category selection
  const allowedAudiences = useMemo(() => {
    if (selectedStaffCategory === 'field') return ['field', 'both'];
    if (selectedStaffCategory === 'administration') return ['administration', 'both'];
    return ['both'];
  }, [selectedStaffCategory]);

  const classifications = useMemo(() => {
    if (!classificationsData?.data) return [];
    
    return classificationsData.data.filter(c => 
      c.isActive && allowedAudiences.includes(c.audience)
    ).sort((a, b) => a.displayOrder - b.displayOrder);
  }, [classificationsData, allowedAudiences]);

  // Auto-calculation logic for enhanced form controls
  const watchedFields = form.watch(["achievedDate", "competencyLevel", "expiryDate"]);
  const [achievedDate, competencyLevel, expiryDate] = watchedFields;

  // Timezone-safe date formatter
  const formatDateYMD = (date: Date): string => {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  };

  // Auto-calculate expiry date when achieved date changes
  useEffect(() => {
    if (!editModalOpen || !selectedRecord) return;
    
    // Add a small delay to avoid interfering with form.reset()
    const timer = setTimeout(() => {
      if (achievedDate && selectedRecord.skill.validityPeriod) {
        const achieved = new Date(achievedDate);
        achieved.setMonth(achieved.getMonth() + selectedRecord.skill.validityPeriod);
        const calculatedExpiry = formatDateYMD(achieved);
        
        // Only update if different to avoid infinite loops
        if (form.getValues("expiryDate") !== calculatedExpiry) {
          form.setValue("expiryDate", calculatedExpiry, { shouldDirty: true });
        }
      } else if (!achievedDate) {
        // Clear expiry when achieved date is cleared
        if (form.getValues("expiryDate")) {
          form.setValue("expiryDate", "", { shouldDirty: true });
        }
      }
    }, 100);
    
    return () => clearTimeout(timer);
  }, [achievedDate, selectedRecord?.skill.validityPeriod, editModalOpen, selectedRecord]);

  // Determine form variant based on classification
  const getFormVariant = (classificationKey: string): 'safety' | 'equipment' | 'foundation' => {
    const safetyKeys = ['lifting-handling-equipment', 'safety-equipment-ppe', 'working-at-heights'];
    const equipmentKeys = ['power-tools-equipment', 'vehicles-transport'];
    
    if (safetyKeys.includes(classificationKey)) return 'safety';
    if (equipmentKeys.includes(classificationKey)) return 'equipment';
    return 'foundation';
  };

  // Custom cell renderer for competency levels with beautiful heatmap colors
const CompetencyCellRenderer = (params: any) => {
  const { value, data, colDef } = params;
  
  if (!value) {
    return (
      <div className="h-full flex items-center justify-center">
        <Badge variant="outline" className="bg-gray-50 text-gray-500 text-xs">
          Not Trained
        </Badge>
      </div>
    );
  }

  const { competencyLevel, ableToUse, expiryDate } = value;
  const isExpired = expiryDate && new Date(expiryDate) < new Date();
  const isExpiringSoon = expiryDate && !isExpired && new Date(expiryDate) < new Date(Date.now() + 90 * 24 * 60 * 60 * 1000);

  // Competency level colors (heatmap style)
  const getCompetencyColor = (level: string) => {
    switch (level) {
      case 'Expert': return 'bg-blue-500 text-white';
      case 'Competent – SOP/Module': return 'bg-green-500 text-white';
      case 'Competent – Supervised': return 'bg-orange-500 text-white';
      case 'In Training (Supervised)': return 'bg-yellow-500 text-white';
      default: return 'bg-gray-400 text-white';
    }
  };

  let badgeColor = getCompetencyColor(competencyLevel);
  let displayText = competencyLevel?.replace('Competent – ', '') || 'Unknown';
  
  // Override with expiry status if needed
  if (isExpired) {
    badgeColor = 'bg-red-500 text-white';
    displayText = 'EXPIRED';
  } else if (isExpiringSoon) {
    badgeColor = 'bg-yellow-600 text-white';
  }

  return (
    <div className="h-full flex items-center justify-center p-1">
      <div 
        className="relative group cursor-pointer"
        title={`${competencyLevel}${ableToUse ? ' - Able to Use' : ''}${expiryDate ? ` - Expires: ${new Date(expiryDate).toLocaleDateString('en-GB')}` : ''}`}
      >
        <Badge className={`text-xs font-medium ${badgeColor} ${isExpired ? 'line-through' : ''}`}>
          {displayText}
        </Badge>
        {ableToUse && !isExpired && (
          <CheckCircle className="w-3 h-3 text-green-600 absolute -top-1 -right-1 bg-white rounded-full" />
        )}
        {isExpired && (
          <AlertTriangle className="w-3 h-3 text-red-600 absolute -top-1 -right-1 bg-white rounded-full" />
        )}
        {isExpiringSoon && !isExpired && (
          <Calendar className="w-3 h-3 text-yellow-600 absolute -top-1 -right-1 bg-white rounded-full" />
        )}
      </div>
    </div>
  );
};

// Staff cell renderer with photo and details
const StaffCellRenderer = (params: any) => {
  const { value } = params;
  if (!value) return null;
  
  const { id, name, jobTitle } = value;
  
  return (
    <div className="flex items-center space-x-3 p-2">
      <div className="w-8 h-8 bg-blue-500 rounded-full flex items-center justify-center text-white text-sm font-medium">
        {name?.charAt(0) || 'U'}
      </div>
      <div className="min-w-0 flex-1">
        <div className="text-sm font-medium truncate">{name}</div>
        <div className="text-xs text-gray-500 truncate">{jobTitle}</div>
      </div>
    </div>
  );
};

// Auto-derive "Able to Use" status based on competency and expiry
  useEffect(() => {
    if (!editModalOpen || !selectedRecord || selectedRecord.skill.audience === 'administration') return;
    
    // Add a small delay to avoid interfering with form.reset()
    const timer = setTimeout(() => {
      const currentDate = new Date();
      const expiryDateObj = expiryDate ? new Date(expiryDate) : null;
      
      const classification = classifications.find(c => c.id === selectedRecord.skill.classificationId);
      const formVariant = classification ? getFormVariant(classification.key) : 'foundation';
      
      // "Able to Use" logic based on form variant:
      // 1. Must have achieved competency (not "Not Trained")
      // 2. Must not be expired (if expiry date exists)
      // 3. Competency requirements vary by variant
      const isCompetent = competencyLevel && competencyLevel !== "Not Trained";
      const isNotExpired = !expiryDateObj || expiryDateObj > currentDate;
      
      // Competency requirements by variant
      const isHighLevel = formVariant === 'safety' ? 
        ["Competent – SOP/Module", "Expert"].includes(competencyLevel) : 
        ["Competent – Supervised", "Competent – SOP/Module", "Expert"].includes(competencyLevel);
      
      // Only show "Able to Use" for safety and equipment variants
      const showsAbleToUse = ['safety', 'equipment'].includes(formVariant);
      const calculatedAbleToUse = showsAbleToUse && isCompetent && isNotExpired && isHighLevel;
      
      // Only update if different to avoid infinite loops
      if (form.getValues("ableToUse") !== calculatedAbleToUse) {
        form.setValue("ableToUse", Boolean(calculatedAbleToUse), { shouldDirty: true });
      }
    }, 100);
    
    return () => clearTimeout(timer);
  }, [competencyLevel, expiryDate, selectedRecord, classifications, editModalOpen, getFormVariant]);

  const modules = useMemo(() => {
    if (!modulesData?.data) return [];
    
    return modulesData.data.filter(m => 
      m.isActive && 
      allowedAudiences.includes(m.audience)
    );
  }, [modulesData, allowedAudiences]);

  // Process matrix data from API (already comes in the correct format)
  const processedMatrixData = useMemo(() => {
    if (!matrixApiData?.data || !staff.length || !modules.length) return {};
    
    const matrix: Record<string, Record<string, TrainingModuleMatrixRow | null>> = {};
    
    // Initialize matrix for all staff and modules
    staff.forEach(s => {
      matrix[s.id] = {};
      modules.forEach(module => {
        matrix[s.id][module.id] = null;
      });
    });

    // Populate with matrix data from API
    matrixApiData.data.forEach(row => {
      if (matrix[row.staffId] && modules.find(m => m.id === row.moduleId)) {
        matrix[row.staffId][row.moduleId] = row;
      }
    });

    return matrix;
  }, [matrixApiData, staff, modules]);

  // Data transformation for AG Grid fancy matrix
  const { rowData, columnDefs } = useMemo(() => {
    if (!staff || !modules || !classifications || !processedMatrixData) {
      return { rowData: [], columnDefs: [] };
    }

    // Filter staff and classifications based on selected category
    let filteredStaff = staff.filter(s => s.isActive);
    let filteredClassifications = classifications.filter(c => c.isActive);
    
    if (selectedStaffCategory === 'field') {
      filteredStaff = filteredStaff.filter(s => s.isFieldStaff);
      filteredClassifications = filteredClassifications.filter(c => c.audience === 'field');
    } else if (selectedStaffCategory === 'administration') {
      filteredStaff = filteredStaff.filter(s => s.isAdministrationStaff);
      filteredClassifications = filteredClassifications.filter(c => c.audience === 'administration');
    }

    // Create row data (one row per staff member)
    const rows = filteredStaff.map(staffMember => {
      const row: any = {
        id: staffMember.id,
        staff: {
          id: staffMember.id,
          name: staffMember.name,
          jobTitle: staffMember.jobTitle,
        }
      };

      // Add module data for each training module
      modules.forEach(module => {
        if (filteredClassifications.some(c => c.id === module.classificationId)) {
          const trainingRecord = processedMatrixData[staffMember.id]?.[module.id];
          row[`module_${module.id}`] = trainingRecord ? {
            competencyLevel: trainingRecord.competencyLevel,
            ableToUse: trainingRecord.ableToUse,
            expiryDate: trainingRecord.expiryDate,
            staffMember,
            module,
            trainingRecord
          } : null;
        }
      });

      return row;
    });

    // Create column definitions with grouped headers
    const cols: any[] = [
      // Pinned staff column
      {
        field: 'staff',
        headerName: 'Staff Member',
        pinned: 'left',
        width: 200,
        cellRenderer: StaffCellRenderer,
        sortable: true,
        filter: true,
      }
    ];

    // Add grouped columns by classification
    filteredClassifications.forEach(classification => {
      const classificationModules = modules.filter(m => m.classificationId === classification.id);
      
      if (classificationModules.length > 0) {
        const groupCol: any = {
          headerName: classification.name,
          headerClass: 'ag-header-group-text-center',
          children: classificationModules.map(module => ({
            field: `module_${module.id}`,
            headerName: `${module.code}`,
            headerTooltip: module.name,
            width: 100,
            cellRenderer: CompetencyCellRenderer,
            onCellClicked: (params: any) => {
              const cellData = params.value;
              if (cellData) {
                handleCellClick(cellData.staffMember, cellData.module, cellData.trainingRecord);
              } else {
                // No training record exists, create new one
                const staffMember = filteredStaff.find(s => s.id === params.data.id);
                const moduleData = modules.find(m => m.id === module.id);
                if (staffMember && moduleData) {
                  handleCellClick(staffMember, moduleData, null);
                }
              }
            },
            cellStyle: { padding: '4px', textAlign: 'center' },
            suppressMenu: true,
          }))
        };
        cols.push(groupCol);
      }
    });

    return { rowData: rows, columnDefs: cols };
  }, [staff, modules, classifications, processedMatrixData, selectedStaffCategory, handleCellClick]);

  // AG Grid event handlers

  const isLoading = staffLoading || classificationsLoading || modulesLoading || matrixLoading || userLoading;

  if (isLoading) {
    return (
      <div className="space-y-6">
        <div className="flex justify-between items-center">
          <Skeleton className="h-8 w-48" />
          <Skeleton className="h-10 w-32" />
        </div>
        <Card>
          <CardHeader>
            <Skeleton className="h-6 w-64" />
            <Skeleton className="h-4 w-96" />
          </CardHeader>
          <CardContent>
            <div className="space-y-4">
              {Array.from({ length: 5 }).map((_, i) => (
                <div key={i} className="flex space-x-4">
                  <Skeleton className="h-10 w-10 rounded-full" />
                  <Skeleton className="h-10 flex-1" />
                </div>
              ))}
            </div>
          </CardContent>
        </Card>
      </div>
    );
  }

  return (
    <div className="space-y-6">
      {/* Page Header */}
      <div className="mb-6 flex justify-between items-start">
        <div>
          <h2 className="text-2xl font-bold text-gray-900">Staff Training Matrix</h2>
          <p className="text-gray-600">Track staff competencies and training records</p>
        </div>
        
        {/* Admin Actions Dropdown - Icon Only */}
        <DropdownMenu>
          <DropdownMenuTrigger asChild>
            <button
              className="p-2 text-gray-500 hover:text-gray-700 hover:bg-gray-100 rounded-md transition-colors"
              data-testid="button-admin-actions"
              title="Admin Actions"
            >
              <Shield className="w-5 h-5" />
            </button>
          </DropdownMenuTrigger>
          <DropdownMenuContent align="end">
            <DropdownMenuItem onClick={() => setShowComplianceDialog(true)} data-testid="menu-item-bulk-compliance">
              <CheckCircle className="w-4 h-4 mr-2" />
              Bulk Compliance Pass
            </DropdownMenuItem>
          </DropdownMenuContent>
        </DropdownMenu>
      </div>

      {/* Controls - Sticky Card */}
      <div className="sticky top-0 z-10 bg-white rounded-lg border border-gray-200 p-4 mb-6 shadow-sm">
        <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
          
          {/* Filter Section */}
          <div className="flex items-center gap-3">
            <span className="text-sm font-medium text-gray-700">Show staff:</span>
            <Select value={selectedStaffCategory} onValueChange={setSelectedStaffCategory}>
              <SelectTrigger className="w-40" data-testid="select-staff-category">
                <SelectValue placeholder="All staff" />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="field">Field Staff</SelectItem>
                <SelectItem value="administration">Administration</SelectItem>
              </SelectContent>
            </Select>
          </div>
          
          {/* Action Buttons */}
          <div className="flex flex-wrap gap-2">
            <button
              onClick={handleSharePointSync}
              disabled={syncMutation.isPending}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-sm text-white bg-blue-600 hover:bg-blue-700 border border-blue-600 rounded-md"
              data-testid="button-sync-sharepoint"
            >
              {syncMutation.isPending ? (
                <RotateCcw className="w-3 h-3 animate-spin" />
              ) : (
                <Upload className="w-3 h-3" />
              )}
              <span>{syncMutation.isPending ? 'Syncing...' : 'Sync from SharePoint'}</span>
            </button>
            
            <button
              onClick={exportToHTML}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-sm text-white bg-purple-600 hover:bg-purple-700 border border-purple-600 rounded-md"
              data-testid="button-export-html"
            >
              <FileText className="w-3 h-3" />
              <span>Export HTML</span>
            </button>
            <button
              onClick={exportToCSV}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-sm text-white bg-green-600 hover:bg-green-700 border border-green-600 rounded-md"
              data-testid="button-export-csv"
            >
              <Download className="w-3 h-3" />
              <span>Export CSV</span>
            </button>
          </div>
          
        </div>
      </div>

      {/* Training Records Matrix */}
      <Card>
        <CardHeader>
          <CardTitle className="flex items-center space-x-2">
            <Award className="w-5 h-5" />
            <span>Training Records</span>
          </CardTitle>
          <CardDescription className="space-y-2">
            <p>Click any cell to view or edit individual training records.</p>
            <div className="flex flex-wrap items-center gap-4 text-sm">
              <div className="flex items-center space-x-2">
                <div className="w-3 h-3 bg-red-500 rounded-full"></div>
                <span>Expired</span>
              </div>
              <div className="flex items-center space-x-2">
                <div className="w-3 h-3 bg-yellow-500 rounded-full"></div>
                <span>Expiring (3 months)</span>
              </div>
              <div className="flex items-center space-x-2">
                <div className="w-3 h-3 bg-green-500 rounded-full"></div>
                <span>Able to Use</span>
              </div>
            </div>
          </CardDescription>
        </CardHeader>
        <CardContent>
          {staff.length === 0 || modules.length === 0 ? (
            <div className="text-center py-8 text-gray-500">
              <Award className="w-12 h-12 mx-auto mb-4 opacity-50" />
              <p>No data to display with current filters</p>
              <p className="text-sm">Try adjusting your filter settings or add some staff/modules first</p>
            </div>
          ) : (
            <div data-testid="training-matrix-table">
              <Accordion type="multiple" defaultValue={['foundation', 'equipment', 'safety-critical']} className="space-y-4">
                {Object.entries(
                  modules.reduce((acc, module) => {
                    const classificationName = classifications.find(c => c.id === module.classificationId)?.name || 'Other';
                    if (!acc[classificationName]) acc[classificationName] = [];
                    acc[classificationName].push(module);
                    return acc;
                  }, {} as Record<string, typeof modules>)
                ).map(([classificationName, classificationModules]) => (
                  <AccordionItem 
                    key={classificationName} 
                    value={classificationName.toLowerCase().replace(/\s+/g, '-')}
                    className="border rounded-lg"
                  >
                    <AccordionTrigger className="px-4 py-3 hover:no-underline">
                      <div className="flex items-center justify-between w-full text-left">
                        <h3 className="text-lg font-semibold">{classificationName}</h3>
                        <div className="flex items-center space-x-2">
                          {(() => {
                            // Check if this classification has any field training modules (non-admin)
                            const hasFieldTraining = classificationModules.some(module => module.audience !== 'administration');
                            
                            if (!hasFieldTraining) {
                              // Admin category - no level badge
                              return null;
                            }
                            
                            const classification = classifications.find(c => c.name === classificationName);
                            const formVariant = classification ? getFormVariant(classification.key) : 'foundation';
                            const category = formVariant === 'safety' ? 'Safety Critical' :
                                          formVariant === 'equipment' ? 'Equipment Level' : 'Foundation Level';
                            
                            return (
                              <Badge 
                                variant={category === 'Safety Critical' ? 'destructive' : category === 'Equipment Level' ? 'default' : 'secondary'} 
                                className={`whitespace-nowrap ${
                                  category === 'Safety Critical' 
                                    ? 'bg-red-100 text-red-800 border-red-200' 
                                    : category === 'Equipment Level' 
                                    ? 'bg-orange-100 text-orange-800 border-orange-200' 
                                    : 'bg-blue-100 text-blue-800 border-blue-200'
                                }`}
                                data-testid={`badge-category-${classification?.key || classificationName.toLowerCase().replace(/\s+/g, '-')}`}
                              >
                                {category}
                              </Badge>
                            );
                          })()}
                          <Badge variant="secondary" className="whitespace-nowrap bg-gray-100 text-gray-800">
                            {classificationModules.length} module{classificationModules.length !== 1 ? 's' : ''}
                          </Badge>
                        </div>
                      </div>
                    </AccordionTrigger>
                    <AccordionContent className="px-4 pb-4">
                      <div className="overflow-auto border rounded-lg">
                        <Table>
                          <TableHeader>
                            <TableRow>
                              <TableHead className="sticky left-0 bg-white border-r min-w-32 sm:min-w-48 z-10 px-2 sm:px-4 text-xs sm:text-sm">
                                Staff Member
                              </TableHead>
                              {classificationModules.map(module => (
                                <TableHead 
                                  key={module.id} 
                                  className="text-center min-w-28 sm:min-w-40 px-1 sm:px-2 py-2 sm:py-3 text-xs sm:text-sm"
                                  title={module.name}
                                  data-testid={`header-module-${module.id}`}
                                >
                                  <div className="text-sm font-medium">
                                    {module.name}
                                  </div>
                                  <div className="text-xs text-gray-500 mt-1">
                                    {module.code}
                                  </div>
                                </TableHead>
                              ))}
                            </TableRow>
                          </TableHeader>
                          <TableBody>
                            {staff.map(staffMember => (
                              <TableRow key={staffMember.id}>
                                <TableCell className="sticky left-0 bg-white border-r z-10 px-2 sm:px-4 py-2 sm:py-4">
                                  <StaffPhotoCell staffId={staffMember.id} name={staffMember.name} jobTitle={staffMember.jobTitle} />
                                </TableCell>
                                {classificationModules.map(module => {
                                  const trainingRecord = processedMatrixData[staffMember.id]?.[module.id];
                                  return (
                                    <TableCell 
                                      key={module.id} 
                                      className="text-center p-1 sm:p-2"
                                      data-testid={`cell-${staffMember.id}-${module.id}`}
                                    >
                                      {trainingRecord ? (
                                        <AbilityBadge 
                                          ableToUse={trainingRecord.ableToUse || false}
                                          competencyLevel={trainingRecord.competencyLevel || ''}
                                          expiryDate={trainingRecord.expiryDate ? String(trainingRecord.expiryDate) : null}
                                          onClick={() => handleCellClick(staffMember, module, trainingRecord)}
                                        />
                                      ) : (
                                        <div 
                                          className="cursor-pointer hover:bg-gray-100 p-2 rounded"
                                          onClick={() => handleCellClick(staffMember, module, null)}
                                          data-testid={`cell-not-trained-${staffMember.id}-${module.id}`}
                                        >
                                          <Badge 
                                            variant="outline" 
                                            className="bg-gray-50 text-gray-500 text-xs"
                                          >
                                            Not Trained
                                          </Badge>
                                        </div>
                                      )}
                                    </TableCell>
                                  );
                                })}
                              </TableRow>
                            ))}
                          </TableBody>
                        </Table>
                      </div>
                    </AccordionContent>
                  </AccordionItem>
                ))}
              </Accordion>
            </div>
          )}
        </CardContent>
      </Card>

      {/* Summary Stats */}
      <Card className="bg-gray-50 border-gray-200">
        <CardContent className="p-4">
          <div className="flex flex-wrap justify-between items-center gap-6 text-sm">
            <div className="flex items-center space-x-2">
              <User className="w-4 h-4 text-blue-600" />
              <span className="font-medium" data-testid="stat-total-staff">{staff.length}</span>
              <span className="text-gray-600">Staff</span>
            </div>
            <div className="flex items-center space-x-2">
              <Award className="w-4 h-4 text-green-600" />
              <span className="font-medium" data-testid="stat-total-modules">{modules.length}</span>
              <span className="text-gray-600">Modules</span>
            </div>
            <div className="flex items-center space-x-2">
              <CheckCircle className="w-4 h-4 text-purple-600" />
              <span className="font-medium" data-testid="stat-total-records">
                {Object.values(processedMatrixData).reduce((total, staffRecords) => 
                  total + Object.values(staffRecords).filter(record => record?.status === 'competent').length, 0
                )}
              </span>
              <span className="text-gray-600">Trained</span>
            </div>
            <div className="flex items-center space-x-2">
              <AlertTriangle className="w-4 h-4 text-red-600" />
              <span className="font-medium" data-testid="stat-expiring-soon">
                {Object.values(processedMatrixData).reduce((total, staffRecords) => 
                  total + Object.values(staffRecords).filter(record => 
                    record?.expiryDate && 
                    new Date(record.expiryDate) < new Date(Date.now() + 90 * 24 * 60 * 60 * 1000) &&
                    new Date(record.expiryDate) > new Date()
                  ).length, 0
                )}
              </span>
              <span className="text-gray-600">Expiring</span>
            </div>
          </div>
        </CardContent>
      </Card>

      {/* Training Record Edit Modal - H&S Policy Compliant */}
      <Dialog open={editModalOpen} onOpenChange={setEditModalOpen}>
        <DialogContent className="sm:max-w-3xl w-[min(768px,100vw-16px)] p-0 max-h-[90vh] overflow-y-auto" data-testid="training-record-modal" aria-describedby="training-record-description">
          <div className="w-full px-3 sm:px-6 py-4 sm:py-6">
            {/* Enhanced Header Context Band */}
          <div className="border-b pb-4 mb-6">
            <div className="flex flex-col sm:flex-row sm:items-start sm:justify-between mb-4 space-y-3 sm:space-y-0">
              <div className="flex items-center space-x-3">
                <div className="w-8 h-8 sm:w-10 sm:h-10 bg-blue-100 rounded-full flex items-center justify-center flex-shrink-0">
                  <User className="w-4 h-4 sm:w-5 sm:h-5 text-blue-600" />
                </div>
                <div className="min-w-0 flex-1">
                  <h3 className="text-base sm:text-lg font-semibold truncate">{selectedRecord?.staff.name}</h3>
                  <p className="text-xs sm:text-sm text-gray-600 truncate">{selectedRecord?.staff.jobTitle || 'Staff Member'}</p>
                </div>
              </div>
              <div className="text-left sm:text-right sm:pr-8">
                <DialogTitle className="text-sm sm:text-base font-medium mb-1">
                  {selectedRecord?.record ? 'Edit Training Record' : 'Add Training Record'}
                </DialogTitle>
                <p id="training-record-description" className="text-xs sm:text-sm text-gray-500">
                  {new Date().toLocaleDateString()}
                </p>
              </div>
            </div>
            
            {/* Module Information Card */}
            {selectedRecord && (() => {
              const classification = classifications.find(c => c.id === selectedRecord.skill.classificationId);
              const formVariant = classification ? getFormVariant(classification.key) : 'foundation';
              const category = selectedRecord.skill.audience === 'administration' ? 'Administration' : 
                             formVariant === 'safety' ? 'Safety Critical' :
                             formVariant === 'equipment' ? 'Equipment Level' : 'Foundation Level';
              const isSafetyCritical = formVariant === 'safety';
              
              return (
                <div className="bg-gray-50 rounded-lg p-3 sm:p-4">
                  <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between mb-3 space-y-2 sm:space-y-0">
                    <div className="flex items-center space-x-3 min-w-0">
                      <Award className="w-4 h-4 sm:w-5 sm:h-5 text-amber-600 flex-shrink-0" />
                      <div className="min-w-0 flex-1">
                        <h4 className="font-medium text-sm sm:text-base truncate">{selectedRecord.skill.name}</h4>
                        <p className="text-xs sm:text-sm text-gray-600">{selectedRecord.skill.code}</p>
                      </div>
                    </div>
                    <div className="flex items-center space-x-2 flex-shrink-0">
                      {(() => {
                        return (
                          <Badge 
                            className={`${
                              category === 'Safety Critical' 
                                ? 'bg-red-100 text-red-800 border-red-200' 
                                : category === 'Equipment Level' 
                                ? 'bg-orange-100 text-orange-800 border-orange-200' 
                                : category === 'Administration'
                                ? 'bg-purple-100 text-purple-800 border-purple-200'
                                : 'bg-blue-100 text-blue-800 border-blue-200'
                            }`}
                          >
                            {isSafetyCritical && <Shield className="w-3 h-3 mr-1" />}
                            {category}
                          </Badge>
                        );
                      })()} 
                    </div>
                  </div>
                  
                  {/* Current Status and Quick Actions */}
                  <div className="flex items-center justify-between">
                    <div className="flex items-center space-x-4">
                      {selectedRecord.record ? (
                        <div className="flex items-center space-x-2">
                          <CheckCircle className="w-4 h-4 text-green-600" />
                          <span className="text-sm font-medium">{selectedRecord.record.competencyLevel}</span>
                          {selectedRecord.record.expiryDate && (
                            <div className="flex items-center space-x-1 text-xs text-gray-500">
                              <Clock className="w-3 h-3" />
                              <span>Expires: {new Date(selectedRecord.record.expiryDate).toLocaleDateString()}</span>
                            </div>
                          )}
                        </div>
                      ) : (
                        <Badge variant="outline" className="text-gray-500">
                          Not Trained
                        </Badge>
                      )}
                    </div>
                    
                    {/* Quick Action Links */}
                    <div className="flex items-center space-x-2">
                      {selectedRecord.skill.sopUrl && (
                        <Button variant="ghost" size="sm" className="h-8 px-2 text-xs">
                          <FileText className="w-3 h-3 mr-1" />
                          SOP
                          <ExternalLink className="w-3 h-3 ml-1" />
                        </Button>
                      )}
                      {selectedRecord.skill.trainingVideoUrl && (
                        <Button variant="ghost" size="sm" className="h-8 px-2 text-xs">
                          <Camera className="w-3 h-3 mr-1" />
                          Video
                          <ExternalLink className="w-3 h-3 ml-1" />
                        </Button>
                      )}
                    </div>
                  </div>
                </div>
              );
            })()}
          </div>

          <Form {...form}>
            <form onSubmit={form.handleSubmit(onSubmit)} className="space-y-8">
              {/* Progressive Disclosure Based on Module Type */}
              {(() => {
                if (!selectedRecord) return null;
                
                const isAdminTraining = selectedRecord.skill.audience === 'administration';
                const classification = classifications.find(c => c.id === selectedRecord.skill.classificationId);
                const formVariant = isAdminTraining ? 'administration' : 
                                  classification ? getFormVariant(classification.key) : 'foundation';
                
                const category = isAdminTraining ? 'Administration' : 
                              formVariant === 'safety' ? 'Safety Critical' :
                              formVariant === 'equipment' ? 'Equipment Level' : 'Foundation Level';
                const isSafetyCritical = formVariant === 'safety';
                const isEquipmentLevel = formVariant === 'equipment';
                const isNonCritical = formVariant === 'foundation';
                
                // For Administration Training: Standardized Layout
                if (isAdminTraining) {
                  return (
                    <div className="space-y-6">
                      <div className="bg-purple-50 p-4 rounded-lg border border-purple-200">
                        <h4 className="font-medium text-purple-800 mb-2">Administration Training Record</h4>
                        <p className="text-sm text-purple-600">Simple competency tracking for administrative skills</p>
                      </div>
                      
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                        <FormField
                          control={form.control}
                          name="competencyLevel"
                          render={({ field }) => (
                            <FormItem>
                              <FormLabel className="text-sm font-medium">Competency Level</FormLabel>
                              <Select onValueChange={field.onChange} value={field.value}>
                                <FormControl>
                                  <SelectTrigger data-testid="select-competency-level" className="h-10">
                                    <SelectValue placeholder="Select competency level" />
                                  </SelectTrigger>
                                </FormControl>
                                <SelectContent>
                                  {Object.values(ADMIN_COMPETENCY_LEVELS).map((level) => (
                                    <SelectItem key={level} value={level}>{level}</SelectItem>
                                  ))}
                                </SelectContent>
                              </Select>
                              <FormMessage />
                            </FormItem>
                          )}
                        />
                        
                        <FormField
                          control={form.control}
                          name="achievedDate"
                          render={({ field }) => (
                            <FormItem>
                              <FormLabel className="text-sm font-medium">Training Completed Date</FormLabel>
                              <FormControl>
                                <Input
                                  type="date"
                                  {...field}
                                  data-testid="input-achieved-date"
                                  className="h-10"
                                />
                              </FormControl>
                              <FormMessage />
                            </FormItem>
                          )}
                        />
                      </div>
                    </div>
                  );
                }
                
                // For Safety Critical: Comprehensive layout
                if (isSafetyCritical) {
                  return (
                    <div className="space-y-6">
                      <div className="bg-red-50 p-3 rounded-lg border border-red-200 flex items-center space-x-2">
                        <Shield className="w-4 h-4 text-red-600 flex-shrink-0" />
                        <p className="text-sm text-red-700 font-medium">Safety Critical — requires annual review</p>
                      </div>
                      
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                        <FormField
                          control={form.control}
                          name="competencyLevel"
                          render={({ field }) => (
                            <FormItem>
                              <FormLabel className="text-sm font-medium">Competency Level</FormLabel>
                              <Select onValueChange={field.onChange} value={field.value}>
                                <FormControl>
                                  <SelectTrigger data-testid="select-competency-level" className="h-10">
                                    <SelectValue placeholder="Select competency level" />
                                  </SelectTrigger>
                                </FormControl>
                                <SelectContent>
                                  {Object.values(FIELD_COMPETENCY_LEVELS).map((level) => (
                                    <SelectItem key={level} value={level}>{level}</SelectItem>
                                  ))}
                                </SelectContent>
                              </Select>
                              <FormMessage />
                            </FormItem>
                          )}
                        />
                        
                        <FormField
                          control={form.control}
                          name="achievedDate"
                          render={({ field }) => (
                            <FormItem>
                              <FormLabel className="text-sm font-medium">Training Date</FormLabel>
                              <FormControl>
                                <Input
                                  type="date"
                                  {...field}
                                  data-testid="input-achieved-date"
                                  className="h-10"
                                />
                              </FormControl>
                              <FormMessage />
                            </FormItem>
                          )}
                        />
                        
                        <FormField
                          control={form.control}
                          name="expiryDate"
                          render={({ field }) => (
                            <FormItem>
                              <FormLabel className="text-sm font-medium">Expiry Date</FormLabel>
                              <FormControl>
                                <Input
                                  type="date"
                                  {...field}
                                  data-testid="input-expiry-date"
                                  className="h-10"
                                />
                              </FormControl>
                              <FormMessage />
                            </FormItem>
                          )}
                        />
                        
                        <FormField
                          control={form.control}
                          name="assessorName"
                          render={({ field }) => (
                            <FormItem>
                              <FormLabel className="text-sm font-medium">Assessor</FormLabel>
                              <FormControl>
                                <Input
                                  {...field}
                                  placeholder="e.g., John Smith"
                                  data-testid="input-assessor-name"
                                  className="h-10"
                                />
                              </FormControl>
                              <FormMessage />
                            </FormItem>
                          )}
                        />
                      </div>
                    </div>
                  );
                }
                
                // For Equipment Level: Simplified layout
                if (isEquipmentLevel) {
                  return (
                    <div className="space-y-6">
                      <div className="bg-orange-50 p-3 rounded-lg border border-orange-200 flex items-center justify-between">
                        <p className="text-sm text-orange-700 font-medium">Equipment training — track competency and expiry</p>
                        <FormField
                          control={form.control}
                          name="ableToUse"
                          render={({ field }) => (
                            <FormItem className="flex flex-row items-center space-x-2 space-y-0">
                              <FormControl>
                                <Checkbox checked={field.value} onCheckedChange={field.onChange} data-testid="checkbox-able-to-use" />
                              </FormControl>
                              <FormLabel className="text-sm font-medium">Able to Use</FormLabel>
                            </FormItem>
                          )}
                        />
                      </div>
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                        <FormField control={form.control} name="competencyLevel" render={({ field }) => (
                          <FormItem>
                            <FormLabel className="text-sm font-medium">Competency Level</FormLabel>
                            <Select onValueChange={field.onChange} value={field.value}>
                              <FormControl><SelectTrigger data-testid="select-competency-level" className="h-10"><SelectValue placeholder="Select level" /></SelectTrigger></FormControl>
                              <SelectContent>{Object.values(FIELD_COMPETENCY_LEVELS).map(level => <SelectItem key={level} value={level}>{level}</SelectItem>)}</SelectContent>
                            </Select>
                            <FormMessage />
                          </FormItem>
                        )} />
                        <FormField control={form.control} name="achievedDate" render={({ field }) => (
                          <FormItem>
                            <FormLabel className="text-sm font-medium">Training Date</FormLabel>
                            <FormControl><Input type="date" {...field} data-testid="input-achieved-date" className="h-10" /></FormControl>
                            <FormMessage />
                          </FormItem>
                        )} />
                        <FormField control={form.control} name="expiryDate" render={({ field }) => (
                          <FormItem>
                            <FormLabel className="text-sm font-medium">Expiry Date</FormLabel>
                            <FormControl><Input type="date" {...field} data-testid="input-expiry-date" className="h-10" /></FormControl>
                            <FormMessage />
                          </FormItem>
                        )} />
                        <FormField control={form.control} name="assessorName" render={({ field }) => (
                          <FormItem>
                            <FormLabel className="text-sm font-medium">Assessor</FormLabel>
                            <FormControl><Input {...field} placeholder="e.g., John Smith" data-testid="input-assessor-name" className="h-10" /></FormControl>
                            <FormMessage />
                          </FormItem>
                        )} />
                      </div>
                    </div>
                  );
                }
                
                // For Foundation Level: Simplified Layout
                if (isNonCritical) {
                  return (
                    <div className="space-y-6">
                      <div className="bg-blue-50 p-3 rounded-lg border border-blue-200 flex items-center justify-between">
                        <p className="text-sm text-blue-700 font-medium">{category} training</p>
                        <FormField control={form.control} name="ableToUse" render={({ field }) => (
                          <FormItem className="flex flex-row items-center space-x-2 space-y-0">
                            <FormControl><Checkbox checked={field.value} onCheckedChange={field.onChange} data-testid="checkbox-able-to-use" /></FormControl>
                            <FormLabel className="text-sm font-medium">Able to Use</FormLabel>
                          </FormItem>
                        )} />
                      </div>
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                        <FormField control={form.control} name="competencyLevel" render={({ field }) => (
                          <FormItem>
                            <FormLabel className="text-sm font-medium">Competency Level</FormLabel>
                            <Select onValueChange={field.onChange} value={field.value}>
                              <FormControl><SelectTrigger data-testid="select-competency-level" className="h-10"><SelectValue placeholder="Select level" /></SelectTrigger></FormControl>
                              <SelectContent>{Object.values(FIELD_COMPETENCY_LEVELS).map(level => <SelectItem key={level} value={level}>{level}</SelectItem>)}</SelectContent>
                            </Select>
                            <FormMessage />
                          </FormItem>
                        )} />
                        <FormField control={form.control} name="achievedDate" render={({ field }) => (
                          <FormItem>
                            <FormLabel className="text-sm font-medium">Training Date</FormLabel>
                            <FormControl><Input type="date" {...field} data-testid="input-achieved-date" className="h-10" /></FormControl>
                            <FormMessage />
                          </FormItem>
                        )} />
                        <FormField control={form.control} name="assessorName" render={({ field }) => (
                          <FormItem>
                            <FormLabel className="text-sm font-medium">Assessor</FormLabel>
                            <FormControl><Input {...field} placeholder="e.g., John Smith" data-testid="input-assessor-name" className="h-10" /></FormControl>
                            <FormMessage />
                          </FormItem>
                        )} />
                      </div>
                    </div>
                  );
                }
                
                return null;
              })()}
              
              
              {/* Notes Section - Always Present */}
              <div className="border-t pt-6">
                <FormField
                  control={form.control}
                  name="notes"
                  render={({ field }) => (
                    <FormItem>
                      <FormLabel className="text-base font-medium">Additional Notes</FormLabel>
                      <FormControl>
                        <Textarea
                          {...field}
                          placeholder="Any additional notes about this training record..."
                          className="resize-none min-h-[100px]"
                          data-testid="textarea-notes"
                        />
                      </FormControl>
                      <p className="text-xs text-gray-500">Optional: Add any relevant comments or observations</p>
                      <FormMessage />
                    </FormItem>
                  )}
                />
              </div>

              
              {/* Sticky Footer with Action Buttons */}
              <div className="sticky bottom-0 bg-white border-t mt-8 p-4 -mx-6 -mb-6 rounded-b-lg">
                <div className="flex items-center justify-end">
                  <div className="flex space-x-3">
                    <Button
                      type="button"
                      variant="outline"
                      onClick={() => {
                        setEditModalOpen(false);
                        setSelectedRecord(null);
                        form.reset();
                      }}
                      data-testid="button-cancel"
                      className="min-w-[100px]"
                    >
                      <X className="w-4 h-4 mr-2" />
                      Cancel
                    </Button>
                    <Button
                      type="submit"
                      disabled={saveTrainingMutation.isPending}
                      data-testid="button-save"
                      className="min-w-[120px]"
                    >
                      {saveTrainingMutation.isPending ? (
                        <>
                          <div className="w-4 h-4 mr-2 animate-spin rounded-full border-2 border-white border-t-transparent" />
                          Saving...
                        </>
                      ) : (
                        <>
                          <Save className="w-4 h-4 mr-2" />
                          Save Record
                        </>
                      )}
                    </Button>
                  </div>
                </div>
              </div>
            </form>
          </Form>
          </div>
        </DialogContent>
      </Dialog>

      {/* Bulk Compliance Update Dialog */}
      <Dialog open={showComplianceDialog} onOpenChange={setShowComplianceDialog}>
        <DialogContent className="sm:max-w-md">
          <DialogHeader>
            <DialogTitle>Bulk Compliance Pass</DialogTitle>
            <DialogDescription>
              Set all field staff to "Competent" status with "Able to Use" on all field training modules.
            </DialogDescription>
          </DialogHeader>
          
          <div className="space-y-4 py-4">
            <div className="space-y-2">
              <Label htmlFor="expiry-date">Expiry Date (Regular Modules)</Label>
              <Input
                id="expiry-date"
                type="date"
                value={complianceExpiryDate}
                onChange={(e) => setComplianceExpiryDate(e.target.value)}
                data-testid="input-compliance-expiry-date"
              />
              <p className="text-sm text-gray-500">
                Regular modules will expire on this date. <strong className="text-red-600">Safety-critical modules will automatically expire in December {new Date().getMonth() >= 11 ? new Date().getFullYear() + 1 : new Date().getFullYear()}</strong> for review before year-end.
              </p>
            </div>
            
            <div className="bg-yellow-50 border border-yellow-200 rounded-md p-3">
              <p className="text-sm text-yellow-800">
                <strong>Warning:</strong> This will update all field staff training records. This action cannot be undone.
              </p>
            </div>
          </div>

          <DialogFooter>
            <Button
              type="button"
              variant="outline"
              onClick={() => setShowComplianceDialog(false)}
              data-testid="button-cancel-compliance"
            >
              Cancel
            </Button>
            <Button
              type="button"
              onClick={handleBulkComplianceUpdate}
              disabled={bulkUpdateMutation.isPending}
              data-testid="button-confirm-compliance"
            >
              {bulkUpdateMutation.isPending ? 'Processing...' : 'Update All Records'}
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </div>
  );
}
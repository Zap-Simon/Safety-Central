import { useState } from "react";
import { useQuery, useMutation } from "@tanstack/react-query";
import { useForm } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import { z } from "zod";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Skeleton } from "@/components/ui/skeleton";
import { Button } from "@/components/ui/button";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { Form, FormControl, FormField, FormItem, FormLabel, FormMessage } from "@/components/ui/form";
import { Input } from "@/components/ui/input";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Textarea } from "@/components/ui/textarea";
import { useToast } from "@/hooks/use-toast";
import { Shield, ExternalLink, ClipboardCheck, ChevronDown, ChevronUp } from "lucide-react";
import { FIELD_COMPETENCY_LEVELS } from "@shared/schema";
import { apiRequest, queryClient } from "@/lib/queryClient";

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
  audience: string;
};

type StaffMember = {
  id: number;
  name: string;
  jobTitle: string;
  isFieldStaff: boolean;
  isActive: boolean;
};

const recordSchema = z.object({
  staffId: z.string().min(1, "Select a staff member"),
  competencyLevel: z.string().min(1, "Select a competency level"),
  achievedDate: z.string().min(1, "Enter the training date"),
  expiryDate: z.string().optional(),
  assessorName: z.string().optional(),
  notes: z.string().optional(),
});

type RecordFormData = z.infer<typeof recordSchema>;

export default function FieldTrainingTab() {
  const { toast } = useToast();
  const [selectedModule, setSelectedModule] = useState<TrainingModule | null>(null);
  const [expandedGroups, setExpandedGroups] = useState<Record<number, boolean>>({});

  const { data: classificationsData, isLoading: classLoading } = useQuery({
    queryKey: ['/api/training-classifications?audience=field'],
  });

  const { data: modulesData, isLoading: modulesLoading } = useQuery({
    queryKey: ['/api/training-modules'],
  });

  const { data: staffData, isLoading: staffLoading } = useQuery({
    queryKey: ['/api/staff'],
  });

  const form = useForm<RecordFormData>({
    resolver: zodResolver(recordSchema),
    defaultValues: {
      staffId: "",
      competencyLevel: "",
      achievedDate: new Date().toISOString().split("T")[0],
      expiryDate: "",
      assessorName: "",
      notes: "",
    },
  });

  const saveMutation = useMutation({
    mutationFn: async (data: RecordFormData) => {
      if (!selectedModule) throw new Error("No module selected");
      const payload = {
        staffId: parseInt(data.staffId),
        skillId: selectedModule.id,
        competencyLevel: data.competencyLevel,
        achievedDate: new Date(data.achievedDate),
        expiryDate: data.expiryDate ? new Date(data.expiryDate) : null,
        assessorName: data.assessorName || null,
        notes: data.notes || null,
        ableToUse: data.competencyLevel === FIELD_COMPETENCY_LEVELS.COMPETENT_SOP_MODULE || data.competencyLevel === FIELD_COMPETENCY_LEVELS.EXPERT,
        status: "Active",
      };
      const res = await apiRequest("POST", "/api/training-records", payload);
      if (!res.ok) throw new Error("Failed to save");
      return res.json();
    },
    onSuccess: () => {
      toast({ title: "Training recorded", description: `${selectedModule?.name} record saved successfully.` });
      queryClient.invalidateQueries({ queryKey: ["/api/training-module-matrix"] });
      queryClient.invalidateQueries({ queryKey: ["/api/training-records"] });
      setSelectedModule(null);
      form.reset();
    },
    onError: () => {
      toast({ title: "Failed to save", variant: "destructive" });
    },
  });

  const onSubmit = (data: RecordFormData) => saveMutation.mutate(data);

  if (classLoading || modulesLoading || staffLoading) {
    return (
      <div className="space-y-4">
        {Array.from({ length: 4 }).map((_, i) => <Skeleton key={i} className="h-32 w-full" />)}
      </div>
    );
  }

  const classifications: any[] = (classificationsData as any)?.data || [];
  const allModules: TrainingModule[] = (modulesData as any)?.data || [];
  const fieldModules = allModules.filter(m => m.isActive && (m.audience === "field" || m.audience === "both"));
  const fieldStaff: StaffMember[] = ((staffData as any)?.data || []).filter((s: StaffMember) => s.isFieldStaff && s.isActive);

  const toggleGroup = (id: number) => setExpandedGroups(prev => ({ ...prev, [id]: !prev[id] }));

  const levelBadgeColor = (level: string) => {
    if (level === "Safety Critical") return "bg-red-100 text-red-800 border-red-200";
    if (level === "Equipment Level") return "bg-orange-100 text-orange-800 border-orange-200";
    return "bg-blue-100 text-blue-800 border-blue-200";
  };

  const modulesByClass = classifications.map(cls => ({
    ...cls,
    modules: fieldModules.filter(m => m.classificationId === cls.id).sort((a, b) => {
      if (a.isSafetyCritical && !b.isSafetyCritical) return -1;
      if (!a.isSafetyCritical && b.isSafetyCritical) return 1;
      return a.displayOrder - b.displayOrder;
    }),
  })).filter(cls => cls.modules.length > 0);

  const getCategoryLabel = (key: string) => {
    if (["lifting-handling-equipment", "safety-equipment-ppe", "working-at-heights"].includes(key)) return "Safety Critical";
    if (["power-tools-equipment", "vehicles-transport"].includes(key)) return "Equipment Level";
    return "Foundation Level";
  };

  return (
    <div className="space-y-4">
      <div className="mb-4">
        <h2 className="text-xl font-bold text-gray-900">Training Topics</h2>
        <p className="text-sm text-gray-600 mt-1">Click any module to record a training completion. Link your SOP documents from the module settings.</p>
      </div>

      {modulesByClass.map(cls => {
        const category = getCategoryLabel(cls.key);
        const isOpen = expandedGroups[cls.id] !== false;
        return (
          <Card key={cls.id} className="overflow-hidden">
            <button
              className="w-full text-left px-4 py-3 flex items-center justify-between hover:bg-gray-50 transition-colors"
              onClick={() => toggleGroup(cls.id)}
            >
              <div className="flex items-center gap-3">
                <span className="font-semibold text-gray-900">{cls.name}</span>
                <Badge className={`text-xs ${levelBadgeColor(category)}`}>{category}</Badge>
                <span className="text-xs text-gray-500">{cls.modules.length} module{cls.modules.length !== 1 ? "s" : ""}</span>
              </div>
              {isOpen ? <ChevronUp className="w-4 h-4 text-gray-400" /> : <ChevronDown className="w-4 h-4 text-gray-400" />}
            </button>

            {isOpen && (
              <CardContent className="pt-0 pb-3">
                <div className="space-y-2">
                  {cls.modules.map((mod: TrainingModule) => (
                    <div
                      key={mod.id}
                      className={`flex items-center justify-between p-3 rounded-lg border ${mod.isSafetyCritical ? "border-red-200 bg-red-50" : "border-gray-200 bg-gray-50"}`}
                    >
                      <div className="flex items-center gap-2 min-w-0 flex-1">
                        {mod.isSafetyCritical && <Shield className="w-3.5 h-3.5 text-red-600 flex-shrink-0" />}
                        <span className="text-xs font-mono text-gray-500 flex-shrink-0">{mod.code}</span>
                        <span className="text-sm font-medium text-gray-900 truncate">{mod.name}</span>
                        {mod.requiresCertification && (
                          <Badge className="text-xs bg-amber-100 text-amber-800 border-amber-200 flex-shrink-0">Cert Required</Badge>
                        )}
                      </div>
                      <div className="flex items-center gap-2 flex-shrink-0 ml-3">
                        {mod.sopUrl && (
                          <a
                            href={mod.sopUrl}
                            target="_blank"
                            rel="noreferrer"
                            className="inline-flex items-center gap-1 text-xs text-blue-600 hover:text-blue-800 font-medium"
                          >
                            SOP <ExternalLink className="w-3 h-3" />
                          </a>
                        )}
                        <Button
                          size="sm"
                          variant="outline"
                          className="h-7 text-xs px-2"
                          onClick={() => {
                            setSelectedModule(mod);
                            form.reset({
                              staffId: "",
                              competencyLevel: "",
                              achievedDate: new Date().toISOString().split("T")[0],
                              expiryDate: "",
                              assessorName: "",
                              notes: "",
                            });
                          }}
                        >
                          <ClipboardCheck className="w-3 h-3 mr-1" />
                          Record
                        </Button>
                      </div>
                    </div>
                  ))}
                </div>
              </CardContent>
            )}
          </Card>
        );
      })}

      <Dialog open={!!selectedModule} onOpenChange={open => { if (!open) { setSelectedModule(null); form.reset(); } }}>
        <DialogContent className="sm:max-w-md">
          <DialogHeader>
            <DialogTitle className="flex items-center gap-2">
              {selectedModule?.isSafetyCritical && <Shield className="w-4 h-4 text-red-600" />}
              Record Training
            </DialogTitle>
            <div className="text-sm text-gray-600">
              <span className="font-medium">{selectedModule?.name}</span>
              <span className="text-gray-400 ml-2">{selectedModule?.code}</span>
            </div>
            {selectedModule?.sopUrl && (
              <a
                href={selectedModule.sopUrl}
                target="_blank"
                rel="noreferrer"
                className="inline-flex items-center gap-1 text-sm text-blue-600 hover:text-blue-800 font-medium w-fit"
              >
                Open SOP <ExternalLink className="w-3 h-3" />
              </a>
            )}
          </DialogHeader>

          <Form {...form}>
            <form onSubmit={form.handleSubmit(onSubmit)} className="space-y-4 pt-2">
              <FormField control={form.control} name="staffId" render={({ field }) => (
                <FormItem>
                  <FormLabel>Staff Member</FormLabel>
                  <Select onValueChange={field.onChange} value={field.value}>
                    <FormControl>
                      <SelectTrigger><SelectValue placeholder="Select staff member" /></SelectTrigger>
                    </FormControl>
                    <SelectContent>
                      {fieldStaff.map(s => (
                        <SelectItem key={s.id} value={String(s.id)}>{s.name}</SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                  <FormMessage />
                </FormItem>
              )} />

              <FormField control={form.control} name="competencyLevel" render={({ field }) => (
                <FormItem>
                  <FormLabel>Competency Level</FormLabel>
                  <Select onValueChange={field.onChange} value={field.value}>
                    <FormControl>
                      <SelectTrigger><SelectValue placeholder="Select level" /></SelectTrigger>
                    </FormControl>
                    <SelectContent>
                      {Object.values(FIELD_COMPETENCY_LEVELS).map(level => (
                        <SelectItem key={level} value={level}>{level}</SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                  <FormMessage />
                </FormItem>
              )} />

              <div className="grid grid-cols-2 gap-3">
                <FormField control={form.control} name="achievedDate" render={({ field }) => (
                  <FormItem>
                    <FormLabel>Training Date</FormLabel>
                    <FormControl><Input type="date" {...field} /></FormControl>
                    <FormMessage />
                  </FormItem>
                )} />

                <FormField control={form.control} name="expiryDate" render={({ field }) => (
                  <FormItem>
                    <FormLabel>Expiry Date</FormLabel>
                    <FormControl><Input type="date" {...field} /></FormControl>
                    <FormMessage />
                  </FormItem>
                )} />
              </div>

              <FormField control={form.control} name="assessorName" render={({ field }) => (
                <FormItem>
                  <FormLabel>Assessor</FormLabel>
                  <FormControl><Input {...field} placeholder="Who assessed this?" /></FormControl>
                  <FormMessage />
                </FormItem>
              )} />

              <FormField control={form.control} name="notes" render={({ field }) => (
                <FormItem>
                  <FormLabel>Notes <span className="text-gray-400 font-normal">(optional)</span></FormLabel>
                  <FormControl><Textarea {...field} placeholder="Any relevant notes..." className="resize-none min-h-[60px]" /></FormControl>
                  <FormMessage />
                </FormItem>
              )} />

              <div className="flex justify-end gap-2 pt-2">
                <Button type="button" variant="outline" onClick={() => { setSelectedModule(null); form.reset(); }}>
                  Cancel
                </Button>
                <Button type="submit" disabled={saveMutation.isPending}>
                  {saveMutation.isPending ? "Saving..." : "Save Record"}
                </Button>
              </div>
            </form>
          </Form>
        </DialogContent>
      </Dialog>
    </div>
  );
}

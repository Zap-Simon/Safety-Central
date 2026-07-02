import { useMemo, useState } from "react";
import { Link } from "wouter";
import { useQuery, useMutation } from "@tanstack/react-query";
import { queryClient } from "@/lib/queryClient";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import { Card, CardContent } from "@/components/ui/card";
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogFooter } from "@/components/ui/dialog";
import { Input } from "@/components/ui/input";
import { Textarea } from "@/components/ui/textarea";
import { Label } from "@/components/ui/label";
import { useToast } from "@/hooks/use-toast";
import {
  ArrowLeft, Search, Plus, Pencil, ShieldCheck, Archive, ArchiveRestore, Loader2, TriangleAlert,
} from "lucide-react";
import type { Hazard } from "@shared/schema";

interface HazardsResponse {
  success: boolean;
  data: Hazard[];
}

const RISK_BADGE: Record<string, string> = {
  H: "bg-red-100 text-red-800 border-red-300",
  M: "bg-yellow-100 text-yellow-800 border-yellow-300",
  L: "bg-green-100 text-green-800 border-green-300",
};

const RISK_LABEL: Record<string, string> = { H: "High", M: "Medium", L: "Low" };

function riskBadge(value: string | undefined, prefix: string) {
  if (!value) return null;
  const key = value.trim().charAt(0).toUpperCase();
  return (
    <Badge variant="outline" className={`text-[10px] px-1.5 py-0 border ${RISK_BADGE[key] || "bg-gray-100 text-gray-600 border-gray-300"}`}>
      {prefix}: {RISK_LABEL[key] || value}
    </Badge>
  );
}

interface HazardFormState {
  id?: number;
  hazardId?: string;
  category: string;
  riskHarm: string;
  initialRisk: string;
  controlType: string;
  controls: string;
  residualRisk: string;
  reviewTrigger: string;
  training: string;
}

const EMPTY_FORM: HazardFormState = {
  category: "",
  riskHarm: "",
  initialRisk: "",
  controlType: "",
  controls: "",
  residualRisk: "",
  reviewTrigger: "",
  training: "",
};

export default function HazardRegister() {
  const { toast } = useToast();
  const [search, setSearch] = useState("");
  const [categoryFilter, setCategoryFilter] = useState<string>("all");
  const [showArchived, setShowArchived] = useState(false);
  const [formOpen, setFormOpen] = useState(false);
  const [form, setForm] = useState<HazardFormState>(EMPTY_FORM);
  const isEditing = form.id !== undefined;

  const { data: resp, isLoading } = useQuery<HazardsResponse>({
    queryKey: ["/api/hazards"],
  });
  const all = resp?.data || [];

  const saveMutation = useMutation({
    mutationFn: async (payload: HazardFormState) => {
      const isUpdate = payload.id !== undefined;
      const { id, hazardId, ...fields } = payload;
      const res = await fetch(isUpdate ? `/api/hazards/${id}` : "/api/hazards", {
        method: isUpdate ? "PUT" : "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(
          isUpdate
            ? fields
            : { ...fields, source: "register", identified: new Date().toISOString().slice(0, 10) },
        ),
      });
      const body = await res.json().catch(() => null);
      if (!res.ok || !body?.success) {
        throw new Error(body?.error || "Couldn't save the hazard.");
      }
      return body.data as Hazard;
    },
    onSuccess: (saved) => {
      queryClient.invalidateQueries({ queryKey: ["/api/hazards"] });
      setFormOpen(false);
      toast({
        title: isEditing ? "Hazard updated" : `Hazard ${saved.hazardId} added`,
        description: isEditing ? undefined : "It's now available to link in Near Miss investigations.",
      });
    },
    onError: (err: Error) => {
      toast({ title: "Save failed", description: err.message, variant: "destructive" });
    },
  });

  const archiveMutation = useMutation({
    mutationFn: async ({ id, archived }: { id: number; archived: boolean }) => {
      const res = await fetch(`/api/hazards/${id}`, {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ archived }),
      });
      const body = await res.json().catch(() => null);
      if (!res.ok || !body?.success) throw new Error(body?.error || "Couldn't update the hazard.");
      return body.data as Hazard;
    },
    onSuccess: (saved) => {
      queryClient.invalidateQueries({ queryKey: ["/api/hazards"] });
      toast({ title: saved.archived ? `Hazard ${saved.hazardId} archived` : `Hazard ${saved.hazardId} restored` });
    },
    onError: (err: Error) => {
      toast({ title: "Update failed", description: err.message, variant: "destructive" });
    },
  });

  const categories = useMemo(() => {
    const seen: string[] = [];
    for (const h of all) {
      if (!seen.includes(h.category)) seen.push(h.category);
    }
    return seen;
  }, [all]);

  const q = search.trim().toLowerCase();
  const visible = useMemo(() => {
    return all.filter((h) => {
      if (!showArchived && h.archived) return false;
      if (categoryFilter !== "all" && h.category !== categoryFilter) return false;
      if (!q) return true;
      return (
        h.riskHarm.toLowerCase().includes(q) ||
        h.hazardId.toLowerCase().includes(q) ||
        h.category.toLowerCase().includes(q) ||
        h.controls.toLowerCase().includes(q) ||
        h.training.toLowerCase().includes(q)
      );
    });
  }, [all, q, categoryFilter, showArchived]);

  const grouped = useMemo(() => {
    const map = new Map<string, Hazard[]>();
    for (const h of visible) {
      if (!map.has(h.category)) map.set(h.category, []);
      map.get(h.category)!.push(h);
    }
    return Array.from(map.entries());
  }, [visible]);

  const activeCount = all.filter((h) => !h.archived).length;
  const highCount = all.filter((h) => !h.archived && h.initialRisk.trim().toUpperCase().startsWith("H")).length;
  const fromInvestigations = all.filter((h) => !h.archived && h.source === "investigation").length;

  const openAdd = () => {
    setForm(EMPTY_FORM);
    setFormOpen(true);
  };

  const openEdit = (h: Hazard) => {
    setForm({
      id: h.id,
      hazardId: h.hazardId,
      category: h.category,
      riskHarm: h.riskHarm,
      initialRisk: h.initialRisk,
      controlType: h.controlType,
      controls: h.controls,
      residualRisk: h.residualRisk,
      reviewTrigger: h.reviewTrigger,
      training: h.training,
    });
    setFormOpen(true);
  };

  const submitForm = () => {
    if (!form.riskHarm.trim() || !form.category.trim()) {
      toast({
        title: "Missing details",
        description: "A hazard needs at least a description and a category.",
        variant: "destructive",
      });
      return;
    }
    saveMutation.mutate(form);
  };

  const set = (field: keyof HazardFormState) => (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) =>
    setForm((f) => ({ ...f, [field]: e.target.value }));

  return (
    <div className="min-h-screen bg-gray-50">
      <div className="max-w-5xl mx-auto px-4 py-6">
        {/* Header */}
        <div className="flex items-center gap-3 mb-1">
          <Link href="/">
            <Button variant="ghost" size="sm" className="text-gray-500 -ml-2">
              <ArrowLeft className="h-4 w-4 mr-1" /> Home
            </Button>
          </Link>
        </div>
        <div className="flex flex-wrap items-start justify-between gap-3 mb-4">
          <div>
            <h1 className="text-xl font-bold text-gray-900 flex items-center gap-2">
              <ShieldCheck className="h-5 w-5 text-orange-600" />
              Operational Hazard Register
            </h1>
            <p className="text-sm text-gray-500 mt-0.5">
              The live register of workplace hazards and their controls. Near Miss investigations link back to these entries.
            </p>
          </div>
          <Button onClick={openAdd} className="bg-orange-600 hover:bg-orange-700 text-white">
            <Plus className="h-4 w-4 mr-1" /> Add Hazard
          </Button>
        </div>

        {/* Stats */}
        <div className="grid grid-cols-3 gap-3 mb-4">
          <Card><CardContent className="p-3 text-center">
            <div className="text-xl font-bold text-gray-900">{activeCount}</div>
            <div className="text-[11px] text-gray-500">Active hazards</div>
          </CardContent></Card>
          <Card><CardContent className="p-3 text-center">
            <div className="text-xl font-bold text-red-600">{highCount}</div>
            <div className="text-[11px] text-gray-500">High initial risk</div>
          </CardContent></Card>
          <Card><CardContent className="p-3 text-center">
            <div className="text-xl font-bold text-blue-600">{fromInvestigations}</div>
            <div className="text-[11px] text-gray-500">Added from investigations</div>
          </CardContent></Card>
        </div>

        {/* Filters */}
        <div className="flex flex-wrap items-center gap-2 mb-4">
          <div className="relative flex-1 min-w-[200px]">
            <Search className="absolute left-2.5 top-1/2 -translate-y-1/2 h-4 w-4 text-gray-400" />
            <Input
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="Search hazards, IDs, controls..."
              className="pl-8 h-9 text-sm bg-white"
            />
          </div>
          <select
            value={categoryFilter}
            onChange={(e) => setCategoryFilter(e.target.value)}
            className="h-9 border border-gray-200 rounded-md px-2 text-sm bg-white focus:outline-none focus:border-orange-400"
          >
            <option value="all">All categories</option>
            {categories.map((c) => (
              <option key={c} value={c}>{c}</option>
            ))}
          </select>
          <Button
            variant={showArchived ? "secondary" : "outline"}
            size="sm"
            className="h-9"
            onClick={() => setShowArchived((v) => !v)}
          >
            <Archive className="h-3.5 w-3.5 mr-1" />
            {showArchived ? "Hiding nothing" : "Show archived"}
          </Button>
        </div>

        {/* Register list */}
        {isLoading ? (
          <div className="flex items-center justify-center py-16 text-gray-400 gap-2">
            <Loader2 className="h-5 w-5 animate-spin" /> Loading hazard register...
          </div>
        ) : grouped.length === 0 ? (
          <div className="py-16 text-center text-gray-400 italic">No hazards match the current filters.</div>
        ) : (
          grouped.map(([cat, list]) => (
            <div key={cat} className="mb-5">
              <h2 className="text-xs font-semibold text-gray-500 uppercase tracking-wide mb-2">
                {cat} <span className="text-gray-300 font-normal">({list.length})</span>
              </h2>
              <div className="space-y-2">
                {list.map((h) => (
                  <Card key={h.id} className={h.archived ? "opacity-60" : ""}>
                    <CardContent className="p-3">
                      <div className="flex items-start justify-between gap-2">
                        <div className="min-w-0">
                          <div className="flex flex-wrap items-center gap-1.5">
                            <span className="text-[11px] font-mono text-blue-700 bg-blue-50 border border-blue-200 rounded px-1.5">
                              {h.hazardId}
                            </span>
                            {riskBadge(h.initialRisk, "Initial")}
                            {riskBadge(h.residualRisk, "Residual")}
                            {h.source === "investigation" && (
                              <Badge variant="outline" className="text-[10px] px-1.5 py-0 border-purple-300 bg-purple-50 text-purple-700">
                                From investigation
                              </Badge>
                            )}
                            {h.archived && (
                              <Badge variant="outline" className="text-[10px] px-1.5 py-0 border-gray-300 bg-gray-100 text-gray-500">
                                Archived
                              </Badge>
                            )}
                          </div>
                          <div className="text-sm font-medium text-gray-900 mt-1">{h.riskHarm}</div>
                          {h.controls && (
                            <div className="text-xs text-gray-500 mt-1 whitespace-pre-wrap">
                              <span className="font-semibold text-gray-600">Controls{h.controlType ? ` (${h.controlType})` : ""}:</span> {h.controls}
                            </div>
                          )}
                          {h.training && (
                            <div className="text-[11px] text-gray-400 mt-1">Training: {h.training}</div>
                          )}
                        </div>
                        <div className="flex flex-col items-end gap-1 shrink-0">
                          <Button variant="ghost" size="sm" className="h-7 px-2 text-gray-400 hover:text-gray-700" onClick={() => openEdit(h)}>
                            <Pencil className="h-3.5 w-3.5" />
                          </Button>
                          <Button
                            variant="ghost"
                            size="sm"
                            className="h-7 px-2 text-gray-400 hover:text-gray-700"
                            disabled={archiveMutation.isPending}
                            onClick={() => archiveMutation.mutate({ id: h.id, archived: !h.archived })}
                            title={h.archived ? "Restore hazard" : "Archive hazard"}
                          >
                            {h.archived ? <ArchiveRestore className="h-3.5 w-3.5" /> : <Archive className="h-3.5 w-3.5" />}
                          </Button>
                        </div>
                      </div>
                    </CardContent>
                  </Card>
                ))}
              </div>
            </div>
          ))
        )}
      </div>

      {/* Add / edit dialog */}
      <Dialog open={formOpen} onOpenChange={setFormOpen}>
        <DialogContent className="max-w-lg max-h-[85vh] overflow-y-auto">
          <DialogHeader>
            <DialogTitle className="text-base">
              {isEditing ? (
                <span className="flex items-center gap-2">
                  Edit hazard
                  <span className="text-[11px] font-mono text-blue-700 bg-blue-50 border border-blue-200 rounded px-1.5">{form.hazardId}</span>
                </span>
              ) : "Add a hazard to the register"}
            </DialogTitle>
          </DialogHeader>
          <div className="space-y-3">
            {!isEditing && (
              <div className="flex items-start gap-2 text-xs text-gray-500 bg-blue-50 border border-blue-100 rounded p-2">
                <TriangleAlert className="h-4 w-4 text-blue-500 shrink-0 mt-0.5" />
                The hazard ID (e.g. CG-HZ-003C) is assigned automatically based on the category.
              </div>
            )}
            <div>
              <Label className="text-xs">Hazard / risk of harm *</Label>
              <Textarea value={form.riskHarm} onChange={set("riskHarm")} rows={2} className="mt-1 text-sm" placeholder="What is the hazard and what harm could it cause?" />
            </div>
            <div>
              <Label className="text-xs">Category *</Label>
              <Input list="register-categories" value={form.category} onChange={set("category")} className="mt-1 text-sm" placeholder="Pick existing or type a new category" />
              <datalist id="register-categories">
                {categories.map((c) => <option key={c} value={c} />)}
              </datalist>
            </div>
            <div className="grid grid-cols-2 gap-3">
              <div>
                <Label className="text-xs">Initial risk</Label>
                <select value={form.initialRisk} onChange={(e) => setForm((f) => ({ ...f, initialRisk: e.target.value }))} className="mt-1 w-full h-9 border border-gray-200 rounded-md px-2 text-sm focus:outline-none focus:border-orange-400">
                  <option value="">Not set</option>
                  <option value="L">Low</option>
                  <option value="M">Medium</option>
                  <option value="H">High</option>
                </select>
              </div>
              <div>
                <Label className="text-xs">Residual risk (after controls)</Label>
                <select value={form.residualRisk} onChange={(e) => setForm((f) => ({ ...f, residualRisk: e.target.value }))} className="mt-1 w-full h-9 border border-gray-200 rounded-md px-2 text-sm focus:outline-none focus:border-orange-400">
                  <option value="">Not set</option>
                  <option value="L">Low</option>
                  <option value="M">Medium</option>
                  <option value="H">High</option>
                </select>
              </div>
            </div>
            <div>
              <Label className="text-xs">Control type</Label>
              <select value={form.controlType} onChange={(e) => setForm((f) => ({ ...f, controlType: e.target.value }))} className="mt-1 w-full h-9 border border-gray-200 rounded-md px-2 text-sm focus:outline-none focus:border-orange-400">
                <option value="">Not set</option>
                <option value="Elimination">Elimination</option>
                <option value="Isolation">Isolation</option>
                <option value="Minimisation">Minimisation</option>
              </select>
            </div>
            <div>
              <Label className="text-xs">Controls / how we manage it</Label>
              <Textarea value={form.controls} onChange={set("controls")} rows={3} className="mt-1 text-sm" placeholder="The controls in place for this hazard..." />
            </div>
            <div>
              <Label className="text-xs">Primary training / competency</Label>
              <Input value={form.training} onChange={set("training")} className="mt-1 text-sm" placeholder="e.g. Manual handling training" />
            </div>
            <div>
              <Label className="text-xs">Review trigger</Label>
              <Input value={form.reviewTrigger} onChange={set("reviewTrigger")} className="mt-1 text-sm" placeholder="e.g. Annual + after incident" />
            </div>
          </div>
          <DialogFooter className="mt-2">
            <Button variant="outline" onClick={() => setFormOpen(false)} disabled={saveMutation.isPending}>Cancel</Button>
            <Button onClick={submitForm} disabled={saveMutation.isPending} className="bg-orange-600 hover:bg-orange-700 text-white">
              {saveMutation.isPending && <Loader2 className="h-4 w-4 mr-1 animate-spin" />}
              {isEditing ? "Save changes" : "Add to register"}
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </div>
  );
}

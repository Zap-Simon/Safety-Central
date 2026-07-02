import { useMemo, useState } from "react";
import { useQuery, useQueryClient } from "@tanstack/react-query";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { Plus, Trash2, Search, ShieldCheck, AlertTriangle, ArrowLeft, Loader2 } from "lucide-react";
import { LIKELIHOODS_LIST, CONSEQUENCES_LIST, getRiskLevelForCell } from "./riskUtils";
import { authService } from "@/auth/authService";
import type { Hazard } from "@shared/schema";

// A hazard row inside a near-miss investigation. Extended additively — legacy
// investigations only carry the first six fields (free-text hazard + control),
// and those rows must keep rendering and saving exactly as before.
export interface HazardRow {
  id: string;
  hazard: string;              // hazard name (registered: the register's risk/harm text)
  likelihood: string;
  consequence: string;
  risk: string;
  control: string;             // legacy free-text control (old rows only)
  hazardRefId?: string;        // register ID, e.g. "CG-HZ-001A"
  category?: string;
  registeredControls?: string; // snapshot of the register's controls at link time
  controlFailure?: string;     // why the existing controls failed / fell short
  correctiveActions?: string;  // corrective / additional controls required
  unregistered?: boolean;      // free-typed hazard the user chose not to register
}

const RISK_COLORS: Record<string, string> = {
  "Low":      "bg-green-100 text-green-800 border-green-300",
  "Moderate": "bg-yellow-100 text-yellow-800 border-yellow-300",
  "High":     "bg-red-100 text-red-800 border-red-300",
  "Extreme":  "bg-gray-900 text-white border-gray-700",
};

interface HazardTableProps {
  rows: HazardRow[];
  onChange: (rows: HazardRow[]) => void;
  readOnly?: boolean;
}

interface HazardsResponse {
  success: boolean;
  data: Hazard[];
}

export default function HazardTable({ rows, onChange, readOnly = false }: HazardTableProps) {
  const queryClient = useQueryClient();
  const [pickerOpen, setPickerOpen] = useState(false);
  const [search, setSearch] = useState("");
  // "pick" = browsing the register, "new" = describing a hazard not in it
  const [mode, setMode] = useState<"pick" | "new">("pick");
  const [newDesc, setNewDesc] = useState("");
  const [newCategory, setNewCategory] = useState("");
  const [newControls, setNewControls] = useState("");
  const [isAdding, setIsAdding] = useState(false);
  const [addError, setAddError] = useState<string | null>(null);

  const { data: hazardsResp, isLoading: hazardsLoading } = useQuery<HazardsResponse>({
    queryKey: ["/api/hazards"],
    staleTime: 60_000,
  });
  const register = useMemo(
    () => (hazardsResp?.data || []).filter((h) => !h.archived),
    [hazardsResp],
  );

  const categories = useMemo(() => {
    const seen = new Map<string, Hazard[]>();
    for (const h of register) {
      if (!seen.has(h.category)) seen.set(h.category, []);
      seen.get(h.category)!.push(h);
    }
    return Array.from(seen.entries());
  }, [register]);

  const q = search.trim().toLowerCase();
  const filteredCategories = useMemo(() => {
    if (!q) return categories;
    return categories
      .map(([cat, list]) => [
        cat,
        list.filter(
          (h) =>
            h.riskHarm.toLowerCase().includes(q) ||
            h.hazardId.toLowerCase().includes(q) ||
            cat.toLowerCase().includes(q) ||
            h.controls.toLowerCase().includes(q),
        ),
      ] as [string, Hazard[]])
      .filter(([, list]) => list.length > 0);
  }, [categories, q]);

  const alreadyLinked = useMemo(
    () => new Set(rows.map((r) => r.hazardRefId).filter(Boolean)),
    [rows],
  );

  const openPicker = () => {
    setSearch("");
    setMode("pick");
    setNewDesc("");
    setNewCategory("");
    setNewControls("");
    setAddError(null);
    setPickerOpen(true);
  };

  const addRegisteredHazard = (h: Hazard) => {
    const newRow: HazardRow = {
      id: `h-${Date.now()}`,
      hazard: h.riskHarm,
      likelihood: "",
      consequence: "",
      risk: "",
      control: "",
      hazardRefId: h.hazardId,
      category: h.category,
      registeredControls: h.controls,
      controlFailure: "",
      correctiveActions: "",
    };
    onChange([...rows, newRow]);
    setPickerOpen(false);
  };

  const startNewHazard = () => {
    setMode("new");
    setNewDesc(search.trim());
    setAddError(null);
  };

  const addNewHazard = async (registerIt: boolean) => {
    const desc = newDesc.trim();
    if (!desc) {
      setAddError("Describe the hazard first.");
      return;
    }
    if (registerIt && !newCategory.trim()) {
      setAddError("Pick or type a category so the hazard files correctly in the register.");
      return;
    }
    setAddError(null);

    if (!registerIt) {
      onChange([
        ...rows,
        {
          id: `h-${Date.now()}`,
          hazard: desc,
          likelihood: "",
          consequence: "",
          risk: "",
          control: "",
          category: newCategory.trim() || undefined,
          controlFailure: "",
          correctiveActions: "",
          unregistered: true,
        },
      ]);
      setPickerOpen(false);
      return;
    }

    setIsAdding(true);
    try {
      const token = await authService.getAccessToken();
      const res = await fetch("/api/hazards", {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}` },
        body: JSON.stringify({
          category: newCategory.trim(),
          riskHarm: desc,
          controls: newControls.trim(),
          source: "investigation",
          identified: new Date().toISOString().slice(0, 10),
        }),
      });
      const body = await res.json().catch(() => null);
      if (!res.ok || !body?.success) {
        setAddError(body?.error || "Couldn't add the hazard to the register. Try again.");
        return;
      }
      const created: Hazard = body.data;
      queryClient.invalidateQueries({ queryKey: ["/api/hazards"] });
      onChange([
        ...rows,
        {
          id: `h-${Date.now()}`,
          hazard: created.riskHarm,
          likelihood: "",
          consequence: "",
          risk: "",
          control: "",
          hazardRefId: created.hazardId,
          category: created.category,
          registeredControls: created.controls,
          controlFailure: "",
          correctiveActions: "",
        },
      ]);
      setPickerOpen(false);
    } catch {
      setAddError("Couldn't reach the server. Check your connection and try again.");
    } finally {
      setIsAdding(false);
    }
  };

  const updateRow = (id: string, field: keyof HazardRow, value: string) => {
    const updated = rows.map((r) => {
      if (r.id !== id) return r;
      const updatedRow = { ...r, [field]: value };
      if (field === "likelihood" || field === "consequence") {
        const likelihood = field === "likelihood" ? value : r.likelihood;
        const consequence = field === "consequence" ? value : r.consequence;
        if (likelihood && consequence) {
          updatedRow.risk = getRiskLevelForCell(likelihood, consequence);
        }
      }
      return updatedRow;
    });
    onChange(updated);
  };

  const removeRow = (id: string) => {
    onChange(rows.filter((r) => r.id !== id));
  };

  // A row is "legacy" if it predates the register link-up: no register reference
  // and never flagged as deliberately unregistered. Those keep the original
  // free-text hazard + control inputs so nothing old breaks.
  const isLegacy = (r: HazardRow) => !r.hazardRefId && !r.unregistered;

  return (
    <div className="space-y-3">
      {rows.length === 0 && (
        <div className="rounded-lg border border-dashed border-gray-300 p-4 text-center text-gray-400 italic text-sm">
          No hazards linked yet. Click "Add Hazard" to pick from the Hazard Register.
        </div>
      )}

      {rows.map((row) => (
        <div key={row.id} className="rounded-lg border border-gray-200 bg-white">
          {/* Header: ID badge + name + risk + remove */}
          <div className="flex items-start gap-2 p-2.5 pb-2">
            <div className="flex-1 min-w-0">
              <div className="flex flex-wrap items-center gap-1.5">
                {row.hazardRefId ? (
                  <Badge variant="outline" className="text-[10px] font-mono border-blue-300 bg-blue-50 text-blue-700 px-1.5 py-0">
                    {row.hazardRefId}
                  </Badge>
                ) : row.unregistered ? (
                  <Badge variant="outline" className="text-[10px] border-amber-300 bg-amber-50 text-amber-700 px-1.5 py-0">
                    <AlertTriangle className="h-2.5 w-2.5 mr-0.5" />
                    Not in register
                  </Badge>
                ) : null}
                {row.category && (
                  <span className="text-[10px] text-gray-400">{row.category}</span>
                )}
                {row.risk && (
                  <Badge className={`text-[10px] font-bold px-1.5 py-0 border ${RISK_COLORS[row.risk] || "bg-gray-100 text-gray-700"}`}>
                    {row.risk.toUpperCase()}
                  </Badge>
                )}
              </div>
              {isLegacy(row) ? (
                <input
                  type="text"
                  value={row.hazard}
                  disabled={readOnly}
                  onChange={(e) => updateRow(row.id, "hazard", e.target.value)}
                  placeholder="Describe hazard..."
                  className="mt-1 w-full border border-gray-200 rounded px-2 py-1 text-xs focus:outline-none focus:border-orange-400"
                />
              ) : (
                <div className="mt-0.5 text-xs font-medium text-gray-800">{row.hazard}</div>
              )}
            </div>
            {!readOnly && (
              <button
                onClick={() => removeRow(row.id)}
                className="text-gray-300 hover:text-red-500 transition-colors shrink-0 mt-0.5"
                title="Remove hazard"
              >
                <Trash2 className="h-3.5 w-3.5" />
              </button>
            )}
          </div>

          {/* Registered controls snapshot (read-only) */}
          {row.registeredControls && (
            <div className="mx-2.5 mb-2 rounded bg-gray-50 border border-gray-100 px-2 py-1.5">
              <div className="flex items-center gap-1 text-[10px] font-semibold text-gray-500 uppercase tracking-wide">
                <ShieldCheck className="h-3 w-3 text-green-600" />
                Registered controls
              </div>
              <div className="text-[11px] text-gray-600 whitespace-pre-wrap mt-0.5">{row.registeredControls}</div>
            </div>
          )}

          {/* Likelihood / consequence */}
          <div className="grid grid-cols-2 gap-2 px-2.5 pb-2">
            <select
              value={row.likelihood}
              disabled={readOnly}
              onChange={(e) => updateRow(row.id, "likelihood", e.target.value)}
              className="w-full border border-gray-200 rounded px-1 py-1 text-xs focus:outline-none focus:border-orange-400"
            >
              <option value="">Likelihood...</option>
              {LIKELIHOODS_LIST.map((l) => (
                <option key={l} value={l}>{l}</option>
              ))}
            </select>
            <select
              value={row.consequence}
              disabled={readOnly}
              onChange={(e) => updateRow(row.id, "consequence", e.target.value)}
              className="w-full border border-gray-200 rounded px-1 py-1 text-xs focus:outline-none focus:border-orange-400"
            >
              <option value="">Consequence...</option>
              {CONSEQUENCES_LIST.map((c) => (
                <option key={c} value={c}>{c}</option>
              ))}
            </select>
          </div>

          {/* Control fields */}
          <div className="px-2.5 pb-2.5 space-y-1.5">
            {isLegacy(row) ? (
              <input
                type="text"
                value={row.control}
                disabled={readOnly}
                onChange={(e) => updateRow(row.id, "control", e.target.value)}
                placeholder="Control measure..."
                className="w-full border border-gray-200 rounded px-2 py-1 text-xs focus:outline-none focus:border-orange-400"
              />
            ) : (
              <>
                {row.hazardRefId && (
                  <textarea
                    value={row.controlFailure || ""}
                    disabled={readOnly}
                    onChange={(e) => updateRow(row.id, "controlFailure", e.target.value)}
                    placeholder="Why did the existing controls fail or fall short?"
                    rows={2}
                    className="w-full border border-gray-200 rounded px-2 py-1 text-xs focus:outline-none focus:border-orange-400 resize-y"
                  />
                )}
                <textarea
                  value={row.correctiveActions || ""}
                  disabled={readOnly}
                  onChange={(e) => updateRow(row.id, "correctiveActions", e.target.value)}
                  placeholder="Corrective / additional controls required..."
                  rows={2}
                  className="w-full border border-gray-200 rounded px-2 py-1 text-xs focus:outline-none focus:border-orange-400 resize-y"
                />
              </>
            )}
          </div>
        </div>
      ))}

      {!readOnly && (
        <Button
          type="button"
          variant="outline"
          onClick={openPicker}
          className="w-full border-dashed border-orange-300 text-orange-600 hover:bg-orange-50 text-xs h-8"
        >
          <Plus className="h-3.5 w-3.5 mr-1" />
          Add Hazard
        </Button>
      )}

      {/* Register picker dialog */}
      <Dialog open={pickerOpen} onOpenChange={setPickerOpen}>
        <DialogContent className="max-w-lg max-h-[85vh] flex flex-col p-0 gap-0">
          <DialogHeader className="px-4 pt-4 pb-2">
            <DialogTitle className="text-sm flex items-center gap-2">
              {mode === "new" && (
                <button
                  onClick={() => setMode("pick")}
                  className="text-gray-400 hover:text-gray-600"
                  title="Back to register"
                >
                  <ArrowLeft className="h-4 w-4" />
                </button>
              )}
              {mode === "pick" ? "Pick a hazard from the register" : "Add a new hazard"}
            </DialogTitle>
          </DialogHeader>

          {mode === "pick" ? (
            <>
              <div className="px-4 pb-2">
                <div className="relative">
                  <Search className="absolute left-2 top-1/2 -translate-y-1/2 h-3.5 w-3.5 text-gray-400" />
                  <input
                    autoFocus
                    type="text"
                    value={search}
                    onChange={(e) => setSearch(e.target.value)}
                    placeholder="Search hazards, categories or IDs..."
                    className="w-full border border-gray-200 rounded pl-7 pr-2 py-1.5 text-xs focus:outline-none focus:border-orange-400"
                  />
                </div>
              </div>
              <div className="flex-1 overflow-y-auto px-4 pb-2 min-h-[200px]">
                {hazardsLoading ? (
                  <div className="flex items-center justify-center py-8 text-gray-400 text-xs gap-2">
                    <Loader2 className="h-4 w-4 animate-spin" /> Loading hazard register...
                  </div>
                ) : filteredCategories.length === 0 ? (
                  <div className="py-8 text-center text-gray-400 text-xs italic">
                    No matching hazards in the register.
                  </div>
                ) : (
                  filteredCategories.map(([cat, list]) => (
                    <div key={cat} className="mb-2">
                      <div className="text-[10px] font-semibold text-gray-500 uppercase tracking-wide sticky top-0 bg-white py-1">
                        {cat}
                      </div>
                      {list.map((h) => {
                        const linked = alreadyLinked.has(h.hazardId);
                        return (
                          <button
                            key={h.hazardId}
                            disabled={linked}
                            onClick={() => addRegisteredHazard(h)}
                            className={`w-full text-left rounded border px-2 py-1.5 mb-1 transition-colors ${
                              linked
                                ? "border-gray-100 bg-gray-50 opacity-50 cursor-not-allowed"
                                : "border-gray-200 hover:border-orange-300 hover:bg-orange-50"
                            }`}
                          >
                            <div className="flex items-center gap-1.5">
                              <span className="text-[10px] font-mono text-blue-700 bg-blue-50 border border-blue-200 rounded px-1">
                                {h.hazardId}
                              </span>
                              {h.initialRisk && (
                                <span className={`text-[9px] font-bold rounded px-1 ${h.initialRisk === "H" ? "bg-red-100 text-red-700" : h.initialRisk === "M" ? "bg-yellow-100 text-yellow-700" : "bg-green-100 text-green-700"}`}>
                                  {h.initialRisk === "H" ? "HIGH" : h.initialRisk === "M" ? "MED" : "LOW"}
                                </span>
                              )}
                              {linked && <span className="text-[9px] text-gray-400">already linked</span>}
                            </div>
                            <div className="text-xs text-gray-800 mt-0.5">{h.riskHarm}</div>
                            {h.controls && (
                              <div className="text-[10px] text-gray-400 truncate mt-0.5">{h.controls}</div>
                            )}
                          </button>
                        );
                      })}
                    </div>
                  ))
                )}
              </div>
              <div className="border-t border-gray-100 p-3">
                <Button
                  type="button"
                  variant="outline"
                  onClick={startNewHazard}
                  className="w-full border-dashed border-amber-300 text-amber-700 hover:bg-amber-50 text-xs h-8"
                >
                  <Plus className="h-3.5 w-3.5 mr-1" />
                  Hazard not in the register? Add a new one
                </Button>
              </div>
            </>
          ) : (
            <div className="px-4 pb-4 space-y-2.5 overflow-y-auto">
              <div>
                <label className="text-[10px] font-semibold text-gray-500 uppercase tracking-wide">Hazard description</label>
                <textarea
                  autoFocus
                  value={newDesc}
                  onChange={(e) => setNewDesc(e.target.value)}
                  placeholder="What is the hazard and what harm could it cause?"
                  rows={2}
                  className="mt-1 w-full border border-gray-200 rounded px-2 py-1.5 text-xs focus:outline-none focus:border-orange-400 resize-y"
                />
              </div>
              <div>
                <label className="text-[10px] font-semibold text-gray-500 uppercase tracking-wide">Category</label>
                <input
                  type="text"
                  list="hazard-categories"
                  value={newCategory}
                  onChange={(e) => setNewCategory(e.target.value)}
                  placeholder="Pick an existing category or type a new one..."
                  className="mt-1 w-full border border-gray-200 rounded px-2 py-1.5 text-xs focus:outline-none focus:border-orange-400"
                />
                <datalist id="hazard-categories">
                  {categories.map(([cat]) => (
                    <option key={cat} value={cat} />
                  ))}
                </datalist>
              </div>
              <div>
                <label className="text-[10px] font-semibold text-gray-500 uppercase tracking-wide">Current controls (if any)</label>
                <textarea
                  value={newControls}
                  onChange={(e) => setNewControls(e.target.value)}
                  placeholder="Any controls already in place for this hazard..."
                  rows={2}
                  className="mt-1 w-full border border-gray-200 rounded px-2 py-1.5 text-xs focus:outline-none focus:border-orange-400 resize-y"
                />
              </div>
              {addError && (
                <div className="text-xs text-red-600 bg-red-50 border border-red-200 rounded px-2 py-1.5">{addError}</div>
              )}
              <div className="space-y-1.5 pt-1">
                <Button
                  type="button"
                  disabled={isAdding}
                  onClick={() => addNewHazard(true)}
                  className="w-full bg-orange-600 hover:bg-orange-700 text-white text-xs h-8"
                >
                  {isAdding ? <Loader2 className="h-3.5 w-3.5 mr-1 animate-spin" /> : <ShieldCheck className="h-3.5 w-3.5 mr-1" />}
                  Add to Hazard Register &amp; use in this investigation
                </Button>
                <Button
                  type="button"
                  variant="outline"
                  disabled={isAdding}
                  onClick={() => addNewHazard(false)}
                  className="w-full text-gray-500 text-xs h-8"
                >
                  Use without adding to the register
                </Button>
                <p className="text-[10px] text-gray-400 text-center">
                  New hazards should normally be added to the register so the whole team benefits.
                </p>
              </div>
            </div>
          )}
        </DialogContent>
      </Dialog>
    </div>
  );
}

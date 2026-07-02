import { useState, useRef, useEffect } from "react";
import { useQuery, useQueryClient } from "@tanstack/react-query";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import { AlertTriangle, Bot, Loader2, CheckCircle, FileText, PenLine, RotateCcw, X, ClipboardList } from "lucide-react";
import RiskMatrix from "./RiskMatrix";
import HazardTable, { type HazardRow } from "./HazardTable";
import { getRiskLevelForCell } from "./riskUtils";
import { authService } from "@/auth/authService";
import { useToast } from "@/hooks/use-toast";
import { format } from "date-fns";
import { getDateGroupKey } from "@shared/dateUtils";

interface NearMissItem {
  id: string;
  title: string;
  description: string;
  secondaryDescription?: string;
  submittedBy: string;
  submittedDate?: string;
  meetingDate: string;
  meetingNotes?: string;
  actionNotes?: string;
  ideaType?: string;
}

interface ResultingAction {
  id: string;
  description: string;
  assignedTo: string;
  completed: boolean;
}

interface ProgressNote {
  id: number;
  nearMissItemId: string;
  content: string;
  author: string | null;
  createdAt: string;
}

interface InvestigationData {
  id?: number;
  nearMissItemId: string;
  itemTitle: string;
  meetingDate: string;
  investigatorName: string;
  siteJob: string;
  eventDate: string;
  eventTime: string;
  eventType: string;
  involvedPersons: string;
  witnesses: string;
  eventDescription: string;
  contributingFactors: string;
  hazards: HazardRow[];
  likelihood: string;
  consequence: string;
  riskLevel: string;
  treatmentGiven: string;
  resultingActions: ResultingAction[];
  investigatorSignature: string | null;
  investigatorSignedAt: string;
  directorName: string;
  directorSignature: string | null;
  signedAt: string;
  status: string;
}

const SECTION_LABELS = ["Event Details", "Involved Persons & Description", "Risk Assessment", "Resulting Actions & Sign-Off"];

const EVENT_TYPES = [
  "Hazard Report",
  "Near Miss",
  "Minor Harm Accident",
  "Harm Accident",
  "Serious Harm Injury",
  "Property Damage",
];

interface Props {
  item: NearMissItem;
  open: boolean;
  onClose: () => void;
}

export default function NearMissInvestigationModal({ item, open, onClose }: Props) {
  const queryClient = useQueryClient();
  const { toast } = useToast();
  const [section, setSection] = useState(0);
  const [isSaving, setIsSaving] = useState(false);
  const [isDrafting, setIsDrafting] = useState<string | null>(null);
  const [isExporting, setIsExporting] = useState(false);
  const [signingRole, setSigningRole] = useState<"investigator" | "approver" | null>(null);
  const [hasDrawn, setHasDrawn] = useState(false);
  const [saveMsg, setSaveMsg] = useState<string | null>(null);
  // Hidden escape hatch: while Ctrl is held a "Close without a report" button
  // appears so an old near miss can be archived without an investigation.
  const [ctrlHeld, setCtrlHeld] = useState(false);
  const [confirmingClose, setConfirmingClose] = useState(false);
  const [isClosingNoReport, setIsClosingNoReport] = useState(false);

  const canvasRef = useRef<HTMLCanvasElement>(null);
  const isDrawingRef = useRef(false);
  const lastPointRef = useRef<{ x: number; y: number } | null>(null);

  // Pre-fill the Date of Event from when the near miss was submitted (the best
  // available estimate of when it happened). Uses the same UTC date key as the
  // rest of the app so it matches the displayed "Submitted" date exactly.
  const submittedKey = getDateGroupKey(item.submittedDate || null);
  const prefilledEventDate = /^\d{4}-\d{2}-\d{2}$/.test(submittedKey) ? submittedKey : "";

  const empty: InvestigationData = {
    nearMissItemId: item.id,
    itemTitle: item.title || "",
    meetingDate: item.meetingDate || "",
    investigatorName: "Simon Hubbard",
    siteJob: "",
    eventDate: prefilledEventDate,
    eventTime: "",
    eventType: item.ideaType || "Near Miss",
    involvedPersons: item.submittedBy || "",
    witnesses: "",
    eventDescription: item.description || "",
    contributingFactors: "",
    hazards: [],
    likelihood: "",
    consequence: "",
    riskLevel: "",
    treatmentGiven: "No treatment required",
    resultingActions: [],
    investigatorSignature: null,
    investigatorSignedAt: "",
    directorName: "Hoani Hunt",
    directorSignature: null,
    signedAt: "",
    status: "Draft",
  };

  const [data, setData] = useState<InvestigationData>(empty);
  const [loaded, setLoaded] = useState(false);

  // These near miss investigation endpoints operate on the app's own database
  // only and are open to any signed-in user, so no Microsoft token is required.
  const authenticatedFetch = async (url: string, options: RequestInit = {}) => {
    return fetch(url, {
      ...options,
      headers: { ...options.headers, "Content-Type": "application/json" },
    });
  };

  const { isLoading, data: existingData } = useQuery({
    queryKey: ["/api/near-miss-investigations", item.id],
    queryFn: async () => {
      const res = await authenticatedFetch(`/api/near-miss-investigations/${encodeURIComponent(item.id)}`);
      if (!res.ok) return null;
      const json = await res.json();
      return json.data || null;
    },
    enabled: open && !loaded,
  });

  useEffect(() => {
    if (!loaded && existingData !== undefined) {
      if (existingData) {
        setData({
          ...existingData,
          // Older drafts saved before the event date was pre-filled may have it
          // blank — fill it from the submitted date so it never starts empty.
          eventDate: existingData.eventDate || prefilledEventDate,
          hazards: typeof existingData.hazards === "string" ? JSON.parse(existingData.hazards) : existingData.hazards || [],
          resultingActions: typeof existingData.resultingActions === "string" ? JSON.parse(existingData.resultingActions) : existingData.resultingActions || [],
        });
        // Persist the backfilled date straight away so saved investigations
        // (and their exports) pick it up without needing a manual re-save.
        // Completed investigations are immutable server-side, so skip those.
        if (!existingData.eventDate && prefilledEventDate && existingData.id && existingData.status !== "Complete") {
          authenticatedFetch(`/api/near-miss-investigations/${existingData.id}`, {
            method: "PUT",
            body: JSON.stringify({ eventDate: prefilledEventDate }),
          }).then(res => {
            if (res.ok) {
              queryClient.invalidateQueries({ queryKey: ["/api/near-miss-investigations", item.id] });
            }
          }).catch(e => console.error("Event date backfill failed:", e));
        }
      }
      setLoaded(true);
    }
  }, [existingData, loaded]);

  // Track whether the Ctrl (or ⌘) key is held to reveal the hidden close button.
  // Reset on key-up, window blur and when the modal closes so it can never get
  // stuck visible (e.g. if focus leaves the tab while Ctrl is down).
  useEffect(() => {
    if (!open) {
      setCtrlHeld(false);
      return;
    }
    const onKeyDown = (e: KeyboardEvent) => { if (e.ctrlKey || e.metaKey) setCtrlHeld(true); };
    const onKeyUp = (e: KeyboardEvent) => { if (!e.ctrlKey && !e.metaKey) setCtrlHeld(false); };
    const reset = () => setCtrlHeld(false);
    window.addEventListener("keydown", onKeyDown);
    window.addEventListener("keyup", onKeyUp);
    window.addEventListener("blur", reset);
    return () => {
      window.removeEventListener("keydown", onKeyDown);
      window.removeEventListener("keyup", onKeyUp);
      window.removeEventListener("blur", reset);
    };
  }, [open]);

  const closeWithoutReport = async () => {
    setIsClosingNoReport(true);
    try {
      const closerName = authService.getCurrentUser()?.name || data.investigatorName || undefined;
      const res = await authenticatedFetch(
        `/api/near-miss-investigations/${encodeURIComponent(item.id)}/close-without-report`,
        { method: "POST", body: JSON.stringify({ name: closerName }) },
      );
      if (res.ok) {
        // Refresh the Actions workload + meeting history so the card immediately
        // leaves the live list and shows up under the Archived (Near Miss) view.
        queryClient.invalidateQueries({ queryKey: ["/api/meeting-history"] });
        queryClient.invalidateQueries({ queryKey: ["/api/near-miss-investigations", item.id] });
        toast({ title: "Near miss closed", description: "Archived without a report. It remains on the register." });
        setConfirmingClose(false);
        onClose();
      } else {
        const err = await res.json().catch(() => ({}));
        toast({ title: "Close failed", description: err.error || "Could not close this near miss.", variant: "destructive" });
      }
    } catch (e) {
      console.error("Close without report failed:", e);
      toast({ title: "Close failed", description: "Something went wrong closing this near miss.", variant: "destructive" });
    } finally {
      setIsClosingNoReport(false);
    }
  };

  const [newNote, setNewNote] = useState("");
  const [isAddingNote, setIsAddingNote] = useState(false);

  const { data: progressNotes = [] } = useQuery<ProgressNote[]>({
    queryKey: ["/api/near-miss-investigations", item.id, "notes"],
    queryFn: async () => {
      const res = await authenticatedFetch(`/api/near-miss-investigations/${encodeURIComponent(item.id)}/notes`);
      if (!res.ok) return [];
      const json = await res.json();
      return json.data || [];
    },
    enabled: open,
  });

  const addNote = async () => {
    const content = newNote.trim();
    if (!content) return;
    setIsAddingNote(true);
    try {
      const res = await authenticatedFetch(`/api/near-miss-investigations/${encodeURIComponent(item.id)}/notes`, {
        method: "POST",
        body: JSON.stringify({ content, author: data.investigatorName || undefined }),
      });
      if (res.ok) {
        setNewNote("");
        queryClient.invalidateQueries({ queryKey: ["/api/near-miss-investigations", item.id, "notes"] });
      }
    } catch (e) {
      console.error("Add note failed:", e);
    } finally {
      setIsAddingNote(false);
    }
  };

  const update = (field: keyof InvestigationData, value: any) => {
    setData(prev => {
      const next = { ...prev, [field]: value };
      if (field === "likelihood" || field === "consequence") {
        const l = field === "likelihood" ? value : prev.likelihood;
        const c = field === "consequence" ? value : prev.consequence;
        if (l && c) next.riskLevel = getRiskLevelForCell(l, c);
      }
      return next;
    });
  };

  const saveToDb = async (overrides?: Partial<InvestigationData>) => {
    setIsSaving(true);
    try {
      // Saving a not-yet-complete investigation moves it into a resumable
      // "In Progress" state so its status reflects that work is underway.
      const nextStatus = data.status === "Complete" ? "Complete" : "In Progress";
      const payload = {
        ...data,
        status: nextStatus,
        ...overrides,
        hazards: JSON.stringify(data.hazards),
        resultingActions: JSON.stringify(data.resultingActions),
      };
      if (overrides?.hazards) payload.hazards = JSON.stringify(overrides.hazards);
      if (overrides?.resultingActions) payload.resultingActions = JSON.stringify(overrides.resultingActions);

      const method = data.id ? "PUT" : "POST";
      const url = data.id ? `/api/near-miss-investigations/${data.id}` : "/api/near-miss-investigations";
      const res = await authenticatedFetch(url, { method, body: JSON.stringify(payload) });
      if (res.ok) {
        const json = await res.json();
        setData(prev => ({ ...prev, id: json.data?.id ?? prev.id, status: nextStatus }));
        setSaveMsg("Saved");
        setTimeout(() => setSaveMsg(null), 2000);
        queryClient.invalidateQueries({ queryKey: ["/api/near-miss-investigations", item.id] });
        queryClient.invalidateQueries({ queryKey: ["/api/meeting-history"] });
      }
    } catch (e) {
      console.error("Save failed:", e);
    } finally {
      setIsSaving(false);
    }
  };

  const draftSection = async (sectionKey: "description" | "contributing" | "actions") => {
    setIsDrafting(sectionKey);
    try {
      const res = await authenticatedFetch("/api/ai-near-miss-draft", {
        method: "POST",
        body: JSON.stringify({
          section: sectionKey,
          context: {
            title: item.title,
            description: item.description,
            secondaryDescription: item.secondaryDescription,
            meetingNotes: item.meetingNotes,
            actionNotes: item.actionNotes,
            eventType: data.eventType,
            siteJob: data.siteJob,
            involvedPersons: data.involvedPersons,
            existing: sectionKey === "description" ? data.eventDescription
              : sectionKey === "contributing" ? data.contributingFactors
              : data.resultingActions.map(a => a.description).join("; "),
          },
        }),
      });
      if (res.ok) {
        const json = await res.json();
        if (sectionKey === "description") update("eventDescription", json.draft);
        else if (sectionKey === "contributing") update("contributingFactors", json.draft);
        else if (sectionKey === "actions") {
          const actions = json.draft
            .split("\n")
            .filter((l: string) => l.trim().startsWith("•") || l.trim().startsWith("-"))
            .map((l: string) => ({
              id: `a-${Date.now()}-${Math.random()}`,
              description: l.replace(/^[•\-]\s*/, "").trim(),
              assignedTo: "",
              completed: false,
            }))
            .filter((a: ResultingAction) => a.description);
          if (actions.length > 0) update("resultingActions", actions);
        }
      }
    } finally {
      setIsDrafting(null);
    }
  };

  const signRole = async (role: "investigator" | "approver") => {
    const canvas = canvasRef.current;
    if (!canvas || !hasDrawn) return;
    const sig = canvas.toDataURL("image/png");
    const when = new Date().toISOString();

    setIsSaving(true);
    try {
      let invId = data.id;
      if (!invId) {
        // Create the record first if it hasn't been saved yet
        const payload = {
          ...data,
          hazards: JSON.stringify(data.hazards),
          resultingActions: JSON.stringify(data.resultingActions),
        };
        const createRes = await authenticatedFetch("/api/near-miss-investigations", {
          method: "POST",
          body: JSON.stringify(payload),
        });
        if (createRes.ok) {
          const json = await createRes.json();
          invId = json.data?.id;
        }
      }

      if (invId) {
        const name = role === "investigator" ? data.investigatorName : data.directorName;
        const completeRes = await authenticatedFetch(`/api/near-miss-investigations/${invId}/complete`, {
          method: "POST",
          body: JSON.stringify({ role, name, signature: sig, signedAt: when }),
        });
        if (completeRes.ok) {
          const json = await completeRes.json().catch(() => ({}));
          const nowComplete = !!json.complete;
          setData(prev => ({
            ...prev,
            id: invId,
            ...(role === "investigator"
              ? { investigatorSignature: sig, investigatorSignedAt: when }
              : { directorSignature: sig, signedAt: when }),
            status: nowComplete ? "Complete" : "In Progress",
          }));
          setSigningRole(null);
          setHasDrawn(false);
          setSaveMsg(nowComplete ? "Signed & Complete — Action ready to close" : "Signature saved");
          setTimeout(() => setSaveMsg(null), 3500);
          queryClient.invalidateQueries({ queryKey: ["/api/near-miss-investigations", item.id] });
          queryClient.invalidateQueries({ queryKey: ["/api/meeting-history"] });
        } else {
          const err = await completeRes.json().catch(() => ({}));
          toast({ title: "Sign-off failed", description: err.error || "Could not save the signature.", variant: "destructive" });
        }
      }
    } catch (e) {
      console.error("Sign-off failed:", e);
      toast({ title: "Sign-off failed", description: "Something went wrong saving the signature.", variant: "destructive" });
    } finally {
      setIsSaving(false);
    }
  };

  const renderSignatureBlock = (
    role: "investigator" | "approver",
    title: string,
    nameField: "investigatorName" | "directorName",
    signatureValue: string | null,
    signedAtValue: string,
  ) => {
    const isSigning = signingRole === role;
    const nameValue = data[nameField];
    const roleLabel = role === "investigator" ? "Investigator" : "Approver / Manager";
    return (
      <div className="bg-gray-50 rounded-lg border border-gray-200 p-4 space-y-3">
        <h4 className="font-semibold text-gray-700 text-sm flex items-center gap-2">
          <PenLine className="h-4 w-4" /> {title}
        </h4>
        <div className="grid grid-cols-2 gap-3">
          <div>
            <label className="label-sm">{roleLabel} Name</label>
            <input className="input-sm" value={nameValue} onChange={e => update(nameField, e.target.value)} disabled={!!signatureValue} />
          </div>
          <div>
            <label className="label-sm">Date Signed</label>
            <input className="input-sm" value={signedAtValue ? signedAtValue.substring(0, 10) : ""} readOnly />
          </div>
        </div>

        {signatureValue ? (
          <div className="space-y-2">
            <p className="text-xs text-green-600 flex items-center gap-1"><CheckCircle className="h-3.5 w-3.5" /> Signed off by {nameValue}</p>
            <img src={signatureValue} alt={`${roleLabel} signature`} className="max-h-16 border-b-2 border-gray-400" />
            {!isComplete && (
              <Button variant="outline" size="sm" onClick={() => {
                if (role === "investigator") { update("investigatorSignature", null); update("investigatorSignedAt", ""); }
                else { update("directorSignature", null); update("signedAt", ""); }
              }}>
                Clear Signature
              </Button>
            )}
          </div>
        ) : isSigning ? (
          <div className="space-y-2">
            <div className="border-2 border-dashed border-blue-300 rounded-xl overflow-hidden bg-white touch-none relative">
              <canvas
                ref={canvasRef} width={480} height={160}
                className="w-full cursor-crosshair block"
                style={{ touchAction: "none", height: "160px" }}
                onMouseDown={handleMouseDown} onMouseMove={handleMouseMove}
                onMouseUp={handleMouseUp} onMouseLeave={handleMouseUp}
                onTouchStart={handleTouchStart} onTouchMove={handleTouchMove}
                onTouchEnd={e => { e.preventDefault(); isDrawingRef.current = false; }}
              />
              {!hasDrawn && (
                <div className="absolute inset-0 flex items-center justify-center pointer-events-none">
                  <span className="text-gray-300 text-sm">{roleLabel} signs here</span>
                </div>
              )}
            </div>
            <div className="flex gap-2">
              <Button variant="outline" onClick={clearCanvas} disabled={!hasDrawn} className="flex-1 text-xs">
                <RotateCcw className="h-3.5 w-3.5 mr-1" /> Clear
              </Button>
              <Button onClick={() => signRole(role)} disabled={!hasDrawn || isSaving} className="flex-[2] bg-green-600 hover:bg-green-700 text-white text-xs">
                <CheckCircle className="h-3.5 w-3.5 mr-1" /> Confirm Signature
              </Button>
              <Button variant="ghost" onClick={() => { setSigningRole(null); setHasDrawn(false); }} className="text-xs">Cancel</Button>
            </div>
          </div>
        ) : (
          <Button
            onClick={() => { setSigningRole(role); setHasDrawn(false); }}
            disabled={signingRole !== null}
            variant="outline"
            className="border-blue-300 text-blue-700 hover:bg-blue-50 text-sm"
          >
            <PenLine className="h-4 w-4 mr-2" /> {roleLabel} Sign-Off
          </Button>
        )}
      </div>
    );
  };

  const exportReport = async () => {
    // Open the tab synchronously inside the click handler. If we call window.open
    // AFTER the await, browsers treat it as a non-user-initiated popup and block
    // it silently — which is why the button previously appeared to "do nothing".
    const reportWindow = window.open("", "_blank");
    if (reportWindow) {
      reportWindow.document.write(
        '<!doctype html><title>Generating report…</title><body style="font-family:system-ui,sans-serif;padding:2rem;color:#374151">Generating report…</body>'
      );
    }
    setIsExporting(true);
    try {
      const payload = {
        ...data,
        hazards: data.hazards,
        resultingActions: data.resultingActions,
      };
      const res = await authenticatedFetch("/api/generate-near-miss-report", {
        method: "POST",
        body: JSON.stringify(payload),
      });
      if (!res.ok) {
        reportWindow?.close();
        toast({
          title: "Export failed",
          description: "Could not generate the report. Please try again.",
          variant: "destructive",
        });
        return;
      }
      const json = await res.json();
      if (!json?.htmlContent) {
        reportWindow?.close();
        toast({
          title: "Export failed",
          description: "The report came back empty. Please try again.",
          variant: "destructive",
        });
        return;
      }
      if (reportWindow) {
        reportWindow.document.open();
        reportWindow.document.write(json.htmlContent);
        reportWindow.document.close();
      } else {
        // Popup was blocked — fall back to downloading the report as a file.
        const blob = new Blob([json.htmlContent], { type: "text/html;charset=utf-8;" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = `near-miss-report-${item.id}.html`;
        a.click();
        setTimeout(() => URL.revokeObjectURL(url), 60000);
      }
    } catch (err) {
      reportWindow?.close();
      toast({
        title: "Export failed",
        description: "Could not generate the report. Please try again.",
        variant: "destructive",
      });
    } finally {
      setIsExporting(false);
    }
  };

  const getCanvasPos = (clientX: number, clientY: number, canvas: HTMLCanvasElement) => {
    const rect = canvas.getBoundingClientRect();
    return { x: (clientX - rect.left) * (canvas.width / rect.width), y: (clientY - rect.top) * (canvas.height / rect.height) };
  };

  const drawLine = (pt: { x: number; y: number }) => {
    const canvas = canvasRef.current;
    if (!canvas || !lastPointRef.current) return;
    const ctx = canvas.getContext("2d")!;
    ctx.beginPath(); ctx.moveTo(lastPointRef.current.x, lastPointRef.current.y);
    ctx.lineTo(pt.x, pt.y); ctx.strokeStyle = "#1f2937"; ctx.lineWidth = 2.5;
    ctx.lineCap = "round"; ctx.lineJoin = "round"; ctx.stroke();
    setHasDrawn(true);
    lastPointRef.current = pt;
  };

  const handleMouseDown = (e: React.MouseEvent<HTMLCanvasElement>) => {
    isDrawingRef.current = true;
    lastPointRef.current = getCanvasPos(e.clientX, e.clientY, canvasRef.current!);
  };

  const handleMouseMove = (e: React.MouseEvent<HTMLCanvasElement>) => {
    if (!isDrawingRef.current) return;
    drawLine(getCanvasPos(e.clientX, e.clientY, canvasRef.current!));
  };

  const handleMouseUp = () => { isDrawingRef.current = false; lastPointRef.current = null; };

  const handleTouchStart = (e: React.TouchEvent<HTMLCanvasElement>) => {
    e.preventDefault();
    isDrawingRef.current = true;
    const t = e.touches[0];
    lastPointRef.current = getCanvasPos(t.clientX, t.clientY, canvasRef.current!);
  };

  const handleTouchMove = (e: React.TouchEvent<HTMLCanvasElement>) => {
    e.preventDefault();
    if (!isDrawingRef.current) return;
    const t = e.touches[0];
    drawLine(getCanvasPos(t.clientX, t.clientY, canvasRef.current!));
  };

  const clearCanvas = () => {
    const canvas = canvasRef.current;
    if (canvas) canvas.getContext("2d")!.clearRect(0, 0, canvas.width, canvas.height);
    setHasDrawn(false);
  };

  const isComplete = data.status === "Complete";

  const AiDraftBtn = ({ sectionKey, label }: { sectionKey: "description" | "contributing" | "actions"; label: string }) => (
    <button
      type="button"
      disabled={isDrafting === sectionKey || isComplete}
      onClick={() => draftSection(sectionKey)}
      className="inline-flex items-center gap-1 text-xs bg-violet-600 hover:bg-violet-700 disabled:bg-gray-200 disabled:text-gray-400 text-white px-2 py-1 rounded"
    >
      {isDrafting === sectionKey ? <><Loader2 className="h-3 w-3 animate-spin" />Drafting…</> : <><Bot className="h-3 w-3" />{label}</>}
    </button>
  );

  const riskBadgeColor = (r: string) =>
    r === "Extreme" ? "bg-black text-white" :
    r === "High" ? "bg-red-500 text-white" :
    r === "Moderate" ? "bg-yellow-400 text-gray-900" :
    "bg-green-500 text-white";

  return (
    <Dialog open={open} onOpenChange={o => { if (!o) onClose(); }}>
      <DialogContent className="max-w-3xl max-h-[92vh] flex flex-col overflow-hidden p-0">
        {/* Header */}
        <div className="bg-gradient-to-r from-orange-600 to-red-700 text-white px-5 py-4 flex-shrink-0">
          {/* pr-8 keeps the status badge clear of the dialog's absolute-positioned close (X) button */}
          <div className="flex items-start justify-between gap-3 pr-8">
            <div className="flex items-start gap-3">
              <AlertTriangle className="h-6 w-6 mt-0.5 flex-shrink-0" />
              <div>
                <DialogTitle className="text-white text-base font-bold leading-tight">Near Miss Investigation</DialogTitle>
                <p className="text-orange-100 text-xs mt-0.5 line-clamp-1">{item.title}</p>
              </div>
            </div>
            <div className="flex items-center gap-2 flex-shrink-0">
              {data.status === "Complete" ? (
                <Badge className="bg-green-400 text-green-900 text-xs font-bold">Complete</Badge>
              ) : data.status === "In Progress" ? (
                <Badge className="bg-amber-300 text-amber-900 text-xs font-bold">In Progress</Badge>
              ) : null}
              {saveMsg && <span className="text-xs text-green-200">{saveMsg}</span>}
            </div>
          </div>

          {/* Section tabs */}
          <div className="flex gap-1 mt-3 overflow-x-auto">
            {SECTION_LABELS.map((lbl, i) => (
              <button
                key={i}
                onClick={() => setSection(i)}
                className={`flex-shrink-0 px-2.5 py-1 rounded-full text-[11px] font-medium transition-all ${
                  section === i ? "bg-white text-orange-700" : "bg-orange-500 bg-opacity-40 text-orange-100 hover:bg-opacity-60"
                }`}
              >
                {i + 1}. {lbl}
              </button>
            ))}
          </div>
        </div>

        {/* Body */}
        <div className="flex-1 overflow-y-auto px-5 py-4 space-y-4">
          {isComplete && (
            <div className="flex items-center gap-2 bg-green-50 border border-green-200 rounded-lg px-3 py-2 text-xs text-green-800 font-medium">
              <CheckCircle className="h-4 w-4 text-green-600 flex-shrink-0" />
              This investigation is complete and locked. Export the report or open it to review.
            </div>
          )}
          {isLoading && !loaded && (
            <div className="flex items-center justify-center h-32">
              <Loader2 className="h-6 w-6 animate-spin text-orange-500" />
            </div>
          )}

          <fieldset disabled={isComplete} className="contents">
          {/* Section 0: Event Details */}
          {section === 0 && (
            <div className="space-y-3">
              <h3 className="font-semibold text-gray-800 text-sm border-b pb-2">1. Event Details</h3>
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="label-sm">Investigator Name</label>
                  <input className="input-sm" value={data.investigatorName} onChange={e => update("investigatorName", e.target.value)} placeholder="Investigator name" />
                </div>
                <div>
                  <label className="label-sm">Site / Job Reference</label>
                  <input className="input-sm" value={data.siteJob} onChange={e => update("siteJob", e.target.value)} placeholder="e.g. Unit 11 – Dover Courts" />
                </div>
                <div>
                  <label className="label-sm">Date of Event</label>
                  <input type="date" className="input-sm" value={data.eventDate} onChange={e => update("eventDate", e.target.value)} />
                </div>
                <div>
                  <label className="label-sm">Time of Event</label>
                  <input type="time" className="input-sm" value={data.eventTime} onChange={e => update("eventTime", e.target.value)} />
                </div>
              </div>
              <div>
                <label className="label-sm">Event Type</label>
                <select className="input-sm" value={data.eventType} onChange={e => update("eventType", e.target.value)}>
                  {EVENT_TYPES.map(t => <option key={t} value={t}>{t}</option>)}
                </select>
              </div>
            </div>
          )}

          {/* Section 1: Involved Persons & Description */}
          {section === 1 && (
            <div className="space-y-3">
              <h3 className="font-semibold text-gray-800 text-sm border-b pb-2">2. Involved Persons & Description</h3>
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="label-sm">Involved Person(s)</label>
                  <input className="input-sm" value={data.involvedPersons} onChange={e => update("involvedPersons", e.target.value)} placeholder="Names of person(s) involved" />
                </div>
                <div>
                  <label className="label-sm">Witnesses</label>
                  <input className="input-sm" value={data.witnesses} onChange={e => update("witnesses", e.target.value)} placeholder="Witness names (or N/A)" />
                </div>
              </div>

              <div>
                <div className="flex items-center justify-between mb-1">
                  <label className="label-sm">How Did This Event Happen?</label>
                  <AiDraftBtn sectionKey="description" label="AI Draft" />
                </div>
                <textarea
                  rows={6}
                  className="input-sm resize-none"
                  value={data.eventDescription}
                  onChange={e => update("eventDescription", e.target.value)}
                  placeholder="Describe the events that led to this incident/near miss occurring..."
                />
              </div>

              <div>
                <div className="flex items-center justify-between mb-1">
                  <label className="label-sm">Contributing Factors (How Was This Allowed to Happen?)</label>
                  <AiDraftBtn sectionKey="contributing" label="AI Draft" />
                </div>
                <textarea
                  rows={5}
                  className="input-sm resize-none"
                  value={data.contributingFactors}
                  onChange={e => update("contributingFactors", e.target.value)}
                  placeholder="What processes were broken, safeguards non-existent, etc.?"
                />
              </div>
            </div>
          )}

          {/* Section 2: Risk Assessment */}
          {section === 2 && (
            <div className="space-y-4">
              <h3 className="font-semibold text-gray-800 text-sm border-b pb-2">3. Risk Assessment</h3>

              <div>
                <p className="text-xs text-gray-500 mb-2">Click a cell to select the likelihood and consequence for the <strong>overall event</strong>:</p>
                <RiskMatrix
                  selectedLikelihood={data.likelihood as any}
                  selectedConsequence={data.consequence as any}
                  onSelect={(likelihood, consequence) => {
                    update("likelihood", likelihood);
                    update("consequence", consequence);
                    update("riskLevel", getRiskLevelForCell(likelihood, consequence));
                  }}
                />
              </div>

              <div>
                <label className="label-sm mb-2 block">Hazard Register</label>
                <p className="text-xs text-gray-500 mb-2">Link the hazards involved from the Operational Hazard Register — new hazards can be added to the register as you go:</p>
                <HazardTable rows={data.hazards} onChange={rows => update("hazards", rows)} />
              </div>

              <div>
                <label className="label-sm">Treatment Given</label>
                <textarea
                  rows={3}
                  className="input-sm resize-none"
                  value={data.treatmentGiven}
                  onChange={e => update("treatmentGiven", e.target.value)}
                  placeholder="What treatment was given to immediately address the issue?"
                />
              </div>
            </div>
          )}

          {/* Section 3: Resulting Actions & Sign-Off */}
          {section === 3 && (
            <div className="space-y-4">
              <h3 className="font-semibold text-gray-800 text-sm border-b pb-2">4. Resulting Actions & Sign-Off</h3>

              <div>
                <div className="flex items-center justify-between mb-2">
                  <label className="label-sm">Resulting Actions</label>
                  <AiDraftBtn sectionKey="actions" label="AI Draft Actions" />
                </div>
                <div className="space-y-2">
                  {data.resultingActions.map((act, i) => (
                    <div key={act.id} className="flex items-start gap-2 bg-gray-50 rounded-lg p-2 border border-gray-200">
                      <input
                        type="checkbox"
                        checked={act.completed}
                        onChange={e => {
                          const updated = data.resultingActions.map((a, idx) => idx === i ? { ...a, completed: e.target.checked } : a);
                          update("resultingActions", updated);
                        }}
                        className="mt-1 h-4 w-4 rounded border-gray-300 text-orange-600"
                      />
                      <div className="flex-1 space-y-1">
                        <input
                          className="input-sm"
                          value={act.description}
                          onChange={e => {
                            const updated = data.resultingActions.map((a, idx) => idx === i ? { ...a, description: e.target.value } : a);
                            update("resultingActions", updated);
                          }}
                          placeholder="Action description..."
                        />
                        <input
                          className="input-sm"
                          value={act.assignedTo}
                          onChange={e => {
                            const updated = data.resultingActions.map((a, idx) => idx === i ? { ...a, assignedTo: e.target.value } : a);
                            update("resultingActions", updated);
                          }}
                          placeholder="Assigned to..."
                        />
                      </div>
                      <button
                        onClick={() => update("resultingActions", data.resultingActions.filter((_, idx) => idx !== i))}
                        className="text-gray-300 hover:text-red-500 mt-1"
                      >
                        <X className="h-4 w-4" />
                      </button>
                    </div>
                  ))}
                  <Button
                    variant="outline"
                    type="button"
                    className="w-full border-dashed border-orange-300 text-orange-600 hover:bg-orange-50 text-xs h-8"
                    onClick={() => update("resultingActions", [...data.resultingActions, { id: `a-${Date.now()}`, description: "", assignedTo: "", completed: false }])}
                  >
                    + Add Action
                  </Button>
                </div>
              </div>

              {/* Dual sign-off — both signatures required before the investigation can be completed */}
              <div className="space-y-3">
                {renderSignatureBlock("investigator", "Investigator Review & Sign-Off", "investigatorName", data.investigatorSignature, data.investigatorSignedAt)}
                {renderSignatureBlock("approver", "Approver / Manager Review & Sign-Off", "directorName", data.directorSignature, data.signedAt)}

                {isComplete ? (
                  <div className="flex items-start gap-2 text-sm text-green-700 bg-green-50 border border-green-200 rounded-lg p-3">
                    <CheckCircle className="h-4 w-4 mt-0.5 shrink-0" />
                    <span>Investigation complete — both signatures captured. The linked Action has been moved to <strong>Ready to Close</strong>.</span>
                  </div>
                ) : (
                  <p className="text-xs text-gray-500 bg-gray-50 border border-gray-200 rounded-lg p-3">
                    Both the <strong>Investigator</strong> and the <strong>Approver / Manager</strong> must sign before this investigation is complete. Completing it moves the linked Action to <strong>Ready to Close</strong>.
                  </p>
                )}
              </div>
            </div>
          )}
          </fieldset>

          {/* Progress history — time-stamped notes added over the life of the investigation.
              Kept at the bottom so it reads as a running log beneath the current step. */}
          <div className="border border-amber-200 rounded-lg overflow-hidden">
            <div className="flex items-center gap-2 px-3 py-2 bg-amber-50 border-b border-amber-200">
              <ClipboardList className="h-3.5 w-3.5 text-amber-600" />
              <span className="text-xs font-semibold text-amber-800">Progress History</span>
              <span className="text-[10px] text-amber-600">{progressNotes.length} {progressNotes.length === 1 ? "entry" : "entries"}</span>
            </div>
            <div className="px-3 py-2 bg-white space-y-2">
              {progressNotes.length === 0 ? (
                <p className="text-xs text-gray-400 italic">No progress notes yet. Add updates and findings as the investigation continues.</p>
              ) : (
                <ul className="space-y-2 max-h-40 overflow-y-auto">
                  {progressNotes.map(note => (
                    <li key={note.id} className="text-xs border-l-2 border-amber-300 pl-2 py-0.5">
                      <div className="flex items-center gap-2 text-[10px] text-gray-400">
                        <span>{format(new Date(note.createdAt), "d MMM yyyy, h:mm a")}</span>
                        {note.author && <span className="text-gray-500">• {note.author}</span>}
                      </div>
                      <p className="text-gray-700 whitespace-pre-wrap">{note.content}</p>
                    </li>
                  ))}
                </ul>
              )}
              {!isComplete && (
                <div className="flex items-start gap-2 pt-1">
                  <textarea
                    rows={2}
                    className="input-sm resize-none flex-1"
                    value={newNote}
                    onChange={e => setNewNote(e.target.value)}
                    placeholder="Add a progress note, update, or finding…"
                  />
                  <Button
                    size="sm"
                    type="button"
                    onClick={addNote}
                    disabled={isAddingNote || !newNote.trim()}
                    className="bg-amber-600 hover:bg-amber-700 text-white text-xs shrink-0"
                  >
                    {isAddingNote ? <Loader2 className="h-3.5 w-3.5 animate-spin" /> : "Add"}
                  </Button>
                </div>
              )}
            </div>
          </div>
        </div>

        {/* Footer */}
        <div className="flex items-center justify-between gap-2 px-5 py-3 border-t border-gray-100 bg-gray-50 flex-shrink-0">
          <div className="flex items-center gap-2">
            {section > 0 && (
              <Button variant="outline" size="sm" onClick={() => setSection(s => s - 1)}>← Back</Button>
            )}
            {section < 3 && (
              <Button
                size="sm"
                disabled={isSaving || isComplete}
                onClick={async () => {
                  await saveToDb();
                  setSection(s => s + 1);
                }}
                className="bg-orange-600 hover:bg-orange-700 text-white"
              >
                {isSaving ? <><Loader2 className="h-3.5 w-3.5 mr-1 animate-spin" />Saving…</> : "Save & Next →"}
              </Button>
            )}
          </div>
          <div className="flex items-center gap-2">
            {/* Hidden escape hatch — only visible while Ctrl/⌘ is held and the
                investigation isn't already complete. Archives the near miss with
                no report. */}
            {ctrlHeld && !isComplete && (
              <Button
                variant="outline"
                size="sm"
                onClick={() => setConfirmingClose(true)}
                className="text-xs border-red-300 text-red-700 hover:bg-red-50"
                data-testid="button-close-without-report"
              >
                <X className="h-3.5 w-3.5 mr-1" />Close without a report
              </Button>
            )}
            <Button
              variant="outline"
              size="sm"
              onClick={exportReport}
              disabled={isExporting}
              className="text-xs"
            >
              {isExporting ? <><Loader2 className="h-3.5 w-3.5 mr-1 animate-spin" />Generating…</> : <><FileText className="h-3.5 w-3.5 mr-1" />Export Report</>}
            </Button>
            {!isComplete && (
              <Button
                size="sm"
                onClick={() => saveToDb()}
                disabled={isSaving}
                className="bg-orange-600 hover:bg-orange-700 text-white"
              >
                {isSaving ? <><Loader2 className="h-3.5 w-3.5 mr-1 animate-spin" />Saving…</> : "Save Progress"}
              </Button>
            )}
            {isComplete && (
              <Badge className="bg-green-500 text-white px-3 py-1 text-xs font-bold">✓ Investigation Complete</Badge>
            )}
          </div>
        </div>

        {/* Confirmation overlay for the "Close without a report" escape hatch */}
        {confirmingClose && (
          <div className="absolute inset-0 z-50 flex items-center justify-center bg-black/40 px-4">
            <div className="bg-white rounded-xl shadow-2xl max-w-sm w-full p-5 space-y-4">
              <div className="flex items-start gap-3">
                <AlertTriangle className="h-6 w-6 text-red-600 flex-shrink-0 mt-0.5" />
                <div>
                  <h3 className="font-bold text-gray-900 text-sm">Close without a report?</h3>
                  <p className="text-xs text-gray-600 mt-1 leading-snug">
                    This archives the near miss with <strong>no investigation</strong> and no meeting sign-off. It stays on the register under Near Miss, marked “Closed without investigation”. Are you sure?
                  </p>
                </div>
              </div>
              <div className="flex justify-end gap-2">
                <Button
                  variant="outline"
                  size="sm"
                  onClick={() => setConfirmingClose(false)}
                  disabled={isClosingNoReport}
                  className="text-xs"
                  data-testid="button-cancel-close-without-report"
                >
                  Cancel
                </Button>
                <Button
                  size="sm"
                  onClick={closeWithoutReport}
                  disabled={isClosingNoReport}
                  className="bg-red-600 hover:bg-red-700 text-white text-xs"
                  data-testid="button-confirm-close-without-report"
                >
                  {isClosingNoReport ? <><Loader2 className="h-3.5 w-3.5 mr-1 animate-spin" />Closing…</> : "Yes, close it"}
                </Button>
              </div>
            </div>
          </div>
        )}
      </DialogContent>
    </Dialog>
  );
}

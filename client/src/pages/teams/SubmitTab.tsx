import { useState, useEffect, useRef } from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import { msalInstance, sharePointRequest, loginRequest } from "@/auth/msalConfig";
import { Button } from "@/components/ui/button";
import { Textarea } from "@/components/ui/textarea";
import { Badge } from "@/components/ui/badge";
import { Card } from "@/components/ui/card";
import {
  AlertTriangle,
  Lightbulb,
  Shield,
  ShoppingCart,
  MessageSquare,
  HelpCircle,
  CheckCircle2,
  Loader2,
  Send,
  ChevronRight,
  RotateCcw,
  Sparkles,
} from "lucide-react";

type Category =
  | "Near Miss"
  | "Safety Observation"
  | "Improvement Idea"
  | "Business Improvement"
  | "Supply Request"
  | "Meeting Agenda Item"
  | "Other";

type ListTarget = "near-miss" | "safety-ideas" | "business-ideas";

interface ClassifyResult {
  category: Category;
  listTarget: ListTarget;
  confidence: number;
  reasoning: string;
  followUpQuestions: string[];
}

const CATEGORY_META: Record<
  Category,
  { icon: React.ReactNode; color: string; bg: string; ring: string; listLabel: string }
> = {
  "Near Miss": {
    icon: <AlertTriangle className="h-5 w-5" />,
    color: "text-red-600",
    bg: "bg-red-50 border-red-200",
    ring: "ring-red-100",
    listLabel: "Near Miss Register",
  },
  "Safety Observation": {
    icon: <Shield className="h-5 w-5" />,
    color: "text-orange-600",
    bg: "bg-orange-50 border-orange-200",
    ring: "ring-orange-100",
    listLabel: "Safety Ideas List",
  },
  "Improvement Idea": {
    icon: <Shield className="h-5 w-5" />,
    color: "text-blue-600",
    bg: "bg-blue-50 border-blue-200",
    ring: "ring-blue-100",
    listLabel: "Safety Ideas List",
  },
  "Business Improvement": {
    icon: <Lightbulb className="h-5 w-5" />,
    color: "text-amber-600",
    bg: "bg-amber-50 border-amber-200",
    ring: "ring-amber-100",
    listLabel: "Business Ideas List",
  },
  "Supply Request": {
    icon: <ShoppingCart className="h-5 w-5" />,
    color: "text-purple-600",
    bg: "bg-purple-50 border-purple-200",
    ring: "ring-purple-100",
    listLabel: "Business Ideas List",
  },
  "Meeting Agenda Item": {
    icon: <MessageSquare className="h-5 w-5" />,
    color: "text-teal-600",
    bg: "bg-teal-50 border-teal-200",
    ring: "ring-teal-100",
    listLabel: "Business Ideas List",
  },
  Other: {
    icon: <HelpCircle className="h-5 w-5" />,
    color: "text-gray-600",
    bg: "bg-gray-50 border-gray-200",
    ring: "ring-gray-100",
    listLabel: "Business Ideas List",
  },
};

const LIST_TYPE_MAP: Record<ListTarget, string> = {
  "near-miss": "Near Miss",
  "safety-ideas": "Safety Ideas",
  "business-ideas": "Business Ideas",
};

const EXAMPLES = [
  "I nearly cut my hand on a broken panel near the cutting table…",
  "The fire exit by the loading bay is blocked with pallets.",
  "We're nearly out of safety gloves in the workshop.",
  "Idea: add a second squeegee station to speed up cleaning.",
];

function decodeJwtPayload(token: string): Record<string, any> {
  try {
    const payload = token.split(".")[1];
    const decoded = atob(payload.replace(/-/g, "+").replace(/_/g, "/"));
    return JSON.parse(decoded);
  } catch {
    return {};
  }
}


export default function SubmitTab() {
  const [authState, setAuthState] = useState<"loading" | "unauthenticated" | "authenticated">("loading");
  const [graphToken, setGraphToken] = useState<string | null>(null);
  const [sharePointToken, setSharePointToken] = useState<string | null>(null);
  const [userName, setUserName] = useState<string>("");

  const [step, setStep] = useState<
    "input" | "classifying" | "followup" | "confirm" | "submitting" | "done"
  >("input");

  const [inputText, setInputText] = useState("");
  const [classifyResult, setClassifyResult] = useState<ClassifyResult | null>(null);
  const [followUpAnswers, setFollowUpAnswers] = useState<string[]>([]);
  const [submitError, setSubmitError] = useState<string>("");
  const [submittedCategory, setSubmittedCategory] = useState<string>("");
  const [exampleIdx, setExampleIdx] = useState(0);
  const [prefetching, setPrefetching] = useState(false);

  const classifyTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const classifyCache = useRef<Map<string, ClassifyResult>>(new Map());
  const classifyInFlight = useRef<Map<string, Promise<ClassifyResult>>>(new Map());

  useEffect(() => {
    initAuth();
  }, []);

  // Rotate the placeholder example to hint at what to type (cheap perceived liveliness)
  useEffect(() => {
    if (step !== "input" || inputText) return;
    const id = setInterval(() => setExampleIdx((i) => (i + 1) % EXAMPLES.length), 3500);
    return () => clearInterval(id);
  }, [step, inputText]);

  // ─── Auth (Teams SSO only — company personal tab, no interactive fallback) ──
  async function initAuth() {
    try {
      await microsoftTeams.app.initialize();

      // Get identity from Teams SSO token
      const ssoToken = await microsoftTeams.authentication.getAuthToken();
      const payload = decodeJwtPayload(ssoToken);
      const loginHint = payload.upn || payload.preferred_username || payload.unique_name || "";
      const displayName = payload.name || loginHint;
      setUserName(displayName);

      // Acquire resource tokens silently using the Teams identity
      await msalInstance.initialize();
      const [graphResp, spResp] = await Promise.all([
        msalInstance.ssoSilent({ loginHint, scopes: loginRequest.scopes }),
        msalInstance.ssoSilent({ loginHint, scopes: sharePointRequest.scopes }),
      ]);
      setGraphToken(graphResp.accessToken);
      setSharePointToken(spResp.accessToken);
      setUserName(graphResp.account?.name || displayName);
      setAuthState("authenticated");
    } catch (err) {
      console.error("Teams SSO auth failed:", err);
      setAuthState("unauthenticated");
    }
  }

  // ─── Classification (eager + de-duped + cached) ──────────────────────────
  // Kicks off in the background while the user is still typing so that, by the
  // time they press Continue, the answer is usually already waiting → instant.
  function runClassify(text: string): Promise<ClassifyResult> {
    const key = text.trim();
    const cached = classifyCache.current.get(key);
    if (cached) return Promise.resolve(cached);
    const inflight = classifyInFlight.current.get(key);
    if (inflight) return inflight;

    const promise = (async () => {
      const resp = await fetch("/api/ai-classify", {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${graphToken}` },
        body: JSON.stringify({ text: key }),
      });
      const data = await resp.json();
      if (!data.success) throw new Error(data.error || "Classification failed");
      const result: ClassifyResult = {
        category: data.category,
        listTarget: data.listTarget,
        confidence: data.confidence,
        reasoning: data.reasoning,
        followUpQuestions: data.followUpQuestions || [],
      };
      classifyCache.current.set(key, result);
      classifyInFlight.current.delete(key);
      return result;
    })();

    classifyInFlight.current.set(key, promise);
    promise.catch(() => classifyInFlight.current.delete(key));
    return promise;
  }

  function handleTextChange(value: string) {
    setInputText(value);
    if (classifyTimer.current) clearTimeout(classifyTimer.current);
    const key = value.trim();
    if (key.length >= 20 && graphToken && !classifyCache.current.has(key)) {
      setPrefetching(true);
      classifyTimer.current = setTimeout(() => {
        runClassify(value)
          .catch(() => {})
          .finally(() => setPrefetching(false));
      }, 900);
    } else {
      setPrefetching(false);
    }
  }

  function applyResult(result: ClassifyResult) {
    setClassifyResult(result);
    setFollowUpAnswers(new Array(result.followUpQuestions.length).fill(""));
    setStep(result.followUpQuestions.length > 0 ? "followup" : "confirm");
  }

  async function handleContinue() {
    const key = inputText.trim();
    if (key.length < 10) return;
    setSubmitError("");

    // Already classified in the background → jump straight there, zero wait.
    const cached = classifyCache.current.get(key);
    if (cached) {
      applyResult(cached);
      return;
    }

    setStep("classifying");
    try {
      const result = await runClassify(inputText);
      applyResult(result);
    } catch (err: any) {
      setSubmitError(err.message || "Could not read your submission. Please try again.");
      setStep("input");
    }
  }

  // ─── Submission (deferred AI title → instant response) ───────────────────
  async function handleSubmit() {
    if (!classifyResult || !sharePointToken) return;
    setStep("submitting");
    setSubmitError("");

    const listType = LIST_TYPE_MAP[classifyResult.listTarget];
    const fullDescription =
      followUpAnswers.length > 0
        ? `${inputText.trim()}\n\n${classifyResult.followUpQuestions
            .map((q, i) => (followUpAnswers[i] ? `${q}\n${followUpAnswers[i]}` : ""))
            .filter(Boolean)
            .join("\n\n")}`
        : inputText.trim();

    const categoryLabel =
      classifyResult.category !== "Business Improvement" && classifyResult.listTarget === "business-ideas"
        ? classifyResult.category
        : undefined;

    try {
      const resp = await fetch("/api/sharepoint/create-item", {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${sharePointToken}` },
        body: JSON.stringify({
          listType,
          deferTitle: true,
          itemData: {
            title: "New Submission",
            description: fullDescription,
            submittedBy: userName,
            ideaType: categoryLabel,
            status: "Submitted",
          },
        }),
      });
      const data = await resp.json();
      if (!data.success) throw new Error(data.error || "Submission failed");
      setSubmittedCategory(classifyResult.category);
      setStep("done");
    } catch (err: any) {
      setSubmitError(err.message || "Submission failed. Please try again.");
      setStep("confirm");
    }
  }

  function reset() {
    setStep("input");
    setInputText("");
    setClassifyResult(null);
    setFollowUpAnswers([]);
    setSubmitError("");
    setSubmittedCategory("");
    setPrefetching(false);
  }

  // ─── Shared chrome ────────────────────────────────────────────────────────
  function Header() {
    return (
      <div className="bg-gradient-to-r from-blue-700 to-blue-500 px-5 pt-6 pb-7 text-white">
        <div className="max-w-lg mx-auto flex items-center gap-3">
          <div className="w-10 h-10 rounded-xl bg-white/15 backdrop-blur flex items-center justify-center">
            <Shield className="h-5 w-5" />
          </div>
          <div>
            <h1 className="text-base font-bold leading-tight">Safety &amp; Ideas</h1>
            <p className="text-xs text-blue-100/90">
              {userName ? `Hi ${userName.split(" ")[0]}` : "Cranfield Glass Christchurch"}
            </p>
          </div>
        </div>
      </div>
    );
  }

  // ─── Render ───────────────────────────────────────────────────────────────
  if (authState === "loading") {
    return (
      <div className="min-h-screen flex items-center justify-center bg-white animate-fade-in">
        <div className="text-center">
          <div className="relative w-12 h-12 mx-auto mb-4">
            <div className="absolute inset-0 rounded-full bg-blue-600 flex items-center justify-center">
              <Shield className="h-6 w-6 text-white" />
            </div>
            <div className="absolute -inset-1 rounded-full border-2 border-blue-600/30 border-t-blue-600 animate-spin" />
          </div>
          <p className="text-gray-500 text-sm">Signing you in automatically…</p>
        </div>
      </div>
    );
  }

  if (authState === "unauthenticated") {
    return (
      <div className="min-h-screen flex items-center justify-center bg-white p-4 animate-fade-in">
        <div className="w-full max-w-sm text-center animate-scale-in">
          <div className="w-14 h-14 bg-blue-100 rounded-2xl flex items-center justify-center mx-auto mb-4">
            <Shield className="h-7 w-7 text-blue-500" />
          </div>
          <h1 className="text-xl font-bold text-gray-900 mb-2">Not available here</h1>
          <p className="text-gray-500 text-sm">
            Please open this tab inside Microsoft Teams. If you're already in Teams, try refreshing the tab.
          </p>
        </div>
      </div>
    );
  }

  if (step === "done") {
    return (
      <div className="min-h-screen flex items-center justify-center bg-white p-4 animate-fade-in">
        <div className="w-full max-w-sm text-center">
          <div className="relative w-20 h-20 mx-auto mb-5">
            <div className="absolute inset-0 rounded-full bg-green-100 animate-pop-in" />
            <div className="absolute inset-0 flex items-center justify-center animate-pop-in">
              <CheckCircle2 className="h-11 w-11 text-green-600" />
            </div>
          </div>
          <h1 className="text-2xl font-bold text-gray-900 mb-2 animate-fade-in-up">Submitted!</h1>
          <p className="text-gray-600 text-sm mb-1 animate-fade-in-up">
            Your <strong>{submittedCategory}</strong> has been recorded.
          </p>
          <p className="text-gray-400 text-xs mb-7 animate-fade-in-up">
            It will appear in the next H&amp;S meeting agenda.
          </p>
          <Button onClick={reset} className="w-full bg-blue-600 hover:bg-blue-700 text-white">
            <RotateCcw className="h-4 w-4 mr-2" />
            Submit another
          </Button>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gray-50">
      <Header />
      <div className="max-w-lg mx-auto px-4 -mt-4 pb-24">
        <Card className="p-5 shadow-sm border-gray-200 animate-fade-in-up">
          {step === "input" && (
            <div className="space-y-4 animate-fade-in">
              <p className="text-sm text-gray-600 font-medium">
                Describe what you saw, a near miss, or an improvement idea.
              </p>
              <div className="relative">
                <Textarea
                  placeholder={EXAMPLES[exampleIdx]}
                  value={inputText}
                  onChange={(e) => handleTextChange(e.target.value)}
                  rows={5}
                  className="resize-none text-base border-gray-200 focus:border-blue-400 transition-colors"
                  autoFocus
                />
                {prefetching && (
                  <div className="absolute bottom-2.5 right-3 flex items-center gap-1.5 text-xs text-blue-500 animate-fade-in">
                    <Sparkles className="h-3.5 w-3.5 animate-pulse" />
                    reading…
                  </div>
                )}
              </div>
              {inputText.trim().length >= 10 && inputText.trim().length < 20 && (
                <p className="text-xs text-gray-400">Keep typing — a bit more detail helps…</p>
              )}
              {submitError && (
                <p className="text-sm text-red-600 bg-red-50 border border-red-200 rounded-lg p-3 animate-fade-in">
                  {submitError}
                </p>
              )}
              <Button
                onClick={handleContinue}
                disabled={inputText.trim().length < 10}
                className="w-full bg-blue-600 hover:bg-blue-700 active:scale-[0.98] transition-transform text-white"
              >
                {classifyCache.current.has(inputText.trim()) ? (
                  <>
                    <Sparkles className="h-4 w-4 mr-2" />
                    Ready — Continue
                  </>
                ) : (
                  <>
                    <ChevronRight className="h-4 w-4 mr-2" />
                    Continue
                  </>
                )}
              </Button>
            </div>
          )}

          {step === "classifying" && <ClassifyingSkeleton />}

          {step === "followup" && classifyResult && (
            <div className="space-y-4 animate-fade-in-up">
              <div className={`p-4 rounded-xl border ring-4 ${CATEGORY_META[classifyResult.category]?.bg} ${CATEGORY_META[classifyResult.category]?.ring}`}>
                <div className={`flex items-center gap-2 ${CATEGORY_META[classifyResult.category]?.color}`}>
                  {CATEGORY_META[classifyResult.category]?.icon}
                  <span className="font-semibold text-sm">{classifyResult.category}</span>
                  <Badge variant="outline" className="ml-auto text-xs bg-white/60">
                    {Math.round(classifyResult.confidence * 100)}% sure
                  </Badge>
                </div>
                <p className="text-xs text-gray-500 mt-2">
                  A couple of quick questions will help us record this accurately.
                </p>
              </div>

              <div className="space-y-3">
                {classifyResult.followUpQuestions.map((question, idx) => (
                  <div key={idx} className="animate-fade-in-up" style={{ animationDelay: `${idx * 60}ms` }}>
                    <label className="block text-sm font-medium text-gray-700 mb-1">{question}</label>
                    <Textarea
                      placeholder="Your answer…"
                      value={followUpAnswers[idx] || ""}
                      onChange={(e) => {
                        const updated = [...followUpAnswers];
                        updated[idx] = e.target.value;
                        setFollowUpAnswers(updated);
                      }}
                      rows={2}
                      className="resize-none text-sm border-gray-200"
                    />
                  </div>
                ))}
              </div>

              <Button
                onClick={() => setStep("confirm")}
                className="w-full bg-blue-600 hover:bg-blue-700 active:scale-[0.98] transition-transform text-white"
              >
                <ChevronRight className="h-4 w-4 mr-2" />
                Review &amp; Submit
              </Button>
              <Button variant="ghost" onClick={reset} className="w-full text-gray-500 text-sm">
                Start over
              </Button>
            </div>
          )}

          {step === "confirm" && classifyResult && (
            <div className="space-y-4 animate-fade-in-up">
              <div className={`p-4 rounded-xl border ${CATEGORY_META[classifyResult.category]?.bg}`}>
                <div className={`flex items-center gap-2 font-semibold text-sm mb-2 ${CATEGORY_META[classifyResult.category]?.color}`}>
                  {CATEGORY_META[classifyResult.category]?.icon}
                  {classifyResult.category}
                </div>
                <p className="text-xs text-gray-500">
                  → Goes to:{" "}
                  <span className="font-medium">{CATEGORY_META[classifyResult.category]?.listLabel}</span>
                </p>
              </div>

              <div className="bg-gray-50 rounded-xl border border-gray-200 p-4">
                <p className="text-xs font-medium text-gray-500 uppercase tracking-wide mb-2">Your submission</p>
                <p className="text-sm text-gray-800 whitespace-pre-wrap">{inputText}</p>
                {followUpAnswers.some((a) => a.trim()) && (
                  <div className="mt-3 pt-3 border-t border-gray-200 space-y-2">
                    {classifyResult.followUpQuestions.map((q, i) =>
                      followUpAnswers[i]?.trim() ? (
                        <div key={i}>
                          <p className="text-xs text-gray-500">{q}</p>
                          <p className="text-sm text-gray-800">{followUpAnswers[i]}</p>
                        </div>
                      ) : null
                    )}
                  </div>
                )}
              </div>

              {submitError && (
                <p className="text-sm text-red-600 bg-red-50 border border-red-200 rounded-lg p-3 animate-fade-in">
                  {submitError}
                </p>
              )}

              <Button
                onClick={handleSubmit}
                className="w-full bg-green-600 hover:bg-green-700 active:scale-[0.98] transition-transform text-white"
              >
                <Send className="h-4 w-4 mr-2" />
                Submit
              </Button>
              <Button
                variant="ghost"
                onClick={() => (classifyResult.followUpQuestions.length > 0 ? setStep("followup") : setStep("input"))}
                className="w-full text-gray-500 text-sm"
              >
                Go back
              </Button>
            </div>
          )}

          {step === "submitting" && (
            <div className="py-10 animate-fade-in">
              <div className="relative h-1.5 w-full bg-gray-100 rounded-full overflow-hidden mb-5">
                <div className="absolute top-0 h-full bg-green-500 rounded-full animate-bar-indeterminate" />
              </div>
              <p className="text-center text-gray-500 text-sm">Saving your submission…</p>
            </div>
          )}
        </Card>
      </div>
    </div>
  );
}

// Skeleton that mirrors the result card layout so the transition feels seamless
function ClassifyingSkeleton() {
  return (
    <div className="space-y-4 animate-fade-in">
      <div className="flex items-center gap-2 text-blue-600">
        <Sparkles className="h-4 w-4 animate-pulse" />
        <span className="text-sm font-medium">Reading your submission…</span>
      </div>
      <div className="relative h-1.5 w-full bg-gray-100 rounded-full overflow-hidden">
        <div className="absolute top-0 h-full bg-blue-500 rounded-full animate-bar-indeterminate" />
      </div>
      <div className="rounded-xl border border-gray-200 p-4 space-y-3">
        <Shimmer className="h-5 w-32" />
        <Shimmer className="h-3 w-full" />
        <Shimmer className="h-3 w-3/4" />
      </div>
      <div className="space-y-2">
        <Shimmer className="h-3 w-40" />
        <Shimmer className="h-9 w-full" />
      </div>
    </div>
  );
}

function Shimmer({ className = "" }: { className?: string }) {
  return (
    <div className={`relative overflow-hidden rounded-md bg-gray-100 ${className}`}>
      <div className="absolute inset-0 -translate-x-full animate-shimmer bg-gradient-to-r from-transparent via-white/70 to-transparent" />
    </div>
  );
}

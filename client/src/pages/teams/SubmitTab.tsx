import { useState, useEffect, useRef } from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import { Button } from "@/components/ui/button";
import { Textarea } from "@/components/ui/textarea";
import { Badge } from "@/components/ui/badge";
import { Card } from "@/components/ui/card";
import { useTeamsTheme } from "@/hooks/useTeamsTheme";
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
  LogIn,
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
  { icon: React.ReactNode; color: string; bg: string; ring: string; darkBg: string; darkRing: string; listLabel: string }
> = {
  "Near Miss": {
    icon: <AlertTriangle className="h-5 w-5" />,
    color: "text-red-600",
    bg: "bg-red-50 border-red-200",
    ring: "ring-red-100",
    darkBg: "bg-red-900/20 border-red-700/40",
    darkRing: "ring-red-900/30",
    listLabel: "Near Miss Register",
  },
  "Safety Observation": {
    icon: <Shield className="h-5 w-5" />,
    color: "text-orange-600",
    bg: "bg-orange-50 border-orange-200",
    ring: "ring-orange-100",
    darkBg: "bg-orange-900/20 border-orange-700/40",
    darkRing: "ring-orange-900/30",
    listLabel: "Safety Ideas List",
  },
  "Improvement Idea": {
    icon: <Shield className="h-5 w-5" />,
    color: "text-blue-500",
    bg: "bg-blue-50 border-blue-200",
    ring: "ring-blue-100",
    darkBg: "bg-blue-900/20 border-blue-700/40",
    darkRing: "ring-blue-900/30",
    listLabel: "Safety Ideas List",
  },
  "Business Improvement": {
    icon: <Lightbulb className="h-5 w-5" />,
    color: "text-amber-500",
    bg: "bg-amber-50 border-amber-200",
    ring: "ring-amber-100",
    darkBg: "bg-amber-900/20 border-amber-700/40",
    darkRing: "ring-amber-900/30",
    listLabel: "Business Ideas List",
  },
  "Supply Request": {
    icon: <ShoppingCart className="h-5 w-5" />,
    color: "text-purple-500",
    bg: "bg-purple-50 border-purple-200",
    ring: "ring-purple-100",
    darkBg: "bg-purple-900/20 border-purple-700/40",
    darkRing: "ring-purple-900/30",
    listLabel: "Business Ideas List",
  },
  "Meeting Agenda Item": {
    icon: <MessageSquare className="h-5 w-5" />,
    color: "text-teal-500",
    bg: "bg-teal-50 border-teal-200",
    ring: "ring-teal-100",
    darkBg: "bg-teal-900/20 border-teal-700/40",
    darkRing: "ring-teal-900/30",
    listLabel: "Business Ideas List",
  },
  Other: {
    icon: <HelpCircle className="h-5 w-5" />,
    color: "text-gray-500",
    bg: "bg-gray-50 border-gray-200",
    ring: "ring-gray-100",
    darkBg: "bg-gray-700/40 border-gray-600/40",
    darkRing: "ring-gray-700/30",
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
  const { isDark } = useTeamsTheme();

  const [authState, setAuthState] = useState<"loading" | "unauthenticated" | "authenticated">("loading");
  const [authError, setAuthError] = useState<string>("");
  const [teamsToken, setTeamsToken] = useState<string | null>(null);
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

  useEffect(() => {
    if (step !== "input" || inputText) return;
    const id = setInterval(() => setExampleIdx((i) => (i + 1) % EXAMPLES.length), 3500);
    return () => clearInterval(id);
  }, [step, inputText]);

  // ─── Auth — Teams SSO only ────────────────────────────────────────────────
  // We acquire a single SSO token via getAuthToken(). The backend exchanges it
  // (on behalf of the user) for Graph / SharePoint access, so the browser never
  // touches MSAL, popups, or redirects — eliminating the iframe SSO races.
  async function initAuth() {
    setAuthState("loading");
    setAuthError("");
    try {
      await microsoftTeams.app.initialize();
      const ssoToken = await microsoftTeams.authentication.getAuthToken();
      const payload = decodeJwtPayload(ssoToken);
      setUserName(payload.name || payload.preferred_username || payload.upn || "");
      setTeamsToken(ssoToken);
      setAuthState("authenticated");
    } catch (err: any) {
      const msg = `Teams sign-in failed: ${err?.message || String(err)}`;
      console.error(msg, err);
      setAuthError(msg);
      setAuthState("unauthenticated");
    }
  }

  // ─── Classification ───────────────────────────────────────────────────────
  function runClassify(text: string): Promise<ClassifyResult> {
    const key = text.trim();
    const cached = classifyCache.current.get(key);
    if (cached) return Promise.resolve(cached);
    const inflight = classifyInFlight.current.get(key);
    if (inflight) return inflight;

    const promise = (async () => {
      const resp = await fetch("/api/ai-classify", {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${teamsToken}` },
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
    if (key.length >= 20 && teamsToken && !classifyCache.current.has(key)) {
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
    const cached = classifyCache.current.get(key);
    if (cached) { applyResult(cached); return; }
    setStep("classifying");
    try {
      const result = await runClassify(inputText);
      applyResult(result);
    } catch (err: any) {
      setSubmitError(err.message || "Could not read your submission. Please try again.");
      setStep("input");
    }
  }

  // ─── Submission ───────────────────────────────────────────────────────────
  async function handleSubmit() {
    if (!classifyResult || !teamsToken) return;
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
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${teamsToken}` },
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

  // ─── Full-screen states ───────────────────────────────────────────────────
  if (authState === "loading") {
    return (
      <div className={`h-full flex items-center justify-center animate-fade-in ${
        isDark ? "bg-gray-900" : "bg-white"
      }`}>
        <div className="text-center">
          <div className="relative w-12 h-12 mx-auto mb-4">
            <div className="absolute inset-0 rounded-full bg-blue-600 flex items-center justify-center">
              <Shield className="h-6 w-6 text-white" />
            </div>
            <div className="absolute -inset-1 rounded-full border-2 border-blue-600/30 border-t-blue-600 animate-spin" />
          </div>
          <p className={`text-sm ${isDark ? "text-gray-400" : "text-gray-500"}`}>
            Signing you in automatically…
          </p>
        </div>
      </div>
    );
  }

  if (authState === "unauthenticated") {
    return (
      <div className={`h-full flex items-center justify-center p-4 animate-fade-in ${
        isDark ? "bg-gray-900" : "bg-white"
      }`}>
        <div className="w-full max-w-sm text-center animate-scale-in">
          <div className={`w-14 h-14 rounded-2xl flex items-center justify-center mx-auto mb-4 ${
            isDark ? "bg-blue-900/40" : "bg-blue-100"
          }`}>
            <Shield className={`h-7 w-7 ${isDark ? "text-blue-400" : "text-blue-500"}`} />
          </div>
          <h1 className={`text-xl font-bold mb-2 ${isDark ? "text-white" : "text-gray-900"}`}>
            Sign in required
          </h1>
          <p className={`text-sm mb-4 ${isDark ? "text-gray-400" : "text-gray-500"}`}>
            Sign in with your Cranfield Glass Microsoft account to submit ideas and reports.
          </p>
          <Button
            onClick={() => initAuth()}
            className="w-full bg-blue-600 hover:bg-blue-700 text-white mb-3"
          >
            <LogIn className="h-4 w-4 mr-2" />
            Try again
          </Button>
          {authError && (
            <div className="mt-2 p-3 rounded-lg bg-red-50 border border-red-200 text-left">
              <p className="text-xs font-mono text-red-700 break-all">{authError}</p>
            </div>
          )}
        </div>
      </div>
    );
  }

  if (step === "done") {
    return (
      <div className={`h-full flex items-center justify-center p-4 animate-fade-in ${
        isDark ? "bg-gray-900" : "bg-white"
      }`}>
        <div className="w-full max-w-sm text-center">
          <div className="relative w-20 h-20 mx-auto mb-5">
            <div className="absolute inset-0 rounded-full bg-green-100 animate-pop-in" />
            <div className="absolute inset-0 flex items-center justify-center animate-pop-in">
              <CheckCircle2 className="h-11 w-11 text-green-600" />
            </div>
          </div>
          <h1 className={`text-2xl font-bold mb-2 animate-fade-in-up ${isDark ? "text-white" : "text-gray-900"}`}>
            Submitted!
          </h1>
          <p className={`text-sm mb-1 animate-fade-in-up ${isDark ? "text-gray-300" : "text-gray-600"}`}>
            Your <strong>{submittedCategory}</strong> has been recorded.
          </p>
          <p className={`text-xs mb-7 animate-fade-in-up ${isDark ? "text-gray-500" : "text-gray-400"}`}>
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

  // ─── Main form ────────────────────────────────────────────────────────────
  const meta = classifyResult ? CATEGORY_META[classifyResult.category] : null;
  const categoryBg = meta ? (isDark ? meta.darkBg : meta.bg) : "";
  const categoryRing = meta ? (isDark ? meta.darkRing : meta.ring) : "";

  return (
    <div className={`flex flex-col h-full min-h-0 ${isDark ? "bg-gray-900" : "bg-white"}`}>
      {userName && (
        <div className="shrink-0 px-5 pt-1 pb-3">
          <p className={`text-2xl font-bold tracking-tight ${isDark ? "text-white" : "text-gray-900"}`}>
            Hi {userName.split(" ")[0]} <span className="ml-0.5">👋</span>
          </p>
          <p className={`mt-0.5 text-sm ${isDark ? "text-gray-400" : "text-gray-500"}`}>
            Small ideas. Continuous improvement.
          </p>
        </div>
      )}
      <div className="flex-1 min-h-0 overflow-y-auto px-4 pt-1 pb-5">
        <Card className={`w-full max-w-lg mx-auto p-5 shadow-sm animate-fade-in-up ${
          isDark ? "bg-gray-800 border-gray-700" : "bg-white border-gray-200"
        }`}>
          {step === "input" && (
            <div className="space-y-4 animate-fade-in">
              <p className={`text-sm font-medium ${isDark ? "text-gray-300" : "text-gray-600"}`}>
                Describe what you saw, a near miss, or an improvement idea.
              </p>
              <div className="relative">
                <Textarea
                  placeholder={EXAMPLES[exampleIdx]}
                  value={inputText}
                  onChange={(e) => handleTextChange(e.target.value)}
                  rows={5}
                  className={`resize-none text-base transition-colors ${
                    isDark
                      ? "bg-gray-700 border-gray-600 text-white placeholder:text-gray-500 focus:border-blue-500"
                      : "border-gray-200 focus:border-blue-400"
                  }`}
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
                <p className={`text-xs ${isDark ? "text-gray-500" : "text-gray-400"}`}>
                  Keep typing — a bit more detail helps…
                </p>
              )}
              {submitError && (
                <p className={`text-sm border rounded-lg p-3 animate-fade-in ${
                  isDark
                    ? "text-red-400 bg-red-900/20 border-red-700/50"
                    : "text-red-600 bg-red-50 border-red-200"
                }`}>
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

          {step === "classifying" && <ClassifyingSkeleton isDark={isDark} />}

          {step === "followup" && classifyResult && meta && (
            <div className="space-y-4 animate-fade-in-up">
              <div className={`p-4 rounded-xl border ring-4 ${categoryBg} ${categoryRing}`}>
                <div className={`flex items-center gap-2 ${meta.color}`}>
                  {meta.icon}
                  <span className="font-semibold text-sm">{classifyResult.category}</span>
                  <Badge variant="outline" className={`ml-auto text-xs ${isDark ? "bg-gray-900/50 border-gray-600 text-gray-300" : "bg-white/60"}`}>
                    {Math.round(classifyResult.confidence * 100)}% sure
                  </Badge>
                </div>
                <p className={`text-xs mt-2 ${isDark ? "text-gray-400" : "text-gray-500"}`}>
                  A couple of quick questions will help us record this accurately.
                </p>
              </div>

              <div className="space-y-3">
                {classifyResult.followUpQuestions.map((question, idx) => (
                  <div key={idx} className="animate-fade-in-up" style={{ animationDelay: `${idx * 60}ms` }}>
                    <label className={`block text-sm font-medium mb-1 ${isDark ? "text-gray-300" : "text-gray-700"}`}>
                      {question}
                    </label>
                    <Textarea
                      placeholder="Your answer…"
                      value={followUpAnswers[idx] || ""}
                      onChange={(e) => {
                        const updated = [...followUpAnswers];
                        updated[idx] = e.target.value;
                        setFollowUpAnswers(updated);
                      }}
                      rows={2}
                      className={`resize-none text-sm ${
                        isDark
                          ? "bg-gray-700 border-gray-600 text-white placeholder:text-gray-500"
                          : "border-gray-200"
                      }`}
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
              <Button variant="ghost" onClick={reset} className={`w-full text-sm ${isDark ? "text-gray-400 hover:text-gray-300 hover:bg-gray-700" : "text-gray-500"}`}>
                Start over
              </Button>
            </div>
          )}

          {step === "confirm" && classifyResult && meta && (
            <div className="space-y-4 animate-fade-in-up">
              <div className={`p-4 rounded-xl border ${categoryBg}`}>
                <div className={`flex items-center gap-2 font-semibold text-sm mb-2 ${meta.color}`}>
                  {meta.icon}
                  {classifyResult.category}
                </div>
                <p className={`text-xs ${isDark ? "text-gray-400" : "text-gray-500"}`}>
                  → Goes to: <span className="font-medium">{meta.listLabel}</span>
                </p>
              </div>

              <div className={`rounded-xl border p-4 ${isDark ? "bg-gray-700 border-gray-600" : "bg-gray-50 border-gray-200"}`}>
                <p className={`text-xs font-medium uppercase tracking-wide mb-2 ${isDark ? "text-gray-400" : "text-gray-500"}`}>
                  Your submission
                </p>
                <p className={`text-sm whitespace-pre-wrap ${isDark ? "text-gray-100" : "text-gray-800"}`}>
                  {inputText}
                </p>
                {followUpAnswers.some((a) => a.trim()) && (
                  <div className={`mt-3 pt-3 border-t space-y-2 ${isDark ? "border-gray-600" : "border-gray-200"}`}>
                    {classifyResult.followUpQuestions.map((q, i) =>
                      followUpAnswers[i]?.trim() ? (
                        <div key={i}>
                          <p className={`text-xs ${isDark ? "text-gray-400" : "text-gray-500"}`}>{q}</p>
                          <p className={`text-sm ${isDark ? "text-gray-100" : "text-gray-800"}`}>{followUpAnswers[i]}</p>
                        </div>
                      ) : null
                    )}
                  </div>
                )}
              </div>

              {submitError && (
                <p className={`text-sm border rounded-lg p-3 animate-fade-in ${
                  isDark
                    ? "text-red-400 bg-red-900/20 border-red-700/50"
                    : "text-red-600 bg-red-50 border-red-200"
                }`}>
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
                className={`w-full text-sm ${isDark ? "text-gray-400 hover:text-gray-300 hover:bg-gray-700" : "text-gray-500"}`}
              >
                Go back
              </Button>
            </div>
          )}

          {step === "submitting" && (
            <div className="py-10 animate-fade-in">
              <div className="relative h-1.5 w-full bg-gray-100 dark:bg-gray-700 rounded-full overflow-hidden mb-5">
                <div className="absolute top-0 h-full bg-green-500 rounded-full animate-bar-indeterminate" />
              </div>
              <p className={`text-center text-sm ${isDark ? "text-gray-400" : "text-gray-500"}`}>
                Saving your submission…
              </p>
            </div>
          )}
        </Card>
      </div>
    </div>
  );
}

function ClassifyingSkeleton({ isDark }: { isDark: boolean }) {
  return (
    <div className="space-y-4 animate-fade-in">
      <div className="flex items-center gap-2 text-blue-500">
        <Sparkles className="h-4 w-4 animate-pulse" />
        <span className="text-sm font-medium">Reading your submission…</span>
      </div>
      <div className="relative h-1.5 w-full bg-gray-100 dark:bg-gray-700 rounded-full overflow-hidden">
        <div className="absolute top-0 h-full bg-blue-500 rounded-full animate-bar-indeterminate" />
      </div>
      <div className={`rounded-xl border p-4 space-y-3 ${isDark ? "border-gray-700" : "border-gray-200"}`}>
        <Shimmer isDark={isDark} className="h-5 w-32" />
        <Shimmer isDark={isDark} className="h-3 w-full" />
        <Shimmer isDark={isDark} className="h-3 w-3/4" />
      </div>
      <div className="space-y-2">
        <Shimmer isDark={isDark} className="h-3 w-40" />
        <Shimmer isDark={isDark} className="h-9 w-full" />
      </div>
    </div>
  );
}

function Shimmer({ className = "", isDark }: { className?: string; isDark: boolean }) {
  return (
    <div className={`relative overflow-hidden rounded-md ${isDark ? "bg-gray-700" : "bg-gray-100"} ${className}`}>
      <div className={`absolute inset-0 -translate-x-full animate-shimmer bg-gradient-to-r from-transparent ${isDark ? "via-gray-600/50" : "via-white/70"} to-transparent`} />
    </div>
  );
}

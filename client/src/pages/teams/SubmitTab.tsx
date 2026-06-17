import { useState, useEffect, useRef } from "react";
import {
  Button,
  Textarea,
  Card,
  Badge,
  Spinner,
  ProgressBar,
  Text,
  Field,
  Skeleton,
  SkeletonItem,
  MessageBar,
  MessageBarBody,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import {
  Warning24Regular,
  Shield24Regular,
  Lightbulb24Regular,
  Cart24Regular,
  Chat24Regular,
  QuestionCircle24Regular,
  CheckmarkCircle24Filled,
  Send20Regular,
  ChevronRight20Regular,
  ArrowCounterclockwise20Regular,
  Sparkle16Regular,
} from "@fluentui/react-icons";
import {
  TeamsPage,
  TeamsScroll,
  TeamsCenter,
  TeamsFullScreen,
} from "./TeamsPageShell";
import { useTeamsAuth } from "@/hooks/useTeamsAuth";

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

type SubmitStep =
  | "input"
  | "classifying"
  | "followup"
  | "confirm"
  | "submitting"
  | "done";

// The Submit flow persists to localStorage so an in-progress report survives a
// tab switch (Teams remounts the tab) or the app being closed and reopened.
// Bump the version suffix whenever the draft shape changes so stale drafts from
// an older build can never restore into a broken render.
const DRAFT_KEY = "cranfield.submit.draft.v2";

interface SubmitDraft {
  step: SubmitStep;
  inputText: string;
  classifyResult: ClassifyResult | null;
  followUpAnswers: string[];
  submittedCategory: string;
}

// Only the "input" step can be resumed from a raw text draft; "followup"/"confirm"
// additionally require a fully-formed classification (see isValidClassifyResult).
const RESUMABLE_STEPS: SubmitStep[] = ["input", "followup", "confirm"];

const VALID_CATEGORIES: Category[] = [
  "Near Miss",
  "Safety Observation",
  "Improvement Idea",
  "Business Improvement",
  "Supply Request",
  "Meeting Agenda Item",
  "Other",
];
const VALID_LIST_TARGETS: ListTarget[] = ["near-miss", "safety-ideas", "business-ideas"];

// A restored classification is only safe to render in stage 2 if every field the
// confirm/followup UI reads is present and the right type. A draft from an older
// build (e.g. missing followUpQuestions, or an unknown category) would otherwise
// crash `.map(...)` / leave `meta` undefined and blank the tab.
function isValidClassifyResult(r: any): r is ClassifyResult {
  return (
    !!r &&
    typeof r === "object" &&
    VALID_CATEGORIES.includes(r.category) &&
    VALID_LIST_TARGETS.includes(r.listTarget) &&
    typeof r.confidence === "number" &&
    Array.isArray(r.followUpQuestions)
  );
}

// A clean text-only draft: keep what the user typed but start stage 1 fresh.
function inputOnlyDraft(inputText: string): SubmitDraft {
  return { step: "input", inputText, classifyResult: null, followUpAnswers: [], submittedCategory: "" };
}

function loadDraft(): SubmitDraft | null {
  try {
    const raw = localStorage.getItem(DRAFT_KEY);
    if (!raw) return null;
    const d = JSON.parse(raw) as SubmitDraft;
    if (!d || typeof d !== "object" || typeof d.inputText !== "string") return null;
    // Transient in-flight steps can't be resumed (the network call was lost),
    // so drop back to the nearest stable step the user can act on.
    if (d.step === "classifying") d.step = "input";
    if (d.step === "submitting") d.step = d.classifyResult ? "confirm" : "input";
    // A completed submission shouldn't reopen — start fresh next time.
    if (d.step === "done") return null;
    // Unknown/corrupted step: don't try to render it.
    if (!RESUMABLE_STEPS.includes(d.step)) return null;
    // Stage 2 needs a valid classification. If it's missing or from an older
    // shape, fall back to stage 1 so the user keeps their text instead of a
    // blank, stuck card.
    if (d.step === "confirm" || d.step === "followup") {
      if (!isValidClassifyResult(d.classifyResult)) return inputOnlyDraft(d.inputText);
      return {
        step: d.step,
        inputText: d.inputText,
        classifyResult: d.classifyResult,
        followUpAnswers: Array.isArray(d.followUpAnswers) ? d.followUpAnswers : [],
        submittedCategory: typeof d.submittedCategory === "string" ? d.submittedCategory : "",
      };
    }
    return inputOnlyDraft(d.inputText);
  } catch {
    return null;
  }
}

const CATEGORY_META: Record<
  Category,
  { icon: React.ReactNode; fg: string; bg: string; border: string; listLabel: string }
> = {
  "Near Miss": {
    icon: <Warning24Regular />,
    fg: tokens.colorPaletteRedForeground1,
    bg: tokens.colorPaletteRedBackground1,
    border: tokens.colorPaletteRedBorder1,
    listLabel: "Near Miss Register",
  },
  "Safety Observation": {
    icon: <Shield24Regular />,
    fg: tokens.colorPaletteDarkOrangeForeground1,
    bg: tokens.colorPaletteDarkOrangeBackground1,
    border: tokens.colorPaletteDarkOrangeBorder1,
    listLabel: "Safety Ideas List",
  },
  "Improvement Idea": {
    icon: <Shield24Regular />,
    fg: tokens.colorPaletteGreenForeground1,
    bg: tokens.colorPaletteGreenBackground1,
    border: tokens.colorPaletteGreenBorder1,
    listLabel: "Safety Ideas List",
  },
  "Business Improvement": {
    icon: <Lightbulb24Regular />,
    fg: tokens.colorPaletteMarigoldForeground1,
    bg: tokens.colorPaletteMarigoldBackground1,
    border: tokens.colorPaletteMarigoldBorder1,
    listLabel: "Business Ideas List",
  },
  "Supply Request": {
    icon: <Cart24Regular />,
    fg: tokens.colorPaletteBerryForeground1,
    bg: tokens.colorPaletteBerryBackground1,
    border: tokens.colorPaletteBerryBorder1,
    listLabel: "Business Ideas List",
  },
  "Meeting Agenda Item": {
    icon: <Chat24Regular />,
    fg: tokens.colorPaletteLightGreenForeground1,
    bg: tokens.colorPaletteLightGreenBackground1,
    border: tokens.colorPaletteLightGreenBorder1,
    listLabel: "Business Ideas List",
  },
  Other: {
    icon: <QuestionCircle24Regular />,
    fg: tokens.colorNeutralForeground2,
    bg: tokens.colorNeutralBackground3,
    border: tokens.colorNeutralStroke2,
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

// Griffel blocks `border*` 4-side shorthands; set each side longhand.
const thinBorderSides = {
  borderTopWidth: tokens.strokeWidthThin,
  borderRightWidth: tokens.strokeWidthThin,
  borderBottomWidth: tokens.strokeWidthThin,
  borderLeftWidth: tokens.strokeWidthThin,
  borderTopStyle: "solid",
  borderRightStyle: "solid",
  borderBottomStyle: "solid",
  borderLeftStyle: "solid",
} as const;
const allBorderColor = (c: string) => ({
  borderTopColor: c,
  borderRightColor: c,
  borderBottomColor: c,
  borderLeftColor: c,
});

const useStyles = makeStyles({
  inputStack: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
    paddingLeft: tokens.spacingHorizontalL,
    paddingRight: tokens.spacingHorizontalL,
    paddingTop: tokens.spacingVerticalM,
    paddingBottom: tokens.spacingVerticalXXXL,
  },
  helper: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXXS,
  },
  bodyPad: {
    paddingTop: tokens.spacingVerticalM,
    paddingBottom: tokens.spacingVerticalL,
    paddingLeft: tokens.spacingHorizontalL,
    paddingRight: tokens.spacingHorizontalL,
  },
  card: {
    width: "100%",
    maxWidth: "520px",
    marginLeft: "auto",
    marginRight: "auto",
    padding: tokens.spacingHorizontalL,
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalL,
  },
  group: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalL },
  groupTight: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalM },
  textareaWrap: { position: "relative", display: "flex", flexDirection: "column" },
  fullWidth: { width: "100%" },
  field: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalXS },
  banner: {
    padding: tokens.spacingHorizontalM,
    borderTopLeftRadius: tokens.borderRadiusLarge,
    borderTopRightRadius: tokens.borderRadiusLarge,
    borderBottomLeftRadius: tokens.borderRadiusLarge,
    borderBottomRightRadius: tokens.borderRadiusLarge,
    ...thinBorderSides,
  },
  bannerHead: { display: "flex", alignItems: "center", gap: tokens.spacingHorizontalS },
  summaryBox: {
    borderTopLeftRadius: tokens.borderRadiusLarge,
    borderTopRightRadius: tokens.borderRadiusLarge,
    borderBottomLeftRadius: tokens.borderRadiusLarge,
    borderBottomRightRadius: tokens.borderRadiusLarge,
    ...thinBorderSides,
    ...allBorderColor(tokens.colorNeutralStroke2),
    backgroundColor: tokens.colorNeutralBackground2,
    padding: tokens.spacingHorizontalM,
  },
  divider: {
    marginTop: tokens.spacingVerticalM,
    paddingTop: tokens.spacingVerticalM,
    borderTopWidth: tokens.strokeWidthThin,
    borderTopStyle: "solid",
    borderTopColor: tokens.colorNeutralStroke2,
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
  },
  skeletonBox: {
    borderTopLeftRadius: tokens.borderRadiusLarge,
    borderTopRightRadius: tokens.borderRadiusLarge,
    borderBottomLeftRadius: tokens.borderRadiusLarge,
    borderBottomRightRadius: tokens.borderRadiusLarge,
    ...thinBorderSides,
    ...allBorderColor(tokens.colorNeutralStroke2),
    padding: tokens.spacingHorizontalM,
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
  },
  brandRow: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalS,
    color: tokens.colorBrandForeground1,
  },
  submitting: {
    paddingTop: tokens.spacingVerticalXXL,
    paddingBottom: tokens.spacingVerticalXXL,
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalL,
  },
  uppercase: {
    textTransform: "uppercase",
    letterSpacing: "0.04em",
    color: tokens.colorNeutralForeground3,
  },
});

export default function SubmitTab() {
  const styles = useStyles();

  // Auth is shared across both tabs (see TeamsAuthProvider) so switching tabs
  // never re-triggers the "Signing you in…" loader.
  const { authState, teamsToken, userName, authError, retry } = useTeamsAuth();

  // Restore any in-progress draft once on mount (survives tab switch / app close).
  const [draft] = useState(loadDraft);

  const [step, setStep] = useState<SubmitStep>(draft?.step ?? "input");

  const [inputText, setInputText] = useState(draft?.inputText ?? "");
  const [classifyResult, setClassifyResult] = useState<ClassifyResult | null>(
    draft?.classifyResult ?? null,
  );
  const [followUpAnswers, setFollowUpAnswers] = useState<string[]>(
    draft?.followUpAnswers ?? [],
  );
  const [submitError, setSubmitError] = useState<string>("");
  const [submittedCategory, setSubmittedCategory] = useState<string>(
    draft?.submittedCategory ?? "",
  );
  const [exampleIdx, setExampleIdx] = useState(0);
  const [prefetching, setPrefetching] = useState(false);

  const classifyTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const classifyCache = useRef<Map<string, ClassifyResult>>(new Map());
  const classifyInFlight = useRef<Map<string, Promise<ClassifyResult>>>(new Map());
  const mainInputRef = useRef<HTMLTextAreaElement>(null);

  // Persist the durable draft whenever it changes. Transient in-flight steps
  // (classifying/submitting) are skipped so a lost network call never restores
  // into a dead spinner — loadDraft() also normalises them on read.
  useEffect(() => {
    if (step === "classifying" || step === "submitting") return;
    try {
      const isEmpty =
        step === "input" && !inputText.trim() && !classifyResult && !submittedCategory;
      if (isEmpty) {
        localStorage.removeItem(DRAFT_KEY);
        return;
      }
      localStorage.setItem(
        DRAFT_KEY,
        JSON.stringify({ step, inputText, classifyResult, followUpAnswers, submittedCategory }),
      );
    } catch {
      /* storage unavailable — fall back to in-memory only */
    }
  }, [step, inputText, classifyResult, followUpAnswers, submittedCategory]);

  // No auto-focus on mount: tabs remount on every switch, so focusing here would
  // pop the mobile keyboard each time you toggle between Submit and Orders. The
  // user taps the input when they're ready to type.

  useEffect(() => {
    if (step !== "input" || inputText) return;
    const id = setInterval(() => setExampleIdx((i) => (i + 1) % EXAMPLES.length), 3500);
    return () => clearInterval(id);
  }, [step, inputText]);

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
      // Normalise the response so the stage-2 UI is always safe to render: an
      // unknown category would leave `meta` undefined (blank card) and a non-array
      // followUpQuestions would crash `.map(...)`.
      const result: ClassifyResult = {
        category: VALID_CATEGORIES.includes(data.category) ? data.category : "Other",
        listTarget: VALID_LIST_TARGETS.includes(data.listTarget) ? data.listTarget : "business-ideas",
        confidence: typeof data.confidence === "number" ? data.confidence : 0,
        reasoning: typeof data.reasoning === "string" ? data.reasoning : "",
        followUpQuestions: Array.isArray(data.followUpQuestions) ? data.followUpQuestions : [],
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
      // Only flag "reading" once the request is actually in flight (after the
      // pause), not during typing — otherwise the UI implies the AI is working
      // while the user is still writing.
      classifyTimer.current = setTimeout(() => {
        setPrefetching(true);
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
      <TeamsCenter className="animate-fade-in">
        <Spinner size="large" label="Signing you in automatically…" />
      </TeamsCenter>
    );
  }

  if (authState === "unauthenticated") {
    return (
      <TeamsFullScreen
        icon={<Shield24Regular />}
        title="Sign in required"
        description="Sign in with your Cranfield Glass Microsoft account to submit ideas and reports."
        actionLabel="Try again"
        onAction={retry}
        error={authError || undefined}
      />
    );
  }

  if (step === "done") {
    return (
      <TeamsPage>
        <TeamsScroll className={styles.bodyPad}>
        <Card className={`${styles.card} animate-fade-in-up`}>
          <div className={styles.group}>
            <div
              className={styles.banner}
              style={{
                backgroundColor: tokens.colorPaletteGreenBackground1,
                borderColor: tokens.colorPaletteGreenBorder2,
              }}
            >
              <div
                className={styles.bannerHead}
                style={{ color: tokens.colorPaletteGreenForeground1, marginBottom: tokens.spacingVerticalXS }}
              >
                <CheckmarkCircle24Filled />
                <Text weight="semibold" size={400} style={{ color: tokens.colorPaletteGreenForeground1 }}>
                  Submitted!
                </Text>
              </div>
              <Text size={200} style={{ color: tokens.colorNeutralForeground3, display: "block" }}>
                Your <strong>{submittedCategory}</strong> has been recorded.
              </Text>
            </div>

            <div className={styles.summaryBox}>
              <Text size={300} block>
                It will appear in the next H&amp;S meeting agenda.
              </Text>
            </div>

            <Button
              appearance="primary"
              size="large"
              className={styles.fullWidth}
              icon={<ArrowCounterclockwise20Regular />}
              onClick={reset}
            >
              Submit another
            </Button>
          </div>
        </Card>
        </TeamsScroll>
      </TeamsPage>
    );
  }

  // ─── Main form ────────────────────────────────────────────────────────────
  const meta = classifyResult ? CATEGORY_META[classifyResult.category] : null;

  // Input step mirrors Orders: a pinned, keyboard-safe textarea (suggestions
  // stay inside as the cycling placeholder) with the tagline + descriptor below.
  const inputEl = (
    <>
      <div className={styles.textareaWrap}>
        <Textarea
          textarea={{ ref: mainInputRef, rows: 5 }}
          placeholder={EXAMPLES[exampleIdx]}
          value={inputText}
          onChange={(_, data) => handleTextChange(data.value)}
          resize="none"
          size="large"
        />
      </div>

      <div className={styles.helper}>
        <Text size={300} weight="semibold" block>
          Small ideas. Continuous improvement.
        </Text>
        <Text size={200} block style={{ color: tokens.colorNeutralForeground3 }}>
          Describe what you saw, a near miss, or an improvement idea.
        </Text>
      </div>

      {submitError && (
        <MessageBar intent="error" className="animate-fade-in">
          <MessageBarBody>{submitError}</MessageBarBody>
        </MessageBar>
      )}

      {(() => {
        const tooShort = inputText.trim().length < 10;
        const isReady = classifyCache.current.has(inputText.trim());
        const icon = prefetching ? (
          <Spinner size="tiny" />
        ) : isReady ? (
          <Sparkle16Regular />
        ) : (
          <ChevronRight20Regular />
        );
        return (
          <Button
            appearance="primary"
            size="large"
            className={styles.fullWidth}
            disabled={tooShort || prefetching}
            icon={icon}
            onClick={handleContinue}
          >
            {prefetching ? "Reading your note…" : "Continue"}
          </Button>
        );
      })()}
    </>
  );

  // Every step (including "input") lives in the single scroll region. Android's
  // on-screen keyboard shrinks the visual viewport; if the input step were pinned
  // (non-scrolling) the Continue button below the textarea would sit under the
  // keyboard with no way to reach it. A scroll region lets the user scroll the
  // button into view, and the extra bottom padding keeps it clear of the keyboard.
  const cardEl = (
    <Card className={`${styles.card} animate-fade-in-up`}>
          {step === "classifying" && <ClassifyingSkeleton />}

          {step === "followup" && classifyResult && meta && (
            <div className={`${styles.group} animate-fade-in-up`}>
              <div
                className={styles.banner}
                style={{ backgroundColor: meta.bg, borderColor: meta.border }}
              >
                <div className={styles.bannerHead} style={{ color: meta.fg }}>
                  {meta.icon}
                  <Text weight="semibold" size={300} style={{ color: meta.fg }}>
                    {classifyResult.category}
                  </Text>
                  <Badge
                    appearance="tint"
                    color="informative"
                    shape="rounded"
                    style={{ marginLeft: "auto" }}
                  >
                    {Math.round(classifyResult.confidence * 100)}% sure
                  </Badge>
                </div>
                <Text size={200} style={{ color: tokens.colorNeutralForeground3, marginTop: tokens.spacingVerticalXS, display: "block" }}>
                  A couple of quick questions will help us record this accurately.
                </Text>
              </div>

              <div className={styles.groupTight}>
                {classifyResult.followUpQuestions.map((question, idx) => (
                  <Field
                    key={idx}
                    label={question}
                    className="animate-fade-in-up"
                    style={{ animationDelay: `${idx * 60}ms` }}
                  >
                    <Textarea
                      placeholder="Your answer…"
                      value={followUpAnswers[idx] || ""}
                      onChange={(_, data) => {
                        const updated = [...followUpAnswers];
                        updated[idx] = data.value;
                        setFollowUpAnswers(updated);
                      }}
                      textarea={{ rows: 2 }}
                      resize="none"
                    />
                  </Field>
                ))}
              </div>

              <Button
                appearance="primary"
                size="large"
                className={styles.fullWidth}
                icon={<ChevronRight20Regular />}
                onClick={() => setStep("confirm")}
              >
                Review &amp; Submit
              </Button>
              <Button appearance="subtle" className={styles.fullWidth} onClick={reset}>
                Start over
              </Button>
            </div>
          )}

          {step === "confirm" && classifyResult && meta && (
            <div className={`${styles.group} animate-fade-in-up`}>
              <div
                className={styles.banner}
                style={{ backgroundColor: meta.bg, borderColor: meta.border }}
              >
                <div className={styles.bannerHead} style={{ color: meta.fg, marginBottom: tokens.spacingVerticalXS }}>
                  {meta.icon}
                  <Text weight="semibold" size={300} style={{ color: meta.fg }}>
                    {classifyResult.category}
                  </Text>
                </div>
                <Text size={200} style={{ color: tokens.colorNeutralForeground3, display: "block" }}>
                  → Goes to: <strong>{meta.listLabel}</strong>
                </Text>
              </div>

              <div className={styles.summaryBox}>
                <Text size={200} weight="semibold" className={styles.uppercase} block>
                  Your submission
                </Text>
                <Text size={300} block style={{ whiteSpace: "pre-wrap", marginTop: tokens.spacingVerticalS }}>
                  {inputText}
                </Text>
                {followUpAnswers.some((a) => a.trim()) && (
                  <div className={styles.divider}>
                    {classifyResult.followUpQuestions.map((q, i) =>
                      followUpAnswers[i]?.trim() ? (
                        <div key={i}>
                          <Text size={200} block style={{ color: tokens.colorNeutralForeground3 }}>
                            {q}
                          </Text>
                          <Text size={300} block>
                            {followUpAnswers[i]}
                          </Text>
                        </div>
                      ) : null
                    )}
                  </div>
                )}
              </div>

              {submitError && (
                <MessageBar intent="error" className="animate-fade-in">
                  <MessageBarBody>{submitError}</MessageBarBody>
                </MessageBar>
              )}

              <Button
                appearance="primary"
                size="large"
                className={styles.fullWidth}
                icon={<Send20Regular />}
                onClick={handleSubmit}
              >
                Submit
              </Button>
              <Button
                appearance="subtle"
                className={styles.fullWidth}
                onClick={() =>
                  classifyResult.followUpQuestions.length > 0 ? setStep("followup") : setStep("input")
                }
              >
                Go back
              </Button>
            </div>
          )}

          {step === "submitting" && (
            <div className={`${styles.submitting} animate-fade-in`}>
              <ProgressBar color="success" thickness="large" />
              <Text size={300} align="center" style={{ color: tokens.colorNeutralForeground3 }}>
                Saving your submission…
              </Text>
            </div>
          )}
    </Card>
  );

  return (
    <TeamsPage>
      {step === "input" ? (
        <TeamsScroll className={styles.inputStack}>{inputEl}</TeamsScroll>
      ) : (
        <TeamsScroll className={styles.bodyPad}>{cardEl}</TeamsScroll>
      )}
    </TeamsPage>
  );
}

function ClassifyingSkeleton() {
  const styles = useStyles();
  return (
    <div className={`${styles.group} animate-fade-in`}>
      <div className={styles.brandRow}>
        <Sparkle16Regular className="animate-pulse" />
        <Text size={300} weight="semibold">
          Reading your submission…
        </Text>
      </div>
      <ProgressBar />
      <div className={styles.skeletonBox}>
        <Skeleton>
          <SkeletonItem style={{ width: "8rem", height: "1.25rem" }} />
        </Skeleton>
        <Skeleton>
          <SkeletonItem style={{ height: "0.75rem" }} />
        </Skeleton>
        <Skeleton>
          <SkeletonItem style={{ width: "75%", height: "0.75rem" }} />
        </Skeleton>
      </div>
      <div className={styles.field}>
        <Skeleton>
          <SkeletonItem style={{ width: "10rem", height: "0.75rem" }} />
        </Skeleton>
        <Skeleton>
          <SkeletonItem style={{ height: "2.25rem" }} />
        </Skeleton>
      </div>
    </div>
  );
}

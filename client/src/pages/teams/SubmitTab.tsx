import { useState, useEffect, useRef } from "react";
import * as microsoftTeams from "@microsoft/teams-js";
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
  CheckmarkCircle48Filled,
  Send20Regular,
  ChevronRight20Regular,
  ArrowCounterclockwise20Regular,
  Sparkle16Regular,
} from "@fluentui/react-icons";
import {
  TeamsPage,
  TeamsPinned,
  TeamsScroll,
  TeamsCenter,
  TeamsFullScreen,
  useKeyboardSafeFocus,
} from "./TeamsPageShell";

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

function decodeJwtPayload(token: string): Record<string, any> {
  try {
    const payload = token.split(".")[1];
    const decoded = atob(payload.replace(/-/g, "+").replace(/_/g, "/"));
    return JSON.parse(decoded);
  } catch {
    return {};
  }
}

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
  centered: { width: "100%", maxWidth: "340px", textAlign: "center" },
  successChip: {
    width: "72px",
    height: "72px",
    borderRadius: tokens.borderRadiusCircular,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    marginLeft: "auto",
    marginRight: "auto",
    marginBottom: tokens.spacingVerticalL,
    backgroundColor: tokens.colorPaletteGreenBackground1,
    color: tokens.colorPaletteGreenForeground1,
  },
  inputStack: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
    paddingLeft: tokens.spacingHorizontalL,
    paddingRight: tokens.spacingHorizontalL,
    paddingTop: tokens.spacingVerticalM,
    paddingBottom: tokens.spacingVerticalL,
  },
  helper: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXXS,
  },
  bodyPad: {
    paddingTop: tokens.spacingVerticalXS,
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
  prefetch: {
    position: "absolute",
    bottom: tokens.spacingVerticalS,
    right: tokens.spacingHorizontalM,
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalXS,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorBrandForeground1,
  },
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

export default function SubmitTab({ onUser }: { onUser?: (name: string) => void } = {}) {
  const styles = useStyles();

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
  const mainInputRef = useRef<HTMLTextAreaElement>(null);

  useEffect(() => {
    initAuth();
  }, []);

  // Keyboard-safe auto-focus (see TeamsPageShell): focus without the page jumping.
  useKeyboardSafeFocus(mainInputRef, authState === "authenticated" && step === "input");

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
      const name = payload.name || payload.preferred_username || payload.upn || "";
      setUserName(name);
      onUser?.(name);
      setTeamsToken(ssoToken);
      setAuthState("authenticated");
    } catch (err: any) {
      const msg = `Teams sign-in failed: ${err?.message || String(err)}`;
      console.error(msg, err);
      setAuthError(msg);
      setAuthState("unauthenticated");
      onUser?.("");
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
        onAction={() => initAuth()}
        error={authError || undefined}
      />
    );
  }

  if (step === "done") {
    return (
      <TeamsCenter className="animate-fade-in">
        <div className={styles.centered}>
          <div className={`${styles.successChip} animate-pop-in`}>
            <CheckmarkCircle48Filled />
          </div>
          <Text as="h1" size={600} weight="bold" block className="animate-fade-in-up">
            Submitted!
          </Text>
          <Text
            size={300}
            block
            className="animate-fade-in-up"
            style={{ marginTop: tokens.spacingVerticalXS }}
          >
            Your <strong>{submittedCategory}</strong> has been recorded.
          </Text>
          <Text
            size={200}
            block
            className="animate-fade-in-up"
            style={{ color: tokens.colorNeutralForeground3, marginTop: tokens.spacingVerticalXS, marginBottom: tokens.spacingVerticalXXL }}
          >
            It will appear in the next H&amp;S meeting agenda.
          </Text>
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
      </TeamsCenter>
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
        {prefetching && (
          <div className={`${styles.prefetch} animate-fade-in`}>
            <Sparkle16Regular className="animate-pulse" />
            reading…
          </div>
        )}
      </div>

      <div className={styles.helper}>
        <Text size={300} weight="semibold" block>
          Small ideas. Continuous improvement.
        </Text>
        <Text size={200} block style={{ color: tokens.colorNeutralForeground3 }}>
          Describe what you saw, a near miss, or an improvement idea.
        </Text>
      </div>

      {inputText.trim().length >= 10 && inputText.trim().length < 20 && (
        <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
          Keep typing — a bit more detail helps…
        </Text>
      )}

      {submitError && (
        <MessageBar intent="error" className="animate-fade-in">
          <MessageBarBody>{submitError}</MessageBarBody>
        </MessageBar>
      )}

      <Button
        appearance="primary"
        size="large"
        className={styles.fullWidth}
        disabled={inputText.trim().length < 10}
        icon={
          classifyCache.current.has(inputText.trim()) ? (
            <Sparkle16Regular />
          ) : (
            <ChevronRight20Regular />
          )
        }
        onClick={handleContinue}
      >
        {classifyCache.current.has(inputText.trim()) ? "Ready — Continue" : "Continue"}
      </Button>
    </>
  );

  // The card sits in a pinned region during the "input" step so the focused
  // textarea has no scrollable ancestor (keyboard-safe). Longer follow-up /
  // confirm steps move into the single scroll region.
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
        <TeamsPinned className={styles.inputStack}>{inputEl}</TeamsPinned>
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

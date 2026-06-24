import { useState, useRef, useCallback, useEffect } from "react";
import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import {
  Button,
  Text,
  Spinner,
  Badge,
  MessageBar,
  MessageBarBody,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import {
  Signature24Regular,
  Shield24Regular,
  Edit20Regular,
  Desktop20Regular,
  Dismiss20Regular,
  ArrowLeft20Regular,
  CheckmarkCircle24Filled,
  CalendarLtr24Regular,
  ArrowCounterclockwise20Regular,
  DocumentText24Regular,
  DocumentText20Regular,
} from "@fluentui/react-icons";
import {
  TeamsPage,
  TeamsPinned,
  TeamsScroll,
  TeamsCenter,
  TeamsFullScreen,
} from "./TeamsPageShell";
import { useTeamsAuth } from "@/hooks/useTeamsAuth";
import { SectionHeader, MeetingCard, HeroCard } from "./MeetingCards";

type SignatureStatus = "signed" | "remote" | "absent";

interface SignatureRecord {
  status: SignatureStatus;
  signatureData: string | null;
  signedAt: string;
}

type MeetingState = "open" | "upcoming" | "locked";

interface SignMeeting {
  meetingDate: string;
  dateKey: string;
  displayDate: string;
  isPresent: boolean;
  mySignature: SignatureRecord | null;
  state: MeetingState;
}

interface MeetingsResponse {
  success: boolean;
  matched: boolean;
  userName: string;
  attendeeName: string | null;
  role?: string;
  meetings: SignMeeting[];
  error?: string;
}

interface ReadableMeeting {
  meetingDate: string;
  dateKey: string;
  displayDate: string;
}

interface MinutesListResponse {
  success: boolean;
  userName: string;
  meetings: ReadableMeeting[];
  error?: string;
}

interface MinutesHtmlResponse {
  success: boolean;
  displayDate: string;
  html: string;
  error?: string;
}

const useStyles = makeStyles({
  intro: {
    paddingLeft: tokens.spacingHorizontalL,
    paddingRight: tokens.spacingHorizontalL,
    paddingTop: tokens.spacingVerticalM,
    paddingBottom: tokens.spacingVerticalS,
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXXS,
  },
  list: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalL,
    paddingLeft: tokens.spacingHorizontalL,
    paddingRight: tokens.spacingHorizontalL,
    paddingTop: tokens.spacingVerticalS,
    paddingBottom: tokens.spacingVerticalXXL,
  },
  section: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
  },
  // Stacks the two tier-1 heroes (signing + agenda) a touch tighter than the
  // gap between full sections so they read as one priority zone.
  heroStack: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
  },
  signPanel: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalL,
    paddingLeft: tokens.spacingHorizontalL,
    paddingRight: tokens.spacingHorizontalL,
    paddingTop: tokens.spacingVerticalM,
    paddingBottom: tokens.spacingVerticalXXL,
    maxWidth: "520px",
    marginLeft: "auto",
    marginRight: "auto",
    width: "100%",
    boxSizing: "border-box",
  },
  signerBox: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalS,
    padding: tokens.spacingHorizontalM,
    borderRadius: tokens.borderRadiusLarge,
    backgroundColor: tokens.colorNeutralBackground2,
  },
  choiceStack: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalM },
  bigBtn: { height: "52px", justifyContent: "flex-start" },
  canvasWrap: {
    position: "relative",
    borderRadius: tokens.borderRadiusLarge,
    borderTopWidth: tokens.strokeWidthThick,
    borderRightWidth: tokens.strokeWidthThick,
    borderBottomWidth: tokens.strokeWidthThick,
    borderLeftWidth: tokens.strokeWidthThick,
    borderTopStyle: "dashed",
    borderRightStyle: "dashed",
    borderBottomStyle: "dashed",
    borderLeftStyle: "dashed",
    borderTopColor: tokens.colorBrandStroke2,
    borderRightColor: tokens.colorBrandStroke2,
    borderBottomColor: tokens.colorBrandStroke2,
    borderLeftColor: tokens.colorBrandStroke2,
    backgroundColor: "#ffffff",
    overflow: "hidden",
  },
  canvasHint: {
    position: "absolute",
    inset: 0,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    pointerEvents: "none",
    color: tokens.colorNeutralForeground4,
  },
  drawActions: { display: "flex", gap: tokens.spacingHorizontalM },
  doneWrap: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: tokens.spacingVerticalM,
    textAlign: "center",
    maxWidth: "340px",
  },
  minutesBar: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalS,
    paddingLeft: tokens.spacingHorizontalS,
    paddingRight: tokens.spacingHorizontalL,
    paddingTop: tokens.spacingVerticalS,
    paddingBottom: tokens.spacingVerticalS,
    borderBottomWidth: tokens.strokeWidthThin,
    borderBottomStyle: "solid",
    borderBottomColor: tokens.colorNeutralStroke2,
  },
  // Scroll container that hosts the minutes Shadow DOM. Rendering the minutes
  // inline (in a shadow root) rather than in a srcDoc iframe avoids the iOS Teams
  // WebKit bug where a srcDoc iframe stays blank after an auth-triggered WebView
  // resume, while still keeping the minutes' styling isolated from the tab chrome.
  minutesHost: {
    flexGrow: 1,
    minHeight: 0,
    width: "100%",
    overflowY: "auto",
  },
  sigPreview: {
    maxHeight: "56px",
    maxWidth: "100%",
    borderRadius: tokens.borderRadiusMedium,
    borderTopWidth: tokens.strokeWidthThin,
    borderRightWidth: tokens.strokeWidthThin,
    borderBottomWidth: tokens.strokeWidthThin,
    borderLeftWidth: tokens.strokeWidthThin,
    borderTopStyle: "solid",
    borderRightStyle: "solid",
    borderBottomStyle: "solid",
    borderLeftStyle: "solid",
    borderTopColor: tokens.colorNeutralStroke2,
    borderRightColor: tokens.colorNeutralStroke2,
    borderBottomColor: tokens.colorNeutralStroke2,
    borderLeftColor: tokens.colorNeutralStroke2,
  },
  // Hero-style status surface for read-only (locked / upcoming) meetings, mirroring
  // the agenda HeroCard's brand-tint look so the whole tab reads as one family.
  statusHero: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
    paddingTop: tokens.spacingVerticalL,
    paddingBottom: tokens.spacingVerticalL,
    paddingLeft: tokens.spacingHorizontalL,
    paddingRight: tokens.spacingHorizontalL,
    borderRadius: tokens.borderRadiusXLarge,
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorNeutralForeground1,
  },
  statusTop: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalM,
  },
  statusChip: {
    width: "48px",
    height: "48px",
    flexShrink: 0,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    borderRadius: tokens.borderRadiusCircular,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  statusHeadings: { minWidth: 0, display: "flex", flexDirection: "column", gap: "2px" },
  statusEyebrow: {
    textTransform: "uppercase",
    letterSpacing: "0.06em",
    color: tokens.colorBrandForeground1,
  },
  // Signature sits on its own light panel inside the hero so it stays legible on
  // the brand tint and reads as "here's what you signed".
  sigBlock: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXS,
    paddingTop: tokens.spacingVerticalM,
    paddingBottom: tokens.spacingVerticalM,
    paddingLeft: tokens.spacingHorizontalM,
    paddingRight: tokens.spacingHorizontalM,
    borderRadius: tokens.borderRadiusLarge,
    backgroundColor: tokens.colorNeutralBackground1,
  },
});

function statusBadge(status: SignatureStatus) {
  if (status === "signed") return <Badge appearance="tint" color="success">Signed</Badge>;
  if (status === "remote") return <Badge appearance="tint" color="brand">Remote</Badge>;
  return <Badge appearance="tint" color="informative">Marked absent</Badge>;
}

// Friendly relative label for a YYYY-MM-DD date key, computed in local time so
// "Today"/"Tomorrow" match the user's calendar rather than UTC midnight.
function relativeFromToday(dateKey: string): string {
  const [y, m, d] = dateKey.split("-").map(Number);
  if (!y || !m || !d) return "";
  const target = new Date(y, m - 1, d);
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const diff = Math.round((target.getTime() - today.getTime()) / 86400000);
  if (diff === 0) return "Today";
  if (diff === 1) return "Tomorrow";
  if (diff > 1) return `In ${diff} days`;
  if (diff === -1) return "Yesterday";
  return `${Math.abs(diff)} days ago`;
}

function generateRemoteSignatureImage(name: string, date: string): string {
  const canvas = document.createElement("canvas");
  canvas.width = 320;
  canvas.height = 100;
  const ctx = canvas.getContext("2d")!;
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  ctx.font = 'italic 32px Georgia, "Times New Roman", serif';
  ctx.fillStyle = "#1f2937";
  ctx.textBaseline = "middle";
  ctx.fillText(name, 10, 36);
  ctx.font = "11px Arial, sans-serif";
  ctx.fillStyle = "#6b7280";
  ctx.textBaseline = "middle";
  ctx.fillText("Attended Remotely", 10, 66);
  ctx.fillText(date, 10, 84);
  return canvas.toDataURL("image/png");
}

const SELECTED_KEY_STORAGE = "teams-sign-selected-datekey";

function readSelectedKey(): string | null {
  try {
    return localStorage.getItem(SELECTED_KEY_STORAGE);
  } catch {
    return null;
  }
}

function persistSelectedKey(key: string | null) {
  try {
    if (key) localStorage.setItem(SELECTED_KEY_STORAGE, key);
    else localStorage.removeItem(SELECTED_KEY_STORAGE);
  } catch {
    /* ignore storage failures */
  }
}

// Defensive sanitizer for the server-rendered minutes before they are placed in
// the LIVE document (Shadow DOM). The old srcDoc iframe had a `sandbox` without
// allow-scripts that neutralised any unsafe HTML; rendering inline removes that
// barrier, so we re-create it here. The input is an INERT DOMParser document, so
// nothing has executed or loaded yet — we strip scripts, inline event handlers,
// and unsafe URL schemes, then the caller imports the cleaned nodes via DOM APIs
// (never raw innerHTML), so no malformed stored field (e.g. a crafted signature
// data URL) can become active markup.
const SANITIZE_SAFE_URL = /^(https:|data:image\/|mailto:|#|\/)/i;
function sanitizeMinutesBody(body: HTMLElement) {
  // Remove element types that can execute or load arbitrary content.
  body
    .querySelectorAll("script, iframe, object, embed, base, link, meta, form, style, svg")
    .forEach((el) => el.remove());
  body.querySelectorAll("*").forEach((el) => {
    for (const attr of Array.from(el.attributes)) {
      const name = attr.name.toLowerCase();
      if (name.startsWith("on")) {
        el.removeAttribute(attr.name); // strip event handlers (onerror, onload, …)
        continue;
      }
      if ((name === "src" || name === "href" || name === "xlink:href" || name === "srcset") &&
          !SANITIZE_SAFE_URL.test(attr.value.trim())) {
        el.removeAttribute(attr.name); // drop javascript:/unknown-scheme URLs
      }
    }
  });
}

export default function SignTab() {
  const styles = useStyles();
  const qc = useQueryClient();
  const { authState, teamsToken, userName, authError, retry, getToken } = useTeamsAuth();

  // Teams tabs remount on switch, wiping local state. Persist only the chosen
  // meeting (by stable dateKey) so a tab switch mid-flow returns the user to the
  // same meeting; the transient drawing step is intentionally NOT persisted.
  const [selectedKey, setSelectedKey] = useState<string | null>(() => readSelectedKey());
  const [mode, setMode] = useState<"choose" | "draw">("choose");
  const [hasDrawn, setHasDrawn] = useState(false);
  const [actionError, setActionError] = useState<string | null>(null);
  const [justSigned, setJustSigned] = useState<SignatureStatus | null>(null);
  // Transient — tabs remount on switch, so we intentionally don't persist this.
  // The read-only document viewer is shared by past-meeting minutes and the
  // upcoming-meeting agenda — both are the same minutes-style HTML rendered in the
  // same Shadow DOM host; only the endpoint and a few labels differ by `kind`.
  const [minutesView, setMinutesView] = useState<
    { dateKey: string; displayDate: string; kind: "minutes" | "agenda" } | null
  >(null);
  const minutesHostRef = useRef<HTMLDivElement>(null);

  const canvasRef = useRef<HTMLCanvasElement>(null);
  const isDrawing = useRef(false);
  const lastPoint = useRef<{ x: number; y: number } | null>(null);

  const meetingsQuery = useQuery<MeetingsResponse>({
    queryKey: ["/api/teams/sign/meetings", teamsToken ? "token" : "none"],
    enabled: authState === "authenticated",
    queryFn: async () => {
      const token = await getToken();
      const res = await fetch("/api/teams/sign/meetings", {
        headers: { Authorization: `Bearer ${token}` },
      });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "Failed to load meetings");
      return json as MeetingsResponse;
    },
  });

  const signMutation = useMutation({
    mutationFn: async (vars: { meetingDate: string; status: SignatureStatus; signatureData: string | null }) => {
      const token = await getToken();
      const res = await fetch("/api/teams/sign", {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}` },
        body: JSON.stringify({
          meetingDate: vars.meetingDate,
          status: vars.status,
          signatureData: vars.signatureData,
          signedAt: new Date().toISOString(),
        }),
      });
      const json = await res.json();
      if (!res.ok || !json.success) throw new Error(json.error || "Failed to save your signature");
      return json;
    },
    onSuccess: (_data, vars) => {
      setJustSigned(vars.status);
      qc.invalidateQueries({ queryKey: ["/api/teams/sign/meetings"] });
    },
    onError: (err: any) => setActionError(err?.message || "Failed to save your signature"),
  });

  // Whether the signed-in user is NOT on the attendee roster. Such staff (e.g. a
  // remote worker) can't self-sign, but they CAN still read locked minutes, so
  // we fetch the readable-minutes list for them instead of a dead-end message.
  const isUnmatched = !!meetingsQuery.data && !meetingsQuery.data.matched;

  // Readable locked/closed meetings. ANY signed-in Cranfield Glass staff member
  // can read locked minutes — including rostered staff who were absent from a
  // meeting — so this loads for every authenticated user, not just non-roster ones.
  const minutesListQuery = useQuery<MinutesListResponse>({
    queryKey: ["/api/teams/minutes/meetings", teamsToken ? "token" : "none"],
    enabled: authState === "authenticated",
    queryFn: async () => {
      const token = await getToken();
      const res = await fetch("/api/teams/minutes/meetings", {
        headers: { Authorization: `Bearer ${token}` },
      });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "Failed to load meetings");
      return json as MinutesListResponse;
    },
  });

  // The rendered minutes HTML for the meeting currently being read.
  const minutesQuery = useQuery<MinutesHtmlResponse>({
    queryKey: ["/api/teams/doc", minutesView?.kind ?? "none", minutesView?.dateKey ?? "none"],
    enabled: !!minutesView,
    queryFn: async () => {
      const token = await getToken();
      const endpoint = minutesView!.kind === "agenda" ? "agenda" : "minutes";
      const res = await fetch(`/api/teams/${endpoint}/${minutesView!.dateKey}`, {
        headers: { Authorization: `Bearer ${token}` },
      });
      const json = await res.json();
      if (!res.ok || !json.success) throw new Error(json.error || "Failed to load this document");
      return json as MinutesHtmlResponse;
    },
  });

  // Derive the selected meeting from fresh query data so its signature status is
  // never stale, even after a remount restores the persisted dateKey.
  const meetings = meetingsQuery.data?.meetings ?? [];
  const selected = selectedKey ? meetings.find((m) => m.dateKey === selectedKey) ?? null : null;

  // Group for display: open meetings need action, locked ones are read-only
  // history, and upcoming ones surface the next meeting date. Open meetings are
  // sorted newest-first so the tier-1 hero is deterministically the most recent
  // meeting awaiting signature regardless of the order the backend returns them.
  const openMeetings = meetings
    .filter((m) => m.state === "open")
    .sort((a, b) => new Date(b.meetingDate).getTime() - new Date(a.meetingDate).getTime());
  const pastMeetings = meetings.filter((m) => m.state === "locked");
  const upcomingMeetings = meetings
    .filter((m) => m.state === "upcoming")
    .sort((a, b) => new Date(a.meetingDate).getTime() - new Date(b.meetingDate).getTime());
  const nextMeeting = upcomingMeetings[0] ?? null;

  // Split open meetings by whether the user has actually signed yet. Unsigned
  // ones are the urgent action that drives the tier-1 hero; already-signed open
  // meetings just await an admin lock, so they sit quietly in the attendance tier.
  const unsignedOpen = openMeetings.filter((m) => !m.mySignature);
  const signedOpen = openMeetings.filter((m) => m.mySignature);
  // The hero takes the most recent unsigned meeting; any extras list below it.
  const heroSignMeeting = unsignedOpen[0] ?? null;
  const extraUnsignedOpen = unsignedOpen.slice(1);

  const clearCanvas = useCallback(() => {
    const canvas = canvasRef.current;
    const ctx = canvas?.getContext("2d");
    if (canvas && ctx) ctx.clearRect(0, 0, canvas.width, canvas.height);
    setHasDrawn(false);
  }, []);

  useEffect(() => {
    if (mode === "draw") clearCanvas();
  }, [mode, clearCanvas]);

  const getPos = (e: { clientX: number; clientY: number }, canvas: HTMLCanvasElement) => {
    const rect = canvas.getBoundingClientRect();
    return {
      x: (e.clientX - rect.left) * (canvas.width / rect.width),
      y: (e.clientY - rect.top) * (canvas.height / rect.height),
    };
  };

  const startDraw = (point: { x: number; y: number }) => {
    isDrawing.current = true;
    lastPoint.current = point;
  };

  const moveDraw = (point: { x: number; y: number }) => {
    if (!isDrawing.current) return;
    const ctx = canvasRef.current?.getContext("2d");
    if (!ctx || !lastPoint.current) {
      lastPoint.current = point;
      return;
    }
    ctx.beginPath();
    ctx.moveTo(lastPoint.current.x, lastPoint.current.y);
    ctx.lineTo(point.x, point.y);
    ctx.strokeStyle = "#1f2937";
    ctx.lineWidth = 2.5;
    ctx.lineCap = "round";
    ctx.lineJoin = "round";
    ctx.stroke();
    setHasDrawn(true);
    lastPoint.current = point;
  };

  const endDraw = () => {
    isDrawing.current = false;
    lastPoint.current = null;
  };

  function openMeeting(meeting: SignMeeting) {
    setSelectedKey(meeting.dateKey);
    persistSelectedKey(meeting.dateKey);
    setMode("choose");
    setHasDrawn(false);
    setActionError(null);
    setJustSigned(null);
  }

  function backToList() {
    setSelectedKey(null);
    persistSelectedKey(null);
    setMode("choose");
    setActionError(null);
    setJustSigned(null);
  }

  function openMinutes(dateKey: string, displayDate: string) {
    setMinutesView({ dateKey, displayDate, kind: "minutes" });
  }

  function openAgenda(dateKey: string, displayDate: string) {
    setMinutesView({ dateKey, displayDate, kind: "agenda" });
  }

  function closeMinutes() {
    setMinutesView(null);
  }

  // Render the minutes inline via Shadow DOM instead of a srcDoc iframe. The
  // server returns a full HTML document (its own <style> + body); we lift those
  // into a shadow root so the minutes styling stays isolated from the Fluent/Teams
  // chrome (and vice-versa) exactly like the old iframe did — but as an ordinary
  // element in the same document. This avoids the iOS Teams WebKit bug where a
  // srcDoc iframe stays blank after an auth-triggered WebView resume, because
  // there is no separate iframe compositor that can get stuck un-painted.
  useEffect(() => {
    const host = minutesHostRef.current;
    const html = minutesQuery.data?.html;
    if (!host || !minutesView || !html) return;

    // attachShadow can only be called once per element; reuse it on re-render.
    const shadow = host.shadowRoot ?? host.attachShadow({ mode: "open" });

    // DOMParser produces an INERT document: nothing executes and no resources
    // load until we move nodes into the live tree, so we can safely collect the
    // document styles, sanitise the body, and only then import the cleaned nodes.
    const doc = new DOMParser().parseFromString(html, "text/html");
    // Only lift the trusted document <head> stylesheet — never <style> tags that
    // might appear inside the (untrusted) body content.
    const docStyleText = Array.from(doc.head?.querySelectorAll("style") ?? [])
      .map((s) => s.textContent ?? "")
      .join("\n");

    // The document's own CSS targets `body` for its page background/typography;
    // inside a shadow root that maps onto `:host`, so we restate those essentials
    // there. The document's <style> still scopes every other rule to this subtree.
    // textContent (not innerHTML) keeps this a pure stylesheet, never markup.
    const styleEl = document.createElement("style");
    styleEl.textContent =
      ":host{display:block;width:100%;min-height:100%;" +
      "background:#f5f5f5;-webkit-text-size-adjust:100%;" +
      'font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;' +
      "color:#242424;font-size:15px;line-height:1.5;}\n" +
      docStyleText;

    shadow.replaceChildren(styleEl);
    if (doc.body) {
      sanitizeMinutesBody(doc.body);
      // Import the now-sanitised nodes into the live shadow tree via DOM APIs
      // (not innerHTML), so no string is ever re-parsed in the live context.
      for (const node of Array.from(doc.body.childNodes)) {
        shadow.appendChild(document.importNode(node, true));
      }
    }

    return () => {
      if (host.shadowRoot) host.shadowRoot.replaceChildren();
    };
  }, [minutesView, minutesQuery.data?.html]);

  function submit(status: SignatureStatus, signatureData: string | null) {
    if (!selected) return;
    setActionError(null);
    signMutation.mutate({ meetingDate: selected.meetingDate, status, signatureData });
  }

  function confirmDrawn() {
    const canvas = canvasRef.current;
    if (!canvas || !hasDrawn) return;
    submit("signed", canvas.toDataURL("image/png"));
  }

  // ─── Auth gates ────────────────────────────────────────────────────────────
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
        description="Sign in with your Cranfield Glass Microsoft account to sign meeting minutes."
        actionLabel="Try again"
        onAction={retry}
        error={authError || undefined}
      />
    );
  }

  // ─── Read-only meeting minutes viewer ───────────────────────────────────────
  // The minutes HTML is injected into a Shadow DOM root (see the effect above)
  // rather than a srcDoc iframe. Shadow DOM keeps the minutes' own styling
  // isolated from the Teams tab chrome while rendering as a normal in-document
  // element, which paints reliably on iOS Teams mobile — a srcDoc iframe there
  // could stay blank after an auth-triggered WebView resume. The server-rendered
  // minutes are pure, escaped HTML/CSS with no scripts, so nothing executes.
  if (minutesView) {
    const isAgendaView = minutesView.kind === "agenda";
    const docNoun = isAgendaView ? "agenda" : "minutes";
    return (
      <TeamsPage>
        <TeamsPinned className={styles.minutesBar}>
          <Button appearance="subtle" icon={<ArrowLeft20Regular />} onClick={closeMinutes}>
            Back
          </Button>
          <Text size={300} weight="semibold" truncate style={{ minWidth: 0 }}>
            {minutesView.displayDate}
          </Text>
        </TeamsPinned>
        {minutesQuery.isLoading ? (
          <TeamsCenter className="animate-fade-in">
            <Spinner size="large" label={`Loading meeting ${docNoun}…`} />
          </TeamsCenter>
        ) : minutesQuery.isError ? (
          <TeamsCenter className="animate-fade-in">
            <div className={styles.doneWrap}>
              <DocumentText24Regular style={{ fontSize: "48px", color: tokens.colorNeutralForeground4 }} />
              <Text size={400} weight="semibold" style={{ textAlign: "center" }}>
                {isAgendaView ? "Couldn't load the agenda" : "Couldn't load these minutes"}
              </Text>
              <Text size={300} style={{ color: tokens.colorNeutralForeground3, textAlign: "center" }}>
                {(minutesQuery.error as Error)?.message || "Something went wrong. Please try again."}
              </Text>
              <Button appearance="primary" onClick={() => minutesQuery.refetch()} style={{ width: "100%" }}>
                Try again
              </Button>
            </div>
          </TeamsCenter>
        ) : (
          <div
            ref={minutesHostRef}
            className={styles.minutesHost}
            aria-label={`Meeting ${docNoun} ${minutesView.displayDate}`}
          />
        )}
      </TeamsPage>
    );
  }

  // ─── Sign panel for a selected meeting ───────────────────────────────────────
  if (selected) {
    if (justSigned) {
      const label =
        justSigned === "signed"
          ? "Signature saved"
          : justSigned === "remote"
          ? "Marked as attended remotely"
          : "Marked as absent";
      return (
        <TeamsPage>
          <TeamsCenter className="animate-fade-in">
            <div className={`${styles.doneWrap} animate-scale-in`}>
              <CheckmarkCircle24Filled
                style={{ fontSize: "56px", color: tokens.colorPaletteGreenForeground1 }}
              />
              <Text as="h1" size={500} weight="bold">{label}</Text>
              <Text size={300} style={{ color: tokens.colorNeutralForeground3 }}>
                Recorded for {selected.displayDate}. The admin will see this on the meeting minutes.
              </Text>
              <Button appearance="primary" size="large" onClick={backToList} style={{ width: "100%" }}>
                Back to meetings
              </Button>
            </div>
          </TeamsCenter>
        </TeamsPage>
      );
    }

    const existing = selected.mySignature;

    // Upcoming or admin-locked meetings are read-only: there's nothing to sign,
    // so show the status (or that it hasn't happened yet) instead of the pad.
    if (selected.state !== "open") {
      const isUpcoming = selected.state === "upcoming";
      // Locked meetings must reflect HOW the person attended, not just that a
      // record exists: an "absent"/"remote" record should never read as "signed".
      const statusHero = (() => {
        if (isUpcoming) {
          return { icon: <CalendarLtr24Regular style={{ color: tokens.colorBrandForeground1 }} />, eyebrow: "Upcoming meeting" };
        }
        if (existing?.status === "signed") {
          return { icon: <CheckmarkCircle24Filled style={{ color: tokens.colorPaletteGreenForeground1 }} />, eyebrow: "Attendance signed" };
        }
        if (existing?.status === "remote") {
          return { icon: <Desktop20Regular style={{ color: tokens.colorBrandForeground1 }} />, eyebrow: "Attended remotely" };
        }
        if (existing?.status === "absent") {
          return { icon: <Dismiss20Regular style={{ color: tokens.colorNeutralForeground3 }} />, eyebrow: "Marked absent" };
        }
        return { icon: <Shield24Regular style={{ color: tokens.colorBrandForeground1 }} />, eyebrow: "Meeting locked" };
      })();
      return (
        <TeamsPage>
          <TeamsScroll>
            <div className={`${styles.signPanel} animate-fade-in`}>
              <Button
                appearance="subtle"
                icon={<ArrowLeft20Regular />}
                onClick={backToList}
                style={{ alignSelf: "flex-start" }}
              >
                All meetings
              </Button>
              <div className={`${styles.statusHero} animate-fade-in`}>
                <div className={styles.statusTop}>
                  <div className={styles.statusChip}>
                    {statusHero.icon}
                  </div>
                  <div className={styles.statusHeadings}>
                    <Text size={200} weight="semibold" className={styles.statusEyebrow}>
                      {statusHero.eyebrow}
                    </Text>
                    <Text size={500} weight="bold" truncate block>{selected.displayDate}</Text>
                    <Text size={200} block style={{ color: tokens.colorNeutralForeground3 }}>
                      {`${relativeFromToday(selected.dateKey)} · ${
                        isUpcoming ? "sign once it has taken place" : "minutes ready below"
                      }`}
                    </Text>
                  </div>
                </div>

                {existing?.signatureData && (
                  <div className={styles.sigBlock}>
                    <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
                      Your signature
                    </Text>
                    <img src={existing.signatureData} alt="Your signature" className={styles.sigPreview} />
                  </div>
                )}

                {!isUpcoming && (
                  <Button
                    appearance="primary"
                    icon={<DocumentText20Regular />}
                    onClick={() => openMinutes(selected.dateKey, selected.displayDate)}
                    style={{ width: "100%" }}
                  >
                    View meeting minutes
                  </Button>
                )}
              </div>
            </div>
          </TeamsScroll>
        </TeamsPage>
      );
    }

    return (
      <TeamsPage>
        <TeamsScroll>
          <div className={`${styles.signPanel} animate-fade-in`}>
            <Button
              appearance="subtle"
              icon={<ArrowLeft20Regular />}
              onClick={backToList}
              style={{ alignSelf: "flex-start" }}
            >
              All meetings
            </Button>

            <div>
              <Text size={500} weight="bold" block>{selected.displayDate}</Text>
              {existing && (
                <div style={{ marginTop: tokens.spacingVerticalXS }}>{statusBadge(existing.status)}</div>
              )}
            </div>

            <div className={styles.signerBox}>
              <Signature24Regular style={{ color: tokens.colorBrandForeground1, flexShrink: 0 }} />
              <Text size={200}>
                You're signing as <strong>{userName}</strong>. You can only sign your own attendance.
              </Text>
            </div>

            {existing?.signatureData && mode === "choose" && (
              <div>
                <Text size={200} block style={{ color: tokens.colorNeutralForeground3, marginBottom: tokens.spacingVerticalXS }}>
                  Your current signature:
                </Text>
                <img src={existing.signatureData} alt="Your signature" className={styles.sigPreview} />
              </div>
            )}

            {actionError && (
              <MessageBar intent="error">
                <MessageBarBody>{actionError}</MessageBarBody>
              </MessageBar>
            )}

            {mode === "choose" ? (
              <div className={styles.choiceStack}>
                <Text size={300} style={{ color: tokens.colorNeutralForeground3 }}>
                  {existing ? "Update how you attended:" : "How did you attend this meeting?"}
                </Text>
                <Button
                  appearance="primary"
                  className={styles.bigBtn}
                  icon={<Edit20Regular />}
                  disabled={signMutation.isPending}
                  onClick={() => { setActionError(null); setMode("draw"); }}
                >
                  Sign with my finger
                </Button>
                <Button
                  className={styles.bigBtn}
                  icon={<Desktop20Regular />}
                  disabled={signMutation.isPending}
                  onClick={() => submit("remote", generateRemoteSignatureImage(userName, selected.displayDate))}
                >
                  I attended remotely
                </Button>
                <Button
                  className={styles.bigBtn}
                  icon={<Dismiss20Regular />}
                  disabled={signMutation.isPending}
                  onClick={() => submit("absent", null)}
                >
                  I was absent
                </Button>
              </div>
            ) : (
              <div className={styles.choiceStack}>
                <div className={styles.canvasWrap} style={{ touchAction: "none" }}>
                  <canvas
                    ref={canvasRef}
                    width={480}
                    height={220}
                    style={{ width: "100%", height: "220px", display: "block", cursor: "crosshair" }}
                    onMouseDown={(e) => startDraw(getPos(e.nativeEvent, e.currentTarget))}
                    onMouseMove={(e) => moveDraw(getPos(e.nativeEvent, e.currentTarget))}
                    onMouseUp={endDraw}
                    onMouseLeave={endDraw}
                    onTouchStart={(e) => { e.preventDefault(); startDraw(getPos(e.touches[0], e.currentTarget)); }}
                    onTouchMove={(e) => { e.preventDefault(); moveDraw(getPos(e.touches[0], e.currentTarget)); }}
                    onTouchEnd={(e) => { e.preventDefault(); endDraw(); }}
                  />
                  {!hasDrawn && (
                    <div className={styles.canvasHint}>
                      <Text size={300}>Sign here</Text>
                    </div>
                  )}
                </div>
                <div className={styles.drawActions}>
                  <Button
                    icon={<ArrowCounterclockwise20Regular />}
                    onClick={clearCanvas}
                    disabled={!hasDrawn || signMutation.isPending}
                    style={{ flexGrow: 1 }}
                  >
                    Clear
                  </Button>
                  <Button
                    appearance="primary"
                    onClick={confirmDrawn}
                    disabled={!hasDrawn || signMutation.isPending}
                    style={{ flexGrow: 2 }}
                  >
                    {signMutation.isPending ? "Saving…" : "Confirm signature"}
                  </Button>
                </div>
                <Button appearance="subtle" onClick={() => setMode("choose")} disabled={signMutation.isPending}>
                  Cancel
                </Button>
              </div>
            )}
          </div>
        </TeamsScroll>
      </TeamsPage>
    );
  }

  // ─── Loading / error / empty list states ─────────────────────────────────────
  if (meetingsQuery.isLoading) {
    return (
      <TeamsCenter className="animate-fade-in">
        <Spinner size="large" label="Loading recent meetings…" />
      </TeamsCenter>
    );
  }

  if (meetingsQuery.isError) {
    return (
      <TeamsFullScreen
        icon={<Shield24Regular />}
        title="Couldn't load meetings"
        description={(meetingsQuery.error as Error)?.message || "Something went wrong. Please try again."}
        actionLabel="Try again"
        onAction={() => meetingsQuery.refetch()}
      />
    );
  }

  // Non-roster staff (e.g. a remote worker) can't self-sign, but any signed-in
  // Cranfield Glass staff member can still READ locked meeting minutes — so show
  // them the readable-minutes list instead of a dead-end "not found" message.
  if (isUnmatched) {
    const readable = minutesListQuery.data?.meetings ?? [];
    return (
      <TeamsPage>
        <TeamsPinned className={styles.intro}>
          <Text size={400} weight="bold">Meeting minutes</Text>
          <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
            You're not on the attendee list, so there's nothing to sign — but you can read the minutes of past meetings here.
          </Text>
        </TeamsPinned>
        <TeamsScroll>
          {minutesListQuery.isLoading ? (
            <TeamsCenter className="animate-fade-in">
              <Spinner size="large" label="Loading meeting minutes…" />
            </TeamsCenter>
          ) : minutesListQuery.isError ? (
            <TeamsCenter className="animate-fade-in">
              <div className={styles.doneWrap}>
                <DocumentText24Regular style={{ fontSize: "48px", color: tokens.colorNeutralForeground4 }} />
                <Text size={400} weight="semibold" style={{ textAlign: "center" }}>Couldn't load meetings</Text>
                <Text size={300} style={{ color: tokens.colorNeutralForeground3, textAlign: "center" }}>
                  {(minutesListQuery.error as Error)?.message || "Something went wrong. Please try again."}
                </Text>
                <Button appearance="primary" onClick={() => minutesListQuery.refetch()} style={{ width: "100%" }}>
                  Try again
                </Button>
              </div>
            </TeamsCenter>
          ) : readable.length === 0 ? (
            <TeamsCenter>
              <div className={styles.doneWrap}>
                <CalendarLtr24Regular style={{ fontSize: "48px", color: tokens.colorNeutralForeground4 }} />
                <Text size={400} weight="semibold" style={{ textAlign: "center" }}>No minutes to read yet</Text>
                <Text size={300} style={{ color: tokens.colorNeutralForeground3, textAlign: "center" }}>
                  Once a meeting has been held and locked by an admin, its minutes will appear here.
                </Text>
              </div>
            </TeamsCenter>
          ) : (
            <div className={styles.list}>
              <div className={styles.section}>
                <SectionHeader
                  icon={<DocumentText24Regular style={{ fontSize: "16px", color: tokens.colorNeutralForeground3 }} />}
                  label="Meeting minutes"
                />
                {readable.map((m) => (
                  <MeetingCard
                    key={m.dateKey}
                    icon={<DocumentText24Regular />}
                    iconTone="neutral"
                    title={m.displayDate}
                    subtitle="Tap to read minutes"
                    onClick={() => openMinutes(m.dateKey, m.displayDate)}
                  />
                ))}
              </div>
            </div>
          )}
        </TeamsScroll>
      </TeamsPage>
    );
  }

  // Locked meetings this user may read. De-duplicate against the meetings already
  // shown above (open, attended, and the upcoming next meeting) so a single
  // meeting is never listed twice on the tab.
  const shownKeys = new Set<string>([
    ...openMeetings.map((m) => m.dateKey),
    ...pastMeetings.map((m) => m.dateKey),
    ...upcomingMeetings.map((m) => m.dateKey),
  ]);
  const readableMinutes = (minutesListQuery.data?.meetings ?? []).filter(
    (m) => !shownKeys.has(m.dateKey),
  );

  // Attendance tier groups the open-but-already-signed meetings (awaiting an
  // admin lock) with the locked meetings the user attended — both are "your
  // attendance is recorded" rows, just at different lifecycle stages.
  const hasAttendance = signedOpen.length > 0 || pastMeetings.length > 0;
  const hasAnything =
    !!heroSignMeeting ||
    !!nextMeeting ||
    extraUnsignedOpen.length > 0 ||
    hasAttendance ||
    readableMinutes.length > 0;

  return (
    <TeamsPage>
      <TeamsScroll>
        {!hasAnything ? (
          <TeamsCenter>
            <div className={styles.doneWrap}>
              <CalendarLtr24Regular style={{ fontSize: "48px", color: tokens.colorNeutralForeground4 }} />
              <Text size={400} weight="semibold" style={{ textAlign: "center" }}>You're all caught up</Text>
              <Text size={300} style={{ color: tokens.colorNeutralForeground3, textAlign: "center" }}>
                There are no meetings to sign right now. Once the next H&amp;S meeting is scheduled it will appear here.
              </Text>
            </div>
          </TeamsCenter>
        ) : (
          <div className={styles.list}>
            {/* ─── Tier 1: priority zone ─────────────────────────────────── */}
            {/* Signing is the highest-importance action, so when something needs
                signing it takes the solid-brand hero. The agenda still appears,
                but as the brand-tint hero just beneath it. With nothing to sign,
                the agenda is promoted to the solid hero. */}
            {(heroSignMeeting || nextMeeting) && (
              <div className={styles.heroStack}>
                {heroSignMeeting && (
                  <HeroCard
                    tone="solid"
                    icon={<Signature24Regular />}
                    eyebrow="Action needed · sign"
                    title={heroSignMeeting.displayDate}
                    subtitle={`${relativeFromToday(heroSignMeeting.dateKey)} · confirm how you attended`}
                    actionLabel="Sign your attendance"
                    onClick={() => openMeeting(heroSignMeeting)}
                  />
                )}
                {nextMeeting && (
                  <HeroCard
                    tone={heroSignMeeting ? "tint" : "solid"}
                    icon={<CalendarLtr24Regular />}
                    eyebrow="Next meeting · agenda"
                    title={nextMeeting.displayDate}
                    subtitle={relativeFromToday(nextMeeting.dateKey)}
                    actionLabel="View the agenda"
                    onClick={() => openAgenda(nextMeeting.dateKey, nextMeeting.displayDate)}
                  />
                )}
              </div>
            )}

            {/* Any further unsigned meetings beyond the hero. */}
            {extraUnsignedOpen.length > 0 && (
              <div className={styles.section}>
                <SectionHeader
                  icon={<Signature24Regular style={{ fontSize: "16px", color: tokens.colorNeutralForeground3 }} />}
                  label="Also ready to sign"
                />
                {extraUnsignedOpen.map((m) => (
                  <MeetingCard
                    key={m.dateKey}
                    icon={<Signature24Regular />}
                    iconTone="brand"
                    title={m.displayDate}
                    subtitle="Tap to sign your attendance"
                    onClick={() => openMeeting(m)}
                  />
                ))}
              </div>
            )}

            {/* ─── Tier 2: signing & attendance ──────────────────────────── */}
            {hasAttendance && (
              <div className={styles.section}>
                <SectionHeader
                  icon={<CheckmarkCircle24Filled style={{ fontSize: "16px", color: tokens.colorNeutralForeground3 }} />}
                  label="Your attendance"
                />
                {signedOpen.map((m) => (
                  <MeetingCard
                    key={m.dateKey}
                    icon={<CheckmarkCircle24Filled />}
                    iconTone="success"
                    title={m.displayDate}
                    subtitle="Awaiting admin lock · tap to update"
                    trailing={m.mySignature ? statusBadge(m.mySignature.status) : undefined}
                    onClick={() => openMeeting(m)}
                  />
                ))}
                {pastMeetings.map((m) => (
                  <MeetingCard
                    key={m.dateKey}
                    icon={<CheckmarkCircle24Filled />}
                    iconTone="success"
                    title={m.displayDate}
                    subtitle="Tap to view the minutes"
                    trailing={
                      m.mySignature
                        ? statusBadge(m.mySignature.status)
                        : <Badge appearance="tint" color="success">Attended</Badge>
                    }
                    onClick={() => openMinutes(m.dateKey, m.displayDate)}
                  />
                ))}
              </div>
            )}

            {/* ─── Tier 3: the quiet minutes archive ─────────────────────── */}
            {readableMinutes.length > 0 && (
              <div className={styles.section}>
                <SectionHeader
                  icon={<DocumentText24Regular style={{ fontSize: "16px", color: tokens.colorNeutralForeground3 }} />}
                  label={hasAttendance ? "Minutes archive" : "Meeting minutes"}
                />
                {readableMinutes.map((m) => (
                  <MeetingCard
                    key={m.dateKey}
                    icon={<DocumentText24Regular />}
                    iconTone="neutral"
                    title={m.displayDate}
                    subtitle="Tap to read minutes"
                    onClick={() => openMinutes(m.dateKey, m.displayDate)}
                  />
                ))}
              </div>
            )}
          </div>
        )}
      </TeamsScroll>
    </TeamsPage>
  );
}

import { useState, useRef, useCallback, useEffect } from "react";
import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import {
  Button,
  Card,
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
  PersonQuestionMark24Regular,
  Edit20Regular,
  Desktop20Regular,
  Dismiss20Regular,
  ArrowLeft20Regular,
  CheckmarkCircle24Filled,
  CalendarLtr24Regular,
  ArrowCounterclockwise20Regular,
} from "@fluentui/react-icons";
import {
  TeamsPage,
  TeamsPinned,
  TeamsScroll,
  TeamsCenter,
  TeamsFullScreen,
} from "./TeamsPageShell";
import { useTeamsAuth } from "@/hooks/useTeamsAuth";

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
  sectionHeader: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalS,
    paddingLeft: tokens.spacingHorizontalXS,
  },
  nextCard: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    textAlign: "center",
    gap: tokens.spacingVerticalS,
    padding: tokens.spacingHorizontalL,
    backgroundColor: tokens.colorBrandBackground2,
  },
  nextIcon: {
    width: "44px",
    height: "44px",
    flexShrink: 0,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    borderRadius: tokens.borderRadiusCircular,
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
  },
  meetingCard: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalM,
    padding: tokens.spacingHorizontalL,
    cursor: "pointer",
  },
  meetingCardStatic: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalM,
    padding: tokens.spacingHorizontalL,
  },
  meetingIcon: {
    width: "40px",
    height: "40px",
    flexShrink: 0,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    borderRadius: tokens.borderRadiusCircular,
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground1,
  },
  meetingBody: { flexGrow: 1, minWidth: 0, display: "flex", flexDirection: "column", gap: "2px" },
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

  // Derive the selected meeting from fresh query data so its signature status is
  // never stale, even after a remount restores the persisted dateKey.
  const meetings = meetingsQuery.data?.meetings ?? [];
  const selected = selectedKey ? meetings.find((m) => m.dateKey === selectedKey) ?? null : null;

  // Group for display: open meetings need action, locked ones are read-only
  // history, and upcoming ones surface the next meeting date.
  const openMeetings = meetings.filter((m) => m.state === "open");
  const pastMeetings = meetings.filter((m) => m.state === "locked");
  const upcomingMeetings = meetings
    .filter((m) => m.state === "upcoming")
    .sort((a, b) => new Date(a.meetingDate).getTime() - new Date(b.meetingDate).getTime());
  const nextMeeting = upcomingMeetings[0] ?? null;

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
                <Text size={200} block style={{ color: tokens.colorNeutralForeground3, marginTop: "2px" }}>
                  {relativeFromToday(selected.dateKey)}
                </Text>
                {existing && (
                  <div style={{ marginTop: tokens.spacingVerticalS }}>{statusBadge(existing.status)}</div>
                )}
              </div>
              <MessageBar intent={isUpcoming ? "info" : "success"}>
                <MessageBarBody>
                  {isUpcoming
                    ? "This meeting hasn't happened yet. You'll be able to sign your attendance once it has taken place."
                    : "This meeting has been locked by an admin, so your attendance is now final and can't be changed."}
                </MessageBarBody>
              </MessageBar>
              {existing?.signatureData && (
                <div>
                  <Text size={200} block style={{ color: tokens.colorNeutralForeground3, marginBottom: tokens.spacingVerticalXS }}>
                    Your signature:
                  </Text>
                  <img src={existing.signatureData} alt="Your signature" className={styles.sigPreview} />
                </div>
              )}
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

  if (meetingsQuery.data && !meetingsQuery.data.matched) {
    return (
      <TeamsFullScreen
        icon={<PersonQuestionMark24Regular />}
        title="We couldn't find you on the attendee list"
        description="Only people on the H&S meeting attendee list can self-sign. Please ask an admin to add you, then try again."
        actionLabel="Try again"
        onAction={() => meetingsQuery.refetch()}
      />
    );
  }

  return (
    <TeamsPage>
      <TeamsPinned className={styles.intro}>
        <Text size={400} weight="bold">Sign meeting minutes</Text>
        <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
          Sign recent meetings and see the ones you've attended.
        </Text>
      </TeamsPinned>
      <TeamsScroll>
        {openMeetings.length === 0 && upcomingMeetings.length === 0 && pastMeetings.length === 0 ? (
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
            {nextMeeting && (
              <Card className={`${styles.nextCard} animate-fade-in-up`}>
                <div className={styles.nextIcon}>
                  <CalendarLtr24Regular />
                </div>
                <div style={{ display: "flex", flexDirection: "column", gap: "2px", width: "100%" }}>
                  <Text size={200} weight="semibold" style={{ color: tokens.colorBrandForeground1 }} block>
                    Next meeting
                  </Text>
                  <Text size={400} weight="bold" block>{nextMeeting.displayDate}</Text>
                  <Text size={200} style={{ color: tokens.colorNeutralForeground3 }} block>
                    {relativeFromToday(nextMeeting.dateKey)} · sign after the meeting
                  </Text>
                </div>
              </Card>
            )}

            {openMeetings.length > 0 && (
              <div className={styles.section}>
                <div className={styles.sectionHeader}>
                  <Signature24Regular style={{ fontSize: "16px", color: tokens.colorNeutralForeground3 }} />
                  <Text size={200} weight="semibold" style={{ color: tokens.colorNeutralForeground3 }}>
                    Ready to sign
                  </Text>
                </div>
                {openMeetings.map((m) => (
                  <Card key={m.dateKey} className={`${styles.meetingCard} animate-fade-in-up`} onClick={() => openMeeting(m)}>
                    <div className={styles.meetingIcon}>
                      {m.mySignature ? <CheckmarkCircle24Filled /> : <Signature24Regular />}
                    </div>
                    <div className={styles.meetingBody}>
                      <Text size={300} weight="semibold" truncate block>{m.displayDate}</Text>
                      {m.mySignature ? (
                        statusBadge(m.mySignature.status)
                      ) : (
                        <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>Tap to sign</Text>
                      )}
                    </div>
                  </Card>
                ))}
              </div>
            )}

            {pastMeetings.length > 0 && (
              <div className={styles.section}>
                <div className={styles.sectionHeader}>
                  <CheckmarkCircle24Filled style={{ fontSize: "16px", color: tokens.colorNeutralForeground3 }} />
                  <Text size={200} weight="semibold" style={{ color: tokens.colorNeutralForeground3 }}>
                    Meetings you attended
                  </Text>
                </div>
                {pastMeetings.map((m) => (
                  <Card key={m.dateKey} className={`${styles.meetingCardStatic} animate-fade-in-up`} onClick={() => openMeeting(m)}>
                    <div className={styles.meetingIcon}>
                      <CheckmarkCircle24Filled />
                    </div>
                    <div className={styles.meetingBody}>
                      <Text size={300} weight="semibold" truncate block>{m.displayDate}</Text>
                      {m.mySignature ? (
                        statusBadge(m.mySignature.status)
                      ) : (
                        <Badge appearance="tint" color="success">Attended</Badge>
                      )}
                    </div>
                  </Card>
                ))}
              </div>
            )}
          </div>
        )}
      </TeamsScroll>
    </TeamsPage>
  );
}

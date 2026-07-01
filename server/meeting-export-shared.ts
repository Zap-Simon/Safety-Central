/**
 * Shared building blocks for every meeting-minutes export (HTML/PDF, CSV,
 * Markdown, Word). Keeping the column semantics here means all export formats
 * stay in lockstep — any future export that uses these helpers automatically
 * gets the compliant labelling:
 *   - Agenda Item    = the original submission text (description + how it happened)
 *   - Discussion     = the meeting notes only (rendered by each format directly)
 *   - Action Required = the REAL actioned-system fields, never fabricated text
 */

export interface ActionLine {
  label: string;
  value: string;
}

/**
 * The original submission IS the agenda item. Combines the description with the
 * "How it happened" follow-up (used for Near Miss reports) into one block.
 */
export function buildAgendaSubmissionText(item: any): string {
  let text = item?.description || '';
  if (item?.secondaryDescription) {
    text += text ? `\n\nHow it happened: ${item.secondaryDescription}` : item.secondaryDescription;
  }
  return text;
}

/**
 * The "Action Required" column, sourced strictly from the real actioned system.
 * No fabricated / boilerplate text — only what was actually recorded.
 *
 * Returns structured lines so each export format can render them in its own way
 * (HTML divs, Markdown bold, Word runs, CSV separated text) while the wording
 * and rules stay identical everywhere.
 */
// The status the action tracker writes by default the moment an item is
// actioned. On its own it is NOT real progress — it just means the item has
// been picked up to move across to the Actions board. Treating it as progress
// is what made actioned items read a misleading "Not Started" in the export.
const PLACEHOLDER_ACTION_STATUS = 'Not Started';

/**
 * True when the item has been picked up / moved to the Actions board. SharePoint
 * uses "Actioned" (the local tracker may surface it as "Actions"); the Near Miss
 * list in particular doesn't always round-trip that status, so the presence of a
 * local action record (even just the bare placeholder) also counts as actioned.
 */
export function isItemActioned(item: any): boolean {
  const status = item?.status || '';
  return status === 'Actioned' || status === 'Actions'
    || !!(item?.actionStatus || item?.actionPriority || item?.actionStartDate);
}

/**
 * True when a human has actually recorded action progress: an assignee, a due
 * date, notes, or a real (non-placeholder) status. The bare "Not Started"
 * placeholder that actioning writes does NOT count as progress.
 */
function hasRealActionProgress(item: any): boolean {
  const hasRealStatus = !!(item?.actionStatus && item.actionStatus !== PLACEHOLDER_ACTION_STATUS);
  return !!(item?.actionAssignedTo || item?.actionDueDate || item?.actionNotes || hasRealStatus);
}

export function buildActionRequiredLines(item: any): ActionLine[] {
  const status = item?.status || '';
  const isClosed = status === 'Closed';
  const isNearMiss = item?.type === 'Near Miss';

  if (hasRealActionProgress(item)) {
    const lines: ActionLine[] = [];
    if (item.actionAssignedTo) lines.push({ label: 'Assigned to', value: String(item.actionAssignedTo) });
    if (item.actionStatus) lines.push({ label: 'Status', value: String(item.actionStatus) });
    if (item.actionDueDate) lines.push({ label: 'Due', value: new Date(item.actionDueDate).toLocaleDateString('en-NZ') });
    if (item.actionNotes) {
      // For a closed item the notes capture what was done / the outcome;
      // for an in-flight item they describe the action still to do.
      lines.push({ label: isClosed ? 'Outcome' : 'Action', value: String(item.actionNotes) });
    }
    return lines;
  }

  // Discussed and closed on the spot with no follow-up action recorded.
  if (isClosed) {
    return [{ label: '', value: 'Discussed and closed — no action required.' }];
  }

  // Actioned but no real progress recorded yet — only the default placeholder is
  // present (e.g. just picked up at the meeting, or actioned before action
  // tracking existed). Show a meaningful status instead of the misleading
  // "Not Started". A Near Miss is moved across for its formal investigation;
  // Safety/Business ideas are simply picked up. Once genuine progress is added
  // (assignee, due, notes, or a real status) the full breakdown above replaces
  // this automatically.
  if (isItemActioned(item)) {
    return isNearMiss
      ? [{ label: 'Status', value: 'Moved to Action Board for Investigation' }]
      : [{ label: 'Status', value: 'Actioned' }];
  }

  // Still open / in discussion with no action captured yet.
  return [{ label: '', value: '—' }];
}

/**
 * The item's overall status as it should READ in the minutes, kept coherent with
 * the "Action Required" column. A Near Miss that has been moved to the Actions
 * board reads "Moved to Action Board for Investigation" instead of a bare
 * "Actioned" (or a stale "Submitted" when SharePoint didn't round-trip the
 * status). Terminal statuses (Closed/Completed) and all other types are left
 * exactly as-is.
 */
export function getDisplayItemStatus(item: any): string {
  const rawStatus = item?.status || 'Submitted';
  if (item?.type !== 'Near Miss') return rawStatus;

  const isActionedStatus = rawStatus === 'Actioned' || rawStatus === 'Actions';
  // SharePoint's Near Miss list doesn't always read the "Actioned" status back,
  // so a still-"Submitted" item that has a local action record was also moved.
  const movedViaLocalRecord = rawStatus === 'Submitted' && isItemActioned(item);

  if (isActionedStatus || movedViaLocalRecord) {
    return 'Moved to Action Board for Investigation';
  }
  return rawStatus;
}

/**
 * True when the action lines are just the empty "—" placeholder (open item with
 * nothing recorded). Lets formats like Word/CSV omit it entirely.
 */
export function isEmptyActionPlaceholder(lines: ActionLine[]): boolean {
  return lines.length === 1 && lines[0].label === '' && lines[0].value === '—';
}

/**
 * Plain-text rendering of the action lines (CSV, Word fallbacks). Returns an
 * empty string for the "—" placeholder so columns stay clean.
 */
export function actionRequiredPlainText(item: any, separator = ' | '): string {
  const lines = buildActionRequiredLines(item);
  if (isEmptyActionPlaceholder(lines)) return '';
  return lines.map((l) => (l.label ? `${l.label}: ${l.value}` : l.value)).join(separator);
}

/**
 * The standalone "Ready to Close" actions section shared by every export format.
 *
 * These are actions a person has completed and parked at "Ready to Close" — they
 * still need a group review + sign-off at the meeting to be formally closed. The
 * same action stays "Ready to Close" until it's actually closed, so it legitimately
 * re-appears across consecutive meetings; that's why we surface the due date with
 * each one (it can sit in the future). This section is built from the FULL meeting
 * dataset, NOT the single selected meeting, so every export shows all outstanding
 * ready-to-close actions regardless of which meeting is being exported.
 */
export interface ReadyToCloseAction {
  id: string;
  title: string;
  type: string;
  assignedTo: string;
  dueDate: string;
  outcome: string;
  submittedBy: string;
}

export function isReadyToCloseAction(item: any): boolean {
  return String(item?.actionStatus || '').trim() === 'Ready to Close';
}

export function buildReadyToCloseActions(allItems: any[]): ReadyToCloseAction[] {
  if (!Array.isArray(allItems)) return [];
  const seen = new Set<string>();
  const result: ReadyToCloseAction[] = [];
  for (const item of allItems) {
    if (!isReadyToCloseAction(item)) continue;
    const id = String(item?.id || '');
    // IDs are list-local in SharePoint (the same numeric id can exist in the
    // Near Miss, Safety and Business lists), so dedupe on type + id, never the
    // bare id, or a ready-to-close action from another list would be dropped.
    const dedupeKey = id ? `${item?.type || ''}::${id}` : '';
    if (dedupeKey && seen.has(dedupeKey)) continue;
    if (dedupeKey) seen.add(dedupeKey);
    result.push({
      id,
      title: item?.title || `${item?.type || 'Action'} Item`,
      type: item?.type || '',
      assignedTo: item?.actionAssignedTo ? String(item.actionAssignedTo) : '',
      dueDate: item?.actionDueDate ? new Date(item.actionDueDate).toLocaleDateString('en-NZ') : '',
      outcome: item?.actionNotes ? String(item.actionNotes) : '',
      submittedBy: item?.submittedBy ? String(item.submittedBy) : '',
    });
  }
  return result.sort((a, b) => a.title.localeCompare(b.title));
}

/**
 * Document control for quality/compliance. Every human-readable export carries a
 * version number, the issue date (when it was generated) and a review date one
 * year on, so any printed or shared copy has a clear "valid until" for the safety
 * management system. Review date is computed from NZ "today" (Pacific/Auckland)
 * so it never drifts a day around UTC midnight.
 */
export interface DocumentVersionInfo {
  version: string;
  issued: string;
  reviewBy: string;
}

export function buildDocumentVersionInfo(issuedDate: string, version = "1.0"): DocumentVersionInfo {
  const nzToday = new Date().toLocaleDateString("en-CA", { timeZone: "Pacific/Auckland" }); // yyyy-mm-dd
  const [y, m, d] = nzToday.split("-").map(Number);
  const reviewBy = new Date(Date.UTC(y + 1, m - 1, d)).toLocaleDateString("en-NZ", {
    day: "2-digit", month: "long", year: "numeric", timeZone: "UTC",
  });
  return { version, issued: issuedDate, reviewBy };
}

/** Inline-styled "Document control" footer for any HTML/PDF export. */
export function documentVersionFooterHTML(issuedDate: string, version = "1.0"): string {
  const info = buildDocumentVersionInfo(issuedDate, version);
  return `<div class="doc-control" style="margin-top:24px;padding-top:10px;border-top:1px solid #e5e7eb;display:flex;flex-wrap:wrap;gap:6px 24px;justify-content:center;font-size:8.5pt;color:#6b7280;">`
    + `<span><strong>Document version:</strong> ${info.version}</span>`
    + `<span><strong>Issued:</strong> ${info.issued}</span>`
    + `<span><strong>Next review:</strong> ${info.reviewBy}</span>`
    + `</div>`;
}

/** "Document control" footer line for any Markdown export. */
export function documentVersionFooterMarkdown(issuedDate: string, version = "1.0"): string {
  const info = buildDocumentVersionInfo(issuedDate, version);
  return `\n**Document version:** ${info.version}  ·  **Issued:** ${info.issued}  ·  **Next review:** ${info.reviewBy}\n`;
}

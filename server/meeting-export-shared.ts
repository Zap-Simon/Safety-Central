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

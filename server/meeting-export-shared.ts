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
export function buildActionRequiredLines(item: any): ActionLine[] {
  const isClosed = item?.status === 'Closed';
  const hasActionData = !!(item?.actionAssignedTo || item?.actionStatus || item?.actionDueDate || item?.actionNotes);

  if (hasActionData) {
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

  // Still open / in discussion with no action captured yet.
  return [{ label: '', value: '—' }];
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

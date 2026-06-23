/**
 * Actions Report export engine.
 *
 * Rebuilds the Actions page export on the SAME professional foundation the
 * meeting-minutes exports use: a Cranfield-branded, A4 print-ready HTML report
 * (Paged.js "Page X of Y" footers, repeating table headers, break-inside
 * protection), plus matching CSV, Markdown and Word outputs so the format set
 * stays in lockstep with the meeting export.
 *
 * Unlike the meeting minutes (which are organised by meeting) this report is
 * action-centric: every row/block is a single tracked action carrying its full
 * lifecycle — On Hold (+ revisit date), Ready to Close, due-date analytics,
 * progress/outcome notes, Near Miss investigation details and the per-action
 * activity history.
 */

import {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  BorderStyle, WidthType, AlignmentType,
} from "docx";
import { PAGEDJS_BASE64 } from "./assets/pagedjs";

const PAGEDJS_SCRIPT = Buffer.from(PAGEDJS_BASE64, "base64").toString("utf-8");

// ─── Shared types ───────────────────────────────────────────────────────────

export interface ActionActivity {
  entryType: string;
  content: string;
  author: string | null;
  createdAt: string;
}

export interface ActionInvestigation {
  riskLevel?: string;
  investigatorName?: string;
  investigatorSignature?: string | null;
  investigatorSignedAt?: string | null;
  directorName?: string | null;
  directorSignature?: string | null;
  signedAt?: string | null;
  status?: string;
}

export interface ActionExportItem {
  id: string;
  title?: string;
  description?: string;
  type: string;
  status?: string;
  meetingDate?: string;
  submittedBy?: string;
  actionAssignedTo?: string;
  actionStatus?: string;
  actionPriority?: string;
  actionStartDate?: string;
  actionDueDate?: string;
  reconsiderDate?: string;
  actionNotes?: string;
  activity?: ActionActivity[];
  investigation?: ActionInvestigation | null;
}

export interface ActionsStats {
  total: number;
  open: number;
  completed: number;
  overdue: number;
  highPriority: number;
}

// ─── Shared colour + label helpers (kept in lockstep across every format) ─────

const TYPE_COLORS: Record<string, string> = {
  "Business Ideas": "#2563eb",
  "Safety Ideas": "#dc2626",
  "Near Miss": "#ea580c",
};

function typeColor(type: string): string {
  return TYPE_COLORS[type] || "#6b7280";
}

const ARCHIVED_STATUSES = new Set(["Completed", "Ready to Close"]);

function isArchived(status?: string): boolean {
  return ARCHIVED_STATUSES.has(status || "");
}

function esc(value: unknown): string {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** Date only, in NZ time so a date entered during the NZ day never rolls back. */
function fmtDate(value?: string | null): string {
  if (!value) return "";
  const d = new Date(value);
  if (isNaN(d.getTime())) return String(value);
  return d.toLocaleDateString("en-NZ", {
    day: "2-digit", month: "short", year: "numeric", timeZone: "Pacific/Auckland",
  });
}

/** Date + time in NZ, used for the activity timeline. */
function fmtDateTime(value?: string | null): string {
  if (!value) return "";
  const d = new Date(value);
  if (isNaN(d.getTime())) return String(value);
  return d.toLocaleString("en-NZ", {
    day: "2-digit", month: "short", year: "numeric",
    hour: "2-digit", minute: "2-digit", timeZone: "Pacific/Auckland",
  });
}

/**
 * The displayed status, keeping On Hold's revisit date and Ready to Close
 * visible. Mirrors what the Actions page itself shows.
 */
function displayStatus(item: ActionExportItem): string {
  const status = item.actionStatus || "Not Started";
  if (status === "On Hold" && item.reconsiderDate) {
    return `On Hold (revisit ${fmtDate(item.reconsiderDate)})`;
  }
  return status;
}

/**
 * Due-date analytics matching the on-page badges: Overdue / Due Today / N days.
 * Suppressed for archived (Completed / Ready to Close) actions, exactly like the
 * page hides the urgency badge once an action is finished.
 */
function dueStatus(item: ActionExportItem): { label: string; tone: "overdue" | "today" | "soon" | "ok" } | null {
  if (!item.actionDueDate || isArchived(item.actionStatus)) return null;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const due = new Date(item.actionDueDate);
  if (isNaN(due.getTime())) return null;
  due.setHours(0, 0, 0, 0);
  const diff = Math.ceil((due.getTime() - today.getTime()) / 86400000);
  if (diff < 0) return { label: "Overdue", tone: "overdue" };
  if (diff === 0) return { label: "Due Today", tone: "today" };
  if (diff <= 3) return { label: `${diff} day${diff === 1 ? "" : "s"} left`, tone: "soon" };
  return { label: `${diff} days left`, tone: "ok" };
}

/** Human-friendly label for an activity entry's type. */
function activityTypeLabel(entryType: string): string {
  switch (entryType) {
    case "note": return "Note";
    case "status": return "Status";
    case "priority": return "Priority";
    case "due_date": return "Due date";
    case "start_date": return "Start date";
    case "assigned": return "Assigned";
    default: return entryType ? entryType.replace(/_/g, " ") : "Update";
  }
}

/** True when a near miss investigation carries anything worth showing. */
function hasInvestigation(item: ActionExportItem): boolean {
  const inv = item.investigation;
  if (!inv) return false;
  return !!(inv.riskLevel || inv.investigatorName || inv.investigatorSignature
    || inv.directorSignature || inv.directorName || (inv.status && inv.status !== "Draft"));
}

function investigatorSignOff(inv: ActionInvestigation): string {
  if (inv.investigatorSignature) {
    return `Signed${inv.investigatorName ? ` by ${inv.investigatorName}` : ""}${inv.investigatorSignedAt ? ` on ${fmtDate(inv.investigatorSignedAt)}` : ""}`;
  }
  return "Not yet signed";
}

function directorSignOff(inv: ActionInvestigation): string {
  if (inv.directorSignature) {
    return `Signed${inv.directorName ? ` by ${inv.directorName}` : ""}${inv.signedAt ? ` on ${fmtDate(inv.signedAt)}` : ""}`;
  }
  return "Awaiting sign-off";
}

// ─── Breakdown helpers for the summary dashboard ──────────────────────────────

function countBy<T>(items: T[], key: (i: T) => string): Array<[string, number]> {
  const counts: Record<string, number> = {};
  for (const item of items) {
    const k = key(item) || "—";
    counts[k] = (counts[k] || 0) + 1;
  }
  return Object.entries(counts).sort((a, b) => b[1] - a[1]);
}

// ─── HTML report ──────────────────────────────────────────────────────────────

export function generateActionsReportHTML(
  items: ActionExportItem[],
  stats: ActionsStats,
  currentDate: string,
): string {
  const statusBreakdown = countBy(items, (i) => i.actionStatus || "Not Started");
  const typeBreakdown = countBy(items, (i) => i.type || "—");
  const priorityBreakdown = countBy(items, (i) => i.actionPriority || "Not Set");

  const overviewRows = items.map((item, index) => {
    const due = dueStatus(item);
    return `
      <tr>
        <td style="text-align:center;">${index + 1}</td>
        <td>
          <div class="type-badge" style="background-color:${typeColor(item.type)};">${esc(item.type)}</div>
          <div class="action-title">${esc(item.title || `${item.type} Action`)}</div>
          <div class="muted"><strong>Assigned:</strong> ${esc(item.actionAssignedTo || "Unassigned")}</div>
        </td>
        <td>${priorityBadge(item.actionPriority)}</td>
        <td>${statusBadge(item)}</td>
        <td>
          ${item.actionDueDate ? `<div>${esc(fmtDate(item.actionDueDate))}</div>` : '<div class="muted">Not set</div>'}
          ${due ? `<div class="due-badge due-${due.tone}">${esc(due.label)}</div>` : ""}
        </td>
      </tr>`;
  }).join("");

  const detailBlocks = items.map((item, index) => renderDetailBlock(item, index)).join("");

  return `
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Cranfield Glass - Actions Report</title>
  <style>
    @page {
      size: A4;
      margin: 1.1cm 1.4cm 1.5cm 1.4cm;
      @bottom-center {
        content: "Cranfield Glass Christchurch  |  Health & Safety Action Report";
        font-family: Arial, sans-serif; font-size: 8pt; color: #9ca3af;
      }
      @bottom-right {
        content: "Page " counter(page) " of " counter(pages);
        font-family: Arial, sans-serif; font-size: 8pt; color: #9ca3af;
      }
    }
    body { font-family: Arial, sans-serif; font-size: 11pt; line-height: 1.4; color: #333; background:#fff; margin:0; padding:0; }
    .container { width:100%; margin:0; padding:0; }
    @media screen {
      body { background:#f3f4f6; }
      .pagedjs_pages { margin:0 auto; }
      .pagedjs_page { margin:0 auto 0.5cm auto !important; background:#fff; box-shadow:0 0 0.4cm rgba(0,0,0,0.15); }
    }
    @media print { .print-button { display:none !important; } }
    .print-button { position:fixed; top:20px; right:20px; background:#3b82f6; color:#fff; border:none; padding:12px 24px; border-radius:6px; cursor:pointer; font-size:14px; font-weight:500; box-shadow:0 2px 4px rgba(0,0,0,0.1); z-index:1000; }
    .print-button:hover { background:#2563eb; }

    .header { text-align:center; margin-bottom:24px; border-bottom:2px solid #3b82f6; padding-bottom:18px; }
    .company-name { font-size:24pt; font-weight:bold; background:linear-gradient(135deg,#14b8a6,#06b6d4,#3b82f6); -webkit-background-clip:text; -webkit-text-fill-color:transparent; background-clip:text; margin-bottom:5px; }
    .document-title { font-size:18pt; font-weight:bold; color:#333; }

    .meeting-info { display:flex; justify-content:space-between; margin-bottom:22px; padding:15px; background:#f8f9fa; border:1px solid #dee2e6; }
    .info-block { flex:1; }
    .info-label { font-weight:bold; margin-bottom:5px; }

    .section-header { font-size:14pt; font-weight:bold; color:#3b82f6; margin-bottom:15px; border-bottom:1px solid #3b82f6; padding-bottom:5px; }

    /* Summary dashboard (mirrors the meeting analytics dashboard styling) */
    .analytics-dashboard { margin-bottom:22px; background:#fff; border:1px solid #e5e7eb; border-radius:8px; overflow:hidden; page-break-inside:avoid; break-inside:avoid; }
    .analytics-bar { display:flex; align-items:center; justify-content:space-between; gap:16px; padding:10px 16px; background:#f8fafc; border-bottom:1px solid #e5e7eb; }
    .analytics-heading { display:flex; flex-direction:column; }
    .analytics-h1 { font-size:13pt; font-weight:bold; color:#1f2937; line-height:1.2; }
    .analytics-sub { font-size:7.5pt; color:#9ca3af; margin-top:2px; }
    .analytics-metrics { display:flex; gap:10px; flex-wrap:wrap; }
    .metric-chip { display:flex; flex-direction:column; align-items:center; justify-content:center; min-width:74px; padding:6px 12px; background:#fff; border:1px solid #e5e7eb; border-radius:6px; }
    .metric-chip-num { font-size:16pt; font-weight:bold; color:#1f2937; line-height:1.1; }
    .metric-chip-num.accent-blue { color:#2563eb; }
    .metric-chip-num.accent-green { color:#16a34a; }
    .metric-chip-num.accent-red { color:#dc2626; }
    .metric-chip-num.accent-amber { color:#d97706; }
    .metric-chip-label { font-size:7pt; color:#6b7280; text-transform:uppercase; letter-spacing:0.3px; margin-top:2px; text-align:center; white-space:nowrap; }
    .analytics-breakdowns { display:grid; grid-template-columns:repeat(3,1fr); gap:0; }
    .breakdown-col { padding:10px 16px; border-right:1px solid #f1f5f9; }
    .breakdown-col:last-child { border-right:none; }
    .breakdown-title { font-size:8pt; font-weight:bold; color:#374151; text-transform:uppercase; letter-spacing:0.4px; margin-bottom:6px; padding-bottom:4px; border-bottom:1px solid #e5e7eb; }
    .bd-row { display:flex; justify-content:space-between; align-items:center; gap:8px; padding:2px 0; font-size:9pt; }
    .bd-label { color:#4b5563; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
    .bd-count { font-weight:bold; color:#1f2937; font-variant-numeric:tabular-nums; }

    /* Overview table */
    .items-table { width:100%; border-collapse:collapse; margin-bottom:28px; border:1px solid #dee2e6; table-layout:fixed; }
    .items-table th { background:#f8f9fa; padding:10px 12px; text-align:left; font-weight:bold; border:1px solid #dee2e6; font-size:10.5pt; overflow-wrap:anywhere; }
    .items-table td { padding:10px 12px; border:1px solid #dee2e6; vertical-align:top; font-size:10pt; overflow-wrap:anywhere; }
    .items-table th:nth-child(1), .items-table td:nth-child(1) { width:5%; }
    .items-table th:nth-child(2), .items-table td:nth-child(2) { width:45%; }
    .items-table th:nth-child(3), .items-table td:nth-child(3) { width:13%; }
    .items-table th:nth-child(4), .items-table td:nth-child(4) { width:20%; }
    .items-table th:nth-child(5), .items-table td:nth-child(5) { width:17%; }
    .items-table tr { page-break-inside:avoid; break-inside:avoid; }
    .items-table thead { display:table-header-group; }

    .type-badge { display:inline-block; padding:3px 7px; border-radius:4px; color:#fff; font-size:8pt; font-weight:bold; margin-bottom:4px; }
    .action-title { font-weight:bold; margin-bottom:3px; }
    .muted { color:#6b7280; font-size:9pt; }

    .badge { display:inline-block; padding:3px 8px; border-radius:10px; font-size:8.5pt; font-weight:bold; }
    .due-badge { display:inline-block; margin-top:3px; padding:1px 7px; border-radius:10px; font-size:8pt; font-weight:bold; }
    .due-overdue { background:#fee2e2; color:#b91c1c; }
    .due-today { background:#ffedd5; color:#c2410c; }
    .due-soon { background:#fef3c7; color:#b45309; }
    .due-ok { background:#f1f5f9; color:#475569; }

    /* Detail blocks — never split a single action across a page */
    .action-block { border:1px solid #e5e7eb; border-left:4px solid #6b7280; border-radius:6px; padding:14px 16px; margin-bottom:14px; page-break-inside:avoid; break-inside:avoid; }
    .ab-head { display:flex; justify-content:space-between; align-items:flex-start; gap:12px; margin-bottom:8px; }
    .ab-title { font-size:12pt; font-weight:bold; color:#111827; }
    .ab-grid { display:grid; grid-template-columns:repeat(2,1fr); gap:4px 24px; margin:8px 0; font-size:10pt; }
    .ab-field { color:#374151; }
    .ab-field strong { color:#111827; }
    .ab-notes { background:#f9fafb; border:1px solid #eef2f7; border-radius:6px; padding:10px; margin-top:8px; font-size:10pt; white-space:pre-wrap; }
    .ab-sub-title { font-size:9pt; font-weight:bold; text-transform:uppercase; letter-spacing:0.4px; color:#6b7280; margin:12px 0 6px; }
    .inv-box { background:#fff7ed; border:1px solid #fed7aa; border-radius:6px; padding:10px; font-size:10pt; }
    .inv-box .ab-field { margin-bottom:2px; }
    .timeline { list-style:none; margin:0; padding:0; }
    .timeline li { padding:5px 0 5px 12px; border-left:2px solid #e5e7eb; margin-left:2px; font-size:9.5pt; }
    .timeline .tl-meta { color:#6b7280; font-size:8.5pt; }

    @media print { body { font-size:11pt; } .analytics-breakdowns { grid-template-columns:repeat(3,1fr); } }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <div class="company-name">Cranfield Glass Christchurch</div>
      <div class="document-title">Health &amp; Safety Action Report</div>
    </div>

    <div class="meeting-info">
      <div class="info-block">
        <div class="info-label">Report:</div>
        <div>Action Management Tracker</div>
      </div>
      <div class="info-block">
        <div class="info-label">Generated:</div>
        <div>${esc(currentDate)}</div>
      </div>
      <div class="info-block">
        <div class="info-label">Actions in report:</div>
        <div>${items.length}</div>
      </div>
    </div>

    <div class="analytics-dashboard">
      <div class="analytics-bar">
        <div class="analytics-heading">
          <span class="analytics-h1">Action Summary</span>
          <span class="analytics-sub">Generated ${esc(currentDate)}</span>
        </div>
        <div class="analytics-metrics">
          <div class="metric-chip"><span class="metric-chip-num">${stats.total}</span><span class="metric-chip-label">Total</span></div>
          <div class="metric-chip"><span class="metric-chip-num accent-blue">${stats.open}</span><span class="metric-chip-label">Open</span></div>
          <div class="metric-chip"><span class="metric-chip-num accent-green">${stats.completed}</span><span class="metric-chip-label">Completed</span></div>
          <div class="metric-chip"><span class="metric-chip-num accent-red">${stats.overdue}</span><span class="metric-chip-label">Overdue</span></div>
          <div class="metric-chip"><span class="metric-chip-num accent-amber">${stats.highPriority}</span><span class="metric-chip-label">High Priority</span></div>
        </div>
      </div>
      <div class="analytics-breakdowns">
        <div class="breakdown-col">
          <div class="breakdown-title">Status</div>
          ${statusBreakdown.map(([label, count]) => `<div class="bd-row"><span class="bd-label">${esc(label)}</span><span class="bd-count">${count}</span></div>`).join("") || '<div class="bd-row"><span class="bd-label muted">No actions</span></div>'}
        </div>
        <div class="breakdown-col">
          <div class="breakdown-title">Type</div>
          ${typeBreakdown.map(([label, count]) => `<div class="bd-row"><span class="bd-label">${esc(label)}</span><span class="bd-count">${count}</span></div>`).join("") || '<div class="bd-row"><span class="bd-label muted">—</span></div>'}
        </div>
        <div class="breakdown-col">
          <div class="breakdown-title">Priority</div>
          ${priorityBreakdown.map(([label, count]) => `<div class="bd-row"><span class="bd-label">${esc(label)}</span><span class="bd-count">${count}</span></div>`).join("") || '<div class="bd-row"><span class="bd-label muted">—</span></div>'}
        </div>
      </div>
    </div>

    <div class="section-header">I. Actions Overview</div>
    ${items.length === 0 ? '<p class="muted">No actions match the current filters.</p>' : `
    <table class="items-table">
      <colgroup><col style="width:5%;"><col style="width:45%;"><col style="width:13%;"><col style="width:20%;"><col style="width:17%;"></colgroup>
      <thead>
        <tr><th>#</th><th>Action</th><th>Priority</th><th>Status</th><th>Due</th></tr>
      </thead>
      <tbody>${overviewRows}</tbody>
    </table>`}

    ${items.length > 0 ? `<div class="section-header" style="break-before:page;">II. Action Details</div>${detailBlocks}` : ""}
  </div>

  <script>
    function addPrintButton() {
      if (document.querySelector('.print-button')) return;
      var btn = document.createElement('button');
      btn.className = 'print-button';
      btn.textContent = '\u{1F5A8}\uFE0F Print Actions Report';
      btn.addEventListener('click', function () { window.print(); });
      document.body.appendChild(btn);
    }
    window.PagedConfig = { auto: true, after: addPrintButton };
    setTimeout(addPrintButton, 8000);
  </script>
  <script>${PAGEDJS_SCRIPT}</script>
</body>
</html>`;
}

function priorityBadge(priority?: string): string {
  const p = priority || "Not Set";
  const colors: Record<string, string> = {
    High: "background:#fee2e2;color:#b91c1c;",
    Medium: "background:#fef3c7;color:#b45309;",
    Low: "background:#dcfce7;color:#15803d;",
  };
  return `<span class="badge" style="${colors[p] || "background:#f1f5f9;color:#475569;"}">${esc(p)}</span>`;
}

function statusBadge(item: ActionExportItem): string {
  const status = item.actionStatus || "Not Started";
  const colors: Record<string, string> = {
    Completed: "background:#dcfce7;color:#15803d;",
    "In Progress": "background:#dbeafe;color:#1d4ed8;",
    "On Hold": "background:#ffedd5;color:#c2410c;",
    "Ready to Close": "background:#d1fae5;color:#047857;",
  };
  const style = colors[status] || "background:#f1f5f9;color:#475569;";
  let html = `<span class="badge" style="${style}">${esc(status)}</span>`;
  if (status === "On Hold" && item.reconsiderDate) {
    html += `<div class="muted" style="margin-top:3px;">Revisit ${esc(fmtDate(item.reconsiderDate))}</div>`;
  }
  return html;
}

function renderDetailBlock(item: ActionExportItem, index: number): string {
  const due = dueStatus(item);
  const priorityBorder: Record<string, string> = {
    High: "#dc2626", Medium: "#d97706", Low: "#16a34a",
  };
  const border = isArchived(item.actionStatus) ? "#16a34a" : (priorityBorder[item.actionPriority || ""] || "#6b7280");

  const activity = (item.activity || []).slice().sort(
    (a, b) => new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime(),
  );

  const investigation = hasInvestigation(item) && item.investigation ? `
    <div class="ab-sub-title">Near Miss Investigation</div>
    <div class="inv-box">
      ${item.investigation.riskLevel ? `<div class="ab-field"><strong>Risk level:</strong> ${esc(item.investigation.riskLevel)}</div>` : ""}
      ${item.investigation.investigatorName ? `<div class="ab-field"><strong>Investigator:</strong> ${esc(item.investigation.investigatorName)}</div>` : ""}
      <div class="ab-field"><strong>Investigator sign-off:</strong> ${esc(investigatorSignOff(item.investigation))}</div>
      <div class="ab-field"><strong>Director sign-off:</strong> ${esc(directorSignOff(item.investigation))}</div>
      ${item.investigation.status ? `<div class="ab-field"><strong>Investigation status:</strong> ${esc(item.investigation.status)}</div>` : ""}
    </div>` : "";

  const timeline = activity.length > 0 ? `
    <div class="ab-sub-title">Activity History</div>
    <ul class="timeline">
      ${activity.map((a) => `
      <li>
        <div><strong>${esc(activityTypeLabel(a.entryType))}:</strong> ${esc(a.content)}</div>
        <div class="tl-meta">${esc(fmtDateTime(a.createdAt))}${a.author ? ` &middot; ${esc(a.author)}` : ""}</div>
      </li>`).join("")}
    </ul>` : "";

  return `
  <div class="action-block" style="border-left-color:${border};">
    <div class="ab-head">
      <div>
        <div class="type-badge" style="background-color:${typeColor(item.type)};">${esc(item.type)}</div>
        <div class="ab-title">${index + 1}. ${esc(item.title || `${item.type} Action`)}</div>
      </div>
      <div style="text-align:right;white-space:nowrap;">
        ${statusBadge(item)}
        ${due ? `<div class="due-badge due-${due.tone}" style="margin-top:4px;">${esc(due.label)}</div>` : ""}
      </div>
    </div>
    <div class="ab-grid">
      <div class="ab-field"><strong>Priority:</strong> ${esc(item.actionPriority || "Not Set")}</div>
      <div class="ab-field"><strong>Assigned to:</strong> ${esc(item.actionAssignedTo || "Unassigned")}</div>
      <div class="ab-field"><strong>Start date:</strong> ${esc(fmtDate(item.actionStartDate) || "Not set")}</div>
      <div class="ab-field"><strong>Due date:</strong> ${esc(fmtDate(item.actionDueDate) || "Not set")}</div>
      <div class="ab-field"><strong>Type:</strong> ${esc(item.type)}</div>
      <div class="ab-field"><strong>Submitted by:</strong> ${esc(item.submittedBy || "Unknown")}</div>
    </div>
    ${item.actionNotes ? `<div class="ab-notes"><strong>${isArchived(item.actionStatus) ? "Outcome" : "Progress notes"}:</strong> ${esc(item.actionNotes)}</div>` : ""}
    ${investigation}
    ${timeline}
  </div>`;
}

// ─── CSV ──────────────────────────────────────────────────────────────────────

export function generateActionsCSV(items: ActionExportItem[]): string {
  const headers = [
    "Title", "Type", "Priority", "Status", "Revisit Date", "Assigned To",
    "Start Date", "Due Date", "Due Status", "Progress/Outcome Notes",
    "Risk Level", "Investigator", "Investigator Sign-off", "Director Sign-off",
    "Submitted By", "Activity History",
  ];

  const rows = items.map((item) => {
    const due = dueStatus(item);
    const inv = hasInvestigation(item) && item.investigation ? item.investigation : null;
    const activity = (item.activity || [])
      .slice()
      .sort((a, b) => new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime())
      .map((a) => `[${fmtDateTime(a.createdAt)}${a.author ? ` ${a.author}` : ""}] ${activityTypeLabel(a.entryType)}: ${a.content}`)
      .join(" \u2022 ");

    return [
      item.title || "",
      item.type || "",
      item.actionPriority || "Not Set",
      item.actionStatus || "Not Started",
      item.actionStatus === "On Hold" ? fmtDate(item.reconsiderDate) : "",
      item.actionAssignedTo || "Unassigned",
      fmtDate(item.actionStartDate),
      fmtDate(item.actionDueDate),
      due ? due.label : "",
      item.actionNotes || "",
      inv?.riskLevel || "",
      inv?.investigatorName || "",
      inv ? investigatorSignOff(inv) : "",
      inv ? directorSignOff(inv) : "",
      item.submittedBy || "",
      activity,
    ];
  });

  return [headers, ...rows]
    .map((row) => row.map((cell) => `"${String(cell ?? "").replace(/"/g, '""')}"`).join(","))
    .join("\n");
}

// ─── Markdown ─────────────────────────────────────────────────────────────────

export function generateActionsMarkdown(
  items: ActionExportItem[],
  stats: ActionsStats,
  currentDate: string,
): string {
  let md = `# CRANFIELD GLASS CHRISTCHURCH
## Health & Safety Action Report

**Generated:** ${currentDate}  
**Actions in report:** ${items.length}

---

## Action Summary

- **Total Actions:** ${stats.total}
- **Open:** ${stats.open}
- **Completed:** ${stats.completed}
- **Overdue:** ${stats.overdue}
- **High Priority:** ${stats.highPriority}

---

## Actions Overview

| # | Action | Type | Priority | Status | Assigned | Due | Due Status |
| --- | --- | --- | --- | --- | --- | --- | --- |
`;

  const cell = (v: unknown) => String(v ?? "").replace(/\|/g, "\\|").replace(/\n+/g, " ");
  items.forEach((item, i) => {
    const due = dueStatus(item);
    md += `| ${i + 1} | ${cell(item.title || `${item.type} Action`)} | ${cell(item.type)} | ${cell(item.actionPriority || "Not Set")} | ${cell(displayStatus(item))} | ${cell(item.actionAssignedTo || "Unassigned")} | ${cell(fmtDate(item.actionDueDate) || "Not set")} | ${cell(due ? due.label : "")} |\n`;
  });

  if (items.length === 0) {
    md += `\n*No actions match the current filters.*\n`;
    return md;
  }

  md += `\n---\n\n## Action Details\n\n`;

  items.forEach((item, i) => {
    md += `### ${i + 1}. ${item.title || `${item.type} Action`}\n\n`;
    md += `**Type:** ${item.type}  \n`;
    md += `**Priority:** ${item.actionPriority || "Not Set"}  \n`;
    md += `**Status:** ${displayStatus(item)}  \n`;
    md += `**Assigned to:** ${item.actionAssignedTo || "Unassigned"}  \n`;
    md += `**Start date:** ${fmtDate(item.actionStartDate) || "Not set"}  \n`;
    md += `**Due date:** ${fmtDate(item.actionDueDate) || "Not set"}  \n`;
    const due = dueStatus(item);
    if (due) md += `**Due status:** ${due.label}  \n`;
    md += `**Submitted by:** ${item.submittedBy || "Unknown"}  \n`;

    if (item.actionNotes) {
      md += `\n**${isArchived(item.actionStatus) ? "Outcome" : "Progress notes"}:** ${item.actionNotes}\n`;
    }

    if (hasInvestigation(item) && item.investigation) {
      const inv = item.investigation;
      md += `\n**Near Miss Investigation**  \n`;
      if (inv.riskLevel) md += `- Risk level: ${inv.riskLevel}\n`;
      if (inv.investigatorName) md += `- Investigator: ${inv.investigatorName}\n`;
      md += `- Investigator sign-off: ${investigatorSignOff(inv)}\n`;
      md += `- Director sign-off: ${directorSignOff(inv)}\n`;
      if (inv.status) md += `- Investigation status: ${inv.status}\n`;
    }

    const activity = (item.activity || [])
      .slice()
      .sort((a, b) => new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime());
    if (activity.length > 0) {
      md += `\n**Activity History**  \n`;
      activity.forEach((a) => {
        md += `- *${fmtDateTime(a.createdAt)}${a.author ? ` · ${a.author}` : ""}* — **${activityTypeLabel(a.entryType)}:** ${a.content}\n`;
      });
    }

    md += `\n---\n\n`;
  });

  md += `*This document was automatically generated from the Cranfield Glass Health & Safety Management System on ${currentDate}.*\n`;
  return md;
}

// ─── Word ─────────────────────────────────────────────────────────────────────

const WORD_BORDER = { style: BorderStyle.SINGLE, size: 1, color: "D1D5DB" } as const;

function wordCell(text: string, opts: { bold?: boolean; header?: boolean; width?: number } = {}): TableCell {
  return new TableCell({
    children: [new Paragraph({ children: [new TextRun({ text, bold: opts.bold || opts.header, size: 18, color: opts.header ? "374151" : "4B5563" })] })],
    shading: opts.header ? { fill: "F3F4F6" } : undefined,
    width: opts.width ? { size: opts.width, type: WidthType.PERCENTAGE } : undefined,
    margins: { top: 60, bottom: 60, left: 80, right: 80 },
  });
}

export async function generateActionsWordDoc(
  items: ActionExportItem[],
  stats: ActionsStats,
  currentDate: string,
): Promise<Buffer> {
  const children: any[] = [];

  children.push(
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "CRANFIELD GLASS CHRISTCHURCH", bold: true, size: 44, color: "1F2937" })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Health & Safety Action Report", bold: true, size: 30, color: "1F2937" })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [new TextRun({ text: `Generated ${currentDate}  ·  ${items.length} action${items.length === 1 ? "" : "s"}`, size: 20, color: "6B7280", italics: true })] }),
  );

  // Summary line
  children.push(new Paragraph({
    spacing: { after: 200 },
    children: [new TextRun({ text: `Summary — Total: ${stats.total}   Open: ${stats.open}   Completed: ${stats.completed}   Overdue: ${stats.overdue}   High Priority: ${stats.highPriority}`, size: 20, color: "374151", bold: true })],
  }));

  if (items.length === 0) {
    children.push(new Paragraph({ children: [new TextRun({ text: "No actions match the current filters.", size: 20, color: "6B7280", italics: true })] }));
  } else {
    children.push(new Paragraph({ spacing: { before: 120, after: 120 }, children: [new TextRun({ text: "Actions Overview", bold: true, size: 26, color: "1F2937" })] }));

    const headerRow = new TableRow({
      tableHeader: true,
      children: [
        wordCell("#", { header: true, width: 5 }),
        wordCell("Action", { header: true, width: 35 }),
        wordCell("Type", { header: true, width: 14 }),
        wordCell("Priority", { header: true, width: 12 }),
        wordCell("Status", { header: true, width: 18 }),
        wordCell("Due", { header: true, width: 16 }),
      ],
    });

    const rows = items.map((item, i) => {
      const due = dueStatus(item);
      const dueText = [fmtDate(item.actionDueDate) || "Not set", due ? `(${due.label})` : ""].filter(Boolean).join(" ");
      return new TableRow({
        children: [
          wordCell(String(i + 1)),
          wordCell(item.title || `${item.type} Action`),
          wordCell(item.type || ""),
          wordCell(item.actionPriority || "Not Set"),
          wordCell(displayStatus(item)),
          wordCell(dueText),
        ],
      });
    });

    children.push(new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: { top: WORD_BORDER, bottom: WORD_BORDER, left: WORD_BORDER, right: WORD_BORDER, insideHorizontal: WORD_BORDER, insideVertical: WORD_BORDER },
      rows: [headerRow, ...rows],
    }));

    children.push(new Paragraph({ spacing: { before: 280, after: 120 }, pageBreakBefore: true, children: [new TextRun({ text: "Action Details", bold: true, size: 26, color: "1F2937" })] }));

    items.forEach((item, i) => {
      const due = dueStatus(item);
      children.push(new Paragraph({ spacing: { before: 160, after: 60 }, keepNext: true, children: [new TextRun({ text: `${i + 1}. ${item.title || `${item.type} Action`}`, bold: true, size: 24, color: "111827" })] }));

      const field = (label: string, value: string) => new Paragraph({ spacing: { after: 20 }, children: [
        new TextRun({ text: `${label}: `, bold: true, size: 20, color: "374151" }),
        new TextRun({ text: value, size: 20, color: "4B5563" }),
      ] });

      children.push(field("Type", item.type || ""));
      children.push(field("Priority", item.actionPriority || "Not Set"));
      children.push(field("Status", displayStatus(item)));
      children.push(field("Assigned to", item.actionAssignedTo || "Unassigned"));
      children.push(field("Start date", fmtDate(item.actionStartDate) || "Not set"));
      children.push(field("Due date", fmtDate(item.actionDueDate) || "Not set"));
      if (due) children.push(field("Due status", due.label));
      children.push(field("Submitted by", item.submittedBy || "Unknown"));

      if (item.actionNotes) {
        children.push(field(isArchived(item.actionStatus) ? "Outcome" : "Progress notes", item.actionNotes));
      }

      if (hasInvestigation(item) && item.investigation) {
        const inv = item.investigation;
        children.push(new Paragraph({ spacing: { before: 80, after: 40 }, children: [new TextRun({ text: "Near Miss Investigation", bold: true, size: 20, color: "9A3412" })] }));
        if (inv.riskLevel) children.push(field("Risk level", inv.riskLevel));
        if (inv.investigatorName) children.push(field("Investigator", inv.investigatorName));
        children.push(field("Investigator sign-off", investigatorSignOff(inv)));
        children.push(field("Director sign-off", directorSignOff(inv)));
        if (inv.status) children.push(field("Investigation status", inv.status));
      }

      const activity = (item.activity || [])
        .slice()
        .sort((a, b) => new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime());
      if (activity.length > 0) {
        children.push(new Paragraph({ spacing: { before: 80, after: 40 }, children: [new TextRun({ text: "Activity History", bold: true, size: 20, color: "374151" })] }));
        activity.forEach((a) => {
          children.push(new Paragraph({ spacing: { after: 20 }, bullet: { level: 0 }, children: [
            new TextRun({ text: `${activityTypeLabel(a.entryType)}: `, bold: true, size: 18, color: "374151" }),
            new TextRun({ text: a.content, size: 18, color: "4B5563" }),
            new TextRun({ text: `  (${fmtDateTime(a.createdAt)}${a.author ? ` · ${a.author}` : ""})`, size: 16, color: "9CA3AF", italics: true }),
          ] }));
        });
      }
    });
  }

  const doc = new Document({ sections: [{ properties: {}, children }] });
  return await Packer.toBuffer(doc);
}

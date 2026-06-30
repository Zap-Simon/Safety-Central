/**
 * Near Miss Register export engine.
 *
 * Turns the Near Miss cards into a single, Cranfield-branded "Near Miss Register"
 * built on the SAME professional foundation the meeting-minutes and Actions
 * reports use: an A4 print-ready HTML report (Paged.js "Page X of Y" footers,
 * repeating table headers, break-inside protection) plus matching CSV, Markdown
 * and Word outputs so the format set stays in lockstep.
 *
 * Unlike the Actions report (which only carries items that became tracked
 * actions) the register lists EVERY Near Miss card and folds in its full
 * investigation: event details, the what/how narrative, contributing factors,
 * the risk assessment, the hazards table, resulting actions and the dual
 * sign-off. It mirrors the SharePoint Near Miss - Accident Safety Register.
 */

import {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  BorderStyle, WidthType, AlignmentType,
} from "docx";
import { PAGEDJS_BASE64 } from "./assets/pagedjs";

const PAGEDJS_SCRIPT = Buffer.from(PAGEDJS_BASE64, "base64").toString("utf-8");

// ─── Shared types ───────────────────────────────────────────────────────────

export interface NearMissHazard {
  hazard?: string;
  likelihood?: string;
  consequence?: string;
  risk?: string;
  control?: string;
}

export interface NearMissResultingAction {
  description?: string;
  assignedTo?: string;
  completed?: boolean;
}

export interface NearMissInvestigationData {
  investigatorName?: string;
  siteJob?: string;
  eventDate?: string;
  eventTime?: string;
  eventType?: string;
  involvedPersons?: string;
  witnesses?: string;
  eventDescription?: string;
  contributingFactors?: string;
  likelihood?: string;
  consequence?: string;
  riskLevel?: string;
  treatmentGiven?: string;
  hazards?: NearMissHazard[];
  resultingActions?: NearMissResultingAction[];
  investigatorSignature?: string | null;
  investigatorSignedAt?: string | null;
  directorName?: string | null;
  directorSignature?: string | null;
  signedAt?: string | null;
  status?: string;
}

export interface NearMissRegisterItem {
  id: string;
  title?: string;
  description?: string;          // What happened
  secondaryDescription?: string; // How it happened
  status?: string;               // meeting status (Submitted / Actioned)
  actionStatus?: string;
  meetingDate?: string;
  submittedBy?: string;
  submittedDate?: string;
  eventType?: string;
  investigation?: NearMissInvestigationData | null;
}

export interface NearMissRegisterStats {
  total: number;
  investigated: number;   // investigation Complete (both sign-offs)
  inProgress: number;     // investigation started but not complete
  notStarted: number;     // no investigation record
  highRisk: number;       // risk level High or Extreme
  closed: number;         // actionStatus Completed / Ready to Close
}

// ─── Helpers (kept in lockstep with the other export engines) ─────────────────

const ARCHIVED_STATUSES = new Set(["Completed", "Ready to Close"]);

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

/** Near Miss risk palette — identical to the single-item investigation report. */
function riskBg(r?: string): string {
  switch (r) {
    case "Extreme": return "#000000";
    case "High": return "#ef4444";
    case "Moderate": return "#eab308";
    case "Low": return "#22c55e";
    default: return "#e5e7eb";
  }
}
function riskFg(r?: string): string {
  if (!r) return "#374151";
  return r === "Moderate" ? "#000000" : "#ffffff";
}

function isArchived(status?: string): boolean {
  return ARCHIVED_STATUSES.has(status || "");
}

/** Investigation progress for a card. */
function investigationState(item: NearMissRegisterItem): { label: string; tone: "done" | "progress" | "draft" | "none" } {
  const inv = item.investigation;
  if (!inv) return { label: "Not started", tone: "none" };
  const s = inv.status || "Draft";
  if (s === "Complete") return { label: "Complete", tone: "done" };
  if (s === "In Progress") return { label: "In progress", tone: "progress" };
  return { label: "Draft", tone: "draft" };
}

function meetingStatusLabel(item: NearMissRegisterItem): string {
  if (isArchived(item.actionStatus)) return item.actionStatus === "Ready to Close" ? "Ready to Close" : "Closed";
  return item.status || "Submitted";
}

function investigatorSignOff(inv: NearMissInvestigationData): string {
  if (inv.investigatorSignature) {
    return `Signed${inv.investigatorName ? ` by ${inv.investigatorName}` : ""}${inv.investigatorSignedAt ? ` on ${fmtDate(inv.investigatorSignedAt)}` : ""}`;
  }
  return inv.investigatorName ? `${inv.investigatorName} (not yet signed)` : "Not yet signed";
}

function approverSignOff(inv: NearMissInvestigationData): string {
  if (inv.directorSignature) {
    return `Signed${inv.directorName ? ` by ${inv.directorName}` : ""}${inv.signedAt ? ` on ${fmtDate(inv.signedAt)}` : ""}`;
  }
  return inv.directorName ? `${inv.directorName} (awaiting sign-off)` : "Awaiting sign-off";
}

function eventTypeOf(item: NearMissRegisterItem): string {
  return item.investigation?.eventType || item.eventType || "Near Miss";
}

function countBy<T>(items: T[], key: (i: T) => string): Array<[string, number]> {
  const counts: Record<string, number> = {};
  for (const item of items) {
    const k = key(item) || "—";
    counts[k] = (counts[k] || 0) + 1;
  }
  return Object.entries(counts).sort((a, b) => b[1] - a[1]);
}

// ─── Stats ────────────────────────────────────────────────────────────────────

export function computeNearMissRegisterStats(items: NearMissRegisterItem[]): NearMissRegisterStats {
  const stats: NearMissRegisterStats = {
    total: items.length, investigated: 0, inProgress: 0, notStarted: 0, highRisk: 0, closed: 0,
  };
  for (const item of items) {
    const state = investigationState(item);
    if (state.tone === "done") stats.investigated += 1;
    else if (state.tone === "none") stats.notStarted += 1;
    else stats.inProgress += 1;

    const risk = item.investigation?.riskLevel;
    if (risk === "High" || risk === "Extreme") stats.highRisk += 1;
    if (isArchived(item.actionStatus)) stats.closed += 1;
  }
  return stats;
}

// ─── HTML report ──────────────────────────────────────────────────────────────

export function generateNearMissRegisterHTML(
  items: NearMissRegisterItem[],
  stats: NearMissRegisterStats,
  currentDate: string,
  dateRangeLabel?: string,
): string {
  const statusBreakdown = countBy(items, (i) => meetingStatusLabel(i));
  const typeBreakdown = countBy(items, (i) => eventTypeOf(i));
  const riskBreakdown = countBy(items, (i) => i.investigation?.riskLevel || "Not assessed");

  const overviewRows = items.map((item, index) => {
    const state = investigationState(item);
    const risk = item.investigation?.riskLevel;
    return `
      <tr>
        <td style="text-align:center;">${index + 1}</td>
        <td>
          <div class="nm-title">${esc(item.title || "Near Miss")}</div>
          <div class="muted"><strong>By:</strong> ${esc(item.submittedBy || "Unknown")}</div>
        </td>
        <td>${esc(fmtDate(item.investigation?.eventDate) || fmtDate(item.submittedDate) || "—")}</td>
        <td>${esc(eventTypeOf(item))}</td>
        <td>${risk ? `<span class="risk-badge" style="background:${riskBg(risk)};color:${riskFg(risk)};">${esc(risk)}</span>` : '<span class="muted">—</span>'}</td>
        <td><span class="inv-badge inv-${state.tone}">${esc(state.label)}</span></td>
        <td>${esc(meetingStatusLabel(item))}</td>
      </tr>`;
  }).join("");

  const detailBlocks = items.map((item, index) => renderDetailBlock(item, index)).join("");

  return `
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Cranfield Glass - Near Miss Register</title>
  <style>
    @page {
      size: A4;
      margin: 1.1cm 1.4cm 1.5cm 1.4cm;
      @bottom-center {
        content: "Cranfield Glass Christchurch  |  Near Miss - Accident Safety Register";
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
    .print-button { position:fixed; top:20px; right:20px; background:#ea580c; color:#fff; border:none; padding:12px 24px; border-radius:6px; cursor:pointer; font-size:14px; font-weight:500; box-shadow:0 2px 4px rgba(0,0,0,0.1); z-index:1000; }
    .print-button:hover { background:#c2410c; }

    .header { text-align:center; margin-bottom:24px; border-bottom:2px solid #ea580c; padding-bottom:18px; }
    .company-name { font-size:24pt; font-weight:bold; color:#1e3a5f; margin-bottom:5px; }
    .document-title { font-size:18pt; font-weight:bold; color:#ea580c; }
    .document-sub { font-size:9pt; color:#6b7280; margin-top:4px; }

    .meeting-info { display:flex; justify-content:space-between; margin-bottom:22px; padding:15px; background:#f8f9fa; border:1px solid #dee2e6; }
    .info-block { flex:1; }
    .info-label { font-weight:bold; margin-bottom:5px; }

    .section-header { font-size:14pt; font-weight:bold; color:#ea580c; margin-bottom:15px; border-bottom:1px solid #ea580c; padding-bottom:5px; }

    /* Summary dashboard */
    .analytics-dashboard { margin-bottom:22px; background:#fff; border:1px solid #e5e7eb; border-radius:8px; overflow:hidden; page-break-inside:avoid; break-inside:avoid; }
    .analytics-bar { display:flex; align-items:center; justify-content:space-between; gap:16px; padding:10px 16px; background:#fff7ed; border-bottom:1px solid #fed7aa; }
    .analytics-heading { display:flex; flex-direction:column; }
    .analytics-h1 { font-size:13pt; font-weight:bold; color:#1f2937; line-height:1.2; }
    .analytics-sub { font-size:7.5pt; color:#9ca3af; margin-top:2px; }
    .analytics-metrics { display:flex; gap:10px; flex-wrap:wrap; }
    .metric-chip { display:flex; flex-direction:column; align-items:center; justify-content:center; min-width:74px; padding:6px 12px; background:#fff; border:1px solid #e5e7eb; border-radius:6px; }
    .metric-chip-num { font-size:16pt; font-weight:bold; color:#1f2937; line-height:1.1; }
    .metric-chip-num.accent-orange { color:#ea580c; }
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
    .items-table th { background:#f8f9fa; padding:9px 10px; text-align:left; font-weight:bold; border:1px solid #dee2e6; font-size:9.5pt; overflow-wrap:anywhere; }
    .items-table td { padding:9px 10px; border:1px solid #dee2e6; vertical-align:top; font-size:9.5pt; overflow-wrap:anywhere; }
    .items-table th:nth-child(1), .items-table td:nth-child(1) { width:4%; }
    .items-table th:nth-child(2), .items-table td:nth-child(2) { width:34%; }
    .items-table th:nth-child(3), .items-table td:nth-child(3) { width:13%; }
    .items-table th:nth-child(4), .items-table td:nth-child(4) { width:13%; }
    .items-table th:nth-child(5), .items-table td:nth-child(5) { width:11%; }
    .items-table th:nth-child(6), .items-table td:nth-child(6) { width:13%; }
    .items-table th:nth-child(7), .items-table td:nth-child(7) { width:12%; }
    .items-table tr { page-break-inside:avoid; break-inside:avoid; }
    .items-table thead { display:table-header-group; }

    .nm-title { font-weight:bold; margin-bottom:3px; }
    .muted { color:#6b7280; font-size:9pt; }

    .risk-badge { display:inline-block; padding:2px 8px; border-radius:4px; font-size:8.5pt; font-weight:bold; }
    .inv-badge { display:inline-block; padding:3px 8px; border-radius:10px; font-size:8.5pt; font-weight:bold; }
    .inv-done { background:#dcfce7; color:#15803d; }
    .inv-progress { background:#dbeafe; color:#1d4ed8; }
    .inv-draft { background:#fef3c7; color:#b45309; }
    .inv-none { background:#f1f5f9; color:#64748b; }

    /* Detail blocks — never split a single near miss across a page */
    .nm-block { border:1px solid #e5e7eb; border-left:4px solid #ea580c; border-radius:6px; padding:14px 16px; margin-bottom:14px; page-break-inside:avoid; break-inside:avoid; }
    .nm-head { display:flex; justify-content:space-between; align-items:flex-start; gap:12px; margin-bottom:8px; }
    .nm-block-title { font-size:12pt; font-weight:bold; color:#111827; }
    .nm-grid { display:grid; grid-template-columns:repeat(2,1fr); gap:4px 24px; margin:8px 0; font-size:10pt; }
    .nm-field { color:#374151; }
    .nm-field strong { color:#111827; }
    .nm-notes { background:#f9fafb; border:1px solid #eef2f7; border-radius:6px; padding:10px; margin-top:8px; font-size:10pt; white-space:pre-wrap; }
    .nm-sub-title { font-size:9pt; font-weight:bold; text-transform:uppercase; letter-spacing:0.4px; color:#9a3412; margin:12px 0 6px; }
    .inv-box { background:#fff7ed; border:1px solid #fed7aa; border-radius:6px; padding:10px; font-size:10pt; }
    .inv-box .nm-field { margin-bottom:2px; }
    .sub-table { width:100%; border-collapse:collapse; font-size:9pt; margin-top:4px; }
    .sub-table th { background:#1e3a5f; color:#fff; padding:6px 8px; text-align:left; font-size:8.5pt; }
    .sub-table td { padding:6px 8px; border-bottom:1px solid #e5e7eb; vertical-align:top; }

    @media print { body { font-size:11pt; } .analytics-breakdowns { grid-template-columns:repeat(3,1fr); } }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <div class="company-name">Cranfield Glass Christchurch</div>
      <div class="document-title">Near Miss - Accident Safety Register</div>
      <div class="document-sub">A complete register of reported near misses and their investigations</div>
    </div>

    <div class="meeting-info">
      <div class="info-block">
        <div class="info-label">Report:</div>
        <div>Near Miss Register</div>
      </div>
      <div class="info-block">
        <div class="info-label">Generated:</div>
        <div>${esc(currentDate)}</div>
      </div>
      ${dateRangeLabel ? `<div class="info-block">
        <div class="info-label">Period covered:</div>
        <div>${esc(dateRangeLabel)}</div>
      </div>` : ''}
      <div class="info-block">
        <div class="info-label">Near misses in register:</div>
        <div>${items.length}</div>
      </div>
    </div>

    <div class="analytics-dashboard">
      <div class="analytics-bar">
        <div class="analytics-heading">
          <span class="analytics-h1">Register Summary</span>
          <span class="analytics-sub">Generated ${esc(currentDate)}</span>
        </div>
        <div class="analytics-metrics">
          <div class="metric-chip"><span class="metric-chip-num">${stats.total}</span><span class="metric-chip-label">Total</span></div>
          <div class="metric-chip"><span class="metric-chip-num accent-green">${stats.investigated}</span><span class="metric-chip-label">Investigated</span></div>
          <div class="metric-chip"><span class="metric-chip-num accent-amber">${stats.inProgress}</span><span class="metric-chip-label">In Progress</span></div>
          <div class="metric-chip"><span class="metric-chip-num">${stats.notStarted}</span><span class="metric-chip-label">Not Started</span></div>
          <div class="metric-chip"><span class="metric-chip-num accent-red">${stats.highRisk}</span><span class="metric-chip-label">High / Extreme</span></div>
        </div>
      </div>
      <div class="analytics-breakdowns">
        <div class="breakdown-col">
          <div class="breakdown-title">Meeting Status</div>
          ${statusBreakdown.map(([label, count]) => `<div class="bd-row"><span class="bd-label">${esc(label)}</span><span class="bd-count">${count}</span></div>`).join("") || '<div class="bd-row"><span class="bd-label muted">None</span></div>'}
        </div>
        <div class="breakdown-col">
          <div class="breakdown-title">Event Type</div>
          ${typeBreakdown.map(([label, count]) => `<div class="bd-row"><span class="bd-label">${esc(label)}</span><span class="bd-count">${count}</span></div>`).join("") || '<div class="bd-row"><span class="bd-label muted">—</span></div>'}
        </div>
        <div class="breakdown-col">
          <div class="breakdown-title">Risk Level</div>
          ${riskBreakdown.map(([label, count]) => `<div class="bd-row"><span class="bd-label">${esc(label)}</span><span class="bd-count">${count}</span></div>`).join("") || '<div class="bd-row"><span class="bd-label muted">—</span></div>'}
        </div>
      </div>
    </div>

    <div class="section-header">I. Near Miss Register</div>
    ${items.length === 0 ? '<p class="muted">No near misses match the current filters.</p>' : `
    <table class="items-table">
      <colgroup><col style="width:4%;"><col style="width:34%;"><col style="width:13%;"><col style="width:13%;"><col style="width:11%;"><col style="width:13%;"><col style="width:12%;"></colgroup>
      <thead>
        <tr><th>#</th><th>Event</th><th>Event Date</th><th>Type</th><th>Risk</th><th>Investigation</th><th>Meeting</th></tr>
      </thead>
      <tbody>${overviewRows}</tbody>
    </table>`}

    ${items.length > 0 ? `<div class="section-header" style="break-before:page;">II. Near Miss Details</div>${detailBlocks}` : ""}
  </div>

  <script>
    function addPrintButton() {
      if (document.querySelector('.print-button')) return;
      var btn = document.createElement('button');
      btn.className = 'print-button';
      btn.textContent = '\u{1F5A8}\uFE0F Print Near Miss Register';
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

function renderDetailBlock(item: NearMissRegisterItem, index: number): string {
  const inv = item.investigation;
  const state = investigationState(item);
  const risk = inv?.riskLevel;

  const hazards = inv?.hazards || [];
  const actions = inv?.resultingActions || [];

  const hazardTable = hazards.length > 0 ? `
    <div class="nm-sub-title">Hazards Identified</div>
    <table class="sub-table">
      <thead><tr><th style="width:34%;">Hazard</th><th style="width:14%;">Likelihood</th><th style="width:14%;">Consequence</th><th style="width:12%;">Risk</th><th style="width:26%;">Control</th></tr></thead>
      <tbody>
        ${hazards.map((h) => `<tr>
          <td>${esc(h.hazard) || "—"}</td>
          <td>${esc(h.likelihood) || "—"}</td>
          <td>${esc(h.consequence) || "—"}</td>
          <td>${h.risk ? `<span class="risk-badge" style="background:${riskBg(h.risk)};color:${riskFg(h.risk)};">${esc(h.risk)}</span>` : "—"}</td>
          <td>${esc(h.control) || "—"}</td>
        </tr>`).join("")}
      </tbody>
    </table>` : "";

  const actionTable = actions.length > 0 ? `
    <div class="nm-sub-title">Resulting Actions</div>
    <table class="sub-table">
      <thead><tr><th style="width:58%;">Action</th><th style="width:27%;">Assigned to</th><th style="width:15%;">Status</th></tr></thead>
      <tbody>
        ${actions.map((a) => `<tr>
          <td>${esc(a.description) || "—"}</td>
          <td>${esc(a.assignedTo) || "—"}</td>
          <td>${a.completed ? '<span class="risk-badge" style="background:#22c55e;color:#fff;">Done</span>' : '<span class="risk-badge" style="background:#e5e7eb;color:#374151;">Pending</span>'}</td>
        </tr>`).join("")}
      </tbody>
    </table>` : "";

  const investigationDetail = inv ? `
    <div class="nm-grid">
      <div class="nm-field"><strong>Investigator:</strong> ${esc(inv.investigatorName) || "—"}</div>
      <div class="nm-field"><strong>Site / Job:</strong> ${esc(inv.siteJob) || "—"}</div>
      <div class="nm-field"><strong>Event date:</strong> ${esc(fmtDate(inv.eventDate)) || "—"}</div>
      <div class="nm-field"><strong>Event time:</strong> ${esc(inv.eventTime) || "—"}</div>
      <div class="nm-field"><strong>Involved persons:</strong> ${esc(inv.involvedPersons) || "—"}</div>
      <div class="nm-field"><strong>Witnesses:</strong> ${esc(inv.witnesses) || "—"}</div>
    </div>
    ${inv.eventDescription ? `<div class="nm-notes"><strong>Investigation findings:</strong> ${esc(inv.eventDescription)}</div>` : ""}
    ${inv.contributingFactors ? `<div class="nm-notes"><strong>Contributing factors:</strong> ${esc(inv.contributingFactors)}</div>` : ""}
    <div class="nm-sub-title">Risk Assessment</div>
    <div class="inv-box">
      <div class="nm-field"><strong>Likelihood:</strong> ${esc(inv.likelihood) || "—"}</div>
      <div class="nm-field"><strong>Consequence:</strong> ${esc(inv.consequence) || "—"}</div>
      <div class="nm-field"><strong>Risk level:</strong> ${risk ? `<span class="risk-badge" style="background:${riskBg(risk)};color:${riskFg(risk)};">${esc(risk)}</span>` : "Not assessed"}</div>
    </div>
    ${inv.treatmentGiven ? `<div class="nm-notes"><strong>Treatment given:</strong> ${esc(inv.treatmentGiven)}</div>` : ""}
    ${hazardTable}
    ${actionTable}
    <div class="nm-sub-title">Sign-off</div>
    <div class="inv-box">
      <div class="nm-field"><strong>Investigator:</strong> ${esc(investigatorSignOff(inv))}</div>
      <div class="nm-field"><strong>Approver / Manager:</strong> ${esc(approverSignOff(inv))}</div>
      <div class="nm-field"><strong>Investigation status:</strong> ${esc(inv.status || "Draft")}</div>
    </div>` : '<div class="nm-notes muted">No investigation has been recorded for this near miss yet.</div>';

  return `
  <div class="nm-block">
    <div class="nm-head">
      <div>
        <div class="nm-block-title">${index + 1}. ${esc(item.title || "Near Miss")}</div>
        <div class="muted">${esc(eventTypeOf(item))} &middot; Submitted by ${esc(item.submittedBy || "Unknown")}${item.submittedDate ? ` &middot; ${esc(fmtDate(item.submittedDate))}` : ""}</div>
      </div>
      <div style="text-align:right;white-space:nowrap;">
        <span class="inv-badge inv-${state.tone}">${esc(state.label)}</span>
        ${risk ? `<div class="risk-badge" style="margin-top:4px;background:${riskBg(risk)};color:${riskFg(risk)};">${esc(risk)} risk</div>` : ""}
      </div>
    </div>
    <div class="nm-grid">
      <div class="nm-field"><strong>Meeting status:</strong> ${esc(meetingStatusLabel(item))}</div>
      <div class="nm-field"><strong>Meeting date:</strong> ${esc(fmtDate(item.meetingDate)) || "—"}</div>
    </div>
    ${item.description ? `<div class="nm-notes"><strong>What happened:</strong> ${esc(item.description)}</div>` : ""}
    ${item.secondaryDescription ? `<div class="nm-notes"><strong>How it happened:</strong> ${esc(item.secondaryDescription)}</div>` : ""}
    <div class="nm-sub-title">Investigation</div>
    ${investigationDetail}
  </div>`;
}

// ─── CSV ──────────────────────────────────────────────────────────────────────

export function generateNearMissRegisterCSV(items: NearMissRegisterItem[]): string {
  const headers = [
    "Title", "Event Type", "Event Date", "Submitted By", "Submitted Date",
    "Meeting Date", "Meeting Status", "What Happened", "How It Happened",
    "Investigation Status", "Investigator", "Site/Job", "Involved Persons", "Witnesses",
    "Investigation Findings", "Contributing Factors", "Likelihood", "Consequence",
    "Risk Level", "Treatment Given", "Hazards", "Resulting Actions",
    "Investigator Sign-off", "Approver Sign-off",
  ];

  const rows = items.map((item) => {
    const inv = item.investigation;
    const state = investigationState(item);
    const hazards = (inv?.hazards || [])
      .map((h) => `${h.hazard || ""} (${h.risk || "—"}; control: ${h.control || "—"})`)
      .join(" • ");
    const actions = (inv?.resultingActions || [])
      .map((a) => `${a.description || ""} → ${a.assignedTo || "Unassigned"} [${a.completed ? "Done" : "Pending"}]`)
      .join(" • ");

    return [
      item.title || "",
      eventTypeOf(item),
      fmtDate(inv?.eventDate),
      item.submittedBy || "",
      fmtDate(item.submittedDate),
      fmtDate(item.meetingDate),
      meetingStatusLabel(item),
      item.description || "",
      item.secondaryDescription || "",
      state.label,
      inv?.investigatorName || "",
      inv?.siteJob || "",
      inv?.involvedPersons || "",
      inv?.witnesses || "",
      inv?.eventDescription || "",
      inv?.contributingFactors || "",
      inv?.likelihood || "",
      inv?.consequence || "",
      inv?.riskLevel || "",
      inv?.treatmentGiven || "",
      hazards,
      actions,
      inv ? investigatorSignOff(inv) : "",
      inv ? approverSignOff(inv) : "",
    ];
  });

  return [headers, ...rows]
    .map((row) => row.map((cell) => `"${String(cell ?? "").replace(/"/g, '""')}"`).join(","))
    .join("\n");
}

// ─── Markdown ─────────────────────────────────────────────────────────────────

export function generateNearMissRegisterMarkdown(
  items: NearMissRegisterItem[],
  stats: NearMissRegisterStats,
  currentDate: string,
  dateRangeLabel?: string,
): string {
  let md = `# CRANFIELD GLASS CHRISTCHURCH
## Near Miss - Accident Safety Register

**Generated:** ${currentDate}  
${dateRangeLabel ? `**Period covered:** ${dateRangeLabel}  \n` : ''}**Near misses in register:** ${items.length}

---

## Register Summary

- **Total:** ${stats.total}
- **Investigated:** ${stats.investigated}
- **In Progress:** ${stats.inProgress}
- **Not Started:** ${stats.notStarted}
- **High / Extreme risk:** ${stats.highRisk}

---

## Register Overview

| # | Event | Event Date | Type | Risk | Investigation | Meeting |
| --- | --- | --- | --- | --- | --- | --- |
`;

  const cell = (v: unknown) => String(v ?? "").replace(/\|/g, "\\|").replace(/\n+/g, " ");
  items.forEach((item, i) => {
    const state = investigationState(item);
    md += `| ${i + 1} | ${cell(item.title || "Near Miss")} | ${cell(fmtDate(item.investigation?.eventDate) || fmtDate(item.submittedDate))} | ${cell(eventTypeOf(item))} | ${cell(item.investigation?.riskLevel || "—")} | ${cell(state.label)} | ${cell(meetingStatusLabel(item))} |\n`;
  });

  if (items.length === 0) {
    md += `\n*No near misses match the current filters.*\n`;
    return md;
  }

  md += `\n---\n\n## Near Miss Details\n\n`;

  items.forEach((item, i) => {
    const inv = item.investigation;
    md += `### ${i + 1}. ${item.title || "Near Miss"}\n\n`;
    md += `**Event type:** ${eventTypeOf(item)}  \n`;
    md += `**Submitted by:** ${item.submittedBy || "Unknown"}  \n`;
    md += `**Meeting status:** ${meetingStatusLabel(item)}  \n`;
    md += `**Meeting date:** ${fmtDate(item.meetingDate) || "—"}  \n`;
    if (item.description) md += `\n**What happened:** ${item.description}\n`;
    if (item.secondaryDescription) md += `\n**How it happened:** ${item.secondaryDescription}\n`;

    if (inv) {
      md += `\n**Investigation**  \n`;
      if (inv.investigatorName) md += `- Investigator: ${inv.investigatorName}\n`;
      if (inv.siteJob) md += `- Site / Job: ${inv.siteJob}\n`;
      if (inv.eventDate) md += `- Event date: ${fmtDate(inv.eventDate)}\n`;
      if (inv.eventTime) md += `- Event time: ${inv.eventTime}\n`;
      if (inv.involvedPersons) md += `- Involved persons: ${inv.involvedPersons}\n`;
      if (inv.witnesses) md += `- Witnesses: ${inv.witnesses}\n`;
      if (inv.eventDescription) md += `- Findings: ${inv.eventDescription}\n`;
      if (inv.contributingFactors) md += `- Contributing factors: ${inv.contributingFactors}\n`;
      if (inv.likelihood || inv.consequence || inv.riskLevel) {
        md += `- Risk: ${inv.likelihood || "—"} × ${inv.consequence || "—"} = ${inv.riskLevel || "Not assessed"}\n`;
      }
      if (inv.treatmentGiven) md += `- Treatment given: ${inv.treatmentGiven}\n`;
      (inv.hazards || []).forEach((h) => {
        md += `- Hazard: ${h.hazard || "—"} (${h.risk || "—"}; control: ${h.control || "—"})\n`;
      });
      (inv.resultingActions || []).forEach((a) => {
        md += `- Action: ${a.description || "—"} → ${a.assignedTo || "Unassigned"} [${a.completed ? "Done" : "Pending"}]\n`;
      });
      md += `- Investigator sign-off: ${investigatorSignOff(inv)}\n`;
      md += `- Approver sign-off: ${approverSignOff(inv)}\n`;
      md += `- Investigation status: ${inv.status || "Draft"}\n`;
    } else {
      md += `\n*No investigation has been recorded for this near miss yet.*\n`;
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

export async function generateNearMissRegisterWordDoc(
  items: NearMissRegisterItem[],
  stats: NearMissRegisterStats,
  currentDate: string,
  dateRangeLabel?: string,
): Promise<Buffer> {
  const children: any[] = [];

  children.push(
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "CRANFIELD GLASS CHRISTCHURCH", bold: true, size: 44, color: "1F2937" })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Near Miss - Accident Safety Register", bold: true, size: 30, color: "EA580C" })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: dateRangeLabel ? 40 : 240 }, children: [new TextRun({ text: `Generated ${currentDate}  ·  ${items.length} near miss${items.length === 1 ? "" : "es"}`, size: 20, color: "6B7280", italics: true })] }),
  );

  if (dateRangeLabel) {
    children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [new TextRun({ text: `Period covered: ${dateRangeLabel}`, size: 20, color: "6B7280", italics: true })] }));
  }

  children.push(new Paragraph({
    spacing: { after: 200 },
    children: [new TextRun({ text: `Summary — Total: ${stats.total}   Investigated: ${stats.investigated}   In Progress: ${stats.inProgress}   Not Started: ${stats.notStarted}   High / Extreme: ${stats.highRisk}`, size: 20, color: "374151", bold: true })],
  }));

  if (items.length === 0) {
    children.push(new Paragraph({ children: [new TextRun({ text: "No near misses match the current filters.", size: 20, color: "6B7280", italics: true })] }));
  } else {
    children.push(new Paragraph({ spacing: { before: 120, after: 120 }, children: [new TextRun({ text: "Register Overview", bold: true, size: 26, color: "1F2937" })] }));

    const headerRow = new TableRow({
      tableHeader: true,
      children: [
        wordCell("#", { header: true, width: 5 }),
        wordCell("Event", { header: true, width: 33 }),
        wordCell("Event Date", { header: true, width: 14 }),
        wordCell("Type", { header: true, width: 14 }),
        wordCell("Risk", { header: true, width: 12 }),
        wordCell("Investigation", { header: true, width: 22 }),
      ],
    });

    const rows = items.map((item, i) => {
      const state = investigationState(item);
      return new TableRow({
        children: [
          wordCell(String(i + 1)),
          wordCell(item.title || "Near Miss"),
          wordCell(fmtDate(item.investigation?.eventDate) || fmtDate(item.submittedDate) || ""),
          wordCell(eventTypeOf(item)),
          wordCell(item.investigation?.riskLevel || "—"),
          wordCell(state.label),
        ],
      });
    });

    children.push(new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: { top: WORD_BORDER, bottom: WORD_BORDER, left: WORD_BORDER, right: WORD_BORDER, insideHorizontal: WORD_BORDER, insideVertical: WORD_BORDER },
      rows: [headerRow, ...rows],
    }));

    children.push(new Paragraph({ spacing: { before: 280, after: 120 }, pageBreakBefore: true, children: [new TextRun({ text: "Near Miss Details", bold: true, size: 26, color: "1F2937" })] }));

    const field = (label: string, value: string) => new Paragraph({ spacing: { after: 20 }, children: [
      new TextRun({ text: `${label}: `, bold: true, size: 20, color: "374151" }),
      new TextRun({ text: value, size: 20, color: "4B5563" }),
    ] });

    items.forEach((item, i) => {
      const inv = item.investigation;
      children.push(new Paragraph({ spacing: { before: 160, after: 60 }, keepNext: true, children: [new TextRun({ text: `${i + 1}. ${item.title || "Near Miss"}`, bold: true, size: 24, color: "111827" })] }));

      children.push(field("Event type", eventTypeOf(item)));
      children.push(field("Submitted by", item.submittedBy || "Unknown"));
      children.push(field("Meeting status", meetingStatusLabel(item)));
      children.push(field("Meeting date", fmtDate(item.meetingDate) || "—"));
      if (item.description) children.push(field("What happened", item.description));
      if (item.secondaryDescription) children.push(field("How it happened", item.secondaryDescription));

      if (inv) {
        children.push(new Paragraph({ spacing: { before: 80, after: 40 }, children: [new TextRun({ text: "Investigation", bold: true, size: 20, color: "9A3412" })] }));
        if (inv.investigatorName) children.push(field("Investigator", inv.investigatorName));
        if (inv.siteJob) children.push(field("Site / Job", inv.siteJob));
        if (inv.eventDate) children.push(field("Event date", fmtDate(inv.eventDate)));
        if (inv.eventTime) children.push(field("Event time", inv.eventTime));
        if (inv.involvedPersons) children.push(field("Involved persons", inv.involvedPersons));
        if (inv.witnesses) children.push(field("Witnesses", inv.witnesses));
        if (inv.eventDescription) children.push(field("Findings", inv.eventDescription));
        if (inv.contributingFactors) children.push(field("Contributing factors", inv.contributingFactors));
        if (inv.likelihood || inv.consequence || inv.riskLevel) {
          children.push(field("Risk", `${inv.likelihood || "—"} × ${inv.consequence || "—"} = ${inv.riskLevel || "Not assessed"}`));
        }
        if (inv.treatmentGiven) children.push(field("Treatment given", inv.treatmentGiven));
        (inv.hazards || []).forEach((h) => {
          children.push(new Paragraph({ spacing: { after: 20 }, bullet: { level: 0 }, children: [
            new TextRun({ text: `Hazard: `, bold: true, size: 18, color: "374151" }),
            new TextRun({ text: `${h.hazard || "—"} (${h.risk || "—"}; control: ${h.control || "—"})`, size: 18, color: "4B5563" }),
          ] }));
        });
        (inv.resultingActions || []).forEach((a) => {
          children.push(new Paragraph({ spacing: { after: 20 }, bullet: { level: 0 }, children: [
            new TextRun({ text: `Action: `, bold: true, size: 18, color: "374151" }),
            new TextRun({ text: `${a.description || "—"} → ${a.assignedTo || "Unassigned"} [${a.completed ? "Done" : "Pending"}]`, size: 18, color: "4B5563" }),
          ] }));
        });
        children.push(field("Investigator sign-off", investigatorSignOff(inv)));
        children.push(field("Approver sign-off", approverSignOff(inv)));
        children.push(field("Investigation status", inv.status || "Draft"));
      } else {
        children.push(new Paragraph({ spacing: { before: 40, after: 40 }, children: [new TextRun({ text: "No investigation has been recorded for this near miss yet.", size: 18, color: "6B7280", italics: true })] }));
      }
    });
  }

  const doc = new Document({ sections: [{ properties: {}, children }] });
  return await Packer.toBuffer(doc);
}

---
name: Meeting export formats lockstep
description: All four meeting-minutes export formats must share column semantics via one helper module.
---

# Meeting export formats lockstep

The meeting-minutes export exists in four formats and they must agree on column
semantics: **HTML/PDF** and **CSV** (server/routes.ts), **Markdown**
(server/markdown-generator.ts), **Word** (server/template-engine.ts via
word-generator.ts).

**Rule:** the "Action Required" column and the original-submission text are
derived from `server/meeting-export-shared.ts`
(`buildActionRequiredLines`, `buildAgendaSubmissionText`,
`actionRequiredPlainText`, `isEmptyActionPlaceholder`). Any change to action-column
wording or submission composition goes there so all formats stay consistent.

**Compliance semantics (do not regress):**
- Agenda Item = original submission (description + "How it happened" =
  secondaryDescription). Meeting discussion column = meetingNotes only.
- Action Required = ONLY real action fields
  (actionAssignedTo/actionStatus/actionDueDate/actionNotes). actionNotes is
  labelled "Outcome" when status === 'Closed', else "Action".
- Closed with no action data → "Discussed and closed — no action required."
- Open with nothing recorded → "—" placeholder (omitted in CSV/Word).
- NEVER reintroduce fabricated type/status boilerplate or use meetingNotes as a
  synthesized "Decision".

**Why:** earlier exports fabricated action text and mislabelled columns, which
failed the client's compliance requirement.

**Format caveat:** Markdown and Word render the submission as separate labelled
lines rather than calling `buildAgendaSubmissionText` directly, because its
`\n\n` separators render literally in those formats. Only the action column is
literally shared there; submission fields (description + secondaryDescription)
are the same source, rendered natively.

**"Actions Ready to Close" section (all four formats):** built by
`buildReadyToCloseActions(meetingData)` and shows every action at
`actionStatus === 'Ready to Close'`. Two non-negotiable rules:
- Source it from the FULL unfiltered `meetingData`, NOT the per-meeting
  `filteredData`. These actions await a group sign-off, so the same one
  legitimately re-appears in consecutive minutes until it's actually closed.
  The due date is shown with each (it can sit in the future) precisely so a
  recurring action stays trackable across meetings.
- Dedupe on `type + id`, NEVER the bare id. SharePoint ids are list-local — the
  same numeric id exists in the Near Miss / Safety / Business lists — so a
  bare-id dedupe silently drops a ready-to-close action from another list.

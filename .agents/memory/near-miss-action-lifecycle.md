---
name: Near Miss action lifecycle & dual sign-off
description: How a Near Miss investigation drives the linked Action's status, and the dual-signature completion gate.
---

# Near Miss: Action lifecycle is the source of truth

An investigation completing is what advances the linked Action — not the other way round, and not SharePoint status (Submitted/Actioned/Closed only mirrors).

**Dual sign-off gate:** an investigation needs TWO signatures before it is `Complete`:
- Investigator (`investigatorSignature` / `investigatorSignedAt`)
- Approver / Manager (stored in the legacy `directorName` / `directorSignature` / `signedAt` columns — kept under the old names to preserve existing data; UI/export label them "Approver / Manager").

The `POST /complete` endpoint signs ONE role per call (`{role:'investigator'|'approver', name, signature, signedAt}`). It computes `bothSigned` from the existing record + the incoming signature; status becomes `Complete` only when both are present, otherwise `In Progress`.

**Auto-advance is a hard requirement, not best-effort.** When both signatures land, `/complete` upserts the linked action item (`listType:'NearMiss'`, `sharePointItemId = nearMissItemId`) to `actionStatus:'Ready to Close'` BEFORE marking the investigation `Complete`. If the upsert throws, the whole request fails so you never get a completed investigation whose Action never moved. The upsert is idempotent, so retries are safe.

**Why:** user confirmed the Action lifecycle drives everything; a completed-but-not-advanced state is the failure mode to avoid.

**How to apply:**
- Completion only goes through `/complete`; the generic PUT strips an incoming `status:'Complete'` so it can't bypass the gate.
- Never add code that sets an Action to `Completed` just from viewing/opening an investigation — only the meeting closure flow finalises it.
- `upsertActionItem` spreads `...item` into both insert and update; Drizzle skips `undefined` keys in `.set`, so passing only `{listType, sharePointItemId, actionStatus}` updates just that column without clobbering others.

# Near Miss single completion path (the only way through the workflow)
A Near Miss moves: Submitted →(meeting)→ Actioned →(investigation form signed off)→ actionStatus `Ready to Close` →(group sign-off in meeting minutes)→ `Completed`. That is the ONLY sanctioned route.
- **Meetings:** the per-item status dropdown only offers Submitted/Actioned for Near Miss (Closed removed); legacy current status stays selectable just for display.
- **Actions page:** `ActionStatusWorkflow` takes `allowComplete` (default true); pass `allowComplete={type !== 'Near Miss'}` so the Completed step is locked there.
- **Completing happens in the meeting-minutes "Ready to Close" group-review container** via a "Mark Completed" button → `updateActionFields(item,{actionStatus:'Completed'})`.
- **Server guard (don't remove):** `POST /api/action-items` rejects (409) a NearMiss→Completed unless the stored row is already `Ready to Close` (or already `Completed`, to allow idempotent re-saves — the meeting-history payload always re-sends `actionStatus` so blocking equal-state edits would break note edits). **Why:** UI-only restrictions are bypassable; the server is the real gate.

# Sanctioned bypass: "Close without a report" escape hatch
A dedicated route `POST /api/near-miss-investigations/:itemId/close-without-report` is the ONLY sanctioned way to jump a NearMiss straight to `Completed` without an investigation or meeting sign-off. It calls `storage.upsertActionItem` directly (so it skips the `/api/action-items` 409 guard, which stays intact) and appends a stamped `Closed without investigation — {name}, {date}` note to `actionNotes`. Closing sets `actionStatus:'Completed'` → item lands in the Actions "Archived" view (isArchivedStatus = Completed|Ready to Close) and leaves the live workload.
- **UI:** hidden in `NearMissInvestigationModal`; the "Close without a report" button only renders while Ctrl/⌘ is held (window keydown/keyup + blur reset, also reset when modal closes), then a confirm overlay, then the call. Closer name from `authService.getCurrentUser()?.name`, server falls back to caller displayName.
- **Why hidden behind Ctrl:** it's an admin escape hatch for old stuck near misses, not part of the normal flow — must not be discoverable/clickable by accident.

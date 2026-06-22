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

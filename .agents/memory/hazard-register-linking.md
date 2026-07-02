---
name: Operational Hazard Register linking
description: How Near Miss investigation hazards link to the local hazards table and what must stay in lockstep
---

# Operational Hazard Register linking

- The hazard register is the app's own `hazards` table (local DB, not SharePoint), seeded once via `runOnce("hazard-register-seed-v1")`.
- `hazardId` (e.g. `CG-HZ-001A`) is the stable link key: number block = category, letter = hazard within it. It is immutable via the update endpoint — investigations reference it forever.
  **Why:** investigation rows store only a snapshot + `hazardRefId`; rewriting IDs would orphan historical links.
- Investigation hazard rows are stored as JSON inside `near_miss_investigations.hazards` and were extended **additively** (`hazardRefId`, `registeredControls` snapshot, `controlFailure`, `correctiveActions`, `unregistered`). Legacy rows are detected by `!hazardRefId && !unregistered` and keep the old free-text hazard/control editing — never drop that branch.
- `registeredControls` is a snapshot taken at link time on purpose (audit record of what controls existed when the near miss happened); do not "refresh" it from the live register.
- **How to apply:** any change to the hazard row shape must update ALL renderers in lockstep: HazardTable UI, the four register export formats (shared label/control helpers), and the single investigation report.
- Hazard writes are identity-checked server-side (Graph /me from the Bearer token, same as the order-admin pattern): create = any signed-in user, edit/archive = HAZARD_ADMINS allowlist (falls back to ORDER_ADMINS).

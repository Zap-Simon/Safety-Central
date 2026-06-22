---
name: One-time startup data migrations
description: How idempotent one-off data migrations run automatically on deploy/boot, once per database.
---

# Run-once data migrations on server startup

When you need a data fix to run "on deploy" (e.g. reset legacy rows so they follow a new flow), use the startup-migration mechanism instead of a manual script.

- `server/migrations.ts` exposes `runStartupMigrations()`, called from `server/index.ts` after `registerRoutes` and before `listen`.
- Each migration is wrapped in `runOnce(key, fn)` which checks/records the key in the `app_migrations` table (`shared/schema.ts`). The key is only recorded **after** the work succeeds, so a failure is retried on the next boot.
- It runs **once per database** — dev and production each apply it the first time they boot after the deploy that introduced it. Bump the key (e.g. `-v2`) to re-run.

**Why:** an unguarded startup mutation would re-run on every boot and clobber data that legitimately reached the same state later (e.g. re-opening genuinely-completed items each deploy).

**How to apply:** add a new `runOnce("descriptive-key-vN", async () => {...; return summary})` block; never reuse an old key for new logic. Migration failures are caught and logged, not fatal, so a bad migration won't take the server down.

import { and, eq } from "drizzle-orm";
import { db } from "./db";
import { actionItems, appMigrations, hazards } from "@shared/schema";
import { HAZARD_REGISTER_SEED } from "./hazard-register-seed";
import { log } from "./vite";

// One-time data migrations that run on server startup. Each is keyed and recorded
// in the app_migrations table so it runs exactly once per database (development
// and production each apply it the first time they boot after deploy).
async function runOnce(key: string, fn: () => Promise<string>) {
  const already = await db
    .select()
    .from(appMigrations)
    .where(eq(appMigrations.key, key));
  if (already.length > 0) return;

  const summary = await fn();
  await db.insert(appMigrations).values({ key }).onConflictDoNothing();
  log(`migration "${key}" applied: ${summary}`);
}

export async function runStartupMigrations() {
  try {
    // Re-open any Near Miss items that were marked Completed under the old flow so
    // they re-run the new workflow (worked on -> investigation form -> Ready to
    // Close -> signed off in a meeting). They land back as actively worked on.
    await runOnce("nearmiss-reset-completed-v1", async () => {
      const updated = await db
        .update(actionItems)
        .set({ actionStatus: "In Progress", updatedAt: new Date() })
        .where(
          and(
            eq(actionItems.listType, "NearMiss"),
            eq(actionItems.actionStatus, "Completed"),
          ),
        )
        .returning({ id: actionItems.id });
      return `${updated.length} Near Miss item(s) reset to In Progress`;
    });

    // Seed the Operational Hazard Register from the one-time spreadsheet import.
    // The app database is the live register from here on — new hazards are added
    // through the app, so this only ever runs once per database.
    await runOnce("hazard-register-seed-v1", async () => {
      const inserted = await db
        .insert(hazards)
        .values(HAZARD_REGISTER_SEED.map((h) => ({ ...h, source: "register" })))
        .onConflictDoNothing({ target: hazards.hazardId })
        .returning({ id: hazards.id });
      return `${inserted.length} hazard(s) seeded into the Operational Hazard Register`;
    });
  } catch (error) {
    // Don't crash the server if a migration fails — log it so it can be retried
    // on the next boot (the key is only recorded after the work succeeds).
    console.error("Startup migration error:", error);
  }
}

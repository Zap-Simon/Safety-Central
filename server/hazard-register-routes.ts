import type { Express } from "express";
import { z } from "zod";
import { storage } from "./storage";
import { insertHazardSchema } from "@shared/schema";

// ─── Operational Hazard Register routes ──────────────────────────────────────
// The app database is the live hazard register. These endpoints operate on the
// app's own database only (no SharePoint), so like the neighbouring near-miss
// investigation endpoints they are open to any signed-in user of the app.

// Hazard IDs look like "CG-HZ-003B": a zero-padded number shared by every
// hazard in the same category, plus a letter suffix per hazard.
const HAZARD_ID_RE = /^CG-HZ-(\d{3})([A-Z])$/;

export function suggestNextHazardId(
  existing: { hazardId: string; category: string }[],
  category: string,
): string {
  const wanted = category.trim().toLowerCase();
  let maxNumber = 0;
  let categoryNumber: number | null = null;
  let maxLetterForCategory = "";

  for (const h of existing) {
    const m = HAZARD_ID_RE.exec(h.hazardId.trim().toUpperCase());
    if (!m) continue;
    const num = parseInt(m[1], 10);
    if (num > maxNumber) maxNumber = num;
    if (h.category.trim().toLowerCase() === wanted) {
      // Use the category's existing number (they all share one)
      if (categoryNumber === null || num < categoryNumber) categoryNumber = num;
      if (m[2] > maxLetterForCategory) maxLetterForCategory = m[2];
    }
  }

  if (categoryNumber !== null && maxLetterForCategory && maxLetterForCategory < "Z") {
    const nextLetter = String.fromCharCode(maxLetterForCategory.charCodeAt(0) + 1);
    return `CG-HZ-${String(categoryNumber).padStart(3, "0")}${nextLetter}`;
  }
  // New category (or the letters ran out): start a new number block
  return `CG-HZ-${String(maxNumber + 1).padStart(3, "0")}A`;
}

export function registerHazardRegisterRoutes(app: Express): void {
  // GET /api/hazards — full register (archived rows included; client filters)
  app.get("/api/hazards", async (_req, res) => {
    try {
      const all = await storage.getAllHazards();
      res.json({ success: true, data: all });
    } catch (error) {
      console.error("Get hazards error:", error);
      res.status(500).json({ success: false, error: "Failed to load hazard register" });
    }
  });

  // GET /api/hazards/next-id?category=... — suggested ID for a new hazard
  app.get("/api/hazards/next-id", async (req, res) => {
    try {
      const category = String(req.query.category || "").trim();
      if (!category) {
        return res.status(400).json({ success: false, error: "category is required" });
      }
      const all = await storage.getAllHazards();
      res.json({ success: true, data: { hazardId: suggestNextHazardId(all, category) } });
    } catch (error) {
      console.error("Next hazard ID error:", error);
      res.status(500).json({ success: false, error: "Failed to suggest hazard ID" });
    }
  });

  // POST /api/hazards — add a new hazard to the register
  app.post("/api/hazards", async (req, res) => {
    try {
      const parsed = insertHazardSchema
        .extend({
          hazardId: z.string().trim().optional(),
          category: z.string().trim().min(1, "Category is required"),
          riskHarm: z.string().trim().min(1, "Hazard description is required"),
        })
        .parse(req.body);

      const suppliedId = (parsed.hazardId || "").toUpperCase();
      if (suppliedId) {
        if (!HAZARD_ID_RE.test(suppliedId)) {
          return res.status(400).json({ success: false, error: `Hazard ID must look like CG-HZ-001A` });
        }
        const all = await storage.getAllHazards();
        if (all.some((h) => h.hazardId.toUpperCase() === suppliedId)) {
          return res.status(409).json({ success: false, error: `Hazard ID ${suppliedId} is already in the register` });
        }
        const created = await storage.createHazard({ ...parsed, hazardId: suppliedId });
        return res.json({ success: true, data: created });
      }

      // Auto-generated ID: retry once if a concurrent insert grabbed the same ID
      // (hazardId has a unique constraint, so the race surfaces as a DB error).
      for (let attempt = 0; ; attempt++) {
        const all = await storage.getAllHazards();
        const hazardId = suggestNextHazardId(all, parsed.category);
        try {
          const created = await storage.createHazard({ ...parsed, hazardId });
          return res.json({ success: true, data: created });
        } catch (err: any) {
          const isUniqueViolation = err?.code === "23505";
          if (!isUniqueViolation || attempt >= 2) throw err;
        }
      }
    } catch (error) {
      if (error instanceof z.ZodError) {
        return res.status(400).json({ success: false, error: error.errors[0]?.message || "Invalid hazard" });
      }
      console.error("Create hazard error:", error);
      res.status(500).json({ success: false, error: "Failed to add hazard to register" });
    }
  });

  // PUT /api/hazards/:id — update / archive a hazard
  app.put("/api/hazards/:id", async (req, res) => {
    try {
      const id = parseInt(req.params.id, 10);
      if (isNaN(id)) {
        return res.status(400).json({ success: false, error: "Invalid hazard id" });
      }
      const parsed = insertHazardSchema.partial().parse(req.body);
      // Never allow the stable hazardId to be rewritten — investigations link by it
      delete (parsed as Record<string, unknown>).hazardId;
      const updated = await storage.updateHazard(id, parsed);
      if (!updated) {
        return res.status(404).json({ success: false, error: "Hazard not found" });
      }
      res.json({ success: true, data: updated });
    } catch (error) {
      if (error instanceof z.ZodError) {
        return res.status(400).json({ success: false, error: error.errors[0]?.message || "Invalid hazard" });
      }
      console.error("Update hazard error:", error);
      res.status(500).json({ success: false, error: "Failed to update hazard" });
    }
  });
}

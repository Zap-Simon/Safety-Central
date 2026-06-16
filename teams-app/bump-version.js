#!/usr/bin/env node
/**
 * Teams app version bump + repackage script
 *
 * Usage (run from project root):
 *   node teams-app/bump-version.js          # bumps patch  1.0.0 → 1.0.1
 *   node teams-app/bump-version.js patch    # same as above
 *   node teams-app/bump-version.js minor    # bumps minor  1.0.0 → 1.1.0
 *   node teams-app/bump-version.js major    # bumps major  1.0.0 → 2.0.0
 *
 * Output:
 *   teams-app/manifest.json      updated in place
 *   cranfield-safety-ideas.zip   created/replaced in project root
 */

import { readFileSync, writeFileSync, unlinkSync } from "fs";
import { execSync } from "child_process";
import { resolve, dirname } from "path";
import { fileURLToPath } from "url";
import { tmpdir } from "os";

const __dirname = dirname(fileURLToPath(import.meta.url));
const manifestPath = resolve(__dirname, "manifest.json");
const zipPath     = resolve(__dirname, "..", "cranfield-safety-ideas.zip");

// ── Read manifest ──────────────────────────────────────────────────────────
const manifest = JSON.parse(readFileSync(manifestPath, "utf8"));
const current  = manifest.version;
const [major, minor, patch] = current.split(".").map(Number);

// ── Bump ───────────────────────────────────────────────────────────────────
const bump = (process.argv[2] || "patch").toLowerCase();
let next;
if      (bump === "major") next = `${major + 1}.0.0`;
else if (bump === "minor") next = `${major}.${minor + 1}.0`;
else if (bump === "patch") next = `${major}.${minor}.${patch + 1}`;
else {
  console.error(`Unknown bump type "${bump}". Use: patch | minor | major`);
  process.exit(1);
}

manifest.version = next;
writeFileSync(manifestPath, JSON.stringify(manifest, null, 2) + "\n", "utf8");
console.log(`✅  Version bumped: ${current} → ${next}`);

// ── Repackage ──────────────────────────────────────────────────────────────
// Write a small Python helper to a temp file (avoids shell escaping issues)
const pyPath = resolve(tmpdir(), "teams_zip_helper.py");
writeFileSync(pyPath, [
  "import zipfile, os, sys",
  "out_path = sys.argv[1]",
  "src_dir  = sys.argv[2]",
  "files    = ['manifest.json', 'color.png', 'outline.png']",
  "with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zf:",
  "    for f in files:",
  "        zf.write(os.path.join(src_dir, f), f)",
].join("\n"), "utf8");

try {
  execSync(`python3 "${pyPath}" "${zipPath}" "${__dirname}"`, { stdio: "inherit" });
  console.log(`📦  Package ready:  cranfield-safety-ideas.zip  (v${next})`);
  console.log(`\n    Upload at: Teams Admin Center → Teams apps → Manage apps → Update app`);
} catch (err) {
  console.error("❌  Zip step failed:", err.message);
  process.exit(1);
} finally {
  try { unlinkSync(pyPath); } catch {}
}

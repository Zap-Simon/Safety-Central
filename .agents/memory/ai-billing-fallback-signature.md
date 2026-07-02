---
name: AI billing & the "everything classifies as Other" signature
description: OpenAI calls are billed via Replit AI Integrations; a silent quota/auth failure makes every submission classify as "Other" with generic follow-ups.
---

# AI billing & failure signature

**Rule:** All server OpenAI clients must use `AI_INTEGRATIONS_OPENAI_API_KEY` + `AI_INTEGRATIONS_OPENAI_BASE_URL` (Replit-managed billing). Never reintroduce a personal `OPENAI_API_KEY`.

**Why:** The user's personal OpenAI key ran out of credit (429 insufficient_quota) in production (July 2026). The classifier's catch-block falls back to category "Other" + 2 generic follow-up questions, so the failure looked like a prompt/classification bug, not a billing outage.

**How to apply:**
- If users report "everything classifies as Other", suspect the OpenAI call is failing outright — test the API directly before touching prompts or rules.
- Quick test method: write a temp `_tmp_test_classify.ts` in workspace root importing OpenAIService, run `npx tsx`, delete after (scripts in /tmp can't resolve relative imports).
- There are TWO OpenAI client constructions: `server/openai-service.ts` (module-level) and an inline one in the near-miss AI-draft route in `server/routes.ts` — keep both on the same credentials.
- After changing env/credentials, the user must republish for production to pick it up.

---
name: AI billing & the "everything classifies as Other" signature
description: OpenAI billing is user-chosen (personal key vs Replit AI Integrations); a silent quota/auth failure makes every submission classify as "Other" with generic follow-ups.
---

# AI billing & failure signature

**Rule:** The user explicitly chose (July 2026) to bill AI through their PERSONAL `OPENAI_API_KEY`, not Replit AI Integrations. Don't switch billing paths without asking; keep ALL OpenAI client sites on the same credentials. The Replit integration (`AI_INTEGRATIONS_OPENAI_API_KEY` + `AI_INTEGRATIONS_OPENAI_BASE_URL`) remains configured as an easy fallback if their key runs dry again.

**Why:** The user's personal key ran out of credit (429 insufficient_quota) in production; the classifier's catch-block falls back to category "Other" + 2 generic follow-up questions, so the failure looked like a prompt/classification bug, not a billing outage. We briefly switched to Replit AI Integrations, then the user topped up their OpenAI credit and asked to switch back.

**How to apply:**
- If users report "everything classifies as Other", suspect the OpenAI call is failing outright — test the API directly before touching prompts or rules.
- Quick test method: write a temp `_tmp_test_classify.ts` in workspace root importing OpenAIService, run `npx tsx`, delete after (scripts in /tmp can't resolve relative imports).
- There are TWO OpenAI client constructions: `server/openai-service.ts` (module-level) and an inline one in the near-miss AI-draft route in `server/routes.ts` — keep both on the same credentials.
- After changing env/credentials, the user must republish for production to pick it up.

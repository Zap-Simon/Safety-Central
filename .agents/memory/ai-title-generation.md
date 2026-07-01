---
name: AI title generation
description: How submission titles are generated and why near-miss titles used to be bad
---

# AI title generation

- Single title path: `OpenAIService.generateSmartTitle(content, type)` in `server/openai-service.ts` (gpt-4o-mini). Both `create-item` (server/routes.ts) and batch generation use it — change the prompt in one place.
- `content` is the FULL submission text: raw input + follow-up Q&A concatenated (see SubmitTab.tsx `fullDescription`). So follow-up answers like a location/job/site or a person's name are in scope for the model.

**Why titles like "3b Airedale courts" happened:** the old prompt had no rule against using a bare location/name as the title, so the model sometimes promoted a location follow-up answer over the actual incident.

**How to apply:** the system prompt now requires the title to describe the incident/hazard/idea and forbids bare location/job/address/person-name titles. If titles regress, tighten this prompt rather than adding a second title path. Prompt-only control isn't deterministic; a post-generation validator + retry would be the next step if needed. Existing SharePoint titles are NOT retroactively fixed by a prompt change — that needs a one-time backfill re-running generateSmartTitle over stored descriptions.

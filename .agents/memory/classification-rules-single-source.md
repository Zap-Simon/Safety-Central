---
name: Classification single source of truth
description: Where Teams Improve+ idea categories, list routing, and follow-up questions are defined and how to change them safely.
---

# Idea classification rules

All submit-tab categories, their `listTarget`, `listLabel`, AI description, and
follow-up questions live in `shared/classification-rules.ts` (`CATEGORY_RULES`,
`CATEGORY_NAMES`, `LIST_TARGET_LIST_TYPE`).

- `server/openai-service.ts#classifySubmission` builds the gpt-4o-mini prompt by
  looping over `CATEGORY_RULES` — do NOT hard-code categories back into the prompt string.
- The server derives the final `listTarget` from `getCategoryRule(category)` (rule
  is authoritative), so the AI returning a wrong listTarget can't misroute a known category.
- `client/.../teams/SubmitTab.tsx` imports `CATEGORY_NAMES`/`LIST_TARGETS`/`LIST_TARGET_LIST_TYPE`
  and `getCategoryRule().listLabel`. The only thing kept local is `CATEGORY_META`
  (React icons + Fluent colour tokens), which MUST have a key for every Category.

**Why:** categories used to be duplicated across the prompt string and SubmitTab and
drifted — e.g. all "Meeting Agenda Item" submissions routed to Business Ideas even when
about safety/near-miss. Now split into Near Miss / Safety / Business Meeting Agenda Item,
each routing to near-miss / safety-ideas / business-ideas respectively.

**How to apply:** to add/change a category, edit `shared/classification-rules.ts` and add the
matching `CATEGORY_META` entry in SubmitTab. Nothing else. Keep teams-app/README.md category table in sync.

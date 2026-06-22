---
name: Actions page deep-link focus
description: Why deep-linking into the Actions page filters to one card instead of opening a detail modal
---

# Actions page deep-link from Meeting Minutes

Meeting Minutes opens an action via `/actions?itemId=...&type=...` in a new tab.
The deep-link handler must **filter the list down to the single targeted item**
(a `focusedItemId` state short-circuiting `filteredItems`), NOT open the detail
modal (`selectedItem` / `NearMissInvestigationModal`) and NOT scroll into view.

**Why:** Opening the detail modal on a freshly-loaded tab triggers modal-driven
follow-up fetches that pop an **interactive MSAL auth prompt**, and the
scroll-into-view jump is jarring. Filtering avoids both.

**How to apply:** Keep the focus-filter approach; provide a "Show all actions"
banner to clear `focusedItemId`. Don't re-add auto-open-modal/auto-scroll on
deep-link.

Related: "Ready to Close" actions are NOT closed — only strike through titles
when `actionStatus === 'Completed'`, never for `Ready to Close`.

---
name: Persisting Teams tab form state across remounts
description: Why and how the Submit flow survives tab switches / app close
---

Teams personal tabs **remount on every tab switch** (and on app close/reopen), so
any `useState` in a tab is wiped when the user toggles tabs. A multi-step form
(e.g. the Submit flow's input → classifying → followup → confirm → submitting → done
state machine) loses all progress unless its durable state is persisted outside React.

**Rule:** persist the durable draft to `localStorage` (NOT just a context provider —
a provider survives tab switches but not app close, and the user expects close/reopen
to keep their in-progress report).

**How to apply:**
- Lazy-init each `useState` from a single `loadDraft()` read (`const [draft] = useState(loadDraft)`).
- A `useEffect` writes the draft on change, but **skips transient in-flight steps**
  (`classifying`/`submitting`) — restoring those would show a dead spinner since the
  network call is gone.
- On read, normalise transient steps to the nearest actionable one
  (`classifying`→`input`, `submitting`→`confirm` if a result exists else `input`).
- Validate the parsed shape (known step enum, required fields, classifyResult present
  for confirm/followup) and return `null` on anything unexpected so corrupted storage
  can't land the UI on a non-rendered branch.
- Clear the key when the form returns to a truly empty `input` state (and on reset).
- Wrap all storage access in try/catch — degrade to in-memory if storage is unavailable.

**Why:** non-persisted refs (e.g. an in-memory classify cache) are fine to lose on
remount as long as the resume path re-fetches; only the user-entered/derived data
must persist.

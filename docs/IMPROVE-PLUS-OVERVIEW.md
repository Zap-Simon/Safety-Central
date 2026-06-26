# Improve+ — System Overview & Competitive Landscape

*Cranfield Glass Christchurch — Health, Safety & Idea Management*

---

## 1. System at a Glance

**Improve+** is a lightweight staff-engagement and meeting-management tool that lives
**inside Microsoft Teams** and runs entirely on Cranfield's own Microsoft 365 data.

It does three jobs that are normally handled by three separate products:

1. **Captures** business ideas, safety improvements, and near-miss incidents from staff.
2. **Runs the meeting** — turning those submissions into a *live agenda* that becomes the
   *minutes* in the same screen, with attendance and digital sign-off baked in.
3. **Tracks the follow-through** — actions, investigations, and closure, with a clear
   audit trail.

Everything stays in your Microsoft 365 tenant (SharePoint, Teams, Power Automate).
No third-party data silo, no per-seat SaaS subscription, no data leaving the business.

---

## 2. How It Works

```
Staff submission            Live meeting               Follow-through
─────────────────           ────────────               ─────────────
Business idea      ─┐                                  ┌─ Actions tracked
Safety idea        ─┼─►  Live agenda  ──►  Minutes  ──►┤  Investigations
Near-miss incident ─┘     (same screen, real time)     └─ Closure + audit
                                  │
                          Attendance + digital
                          signatures (in-meeting)
```

1. **Capture** — Staff submit ideas or report a near-miss from a Teams tab (or Microsoft
   Forms). Submissions are automatically classified and routed to the right list.
2. **Live agenda** — At the meeting, open items appear as a working agenda. Notes typed
   during the meeting *are* the minutes — there's no separate write-up step afterwards.
3. **Attendance & signatures** — Who attended is recorded and locked; attendees sign off
   on the minutes digitally, inside Teams, as the meeting closes.
4. **Actions & investigations** — Items marked for action are tracked to completion.
   Near-miss incidents carry a formal investigation with dual sign-off before closing.
5. **Export & audit** — Minutes export to professional PDF/HTML/Word/CSV with Cranfield
   branding, page numbering, and a full record for compliance.

---

## 3. What Makes It Distinctive

The individual ingredients all exist somewhere on the market. What's **rare** is the
combination of all four in one purpose-built, tenant-owned tool:

| # | Capability | Why it matters |
|---|------------|----------------|
| 1 | **Domain capture** (ideas + safety + near-miss) with auto-classification | Right item goes to the right list/owner with no manual sorting |
| 2 | **Live agenda → minutes** in the same surface | No double-handling; the meeting writes its own record |
| 3 | **In-meeting digital signatures + attendance lock** | Formal, defensible sign-off without a separate e-sign tool |
| 4 | **Native inside Teams, on your own SharePoint data** | No data silo, no per-seat cost, nothing leaves the tenant |

Most vendors own **one column** of this table and sell it as a standalone product.

---

## 4. Competitive Comparison

A simple feature × vendor view. ● = strong / core capability, ◐ = partial / weak,
○ = not offered.

| Capability | **Improve+** | Decisions (M365 meetings) | Board portals (Diligent, OnBoard, Convene) | EHS tools (SafetyCulture, EcoOnline, Donesafe) | Native M365 (Loop, OneNote, Lists) |
|---|:---:|:---:|:---:|:---:|:---:|
| Lives natively inside Teams / M365 | ● | ● | ○ | ○ | ● |
| Runs on **your own** SharePoint data (no silo) | ● | ◐ | ○ | ○ | ● |
| Live agenda that becomes the minutes | ● | ● | ◐ | ○ | ◐ |
| In-meeting digital signatures / sign-off | ● | ◐ | ● | ○ | ○ |
| Attendance capture + lock | ● | ◐ | ● | ○ | ○ |
| Near-miss / safety incident capture | ● | ○ | ○ | ● | ○ |
| Idea capture + auto-classification & routing | ● | ○ | ○ | ◐ | ○ |
| Investigation workflow with dual sign-off | ● | ○ | ○ | ● | ○ |
| Action tracking to closure | ● | ● | ● | ● | ◐ |
| Branded professional export (PDF/Word/CSV) | ● | ● | ● | ● | ◐ |
| No per-seat SaaS subscription | ● | ○ | ○ | ○ | ● |

### How to read this table

- **Decisions** is the closest comparison for the *meeting* half — it's genuinely
  Teams-native — but knows nothing about near-miss incidents, safety triage, or your
  investigation sign-off.
- **Board portals** nail the *formal sign-off* of minutes, but are heavyweight, costly
  governance tools that don't live in Teams and have no safety capture.
- **EHS tools** own *near-miss and safety* workflows, but are their own ecosystems with
  weak, bolt-on meeting features and your data living on their servers.
- **Native M365** gives you raw building blocks (notes, lists, tasks) but no live
  agenda→minutes→signature pipeline and no domain logic.

To match Improve+ off the shelf you'd need **three separate paid products** — a meeting
tool, a board portal, and an EHS platform — none of which talk to each other or to your
SharePoint the way one purpose-built tool does.

---

## 5. The Honest Take

Improve+ is **not technologically revolutionary** — every individual feature exists
somewhere on the market.

What *is* rare and valuable is the **integration and fit**: a single lightweight,
M365-native tool that collapses three product categories into one workflow for one team,
with **zero data leaving the tenant** and **zero per-seat SaaS bills**.

That "right-sized, fully-owned, fused" quality is exactly what off-the-shelf vendors
**can't** easily deliver — because their business model depends on selling each column of
the comparison table separately.

---

## Appendix — Technical Summary

For a technical audience, the system is built as follows:

- **Frontend** — React 18 + TypeScript, Vite, Tailwind + shadcn/ui, TanStack Query,
  Wouter routing. Mobile-first.
- **Backend** — Node.js + Express, RESTful API, Drizzle ORM over PostgreSQL (Neon),
  PostgreSQL-backed sessions.
- **Microsoft 365 integration** — Microsoft Graph API for SharePoint Lists and document
  libraries; Power Automate for workflow/notifications; Microsoft Forms embedded in Teams.
- **Authentication** — Azure AD. The main website uses MSAL; the Teams tabs use the
  OAuth 2.0 On-Behalf-Of (OBO) flow, so the browser only ever holds a Teams SSO token and
  the server exchanges it for downstream Graph/SharePoint tokens. Shared API endpoints
  detect the token audience and handle each path transparently.
- **AI features** — OpenAI GPT-4o / gpt-4o-mini for title generation, content processing,
  and meeting-notes drafting, with graceful fallbacks.
- **Export engine** — Server-side Paged.js-based rendering for branded, paginated
  PDF/HTML minutes, plus Word and CSV exports sharing a single source of formatting.

---

*Document generated for Cranfield Glass Christchurch. Aimed at management/stakeholders
with a technical appendix.*

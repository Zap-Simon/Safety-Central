---
name: Orders admin authorization
description: How "order admin" (remove item / clear list / mark ordered) access is decided in the Teams Orders tab.
---

# Orders admin gate

Order-admin rights (mark ordered, remove a single item, clear the whole list) are decided by an
**allowlist**, not by the staff-table `role`. The list comes from `process.env.ORDER_ADMINS`
(comma-separated) and falls back to a hardcoded default. Each entry is matched against the caller's
Graph-resolved **email (UPN) OR display name**, in `resolveOrderAdminFromToken` (server/routes.ts).

**Why:** The user defines admins as a specific named set of people, independent of staff roles. The
gate was intentionally switched away from the previous `role === 'admin'|'supervisor'` staff-table
check. Email is preferred over display name because display names are mutable and not guaranteed
unique — matching admin purely by name is a weak access control (flagged in review). Name matching is
kept only as a convenience for the current single entry; move to emails (or a Microsoft 365 group
membership check) as the team list grows.

**How to apply:** To change who is an order admin, edit `ORDER_ADMINS` (prefer full emails). The
client never decides admin status — it only reflects `/api/orders/is-admin`. All mutating order
endpoints re-check server-side. Removing items / clearing are **soft** (status → `archived`), so
nothing is hard-deleted.

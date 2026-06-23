---
name: NZ date-arrival checks
description: "Has the meeting day arrived?" checks must compare against NZ today, not UTC today.
---

Meeting date keys are NZ calendar dates (SharePoint dates are normalised to `YYYY-MM-DDT10:00:00.000Z` where the date part is the NZ day staff see in MS Lists). So any "has this day arrived yet?" comparison must use `getNZTodayKey()` (Pacific/Auckland via `toLocaleDateString('en-CA', { timeZone })`), NOT `getDateGroupKey(new Date())` (UTC).

**Why:** NZ is UTC+12 (NZST) / UTC+13 (NZDT), so UTC lags NZ by half a day. Using UTC today kept early-morning meetings stuck as "upcoming" until ~noon NZ on the meeting day. Using NZ today flips them at NZ midnight, and since NZ is always ahead of UTC it can never flip early.

**How to apply:** In the Teams Sign flow (`GET /api/teams/sign/meetings` open-vs-upcoming classification and the `POST /api/teams/sign` future-date guard) compare meeting keys against `getNZTodayKey()`. Keep client and server using the same rule so a crafted POST can't bypass it. `getDateGroupKey` (UTC) is still correct for grouping/merging same-day keys — only the "is it today yet" comparison needs the NZ today key.

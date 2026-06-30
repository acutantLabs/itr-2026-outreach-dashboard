# Client File Tabs & Expanded Desk — Design Spec
**Date:** 2026-06-30  
**Project:** Sonia Outreach Dashboard — ITR 2026

---

## Overview

Surface each client's SharePoint folder contents directly in the dashboard: as scannable file tabs on every card, as a detailed Files panel in the desk modal, and as a new "Files" facet in the Refine filter. Users can also upload files directly from the dashboard.

---

## Data Source

- **Location:** SharePoint `Documents/ITR 2026/Client Folders/{ClientID} - {ClientName}/`
- **Folder path:** stored as a field on each client record in the tracker list (already loaded via `buildClientsFromSharePoint`)
- **API:** Microsoft Graph `drives/{driveId}/root:/{path}:/children`
- **Standard files:** filename contains `"ais"` or `"26as"` (case-insensitive)
- **Additional files:** anything else in the folder

---

## File Classification

| Type | Match rule | Colour token | Visual |
|------|-----------|-------------|--------|
| Standard | name includes `ais` or `26as` | `--file-std: #0d9488` (teal) | filled tab |
| Additional | everything else | `--file-add: #7c3aed` (violet) | filled tab |

Two new CSS variables added to `:root`. Neither conflicts with existing palette (teal and violet are unused).

---

## Loading Strategy — Staggered Background Fetch + LocalStorage Cache

1. After `buildClientsFromSharePoint()` resolves, kick off a background file scan.
2. Fetch file listings in parallel batches of 15 clients at a time.
3. Each result is stored in `localStorage` keyed `itr_files_{clientId}` with a `ts` timestamp.
4. On load, read cache first — entries younger than 4 hours are used as-is; stale entries are refreshed in the background batch.
5. When a specific client's desk opens, always force-refresh that client's listing (regardless of cache age), then update both the desk panel and the card.
6. Cards update in place as each batch completes — no full grid re-render.

**Cache schema per entry:**
```json
{ "ts": 1719700000000, "files": [{ "name": "alpa ben AIS.pdf", "url": "https://...", "modified": "2026-06-24" }] }
```

---

## Scan Progress Indicator

- A 2px progress bar rendered immediately below the `.red-rule` element.
- Animates from 0% → 100% as batches complete; fades out on finish.
- A muted caption `"scanning files · N / 375"` appears beneath it (`.scan-caption`), hidden when done.
- No modal, no toast, no layout shift.

---

## Card — File Tab Row

Added as a new bottom section on both `.fc` (closed) and `.sc` (open stack) cards.

**Closed card (`.fc`):**
```
├──────────────────────────────┤
│ ▐█▌ ▐█▌  ▐░▌                │  small 6×14px coloured tabs
└──────────────────────────────┘
  teal = standard    violet = additional
  hover → tooltip: filename
  click → opens file URL (new tab), stopPropagation
```

- Max 6 tabs shown; overflow shown as `+N` muted chip.
- No tabs rendered (section hidden) while cache is loading; shimmer placeholder shown for first scan.
- Empty folder: muted `— no files` label.

**Open stack card (`.sc`):** file count badge only (`2S · 1+` meaning 2 standard, 1 additional), no individual tabs (space-constrained).

---

## Desk Modal — Expanded Layout + Files Panel

### Layout expansion
The desk currently uses `grid-template-columns: 268px 1fr`. Expand the left panel to `320px` and increase the overall desk max-width to fit the new Files section comfortably.

### Files panel (new, inside left column, between Contact card and bottom)
```
FILES  ▾  2 standard · 1 additional          [↑ UPLOAD]
┌──────────────────────────────────────────────┐
│ ▐█▌  alpa ben 26as.pdf    PDF  24 Jun  [↗]  │
│ ▐█▌  alpa ben AIS.pdf     PDF  24 Jun  [↗]  │
│ ▐░▌  salary slip june.pdf PDF  28 Jun  [↗]  │
└──────────────────────────────────────────────┘
```
- Collapsible (chevron toggle); defaults open.
- Each row: colour tab, filename (truncated with title attr), type badge, modified date, open-in-SharePoint icon.
- `[↑ UPLOAD]` button: triggers native `<input type="file" multiple>`, uploads via Graph `PUT /drives/{driveId}/root:/{path}/{filename}:/content`, then refreshes the file listing for this client and updates cache + card.
- Upload permission: any signed-in user (no Sonia gate).
- Loading state: spinner replaces list while force-refreshing on desk open.

---

## Refine Panel — Files Facet

New section added to the existing Refine dropdown, below existing facets:

```
FILES
  📁 Has additional files     47
  📁 Standard files only     189
  📁 No files yet            139
```

- Counts computed from the cache (or `—` while scan is incomplete).
- Selecting "Has additional files" filters the grid to only those clients.
- Mutually exclusive single-select within the Files facet; composable with other Refine facets.

---

## New CSS Variables

```css
--file-std: #0d9488;   /* teal — AIS / 26AS standard files */
--file-add: #7c3aed;   /* violet — additional / other files */
```

---

## What is NOT changing

- Header filter tabs (ALL, NEW, ACTION, DUE, REPLY, DONE, WAIT, FLAG)
- Comm chips, desk slips rail, shadow shelf
- SoniaRead gating on comm read-state
- Email composer, template system
- Any SharePoint list write operations

---

## Open Questions (resolved)

- **Colours:** teal for standard, violet for additional (no palette conflicts).
- **Upload permission:** any signed-in user.
- **Folder field name in tracker:** to be confirmed during implementation by inspecting `f` fields in `buildClientsFromSharePoint`; fallback: derive path as `ITR 2026/Client Folders/{clientId} - {clientName}`.

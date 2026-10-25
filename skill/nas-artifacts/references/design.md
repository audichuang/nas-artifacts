# Page design rules

For agents writing pages to publish here. Start from `templates/report`, `templates/dashboard`, `templates/deck` or `templates/checklist` (`cp -r` into a scratch folder, edit, publish); they already follow everything below. Rules first, reason after the dash.

## Basics

- Title: 2-4 words, a name not a sentence (`季度研究報告`), set in `<title>` and the H1. It is the gallery card.
- `<html lang="zh-Hant">` for Traditional Chinese, `<meta name="viewport" content="width=device-width, initial-scale=1">`. UI copy in the user's language.
- One file unless there is data or assets to swap later. Put changing data in `data.js` (sets a global) so a republish can send only that file.
- Revising a page: re-read the whole file first, then edit. Never patch from memory or from a stale copy; the CLI replaces the version with what is on disk.

## Tokens

All colour, spacing and radius come from variables on `:root`. Dark mode redefines the same names under `@media (prefers-color-scheme: dark)`. Set `body { background: var(--bg); color: var(--text) }` explicitly -- an unset background falls back to the browser default (white), so light text lands on white in dark mode.

```css
:root{--bg:#fafaf9;--surface:#fff;--text:#1c1917;--muted:#57534e;--border:#e7e5e4;--accent:#2563eb;
  --c1:#0072b2;--c2:#e69f00;--c3:#009e73;--s1:4px;--s2:8px;--s3:12px;--s4:16px;--s5:24px;--s6:40px;--radius:10px}
@media (prefers-color-scheme:dark){:root{--bg:#171717;--surface:#1f1f1f;--text:#f5f5f4;--muted:#a8a29e;--border:#3a3a3a;--accent:#60a5fa;--c1:#56b4e9}}
```

Copy the full token set from any template instead of inventing one. Do not hard-code hex values in rules; a hard-coded colour is the usual cause of unreadable dark mode. Keep text contrast >= 4.5:1 in both themes.

## Type and spacing

- Font stack: `system-ui,-apple-system,"Segoe UI","PingFang TC","Noto Sans TC","Microsoft JhengHei",sans-serif`. No web fonts unless asked (Google Fonts or the script CDNs) -- system fonts render instantly and CJK web fonts are huge.
- Scale: 13 (captions) / 15-16 (body) / 18 (lead) / 22 (h2) / 26-36 (h1, use `clamp`). Body line-height 1.6-1.7 for Chinese text.
- Spacing only from the 4/8/12/16/24/40 scale. Body column <= 760px for prose, <= 1080px for dashboards.

## Tables

- Wrap in `overflow-x:auto` with `tabindex="0"` and an `aria-label`; the table scrolls, the page does not.
- Right-align numbers, `font-variant-numeric: tabular-nums`, same decimals per column, units in the header.
- Header row visually distinct (background + muted small text). Sticky header only when the table sits in its own vertical scroller (`max-height` on the wrapper); `position:sticky` inside an `overflow-x:auto` wrapper does not stick to the page, so skip it for short tables.
- `white-space:nowrap` on cells keeps rows one line; drop it on long text columns.

## Charts

- Inline SVG with a `viewBox` and `width:100%` so it scales; text in SVG uses the tokens (`fill:var(--muted)`).
- Label both axes (with units), give the chart a `<title>`/`aria-label`, state the takeaway in the caption or heading.
- Palette: Okabe-Ito based and colour-blind-safe (`--c1` blue, `--c2` orange, `--c3` green). Max 3-4 series, never a rainbow. Also distinguish series by something other than colour (dash, direct label, legend order).
- Start bar charts at zero. Grid lines faint (`--border`). No 3D, no gradients, no pie with more than 4 slices.
- External chart libraries only at an exact version from cdnjs/jsdelivr/unpkg (`chart.js@4.4.1`); never `latest` or a range. Plain SVG is usually smaller and always works offline.

## Mobile and interaction

- Must work at 320-414px: 16px side gutter (`padding-inline:16px` on the main container), no horizontal page scroll (`document.documentElement.scrollWidth <= innerWidth`). Check with a 375px screenshot.
- Grids use `repeat(auto-fit,minmax(150px,1fr))`; wide children get `min-width:0` or their own scroller; long strings `overflow-wrap:anywhere`.
- Tap targets >= 44px; inputs `font-size:16px` or larger (smaller makes iOS zoom).
- Never remove focus outlines; use `:focus-visible{outline:2px solid var(--accent);outline-offset:2px}`. Every control reachable and operable by keyboard; use real `<button>`, `<label>`, `<input>`.

## What the sandbox forbids, and the fix

| Forbidden | Why / fix |
|---|---|
| `fetch`, XHR, WebSocket, `EventSource` | `connect-src 'none'`. Inline the data, publish a `data.js` that sets a global, or use `db`. |
| `<a target=_blank>`, `window.open`, navigation away | No popups, frame cannot navigate. Show the URL as text. Same-page `#anchors` and relative links to published files work. |
| `localStorage`, `sessionStorage`, cookies, IndexedDB | Opaque origin throws. Wrap in try/catch if used for a convenience; shared state goes in `db` (`--cap db`). |
| Remote images (`<img src="https://...">`) | Blocked. Use a published file, a `data:` URI, or `artifact asset upload`; draw simple graphics as inline SVG. |
| Script/style from other hosts, unpinned versions | Only cdnjs, jsdelivr, unpkg, cdn.tailwindcss.com, code.jquery.com with an exact version; fonts from Google Fonts or those CDNs. Inline everything else. |
| `window.artifact` calls outside the viewer (opening the content URL directly) | They reject. Catch and show a short message telling the reader to open the page from the Artifacts viewer. |
| Remembering state between visits without `db` | Not possible; do not fake it. Declare `--cap db` and re-render from `db` on load and on every `subscribe` event (including `{type:'db',resync:true}`, which has no collection). |

## Before publishing

Open at 375px and 1280px, light and dark. Check: no console errors, no horizontal scroll at 375, text readable in both themes, tab through the controls, nothing depends on a network call.

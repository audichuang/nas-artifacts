# Pins

Any tailnet reader can pin an artifact so it stays at the top of the gallery. The pin is kept on the server, so every device sees it. The control is the ☆/★ button in the viewer bar; the gallery has no pin control, it only marks pinned cards with ★ and lists them first.

## Sub-features

- `pin`: viewer `#pin` (☆ / ★, `aria-pressed`) → `PATCH /api/artifacts/<id>` `{ pinned }` (no token, viewer header).
- `gallery`: pinned cards carry `.pinned-mark` (★, title 已釘選) and sort first (`ORDER BY pinned DESC, updated_at DESC`), also within a search.
- `CLI`: no pin command; `artifact get --json` and `list --json` carry `pinned`.

## How to get to it (user POV)

- Reader: opens a page and taps ☆ in the bar. Back on the gallery, on any device, the page is at the top with a ★.

## Driving it with Playwright Chromium (two contexts) and the artifact CLI

Preconditions: instance ready.

```bash
node .claude/skills/verify/scripts/journeys/pins.mjs
```

- Publishes A, waits a second, publishes B (newer); titles carry the nonce.
- Gallery, `#search` = nonce, wait for `#status` 共 2 個: order B, A; no `.pinned-mark`.
- Opens A: `#pin` ☆ / `aria-pressed="false"` → click → ★ / `true`; reload → still ★.
- A second browser context (another device): gallery search → order A, B; A's card has `.pinned-mark`, B's not. CLI `get A --json` → `pinned: true`.
- Unpins in the viewer → reload ☆; the other context's gallery is back to B, A with no mark; CLI `pinned: false`. Both artifacts are deleted at the end.

Evidence: `01-gallery-before.png` … `04-gallery-after-unpin.png`, `cli.log`, `result.json`.

## Gotchas

- Always filter the gallery by the nonce: pinned-first ordering spans the whole data dir, so another journey's pinned artifact would otherwise sit on top.
- The two artifacts are published a second apart so `updated_at` orders them deterministically before pinning.
- Pinning does not change `updated_at`; only the pinned flag reorders the list.

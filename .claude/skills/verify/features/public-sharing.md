# Public sharing

The owner can open a single artifact to anyone with its link. Anonymous visitors then reach it on a separate pair of public listeners, read-only and current version only. While an artifact is private, its public URL looks exactly like an id that does not exist. Closing it again takes effect on the next request.

## Sub-features

- `visibility`: `artifact visibility <id> public|private` and `publish --public|--private`. `artifact list` marks `[public]`.
- `not shown == not found`: a private id on the public app returns 404 with the same body as a missing id, after the id is replaced. The public content listener also returns 404.
- `public viewer`: `<public app>/a/<id>` frames `<public content>/a/<id>/v/<n>/` in `iframe.public-frame`, with frame origin `'null'`. There are no write controls (`#comments, #pin, #menu-delete, #toggle-comments` count 0). Relative links work, and a download button (`a.button[download]`) serves the original file.
- `headers`: `Cache-Control: private, no-cache` and `X-Robots-Tag: noindex, nofollow` on the public viewer and the public content listener.
- `nothing else`: anonymous `/` and `/api/artifacts` on the public app return 404.
- `current only`: after v2, `?v=1` and `/v/1/` on the public listeners return 404.
- `tailnet link`: the tailnet viewer shows `#public-link` with the public URL.

## How to get to it (user POV)

- Owner (on the tailnet): `artifact visibility <id> public`, then copies the link from the CLI output or from the viewer's public-link button.
- Anyone: opens `http://localhost:28792/a/<id>` with no account.

## Driving it with Playwright Chromium (fresh anonymous context) and the artifact CLI

Preconditions: instance ready.

```bash
node .claude/skills/verify/scripts/journeys/public-sharing.mjs
```

- Publishes `index.md` (with `[next page](two.md)`) and `two.md` with `--new`, so it starts private.
- A new browser context with no session opens the public URL: 404. A request for `/a/no-such-id` is also 404 and identical once the id is replaced. `<public content>/a/<id>/v/1/index.md` is 404.
- `artifact visibility <id> public` prints `… is public: <public URL>`, and `list` shows `[public]`.
- Reloads: the response headers are saved, `iframe.public-frame h1` renders, frame origin is `'null'` and there are 0 write controls. The download button gives the same sha256 as `index.md`. The `next page` link renders "Second page". `/` and `/api/artifacts` are 404.
- Tailnet viewer: `#public-link` href equals the public URL.
- Edits and runs `publish <dir>` (v2): stdout keeps `public: <url>` and the public page shows "Second edition.". `?v=1` and `/v/1/index.md` are 404.
- `artifact visibility <id> private`: the public app and `/v/2/index.md` on public content are 404, and `list` drops `[public]`.

Evidence: `01-private-is-404.png` … `05-private-again.png`, `public-viewer-headers.json`, `cli.log`, `result.json`.

## Gotchas

- Always use `scripts/artifact`. The bare CLI defaults to production and reads the real token.
- Use a new browser context for the anonymous side. A context that enrolled a passkey is the owner and sees private pages through the session gate.
- A page with capabilities cannot be public. Use a capability-free artifact here.
- `localhost` (public app) and `127.0.0.1` (public content) are different origins, like the tailnet pair. The public pair must also differ from the tailnet pair (the server refuses to start otherwise).
- A republish without `--public`/`--private` keeps the current visibility. `--private` closes the page before the new content is uploaded.

# Delete and shared blobs

An agent deletes an artifact with the CLI (permanent, so it needs `--yes`). Open viewers are told and sent to the gallery, and every URL of the artifact stops serving. File contents are stored once per content hash and shared between versions, artifacts and assets, so deleting one artifact or asset must leave bytes another one still references. Assets are deleted with the publish token through the API; the CLI has no command for it.

## Sub-features

- `delete`: `artifact delete <id>` without `--yes` fails with `pass --yes`; with it prints `deleted <id>`. Viewers get the `deleted` event: `#toast` 這個 artifact 已被刪除, then `/` after 1.5 s.
- `gone`: `<app>/a/<id>`, `<content>/a/<id>/v/<n>/…` → 404; `list` drops it; `get` fails.
- `shared blob`: an identical file in another artifact still serves.
- `asset delete`: `DELETE /api/artifacts/<id>/assets/<assetId>`: token only (401 without it, also with just `x-artifact-client: viewer`), 204 with it; then the asset URL is 404 and `asset list` prints `(no assets)`.

## How to get to it (user POV)

- Agent: `artifact delete <id> --yes` when the user asks to remove a page.
- Reader with the page open: sees the toast and lands on the gallery.
- Asset delete: no CLI, MCP, viewer or bridge entry point exists today; only an HTTP call with the publish token.

## Driving it with the artifact CLI, Node fetch and Playwright Chromium

Preconditions: instance ready.

```bash
node .claude/skills/verify/scripts/journeys/delete-and-blobs.mjs
```

- Publishes A and B, each with a different `index.html` and an identical `shared.txt` (`text/plain` is servable; `application/octet-stream` files are 404 by design). Both `shared.txt` URLs → 200, equal bytes.
- A viewer tab opens A. `delete A` → non-zero, `pass --yes`, A still in `list --json`. `delete A --yes` → `deleted A`.
- The tab shows the toast and lands on `/`. `list` has B, not A; `get A` fails; Node fetch of `<app>/a/A`, `<content>/a/A/v/1/` and A's `shared.txt` → 404; gallery search by nonce shows only B.
- B's `shared.txt` → 200 same bytes; `read B shared.txt` equal; B renders in a fresh browser context.
- `asset.txt` with the same bytes: `asset upload B asset.txt` prints the URL → 200 equal bytes; `asset list` shows it. `DELETE` the asset with no token → 401 (URL still 200), with only the viewer header → 401 (still 200), with `Authorization: Bearer ${V.token}` (the throwaway token) → 204. Then `asset list` → `(no assets)`, asset URL → 404 (Node fetch: assets are served `immutable`, a browser would answer from cache), B's `shared.txt` still 200 with the same bytes.

Evidence: `01-viewer-on-A.png` … `03-gallery-only-B.png`, `cli.log`, `ids.json`, `result.json`.

## Gotchas

- Always use `scripts/artifact`. The bare CLI defaults to production and reads the real token.
- Blob removal from disk is not observed: that would mean reading the store's blob dir, which helpers must not touch. What is proved is the user-visible half: bytes still referenced keep serving after a delete.
- The CLI has no `asset delete` (`asset upload|list` only), MCP has no asset tool, and the bridge exposes `uploads.put/list/subscribe` only. The journey uses the token-only API route with `V.token`; report the missing CLI command to the owner, it is not a server bug.
- Wait for the toast first, then for the URL: the redirect to `/` fires 1.5 s after the toast.

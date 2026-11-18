# Capabilities: a live page

A page published with `--cap db` gets a shared document store through `artifact.db`. One published with `--cap uploads` can store files viewers add through `artifact.uploads`. The page itself stays sandboxed with no network. Every call goes through the viewer's `postMessage` bridge, and the server checks the declared capability again on each call. Every viewer and the CLI see writes live. A page that declared nothing is refused. A page with capabilities can never be made public.

## Sub-features

- `declare`: `artifact publish <path> --cap db,uploads`. `get` shows `capabilities db, uploads`.
- `db live`: `artifact.db.get/set/subscribe` in the page. A write in one tab re-renders the others over the event stream.
- `db CLI`: `artifact db get|set|list|query|update|delete <id> <collection> …` (needs the token, which the wrapper supplies). CLI writes reach open pages live.
- `uploads`: `artifact.uploads.put(blob, name)` returns an asset URL on the content origin. `artifact asset list <id>` shows it with source `page`.
- `undeclared`: on a page without the capability, the bridge rejects with code `capability_not_declared`. A direct viewer-style API call gets 403 `capability_not_declared`.
- `no public`: `artifact visibility <id> public` on such a page fails with 409 `has_capabilities`.

## How to get to it (user POV)

- Agent: publishes a small tool (counter, checklist, form) with `--cap`, then reads or seeds its data with `artifact db …`.
- Readers: open the viewer URL in several tabs or on several devices and click. Everyone sees the same numbers without reloading.

## Driving it with the artifact CLI and Playwright Chromium (two tabs)

Preconditions: instance ready.

```bash
node .claude/skills/verify/scripts/journeys/capabilities-live-page.mjs
```

- Publishes the counter page twice: `--cap db,uploads` and without caps. The page has `#count`, `#bump` (+1), `#up` (upload `note.txt`) and `#err` (shows `e.code`).
- Tab 1 and tab 2 open the capable page and `#count` reads `0`. The tab 1 frame's `origin` is `'null'`.
- Clicks `#bump` in tab 1's frame and tab 2's `#count` becomes `1`. `artifact db get <id> c main` returns `{"data":{"n":1}}`.
- `artifact db set <id> c main '{"n":41}'` makes both tabs show `41`.
- Clicks `#up`: `#upload` shows `<content>/a/<id>/assets/…/note.txt`. Node fetch of that URL → 200 with the body `hello from the page` byte for byte, and `artifact asset list <id>` has a line with source `page` ending in that same URL.
- Tab 3 opens the page without caps: `#err` shows `capability_not_declared`. `PUT <app>/api/artifacts/<id>/db/c/main` with `x-artifact-client: viewer` and no token gets 403 `capability_not_declared`, saved as `direct-api-call.json`.
- `artifact visibility <capId> public` exits 1 with `409 has_capabilities …`, and `get` still reports `private`.

Evidence: `01-tab2-before.png`, `02-tab2-after-bump.png`, `03-tab1-after-cli-set.png`, `04-uploaded.png`, `05-no-capability.png`, `cli.log`, `direct-api-call.json`, `result.json`.

## Gotchas

- Always use `scripts/artifact`. The bare CLI defaults to production and reads the real token.
- Capabilities and public are mutually exclusive. Use a separate artifact for the public-sharing journey.
- The bridge refuses an undeclared capability in the viewer before any request is sent. The direct API call is what proves the server re-checks.
- Page clicks go through `page.frameLocator('#frame').locator('#bump').click()`, a real click in the sandboxed frame. Do not call page functions with `evaluate` when the point is the user path.
- Anyone on the tailnet may use declared capabilities without a token. That is intended (see AGENTS.md), not a finding.

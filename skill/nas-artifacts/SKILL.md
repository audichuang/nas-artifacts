---
name: nas-artifacts
description: Publishes HTML or Markdown pages to the user's self-hosted Artifacts server on their tailnet (artifacts.ttnas.audichuang.app) and manages them. Use when a page, report, dashboard, diagram, explainer or small tool would serve the user better than terminal text, or they ask to publish, update, share or "show me" one; when they ask what viewers commented on a page; or when a page needs shared data or file uploads. Prefer it over claude.ai Artifacts and Orca Artifacts unless the user names those.
---

# nas-artifacts

Pages live on the user's NAS and open only from devices on their tailnet, with no sign-in. Every page gets a stable URL; republishing the same path makes a new version at the same URL, and open viewers switch to it live.

## Publish

```bash
A=<this-skill-directory>/scripts/artifact.mjs
node "$A" publish <file-or-dir> --icon 📊 --description "<one sentence>" [--cap db,uploads] [--label "<what changed>"]
```

- A directory publishes every file in it (dotfiles and `node_modules` skipped); the entry is `index.html`, `index.md`, or the only page. Relative links between the files work.
- The CLI remembers which artifact each local path became, so publishing the same path again **updates** it: files you removed locally are removed from the new version. `--new` forces a separate artifact; `--id` targets a specific one.
- Write the page into the scratchpad or a project folder, never into a published directory you did not create.
- Done when the command prints the URL; give the user that URL.

Credentials: `ARTIFACTS_TOKEN` env, or `~/.config/nas-artifacts/config.json` (set once with `artifact config --token-stdin`). The token is in the 1Password `Agent` vault, item `nas-artifacts publish token`; pipe it, never print it. Reads (list, get, read, comments) need no token.

## Page contract

The page runs in a sandbox on its own origin:

- **No network.** `fetch`/XHR/WebSocket are blocked, and inside the viewer the page cannot navigate away from its own files. Put data inline, in a published `.js` file that sets a global, or in the `db` capability.
- **No popups or external links**: `target=_blank` and `window.open` do nothing; show a URL as text if the reader needs it.
- **Scripts and styles**: inline, published files, or exact-version URLs from `cdnjs.cloudflare.com`, `cdn.jsdelivr.net`, `unpkg.com`, `cdn.tailwindcss.com`, `code.jquery.com`. Fonts: Google Fonts. Images: published files, data URIs, or assets (`artifact asset upload`) — remote image URLs are blocked.
- **No cookies; `localStorage` throws** (the origin is opaque). Shared state goes in `db`.
- Theme with CSS variables and `prefers-color-scheme`; set an explicit `body` background; make it work at phone width.
- Markdown entries are rendered with GitHub-like styling; raw HTML inside them runs under the same sandbox.

## Page capabilities

Declare what the page uses with `--cap` (republishing without `--cap` keeps the previous set). Inside the page, `window.artifact` is available from the first script:

```js
await artifact.info()                         // { id, version, latestVersion, title, capabilities, viewer: { name } }
// db: per-artifact shared database, every viewer sees the same data
await artifact.db.set('tasks', 't1', { title: 'x', done: false })
await artifact.db.update('tasks', 't1', { done: true })   // shallow merge
await artifact.db.get('tasks', 't1')          // { id, data, createdAt, updatedAt } | null
await artifact.db.list('tasks', { limit, after })          // { docs, next }
await artifact.db.query('tasks', { done: true })           // top-level field equality
await artifact.db.batch([{ op: 'set', collection, id, data }, { op: 'delete', collection, id }])  // all or nothing
artifact.db.subscribe(({ collection, id, deleted }) => rerender())  // live, from any viewer
// uploads: files people add
const { url } = await artifact.uploads.put(fileFromInput)
await artifact.uploads.list()
```

Collection and doc ids: 1-128 chars of `A-Za-z0-9_.:-`. Documents up to 256 KB. Render from `db` on load and on `subscribe`, not from a value cached in the page.

From the shell, `artifact db list|get|set|update|delete|query <id> <collection> …` reads and writes the same data with the token (no capability needed), for seeding or reading back what people entered.

## Feedback

Every artifact has viewer comments. `node "$A" comments <id>` prints them with version numbers; read them before revising a page the user says has feedback.

`node "$A" --help` lists every command.

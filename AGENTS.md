# nas-artifacts

Pages agents publish for one person to read on his own devices over Tailscale: no login to view, nothing leaves the NAS. Every change is weighed against that — a feature that needs accounts, a public URL or a third-party service is out of scope unless he asks.

## Security boundary (the part that must not regress)

- Two origins, one process: the app port serves gallery, viewer and API; the content port serves artifact files. The app origin never serves artifact bytes as a page (`/api/.../files/*` is `attachment`).
- Content responses carry the CSP built in `src/content-server.ts`: `sandbox` **without** `allow-same-origin` and without popups, `connect-src 'none'`, and own-file sources scoped to `<content>/a/<id>/` (plus the `/_/` bridge and markdown CSS). New content routes go through that middleware; widening any of these needs a red-first test showing why it stays safe. The allowed hosts are repeated in the CLI's `check` (`SCRIPT_HOSTS` in `artifact.mjs`) and in SKILL.md's page contract; change all three together.
- The viewer sends `frame-src <content>/a/<id>/` so a framed page cannot navigate anywhere else. Rendered pages are `no-cache` so a bridge or CSP change reaches pages browsers already hold — keep it that way. Their ETag covers the bridge, the markdown CSS, the deck runtime (`content/deck.*`) and `RENDER_VERSION` (which counts the renderer libraries' versions); changing what `renderMarkdown` or the PowerPoint reader (`src/pptx.ts`) emits means bumping its tag (`md-`, `pptx-`), or browsers get a 304 for the old render.
- A published `.pptx` is drawn on the server from the file's own masters, layouts and theme; there is no built-in template to fall back on, and none should be added. Its pictures are served from inside the file at `<content>/a/<id>/v/<n>/…/~office/<file>/ppt/media/*` (image types only), inside the same CSP scope; the file is untrusted input, so the zip and XML readers (`src/zip.ts`, `src/xml.ts`) stay bounded and refuse DTDs.
- Page capabilities reach the API only through the viewer's `postMessage` bridge (`public/viewer.js` ↔ `content/bridge.js`). The server re-checks the declared capability on every viewer call; viewer mutations need the `x-artifact-client: viewer` header, which forces a CORS preflight no other origin passes. Messages a page sends unasked (`report`: place, errors) are untrusted too: the viewer acts on one only after `framePath()` maps it to a file of the version on show. Publish, update, delete and asset delete need the Bearer token.
- Anyone on the tailnet may view, comment, pin, report page errors and use declared capabilities without a token. That is intended, not a gap. The error log needs no declared capability either, so it stays bounded (100 rows of 500-char fields per version, de-duplicated) and diagnostic only: nothing reads it back into a page. A slide deck (`*.slides.md`) that declares no capabilities is the only artifact a viewer may version (`PUT …/slides`): only its entry file changes, through `store.update` with `baseVersion`, and since deck Markdown may carry raw HTML and script, a deck with capabilities stays publish-only; before adding `--cap` to a deck, read what viewers wrote into it.

## Storage invariants

- Blobs are content-addressed and shared across versions and artifacts. Write them only through `commitWithBlobs` (validate first, write, then commit; clean up on failure) and remove them only through the per-delete reference check — a blob another version or asset still references must survive.
- Store writes that emit change events go through `transaction()`, which announces them only after COMMIT; emitting from outside it makes viewers see writes that were rolled back.
- Versions are immutable. `PUT` requires `baseVersion` (or `force`); a 409 carries `details.changedFiles`. Keep that contract: agents on other machines rely on it to avoid overwriting each other.

## Two things ship separately

- The **server** ships as an image: push to `main` → CI (`.github/workflows/ci.yml`) runs `tc` + tests and pushes `ghcr.io/audichuang/nas-artifacts:sha-<full sha>`. It is live only after the nas-ops repo pins that digest in `hosts/ttnas/artifacts/compose.yaml` and Komodo deploys it (`nas-container` skill; record it in that folder's `deployments.md`). This repo holds no deploy config.
- The **skill + CLI** (`skill/nas-artifacts/`) ship to every agent on every machine through aghub (`aghub-cli -g source sync audichuang/nas-artifacts --update --yes`). Machines update at different times, so a server change must keep working with the previous CLI, or land the CLI side first.
- The CLI stays a single dependency-free file runnable on Node 20+; agents run it straight from the installed skill folder. `SKILL.md` is read by Claude Code, Codex, omp, grok and Antigravity alike — write it agent-neutral, with no tool names from one harness.

## Verifying

- `pnpm tc` and `node --test test/*.test.ts` are the gate CI runs. The browser e2e (`test/bridge.e2e.test.ts`) needs Chromium (`CHROMIUM_PATH`, or a Playwright headless shell in the default cache) and binds fixed ports 18790/18791/18799 — run one e2e process at a time.
- Anything touching the sandbox, CSP, bridge or capability checks gets a real-Chromium e2e that fails before the fix; unit tests of headers alone have missed browser behaviour here before.
- After a deploy, accept it from a tailnet machine (ttnas itself cannot resolve its own `*.ttnas.audichuang.app` names): publish a scratch page with the CLI, open the viewer, then delete the page.

## Public repo

Everything committed is public. The publish token lives in the 1Password `Agent` vault (`nas-artifacts publish token`) and as a 0400 file on ttnas; it never goes into the repo, CI, tests or logs. Test tokens are throwaway strings.

UI copy is Traditional Chinese (zh-Hant).

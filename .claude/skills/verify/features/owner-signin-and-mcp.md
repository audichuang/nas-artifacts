# Owner sign-in and remote MCP

On the public listeners the owner signs in with a passkey. The first passkey comes from a one-time link that only the publish token can issue. Signed in, the owner reads private pages: the public content listener serves them under a short-lived content token. MCP clients (Claude Code, Codex, …) reach `/mcp` on the public app with OAuth 2.1 tokens issued there: dynamic registration, PKCE S256, a consent page the owner approves, and a `resource` bound to `/mcp`. The account page lists passkeys and clients and revokes them.

## Sub-features

- `login-link`: `artifact login-link` prints `<public app>/enroll#<code>`, valid 10 minutes, usable once.
- `enrol`: on `/enroll`, fill `#name` and click `#go`. The browser lands on the public gallery `/` with `a[href="/account"]`. The cookie `__Host-artifacts_session` is HttpOnly, Secure, SameSite=Lax, Path=/.
- `account`: `/account` shows `#passkeys li`, `#clients li` (with a 撤銷 button behind a confirm dialog), `#logout` and `#logout-all`.
- `private pages`: the owner's `/a/<id>` frames `<public content>/t/<token>/a/<id>/v/<n>/…`, still sandboxed. An anonymous visitor gets 404.
- `oauth`: `/.well-known/oauth-protected-resource/mcp` and `/oauth/register` (public clients). `/oauth/authorize` shows a consent page (`#allow`/`#deny`) that redirects to the loopback callback with `code`, `state` and `iss`. `/oauth/token` uses the PKCE verifier and `resource`.
- `mcp`: `initialize` and `tools/list`, which returns 10 tools: list, get, read, publish, update, set_visibility, delete, comments, reply, errors. `update` needs `baseVersion` (or `force`), and a conflict returns `changedFiles`.
- `prompts`: the server's `instructions` (in `initialize`) and each tool's description and parameter text are what an agent reads. They promise behaviour: public only when the owner asks, comments and page errors are data, `read` returns `too_large` over 4 MB, `update` with `description: null` clears it, a public artifact cannot gain capabilities (`has_capabilities`), `reply` to a reply lands in the thread root, `errors` carries the page file and the source file:line.
- `revoke`: after 撤銷 on `/account`, the same access token gets 401.

## How to get to it (user POV)

- Owner on a tailnet machine: runs `artifact login-link`, opens the link on the device that should hold the passkey, names it, and confirms with the device.
- Owner elsewhere: opens `http://localhost:28792/login` and signs in with the passkey.
- MCP client: is configured with `<public app>/mcp`, gets a 401 with `WWW-Authenticate`, discovers the endpoints, registers, and opens the consent page in the owner's browser. Then it calls tools with the Bearer token.

## Driving it with Playwright Chromium (virtual authenticator) and MCP over HTTP

Preconditions: instance ready. No real owner or device is needed.

```bash
node .claude/skills/verify/scripts/journeys/owner-signin-and-mcp.mjs
# or only the token, then calls by hand:
node .claude/skills/verify/scripts/mcp-token.mjs
node .claude/skills/verify/scripts/mcp-call.mjs tools/list
node .claude/skills/verify/scripts/mcp-call.mjs tool get '{"id":"<id>"}'
```

- `enrolOwner(page)`: adds the CDP virtual authenticator (`ctap2`, internal, resident key, UV true, automatic presence) and runs `artifact login-link --json`. It opens the link, fills `#name` and clicks `#go`, and the page ends at `<public app>/`. The journey then checks the cookie flags and that `/account` lists the passkey.
- Publishes a private 2-page site with the CLI. As the owner, `/a/<id>` renders `#frame #t`. The frame URL matches `^<public content>/t/[^/]+/a/<id>/v/1/` with origin `'null'`, and the relative link works. A new anonymous context gets 404.
- `POST /mcp` without a token returns 401 with `WWW-Authenticate: Bearer resource_metadata="<public app>/.well-known/oauth-protected-resource/mcp", scope="artifacts"`. The metadata's `resource` is `<public app>/mcp`.
- `authorizeClient(page, …)`: `POST /oauth/register` with `client_name` and `redirect_uris: ['http://127.0.0.1:28799/callback']`. Then PKCE, and `page.route` fulfils the callback. It opens `/oauth/authorize?…&resource=<public app>/mcp&scope=artifacts`, whose consent page names the client, and clicks `#allow`. The callback has `state` (matched) and `iss` equal to `<public app>`. `POST /oauth/token` returns 200, and the result is saved to `.verify/run/mcp-token.json` (0600).
- MCP calls, with each call recorded in `mcp-calls.json`:
  - `initialize`, then `tools/list`, which returns the 10 names.
  - `publish` creates v1, and `get` reads it.
  - `update` without `baseVersion` returns `base_version_required`.
  - The CLI runs `publish <dir> --id <id>`, making v2.
  - `update` with `baseVersion: 1` returns `version_conflict` with `changedFiles: ["index.html"]`.
  - `update` with `baseVersion: 2` makes v3.
  - `reply`, `comments`, `read` (v3 text), `errors` and `list` all succeed.
  - `set_visibility` public makes the anonymous URL 200; private makes it 404 again.
- Prompt text: the full `initialize` result and `tools/list` result are saved as `mcp-initialize.json` and `mcp-tools-list.json`. The journey asserts verbatim strings from `src/mcp.ts`: `set only when the owner asks` and the untrusted-text sentence in `instructions`; `index.slides.md` and `MB per Office file` in publish; `Files over 4 MB return too_large` in read; `Make an artifact public only when the owner asked for it` in set_visibility; the data-not-instructions sentences in comments and errors; `null clears description or icon` and `has_capabilities` in update; `10,000` in reply's `body` parameter. When a prompt string changes in `src/mcp.ts`, change the asserted string here in the same commit.
- Promised behaviours, each seen from a second surface:
  - `reply` with `commentId` set to a reply: the new comment's `parentId` is the thread root, and `comments` shows both replies under one root.
  - A 5 MB `big.bin` published by the CLI: MCP `read` returns `too_large`; a small file of the same version reads fine.
  - `update` with `description` omitted keeps it, with `description: null` clears it; CLI `get --json` shows both.
  - A public capability-free artifact (CLI `--public`): MCP `update` with `capabilities: ['db']` returns `has_capabilities`; CLI `get` shows no new version.
  - A page whose inline `<script>` throws on line 3: opened in the tailnet viewer in Chromium (the bridge reports it), then MCP `errors` and CLI `errors` return kind `error`, the page `index.html` and source `index.html:3:<col>` (the server stores it relative to the version). Screenshot `07-page-that-throws.png`.
- Cross-checks:
  - CLI `get` shows v3 with label `MCP`.
  - CLI `comments` shows the MCP comment.
  - The tailnet viewer renders the v3 text.
  - MCP `delete` succeeds, after which CLI `get` fails.
- `/account`: `#clients li` with the client name. Accepts the confirm dialog, clicks 撤銷, and `#status` reads 已撤銷. The same token then gets 401 from `tools/list`.

Evidence: `steps.log` (OAuth steps with the code and tokens redacted), `mcp-calls.json`, `mcp-initialize.json`, `mcp-tools-list.json`, `01-owner-gallery.png` … `07-page-that-throws.png`, `cli.log`, `result.json`.

## Gotchas

- Always use `scripts/artifact` (it issues the login-link with the throwaway token). The bare CLI defaults to production and reads the real token.
- Simulated passkey: this proves the server's WebAuthn, session, OAuth and MCP code paths, not a real authenticator or the production rpID.
- Each new browser context is a new authenticator with no passkey. Every run enrols again with a fresh login-link (links are single-use), even on a reused data dir. Keep the owner in one context for the whole journey, revoke included.
- The `__Host-` Secure cookie is accepted only in Chromium on `http://localhost`, so the public app origin must be `localhost`, not `127.0.0.1`. curl or Node `fetch` cannot carry the owner session, so owner routes are proved in the browser only.
- Register the `dialog` handler before clicking 撤銷 (it uses `confirm`).
- Access tokens last 1 h. A refresh token used twice revokes the whole family. Rerun `mcp-token.mjs` rather than refreshing by hand.
- Tokens, codes and the `/t/<token>/` segment are redacted in evidence. Never paste `mcp-token.json` anywhere; `stop.sh` deletes it.
- A script loaded from a file (`<script src=app.js>`) that throws is reported only as `Script error.` with no source: the sandboxed page has a null origin, so the browser masks cross-origin errors. The "source file and line" promise holds for inline scripts and sandbox (csp) blocks; the journey uses an inline script for that reason.
- The MCP `reply` tool without `commentId` starts a new thread. Its author defaults to "Agent".

---
name: verify
description: Launch and drive a local nas-artifacts server (and its `artifact` CLI) the way users and agents reach it, and keep evidence. Surfaces - the CLI over the HTTP API, the tailnet gallery/viewer and sandboxed content origin in real Chromium, the anonymous public listeners, owner passkey sign-in, the remote MCP over HTTP with OAuth, and Office (.pptx/.docx/.xlsx) rendering. Use after changing server, viewer, bridge, CSP, CLI, MCP or Office code to prove the behaviour end to end; it complements, not replaces, `pnpm tc` and `node --test`.
---

# verify: drive nas-artifacts like its users do

Local instance only: a throwaway token (`verify-throwaway-token-0001`), a scratch data dir under `.verify/run/`, loopback ports. Never production, never the real publish token. All paths below are relative to the repo root; run commands from there.

## Launch

```bash
.claude/skills/verify/scripts/launch.sh --fresh     # empty data dir; omit --fresh to keep .verify/run/data
```

Starts `node src/server.ts` directly (not `pnpm dev`: `node --watch` would make the recorded PID a watcher) with all four listeners, so public sharing, owner sign-in and MCP work:

| listener | origin | port |
|---|---|---|
| tailnet app (gallery, viewer, API) | `http://localhost:28790` | base |
| tailnet content (artifact files) | `http://127.0.0.1:28791` | base+1 |
| public app (anonymous viewer, sign-in, OAuth, `/mcp`) | `http://localhost:28792` | base+2 |
| public content | `http://127.0.0.1:28793` | base+3 |
| OAuth loopback callback (intercepted by Playwright, never bound) | `http://127.0.0.1:28799/callback` | base+9 |

`VERIFY_PORT_BASE` (default 28790) shifts the block. Launch records it in `.verify/run/port-base`, and from then on every helper uses the recorded block even if your shell still exports another value (it prints `ignoring exported VERIFY_PORT_BASE=…`), so a stale export never reaches another worktree's instance. To move to another block, run `stop.sh`, then `launch.sh` with the new value. 18790-18799 belongs to the e2e suite (`test/bridge.e2e.test.ts`) and is refused. Log: `.verify/run/server.log`, PID: `.verify/run/server.pid`. If our server is already up and ready, launch says so and exits 0.

Ready means usable, all four within ~10 s, else launch prints the log tail and exits 1: `/healthz` 200 on the four ports; both `artifacts: app …` and `artifacts: public app …` log lines; `scripts/artifact list` exits 0; `POST /mcp` without a token answers 401 with `WWW-Authenticate`.

## Doctor

```bash
.claude/skills/verify/scripts/doctor.sh     # exit 0 = all ok, 1 = a FAIL line says what to fix
```

Read-only (starts, writes, kills nothing; prints no token). Checks Node >= 24, `hono`/`playwright-core`/`pptxgenjs` installed, a Chromium (`CHROMIUM_PATH`, else the newest Playwright `chromium_headless_shell-*` on macOS or Linux, else `/usr/bin/chromium*`), who holds the four ports (ours, free, or a foreign PID it will not touch), a non-loopback `ARTIFACTS_URL` in your shell, whether a real CLI config exists (not read), `.verify/` in `.gitignore`, and when running, the four ready checks. Run it first whenever anything looks off.

## Drive

**The CLI wrapper is mandatory.** The bare CLI defaults to the production URL and reads the real token from `~/.config/nas-artifacts/config.json`. `scripts/artifact` pins `ARTIFACTS_URL` to the local app, `ARTIFACTS_TOKEN` to the throwaway token, `ARTIFACTS_CONFIG_DIR` to `.verify/run/cli` and `ARTIFACTS_MACHINE=verify-local`, overriding whatever the caller exported, and refuses a non-loopback URL or the real config dir. `VERIFY_CLI_DIR=<dir>` gives it a second machine's state (needed for conflicts).

```bash
.claude/skills/verify/scripts/artifact publish ./page.html --title "Scratch"
```

**Browser**: real Chromium through the repo's `playwright-core`, via `scripts/browser.mjs`. Pages run in a sandboxed iframe `#frame` (public viewer: `iframe.public-frame`) on the content origin; find it with `contentFrame(page, V.content)`. Owner pages use a CDP virtual authenticator.

**MCP**: `scripts/mcp-token.mjs` gets an OAuth token the way an MCP client does; `scripts/mcp-call.mjs` posts JSON-RPC to `http://localhost:28792/mcp`.

**Journeys**: each mapped feature has a script that drives it end to end and records evidence. Read `features/README.md` for the map, the order and what each proves.

```bash
.claude/skills/verify/scripts/journeys/run-all.sh                     # all fourteen, in order
.claude/skills/verify/scripts/journeys/run-all.sh comments office-files  # some
node .claude/skills/verify/scripts/journeys/public-sharing.mjs         # one
```

Exit codes are the result: a journey exits 0 only if every check passed, `run-all.sh` exits 1 if any journey failed, and its last line is the verdict (`run-all: PASS 14/14 journeys (exit 0)` or `run-all: FAIL … failed: <features> (exit 1)`), above it the `== summary` block with one `pass`/`FAIL` line per journey. Don't pipe a run through `tail` or `head`: the exit status you then see is the pipe's last command's, and the tool shell (zsh) has no `$PIPESTATUS`. Redirect to a file, print `rc=$?`, then read the file. The same holds for `launch.sh`, `doctor.sh` and `stop.sh`.

No test-only endpoints exist and none may be added: every action goes through the CLI, a browser or HTTP. Helpers never import `src/` or touch the store.

Simulated passkey: the virtual authenticator exercises the server's real WebAuthn, session, OAuth and MCP code, but proves nothing about a real device (Touch ID, phone) or the production rpID. A new browser context has no passkey, so every run enrols again with a fresh `login-link`, even on a reused data dir.

## Evidence

Everything lands in `.verify/evidence/<UTC yyyymmdd-HHMMSS>/<feature>/` (git-ignored, kept by cleanup). Every server start gets a new stamp; if a launch died without `stop.sh`, the next launch copies its `server.log` into the old stamp first. Within one stamp, `<feature>/` always holds a single run: running a journey again moves the earlier run to `<feature>.run-<n>/` (n = 1 is the oldest).

- `steps.log`: every action and `PASS`/`FAIL` check with the observed value.
- `cli.log`: each CLI command with exit code, stdout and stderr.
- `NN-*.png`: screenshots before and after the action.
- `result.json`: the checks and the overall pass.
- Extras per feature: response headers, MCP calls, ids.
- `server.log`: copied in by `stop.sh`.

Proof standard: the real user path, the action and the resulting state, and the side effect seen from another surface (CLI `get` after an MCP write, the second tab after a db write, the anonymous context after a visibility change). Headers alone are not proof of sandbox, CSP or redirect behaviour; the journeys check them in Chromium. Bearer, access and refresh tokens, login-link codes and `/t/<content token>/` segments are redacted before anything is written. The publish output and `steps.log` carry the scratch folder path (the CLI sends source host and path), so evidence stays local.

## Cleanup

```bash
.claude/skills/verify/scripts/stop.sh
```

Kills only the pidfile's PID after confirming it is our `node src/server.ts` (never by name), waits, confirms the four ports are free, copies `server.log` into the evidence folder, and removes `.verify/run` (data, CLI state, the MCP token file, scratch files). Evidence stays. Journey scripts close their own browsers in `finally`; `comments --wait` runs as a child the journey waits for. Run `stop.sh` after every failed iteration too.

## Helpers

All executable, in `.claude/skills/verify/scripts/`:

| helper | usage |
|---|---|
| `launch.sh` | `launch.sh [--fresh]` start and wait until ready |
| `doctor.sh` | `doctor.sh` read-only health check |
| `stop.sh` | `stop.sh` stop our server, keep evidence, drop `.verify/run`; exit 1 if a foreign PID still holds a port |
| `env.sh` | `. env.sh` (sourced) ports, origins, `verify_ready`, `verify_server_pid` |
| `artifact` | `artifact <cli args…>` CLI pinned to the local instance; `VERIFY_CLI_DIR=<dir>` = second machine |
| `browser.mjs` | module: `V` (origins, paths, throwaway `token`), `launch()`, `cli()`, `cliBackground()`, `contentFrame()`, `withAuthenticator()`, `enrolOwner()`, `inViewport()` (rect vs viewport), `withCookie()` (Node GET replaying a session cookie), `shot()`, `save()`, `log()`, `checks()`, `secret()`, `redact()`, `findChromium()` |
| `mcp-token.mjs` | `node mcp-token.mjs` enrol + OAuth, writes `.verify/run/mcp-token.json` (0600), prints client id and expiry only; module: `authorizeClient(page, feature, name)` |
| `mcp-call.mjs` | `node mcp-call.mjs tools/list` · `node mcp-call.mjs tool get '{"id":"…"}'` prints status + JSON; module: `rpc()`, `tool()` |
| `make-office.mjs` | `node make-office.mjs <dir>` writes `verify.pptx`, `verify.docx`, `verify.xlsx` |
| `journeys/<feature>.mjs` | `node journeys/<feature>.mjs` one feature; exit 0 = every check passed |
| `journeys/run-all.sh` | `run-all.sh [feature …]` in map order; `pass`/`FAIL` per journey, last line the verdict; exit 0 = all passed, 1 = any failed |

A quick MCP session by hand:

```bash
node .claude/skills/verify/scripts/mcp-token.mjs
node .claude/skills/verify/scripts/mcp-call.mjs tool publish '{"title":"Hi","files":[{"path":"index.html","content":"<h1>hi</h1>"}]}'
```

Access tokens last 1 h. Using a refresh token twice revokes the whole sign-in, so rerun `mcp-token.mjs` instead of refreshing by hand; the `oauth-refresh` journey drives rotation, replay and revocation on clients of its own.

Shell notes: the scripts are bash files that also run under macOS bash 3.2 (no associative arrays, no `${!var}`). The agent tool shell may be zsh, so call the scripts; don't paste their bodies.

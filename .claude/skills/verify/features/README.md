# nas-artifacts verification map

This directory is the maintained source for verifying what users and agents can do with nas-artifacts. Read this index before driving the app, then use the matching feature file as the recipe. Each feature has a journey script that does exactly what its file describes.

## Baseline preconditions

- `.claude/skills/verify/scripts/launch.sh --fresh` printed `ready`. `doctor.sh` exits 0.
- Every CLI call goes through `.claude/skills/verify/scripts/artifact`, never the bare CLI, which defaults to production.
- Evidence goes to `.verify/evidence/<stamp>/<feature>/`, where `<stamp>` is the server start time in `.verify/run/stamp`. A rerun of a journey in the same stamp moves the earlier run to `<feature>.run-<n>/`.
- Never drive an instance this run did not start. `doctor.sh` shows whose PID holds the ports.

## Features

| feature | surface | journey in one line | file |
|---|---|---|---|
| publish-and-view | CLI + tailnet gallery/viewer + content origin | check, publish HTML/MD/deck, find in gallery, read sandboxed in the viewer | [publish-and-view.md](publish-and-view.md) |
| versions-and-conflict | CLI (two config dirs) + viewer live update | republish to same URL, diff, second machine's conflict, --force, prune | [versions-and-conflict.md](versions-and-conflict.md) |
| comments | viewer comment panel + CLI | comment in viewer, CLI reads, replies, resolves, `--wait` wakes | [comments.md](comments.md) |
| capabilities-live-page | CLI `--cap`/`db`/`asset` + two viewer tabs via the bridge | shared counter live across tabs and CLI, upload, undeclared refused | [capabilities-live-page.md](capabilities-live-page.md) |
| public-sharing | CLI `visibility` + anonymous public listeners | private is 404, public is read-only sandboxed, current version only, private again | [public-sharing.md](public-sharing.md) |
| owner-signin-and-mcp | public app sign-in, `/account`, OAuth consent, `/mcp` | passkey from login-link, private page under content token, every MCP tool, revoke | [owner-signin-and-mcp.md](owner-signin-and-mcp.md) |
| office-files | CLI publish/read/export + viewer Office rendering | .pptx deck with picture, .docx text, .xlsx tabs, byte-identical downloads | [office-files.md](office-files.md) |
| owner-sessions | public app `/login`, `/account` + Node cookie replays | sign in again, logout, logout-all and passkey removal end sessions and rotate content tokens | [owner-sessions.md](owner-sessions.md) |
| oauth-refresh | `/oauth/token`, `/oauth/revoke`, `/mcp` + `/account` | refresh rotates once, a replay revokes the family, 撤銷 and revoke end a client | [oauth-refresh.md](oauth-refresh.md) |
| deck-edit | viewer deck editor (two tabs) + CLI | 編輯 → live preview → 儲存 makes v2 瀏覽器編輯, stale draft never overwrites, caps deck not editable | [deck-edit.md](deck-edit.md) |
| pins | viewer `#pin` + gallery in two contexts + CLI | ☆→★ survives reload, another device lists it first with ★, unpin reverses | [pins.md](pins.md) |
| phone-viewer-bar | viewer and public app at 390x844 | bar slides away and back on a phone, not on desktop; a page can hide it; sign-in and visibility stay reachable | [phone-viewer-bar.md](phone-viewer-bar.md) |
| page-errors | viewer ⚠ dialog + CLI `errors` | throw, sandbox block, rejection, missing picture with page files; repeats counted; per version | [page-errors.md](page-errors.md) |
| delete-and-blobs | CLI delete/asset delete + viewer; 401 over HTTP | delete needs --yes, viewer told, URLs 404, shared bytes survive; asset delete token-only | [delete-and-blobs.md](delete-and-blobs.md) |

## Full pass, in order

```bash
.claude/skills/verify/scripts/launch.sh --fresh
.claude/skills/verify/scripts/doctor.sh
.claude/skills/verify/scripts/journeys/run-all.sh > .verify/run-all.txt 2>&1; echo "rc=$?"   # the fourteen above, in order
.claude/skills/verify/scripts/stop.sh
```

`run-all.sh` exits 0 only if every journey passed and 1 otherwise; its last line is the verdict and the `== summary` block above it has one `pass`/`FAIL` line per journey. Redirect rather than pipe it (a pipe reports the last command's exit code), then read the file.

The journeys do not depend on each other; each creates its own artifacts with a per-run nonce in titles, so they also run on a reused data dir. Run one with `node .claude/skills/verify/scripts/journeys/<feature>.mjs`.

## Proof and failure reporting

- A journey's `result.json` lists every check with the observed value. A journey exits 0 only if all passed; `run-all.sh` exits 1 if any journey failed.
- When a step fails because the product is wrong, report the failing check and its evidence. Do not loosen the check. Record the reason in that feature's Gotchas.
- A skipped entry point is reported as not run. Do not count it as verified through another path.

## Feature file contract

Each file starts with an H1 and a paragraph describing the behaviour, then exactly four H2s in this order: `Sub-features`, `How to get to it (user POV)`, `Driving it with <harness>`, `Gotchas`.

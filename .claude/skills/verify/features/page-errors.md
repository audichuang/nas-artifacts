# Page errors

What breaks while a page runs in a reader's browser is logged per version, so the agent that published it can fix it. The bridge in the sandboxed page reports script errors, unhandled Promise rejections, files that failed to load and requests the sandbox blocked; the viewer maps each to a file of the version on show and posts it. The viewer shows a ⚠ count; the CLI and MCP list the rows. Repeats bump a count instead of adding rows.

## Sub-features

- `kinds`: `error` 腳本錯誤, `rejection` 未處理的 Promise, `resource` 載入失敗, `csp` 被沙箱擋下.
- `viewer`: `#errors` (hidden at zero) with `#error-count` = number of distinct rows (not the sum of counts); clicking opens `#error-dialog` with `#error-list li.thread` (`.comment-meta` spans: kind label, page file, `×N`; `.comment-body` message; `.error-source`), closed by `#close-errors`.
- `CLI`: `artifact errors <id> [--v N] [--json]`, line format `[kind] <page file> ×N  <message>  (<source>)`, or `(no errors reported for this version)`.
- `per version`: rows belong to the version that was on show.

## How to get to it (user POV)

- Reader: sees ⚠ 4 in the bar, clicks it to see what broke.
- Agent: `artifact errors <id>` after publishing, then fixes and publishes again.

## Driving it with Playwright Chromium and the artifact CLI

Preconditions: instance ready.

```bash
node .claude/skills/verify/scripts/journeys/page-errors.mjs
```

- `index.html`: separate inline scripts on line 4 `throw new Error('verify boom <nonce>')`, line 5 `fetch('<public app>/.well-known/oauth-protected-resource/mcp').catch(() => {})` (a CORS-open target, so only `connect-src 'none'` stops it), line 6 `Promise.reject(new Error('verify reject <nonce>'))`, and `<a id="next" href="two.html">`. `two.html` has `<img src="missing.png">`.
- Viewer: `#errors` appears with `#error-count` 3; clicking `#next` in the frame → 4. The dialog has four rows, one per kind label; 載入失敗 on `two.html`, the other three on `index.html` (`dialog-rows.json`).
- CLI `errors <id> --json` (polled until four rows, `errors-v1.json`): `error` `index.html`, message contains `verify boom`, source `index.html:4:<col>`; `csp` `index.html`, message exactly `connect-src blocked <the fetch URL>`, source `index.html:5:<col>`; `rejection` `index.html`, message `Error: verify reject <nonce>`; `resource` `two.html`, `<img> failed to load`, source `missing.png`. Plain `errors <id>` prints four `[kind] <file> ×1  …` lines.
- Second visit with `page.goto(art.url)`: the three `index.html` rows reach ×2, `two.html` stays ×1, still four rows, `#error-count` still 4.
- Publishes a fixed v2: the viewer shows no ⚠; `errors <id>` prints `(no errors reported for this version)`; `errors <id> --v 1` still lists the four.

Evidence: `01-errors-index.png` … `04-fixed-v2.png`, `dialog-rows.json`, `errors-v1.json`, `errors-v1-second-visit.json`, `cli.log`, `result.json`.

## Gotchas

- Use `page.goto(art.url)` for the second visit, not `reload()`: after the link the address bar names `?path=two.html`, so a reload reopens two.html and the index rows never repeat.
- The rejection row's source is the stack line as Chromium gives it, `at <content>/a/<id>/v/1/:6:24`, a full content URL that `shortSource` does not shorten (it only handles `<content>/…` sources without the `at ` prefix). The journey pins kind, file and message for that row, not its source.
- Script files loaded with `<script src>` that throw are reported as `Script error.` without a source (null-origin sandbox); inline scripts keep file and line. See owner-signin-and-mcp.
- The bridge reports at most 20 errors per page load and batches them for 300 ms, so poll the CLI rather than reading it once.
- If Chromium adds an unexpected row (favicon, extra rejection), find out why before widening an assertion.

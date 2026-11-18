# Versions and conflicts

Publishing again from the same folder makes a new version at the same URL. An open viewer follows it live. When two machines edit the same artifact, the second publish is refused with the list of files that changed remotely and two ways out. Old versions can be pruned. A link to a pruned version opens the current one.

## Sub-features

- `republish`: same path, printed as `(v2, same URL)`. `--label` names the version.
- `live viewer`: an open viewer switches to the new version over the event stream, without a reload, at the same URL.
- `diff`: `artifact diff <id>` shows `-`/`+` lines between versions.
- `checkout`: `artifact checkout <id> <dir>` downloads the current version and remembers it for that config dir.
- `conflict`: a publish based on an older version exits 1, prints `changed remotely since vN: <files>` and the options `checkout` / `--force`, and uploads nothing.
- `force`: `publish --force` overwrites.
- `prune`: `artifact prune <id> --keep 1` keeps only the current version. Then `?v=1` opens the current one.

## How to get to it (user POV)

- Agent A edits its folder and runs `artifact publish <folder>` again. The reader's open tab updates by itself.
- Agent B on another machine runs `artifact checkout <id>`, edits, and publishes. If A published in between, B is told what changed and chooses between checkout and `--force`.
- The owner tidies up with `artifact prune`.

## Driving it with the artifact CLI (two config dirs) and Playwright Chromium

Preconditions: instance ready.

```bash
node .claude/skills/verify/scripts/journeys/versions-and-conflict.mjs
```

- A: `publish <a> --new` gives v1. The viewer shows `#v` "version one". The script sets a `window` marker that a reload would wipe.
- A edits and runs `publish <a> --label second`: stdout `(v2, same URL)`. The viewer shows "version two", the marker is still set, the URL is the same and the frame URL has `/v/2/`.
- `diff <id>` shows `-…version one…` and `+…version two…`.
- B (`VERIFY_CLI_DIR=.verify/run/work/cli-second-<nonce>`): `checkout <id> <b>/copy` checks out v2.
- A: `publish <a>` gives v3.
- B edits and runs `publish <b>`: exit 1, stderr `is at v3, but your copy was based on v2. Nothing was uploaded.`, `changed remotely since v2: index.html`, and options 1 and 2. `get` still says v3.
- B: `publish <b> --force` gives v4. `get` lists v4..v1 and `read` equals B's file.
- `prune <id> --keep 1` gives `dropped v3, v2, v1`. `get` lists only v4. `/a/<id>?v=1` frames `/v/4/`.

Evidence: `cli.log` (with the full refusal text), `01-viewer-v1.png`, `02-viewer-v2-live.png`, `03-pruned-v1-shows-v4.png`, `result.json`.

## Gotchas

- Always use `scripts/artifact`. The bare CLI defaults to production and reads the real token.
- A conflict needs two CLI states. The CLI keeps the base version per path in `ARTIFACTS_CONFIG_DIR/state.json`, so one config dir can never conflict with itself. Use `VERIFY_CLI_DIR` for machine B.
- `checkout` refuses a non-empty dir without `--force`. Point it at a fresh subfolder.
- Proving "no reload" needs a marker on `window`. Playwright's `framenavigated` also fires for same-document URL changes, so counting navigations gives a false failure.
- The viewer follows new versions only while it shows the latest one. A viewer pinned with `?v=N` just gets a toast.

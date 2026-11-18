# Deck editing in the viewer

A reader can edit a capability-free Markdown deck (`*.slides.md` entry) right in the tailnet viewer. Typing previews the slides live without making a version; saving makes a new version labelled 瀏覽器編輯 (plus the reader's name if they gave one when commenting). Saving is based on the version the draft started from, so a draft never overwrites a newer version. A deck that declares capabilities is publish-only.

## Sub-features

- `open`: `#edit` (編輯) shows only for a capability-free Markdown deck on its entry; it opens `#editor` with `#editor-source` (the entry's source), `#editor-status` 根據 v<n>, `#save-deck` disabled.
- `preview`: typing marks it 根據 v<n> · 未儲存, enables `#save-deck`, and after 250 ms `POST …/slides/preview` swaps the rendered deck into the frame.
- `save`: `PUT /api/artifacts/<id>/slides` `{ markdown, baseVersion }` → `#editor-status` 已存成 v<n+1>; label `瀏覽器編輯` (or `瀏覽器編輯 · <name>`). Other viewers update live.
- `conflict`: a stale base → 409 → `#editor-status` 別人已存了新版本 and a `confirm` (OK loads the latest and drops the draft; Cancel keeps it).
- `capabilities`: a deck with capabilities hides `#edit`; the server answers a viewer `PUT …/slides` with 403 `deck_not_editable`.

## How to get to it (user POV)

- Reader: opens a deck, clicks 編輯, fixes a typo on the left while watching the slide on the right, clicks 儲存 (⌘S).
- Agent afterwards: `artifact get <id>` shows the 瀏覽器編輯 version; `artifact checkout <id> <dir>` fetches it before the next edit.

## Driving it with Playwright Chromium (two tabs) and the artifact CLI

Preconditions: instance ready.

```bash
node .claude/skills/verify/scripts/journeys/deck-edit.mjs
```

- Publishes `talk.slides.md` (two slides) with `--new`; entry is `talk.slides.md`, v1. Tabs 1 and 2 open it (each `browser.newPage()` is its own context, so no `artifacts.viewerName`).
- Tab 1: `#edit` visible → click → `#editor-source` equals the published source byte for byte; status `根據 v1`, 儲存 disabled.
- `fill()` the edited source (heading `Edited <nonce>`) → status `根據 v1 · 未儲存`, 儲存 enabled; the frame's deck gains a `.slide` with the edited heading; CLI `get` still v1.
- Click 儲存 → `已存成 v2`. Tab 2's `.slide.active` shows the edit without a reload. CLI `get --json`: `currentVersion 2`, version 2 `label === '瀏覽器編輯'`; `#version option[value="2"]` reads `v2 · 瀏覽器編輯 · 最新`; `artifact checkout <id> <dir> --force` writes the typed bytes.
- Types a further change (status `根據 v2 · 未儲存`), and meanwhile the CLI publishes v3 from the checkout dir. Click 儲存 with `page.once('dialog', d => d.dismiss())` → `#editor-status` 別人已存了新版本; CLI `get` still v3 and `read` returns the CLI's text.
- Same source published with `--cap db`: `#edit` hidden once the deck renders; `PUT <app>/api/artifacts/<id>/slides` with `x-artifact-client: viewer`, no token → 403 `deck_not_editable` (saved as `direct-slides-put.json`); CLI `get` still v1.

Evidence: `01-editor-open.png` … `05-cap-deck-no-edit.png`, `cli.log`, `direct-slides-put.json`, `result.json`.

## Gotchas

- `fill()` leaves the cursor at the end, and the preview follows the cursor, so the frame then shows the last slide: wait for the edited `.slide` to be attached, not for `.slide.active` to contain it.
- The label includes the reader's name if `localStorage['artifacts.viewerName']` is set (from a first comment). A fresh context has none, so the label is exactly `瀏覽器編輯`.
- `artifact checkout` refuses a non-empty dir; the journey passes `--force` into a scratch dir. A publish from that dir sends the checked-out version as base, so it makes v3 cleanly.
- Register the `dialog` handler before clicking 儲存 in the conflict step; dismissing keeps the draft so the status stays readable.
- `.pptx` entries are edited through ONLYOFFICE, which the verify instance does not configure; this journey covers Markdown decks only.

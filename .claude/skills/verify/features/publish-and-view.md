# Publish and view

An agent publishes a file or folder with the CLI and gets a stable viewer URL. The reader finds the page in the gallery and reads it in the viewer. The page runs in a sandboxed iframe on a separate content origin, with no network access and no same-origin powers. Markdown is rendered with a table of contents, and `*.slides.md` files become decks.

## Sub-features

- `check`: `artifact check <path>` reports what the sandbox will block, e.g. `fetch(` -> connect-src 'none'.
- `publish`: `artifact publish <path> --title --icon --description` returns `url`, `contentUrl` and `currentVersion: 1`.
- `gallery`: cards on `/`; `#search` does full-text search over page text.
- `viewer`: `/a/<id>` frames `<content>/a/<id>/v/<n>/` in `#frame`. Relative links move the viewer to `?path=<file>`.
- `sandbox`: frame `origin === 'null'`, `fetch` to the app API fails, CSP `sandbox` without `allow-same-origin`, `connect-src 'none'`. The viewer's `frame-src` is scoped to `<content>/a/<id>/`.
- `top-level redirect`: opening a content page URL as a tab lands on the viewer.
- `markdown`: headings and a `nav.toc` once there are 3 or more h2/h3.
- `deck`: `*.slides.md`. Arrow keys pressed in the viewer turn `.slide.active`.
- `read`: `artifact read <id>` prints the published bytes.

## How to get to it (user POV)

- Agent: `artifact publish ./report/ --title "…"`, then shares the printed URL.
- Reader: opens `http://localhost:28790/`, types in the search box, clicks a card, and clicks links inside the page. A deck turns with arrow keys.

## Driving it with the artifact CLI and Playwright Chromium

Preconditions: instance launched and ready.

```bash
node .claude/skills/verify/scripts/journeys/publish-and-view.mjs
```

What it does, with the observable result of each step:

- Writes a site (index.html with a `fetch('<public app>/.well-known/oauth-protected-resource/mcp')` script that records any `securitypolicyviolation` in a `#leak` marker, and `<a id="next" href="two.html">`), `notes.md` (3 h2s) and `talk.slides.md` (2 slides) under `.verify/run/work/`.
- `artifact check <site>`: one warning naming `fetch(` and connect-src 'none'.
- `artifact publish <site> --new --title … --icon 🧪 --description … --json`: `currentVersion` 1, `url` on `:28790`, `contentUrl` on `:28791`. Then it publishes the Markdown file and the deck, and `artifact list` shows all three.
- Gallery `/`: `.card` with the title. `#search` filled with a unique needle from the page text leaves exactly one `.card`.
- Clicks the card: the viewer response has `content-security-policy: frame-src <content>/a/<id>/`. `#frame #t` renders. The page frame's `origin` is `'null'`. Control: an `about:blank` page (opaque origin, no CSP) fetches the same URL and gets `null 200`, since that route answers `Access-Control-Allow-Origin: *`. In the frame, `#leak` reads `csp:connect-src`, so the CSP blocked the fetch, not CORS. The frame document response CSP has `sandbox allow-scripts …` without `allow-same-origin` and has `connect-src 'none'`.
- Clicks `#next` in the frame: `#two` renders and the viewer URL becomes `?path=two.html`.
- A new tab goes to `<content>/a/<id>/v/1/index.html` and ends on `<app>/a/<id>`.
- `artifact read <id>` is byte-equal to index.html.
- Markdown viewer: `h1` and `nav.toc li` × 3. Deck viewer: `.slide.active` "Slide one", then ArrowRight gives "Slide two".

Evidence: `cli.log`, `headers.json` (viewer and frame document headers), `01-gallery` … `07-deck-slide2.png`, `result.json`.

## Gotchas

- Always use `scripts/artifact`. The bare CLI defaults to the production URL and reads the real token from the user's config file.
- `localhost` (app) and `127.0.0.1` (content) are different origins on purpose. Swapping them breaks the isolation the checks rely on.
- `curl` on a content URL gets a 200 and never shows the top-level redirect. Sandbox, CSP and redirect proof must come from Chromium.
- The frame document is `/a/<id>/v/<n>/` (no `index.html`) when the entry is `index.html`. Match both forms when capturing its response.
- The CLI remembers a published path in `ARTIFACTS_CONFIG_DIR/state.json`. Republishing the same path makes v2 of the same artifact, so the journey passes `--new`.
- Gallery search is debounced. Wait for the card count to settle before counting.
- The in-page fetch must target a URL that sends `Access-Control-Allow-Origin: *`. The tailnet `/api` sends no CORS headers, so a fetch from the opaque-origin frame fails there with or without `connect-src 'none'`. An earlier version of this journey fetched `/api/artifacts`, and its marker still read `blocked` with the CSP widened to `connect-src *`. The control check keeps the target honest.
- Publish output and the gallery card carry the source host (`verify-local`) and the scratch path.

# Phone viewer bar

On a phone-width screen (≤ 720 px) the viewer bar slides away while the reader scrolls down and comes back when they scroll up or reach the top, so the page gets the whole screen. Desktop keeps the bar. The scroll position comes from the page (the bridge's `place` report), which is untrusted, so a page can hide the bar; AGENTS.md therefore keeps controls that must stay reachable (sign-in state, visibility) out of that bar.

## Sub-features

- `slide`: `slideBar(delta, scrollY)` in `public/viewer.js` adds `body.bar-away` when a report moves down > 24 px past scrollY 80, removes it on a move up > 24 px, below 80, or while comments or the editor are open. CSS gives `.viewer-bar` `margin-top: calc(-1 * var(--bar-h))` (92 px on phones).
- `desktop`: wider than 720 px the class is never added.
- `page can hide it`: any page can post a `place` report with a large `scrollY`.
- `sign-in state`: the public gallery's 帳號 link and `/account`'s 登出 / 登出所有裝置 are not in a sliding bar.
- `visibility`: the tailnet gallery card's 公開 mark (`.public-mark`), and in the viewer the 公開 chip `#public-link`. The chip is a sibling after the bar's `<header>`, not inside it: `position: fixed` at the bar's right end (`top: 9px`), and while it shows the bar keeps `--chip-w` of padding on its right for it (`public/app.css`). With the bar away on a phone the chip moves to `top: 6px` and `.viewer` starts `--chip-strip` (46 px) down, so the chip sits in its own strip instead of over the top-right of the page; pages that are not public still get the frame at top 0. The same chip appears in the owner's viewer on the public app.

## How to get to it (user POV)

- Reader on a phone: opens a long page, reads down (bar goes), flicks up (bar returns).
- Owner on a phone: 帳號 on the public gallery, then 登出.

## Driving it with Playwright Chromium at 390x844

Preconditions: instance ready.

```bash
node .claude/skills/verify/scripts/journeys/phone-viewer-bar.mjs
```

All positions come from `inViewport(page, selector)` (`getBoundingClientRect` against the viewport), never `isVisible()`.

- (a) A long page (80 × 60 px paragraphs) at 390 px: bar on screen; `page.mouse.wheel(0, 800)` over `#frame` scrolls the frame (frame `scrollY` > 80), `body.bar-away`, bar bottom ≤ 0, frame top 0; `wheel(0, -400)` → class gone, bar top 0, `#toggle-comments` on screen.
- (b) The same page at 1200 px: after the wheel the bar is still on screen, no `bar-away`.
- (c) A short page whose script, 600 ms after `load`, posts `{ __artifact: 1, report: { type: 'place', path: location.pathname, hash: '', scrollY: 5000 } }`: the bar goes away, and a wheel up does not bring it back. This is the documented property, expected to pass.
- (d) Owner enrolled in a 390 px context: the public gallery's `a[href="/account"]` is on screen, `/account`'s `#logout` and `#logout-all` are on screen and enabled, and neither page has a `.viewer-bar`.
- (e) The (c) page made public with `artifact visibility <id> public`: the tailnet gallery card's `.public-mark` is on screen at 390 px; in the viewer, control first: `#public-link` on screen while the bar shows; then after the page's forged report hides the bar, the chip is still on screen, outside `.viewer-bar`, hit-testable (`elementFromPoint` at its centre returns it), does not intersect the `#frame` rect, and links to the public URL. Both checks also require `body.bar-away`, which the journey waits for with `waitForFunction` (a miss fails the checks rather than throwing). The owner (enrolled in (d)) opens the same page on the public app and gets the same result. All indicators and their rects are saved in `visibility-indicators.json`; the 公開 page's viewer is then shot at 390 px and 1280 px in light and dark.

Evidence: `01-phone-bar-away.png` … `08-owner-viewer-bar-hidden.png`, `09-viewer-public-{390,1280}-{light,dark}.png`, `visibility-indicators.json`, `cli.log`, `result.json`.

## Gotchas

- Check (e) was red until the chip moved out of the bar: in `.bar-actions` it sat at top -44 / bottom -11 at 390 px once the page's forged report hid the bar. (c) still passes on purpose: a page can hide the bar; what (e) guards is that the visibility mark is not in it. Don't move the chip back into `.viewer-bar`, and keep `--chip-w` at least the chip's width or it covers the bar's last control (the e2e test checks for overlap). Before the strip, the chip stayed fixed over the frame's top-right corner with the bar away and took taps meant for the page's own controls there; the e2e test and (e) now check `coversFrame`.
- The forged report must land after the bridge's own: the bridge posts `scrollY: 0` on `DOMContentLoaded`, and a report below 80 removes the class. Hence the `load` + 600 ms delay and the `#sent` marker the journey waits for.
- Wait ~700 ms after a wheel: the bridge throttles scroll reports to 250 ms and the slide transition takes 0.2 s.
- `page.mouse.wheel` over the frame does scroll the cross-origin sandboxed frame in the headless shell (frame `scrollY` 800 observed); no keyboard fallback needed.
- The owner viewer on the public app has the same bar and the same chip; (e) drives it for the chip. The owner has no sign-in control in that bar (the 帳號 link is on the gallery).

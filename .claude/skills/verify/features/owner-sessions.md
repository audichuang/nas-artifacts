# Owner sessions

The owner signs in on the public app with a passkey and signs out again, on one device or everywhere. Signing out ends the session on the server, not just in the browser. Signing out everywhere, and removing a passkey, also rotate the content key, so every `/t/<content token>/` URL a frame already holds stops working at once instead of after its 12 h expiry. A removed passkey can no longer sign in, and the last passkey cannot be removed.

## Sub-features

- `sign in again`: `/login`, `#go` (用 passkey 登入) with an existing passkey lands on `/` (or `data-next`). On failure `#status` reads 已取消，或這台裝置沒有可用的 passkey (no credential) or 這把 passkey 沒有註冊在這裡 (credential unknown to the server).
- `logout`: `/account` `#logout` → `POST /auth/logout` ends this session and lands on `/login`.
- `logout-all`: `#logout-all` (behind `confirm`) → `POST /auth/logout-all` deletes every session and writes a new `content_token_key` in the same transaction.
- `sessions`: `#sessions` counts live sessions.
- `remove passkey`: each `#passkeys li` has a 刪除 button (behind `confirm`), disabled when only one passkey is left. `DELETE /api/account/passkeys/:id` removes it and ends all sessions (rotating the key); the page goes to `/login?next=/account`.
- `passkey provider sync` (WebAuthn Signal API, feature-detected): `/account` load calls `signalAllAcceptedCredentials` with the server's `rpId`, `userId` and full list; a removal calls it with the list the DELETE response returns (the session is already gone); a sign-in the server answers 404 (unknown credential) calls `signalUnknownCredential`.

## How to get to it (user POV)

- Owner on a phone or laptop: opens `<public app>/login` and taps 用 passkey 登入.
- From the gallery: 帳號 → 登出 / 登出所有裝置, or 刪除 next to a lost device's passkey.

## Driving it with Playwright Chromium (virtual authenticator) and Node fetch replays

Preconditions: instance ready.

```bash
node .claude/skills/verify/scripts/journeys/owner-sessions.mjs
```

- Context A enrols `device A <nonce>` (`enrolOwner`). The CLI publishes a private folder (`index.html` with `#t`, `data.txt`). The journey opens `/a/<id>` as the owner and builds `tokenUrl1 = new URL('data.txt', <frame URL>)` under `/t/<token>/a/<id>/v/1/`. Every cookie value and content token is registered with `secret()`.
- Controls: `withCookie(<public app>/a/<id>, cookieA)` (Node `fetch` with an explicit `cookie: __Host-artifacts_session=…` header) → 200; the same GET without a cookie → 404; `tokenUrl1` → 200 with the `data.txt` bytes.
- `#logout` → `/login`. Replaying cookie A on `/a/<id>` and `/api/account` → 404, and the browser gets the 404 page with 擁有者登入. `tokenUrl1` after a single logout is logged as an observation (200: one logout does not rotate the key).
- Same page, `/login` → `#go` → `/` (the authenticator kept the credential). The private page renders; cookie B ≠ A.
- `ctxA.clearCookies()`, sign in again → cookie C; `tokenUrl2` from a fresh frame; `#sessions` ≥ 2. B, C and `tokenUrl2` → 200 before; then `#logout-all` (dialog accepted) → `/login`; B and C → 404; `tokenUrl1` and `tokenUrl2` → 404.
- Sign in again: the passkey survives; the frame's new token URL differs and serves `data.txt`; `#sessions` reads exactly `1`.
- Context B enrols `device B <nonce>` from a second login-link; its cookie and token URL → 200. On page A's `/account` every 刪除 is enabled; accept the dialog and click 刪除 on the device B row → `/login?next=/account`. Device B's cookie, device A's pre-removal cookie and device B's token URL → 404. Device B on `/login` → `#status` 這把 passkey 沒有註冊在這裡, URL stays `/login`. Signal calls are recorded by `spySignals` (an init script that wraps, or defines, the two methods and reports through an exposed binding): the two-device `/account` load lists exactly `/api/account`'s ids, the removal lists them without device B, and device B's refused sign-in signals its id unknown, all with `rpId` `localhost`.
- Context D has the native `parse*OptionsFromJSON` and `toJSON` removed by an init script, so the hand-rolled conversions in `webauthn.js` run (the path older Safari and Firefox take; Chromium otherwise always uses the native ones). One login-link is opened in two tabs; tab 1 enrols `device D <nonce>` and, after `clearCookies()`, signs in again at `/login`; tab 2 then finishes with the spent link → `#status` 這個註冊連結無效或已過期 and a `signalUnknownCredential` for its orphaned credential. Device A signs in → `/account`: device B's row is gone, device A's remains, and each 刪除 is disabled exactly when one row is left.

Evidence: `steps.log` (cookies and tokens redacted), `cli.log`, `01-owner-private-page.png` … `08-account-after-removal.png`, `result.json`.

## Gotchas

- Node `fetch` CAN carry the owner session: the server only reads the `cookie` header, so an explicit `cookie: __Host-artifacts_session=<value>` header works (the controls prove it). What Node cannot do is receive the cookie in the first place (Playwright's context holds it; read it from `context.cookies()`).
- The virtual authenticator belongs to the page's CDP session, not the context: sign in again on the same page that enrolled. `clearCookies()` keeps the authenticator.
- The `passkeys` and `sessions` tables are shared by every journey on the data dir (owner-signin-and-mcp, oauth-refresh and phone-viewer-bar enrol too, and nothing removes those). So passkey rows are found by their nonce name and the count is relative; "刪除 disabled on the last passkey" is asserted as `disabled === (rows <= 1)`, and is the literal one-row state only when this journey runs alone on a fresh data dir.
- Do not add a second passkey with `#add-passkey` on the same authenticator: `excludeCredentials` refuses it (這台裝置已經有這裡的 passkey 了). Use a second browser context as the second device.
- `#logout-all` and passkey removal end every session on the instance. Each journey enrols its own owner, so running this one between others is safe; never point it at an instance someone is using.
- Register the `dialog` handler before clicking `#logout-all` or 刪除 (both use `confirm`).

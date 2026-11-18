# OAuth refresh and revocation

MCP clients keep their sign-in with rotating refresh tokens. Each refresh token works once and comes back with a new access and refresh token. A refresh token that has already been used means someone holds a copy, so presenting it again revokes the whole sign-in (the token family). The owner can also revoke a client on the account page, and a client can revoke its own sign-in with RFC 7009.

## Sub-features

- `refresh`: `POST /oauth/token` form-encoded `grant_type=refresh_token&refresh_token=…&client_id=…&resource=<public app>/mcp` → 200, `Cache-Control: no-store`, new `access_token` and `refresh_token`.
- `replay`: the old refresh token again → 400 `invalid_grant`; every token of that family is deleted (both access tokens 401, the new refresh token `invalid_grant`). `/account` stops listing the client (it lists clients with a live unused token).
- `撤銷`: `/account` `#clients li` → 撤銷 (behind `confirm`) → `DELETE /api/account/clients/:id`; its refresh token → `invalid_grant`, its access token → 401.
- `revoke`: `POST /oauth/revoke` `token=<refresh or access>` → 200 and the family ends.

## How to get to it (user POV)

- Claude Code or Codex refresh in the background when the 1 h access token expires; the owner never sees it.
- Owner: 帳號 → 已授權的應用程式 → 撤銷 for a client he no longer uses.

## Driving it with MCP over HTTP and Playwright Chromium (/account)

Preconditions: instance ready.

```bash
node .claude/skills/verify/scripts/journeys/oauth-refresh.mjs
```

- Enrols an owner, `authorizeClient(page, F, 'refresh X <nonce>')` → A1/R1. Every call passes its token explicitly to `rpc('tools/list', {}, token)`: `authorizeClient` overwrites `mcp-token.json` each time, and `rpc()` without a token would silently use the newest client's.
- `tools/list` with A1 → 200. Refresh R1 → 200 `no-store`, A2 ≠ A1, R2 ≠ R1; A2 → 200. (A1 after a normal refresh is logged, not asserted: access tokens live out their hour.) `/account` lists X.
- Replay R1 → 400 `invalid_grant`; then A2 and A1 → 401, refresh R2 → 400 `invalid_grant`; `/account` after a reload has no row for X.
- Client Y → A3/R3, A3 → 200; 撤銷 on `/account` (dialog accepted) → `#status` 已撤銷, row detached; refresh R3 → 400 `invalid_grant`, A3 → 401.
- Client Z: `POST /oauth/revoke` with its refresh token → 200; its access token → 401.

Evidence: `steps.log` (tokens redacted), `oauth-calls.json` (each call's status and error, no tokens), `01-account-client-x.png` … `03-account-client-y-revoked.png`, `result.json`.

## Gotchas

- There is no grace period for a second use of a refresh token: a client that retries a refresh whose answer it lost loses the sign-in. That is the design (src/oauth.ts), not a finding.
- Tokens are registered with `secret()` as soon as they arrive, so they never reach evidence.
- The client list on `/account` only shows clients with a live, unused token, so a revoked family disappears from it even though the client registration remains.

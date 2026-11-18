#!/usr/bin/env node
// Journey: OAuth refresh-token rotation for the remote MCP. A refresh works once; replaying the rotated-out
// refresh token revokes the whole sign-in (both access tokens and the new refresh token). Revoking a client on
// /account kills its refresh token; POST /oauth/revoke ends a sign-in too.
// Usage: node .claude/skills/verify/scripts/journeys/oauth-refresh.mjs
import { V, checks, enrolOwner, launch, log, save, secret, shot } from '../browser.mjs'
import { rpc } from '../mcp-call.mjs'
import { MCP_URL, authorizeClient } from '../mcp-token.mjs'

const F = 'oauth-refresh'
const c = checks(F)
const nonce = Date.now().toString(36)
const calls = []

/** POST /oauth/token grant_type=refresh_token, as an MCP client does. Tokens in the reply are registered as secrets. */
async function refresh(refreshToken, clientId, label) {
  const response = await fetch(`${V.publicApp}/oauth/token`, {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ grant_type: 'refresh_token', refresh_token: refreshToken, client_id: clientId, resource: MCP_URL })
  })
  const body = await response.json().catch(() => null)
  secret(body?.access_token)
  secret(body?.refresh_token)
  const out = { status: response.status, cacheControl: response.headers.get('cache-control'), body }
  calls.push({ call: `refresh (${label})`, status: out.status, cacheControl: out.cacheControl, error: body?.error ?? null })
  log(F, `POST /oauth/token refresh_token (${label}) -> ${out.status} ${body?.error ?? 'tokens'}`)
  return out
}
/** tools/list with an explicit access token (never the saved mcp-token.json). */
async function list(token, label) {
  const out = await rpc('tools/list', {}, token)
  calls.push({ call: `tools/list (${label})`, status: out.status })
  return out.status
}

let browser
try {
  browser = await launch()
  const ctx = await browser.newContext({ viewport: { width: 1200, height: 800 } })
  const page = await ctx.newPage()
  await enrolOwner(page, F, `verify device ${nonce}`)

  // Client X: one refresh, then a replay of the old refresh token.
  const nameX = `refresh X ${nonce}`
  const x = await authorizeClient(page, F, nameX)
  const A1 = x.tokens.access_token
  const R1 = x.tokens.refresh_token
  c.ok((await list(A1, 'A1')) === 200, 'client X: tools/list with access token A1 -> 200')
  const first = await refresh(R1, x.clientId, 'R1, first use')
  const A2 = first.body?.access_token
  const R2 = first.body?.refresh_token
  c.ok(first.status === 200 && first.cacheControl === 'no-store', 'refresh with R1 -> 200, Cache-Control: no-store', `${first.status} ${first.cacheControl}`)
  c.ok(A2 && R2 && A2 !== A1 && R2 !== R1, 'the refresh returns a new access token A2 and a new refresh token R2')
  c.ok((await list(A2, 'A2')) === 200, 'tools/list with A2 -> 200')
  log(F, `observation: A1 after a normal refresh -> ${await list(A1, 'A1 after refresh')} (access tokens live out their hour)`)
  await page.goto(`${V.publicApp}/account`)
  await page.locator('#clients li', { hasText: nameX }).waitFor({ timeout: 5000 })
  c.ok(true, '/account lists client X while its sign-in is live')
  await shot(page, F, '01-account-client-x')

  const replay = await refresh(R1, x.clientId, 'R1 replayed')
  c.ok(replay.status === 400 && replay.body?.error === 'invalid_grant', 'replaying R1 -> 400 invalid_grant', `${replay.status} ${replay.body?.error}`)
  const a2After = await list(A2, 'A2 after replay')
  const a1After = await list(A1, 'A1 after replay')
  c.ok(a2After === 401 && a1After === 401, 'after the replay both access tokens get 401 (the family is revoked)', `A2 ${a2After} A1 ${a1After}`)
  const r2After = await refresh(R2, x.clientId, 'R2 after replay')
  c.ok(r2After.status === 400 && r2After.body?.error === 'invalid_grant', 'after the replay the new refresh token R2 -> 400 invalid_grant', `${r2After.status} ${r2After.body?.error}`)
  await page.reload()
  await page.locator('#clients li').first().waitFor({ timeout: 5000 })
  const xRows = await page.locator('#clients li', { hasText: nameX }).count()
  c.ok(xRows === 0, '/account no longer lists client X (no live grant left)', xRows)
  await shot(page, F, '02-account-after-replay')

  // Client Y: revoked from the account page.
  const nameY = `refresh Y ${nonce}`
  const y = await authorizeClient(page, F, nameY)
  const A3 = y.tokens.access_token
  const R3 = y.tokens.refresh_token
  c.ok((await list(A3, 'A3')) === 200, 'client Y: tools/list with A3 -> 200')
  await page.goto(`${V.publicApp}/account`)
  const rowY = page.locator('#clients li', { hasText: nameY })
  await rowY.waitFor({ timeout: 5000 })
  page.once('dialog', (dialog) => dialog.accept())
  await rowY.locator('button', { hasText: '撤銷' }).click()
  await page.locator('#status', { hasText: '已撤銷' }).waitFor({ timeout: 5000 })
  await rowY.waitFor({ state: 'detached', timeout: 5000 })
  c.ok(true, '撤銷 on /account: #status 已撤銷 and the row goes')
  await shot(page, F, '03-account-client-y-revoked')
  const r3After = await refresh(R3, y.clientId, 'R3 after 撤銷')
  c.ok(r3After.status === 400 && r3After.body?.error === 'invalid_grant', "after 撤銷, client Y's refresh token -> 400 invalid_grant", `${r3After.status} ${r3After.body?.error}`)
  c.ok((await list(A3, 'A3 after 撤銷')) === 401, "after 撤銷, client Y's access token -> 401")

  // Client Z: the client itself revokes with RFC 7009.
  const z = await authorizeClient(page, F, `refresh Z ${nonce}`)
  c.ok((await list(z.tokens.access_token, 'A4')) === 200, 'client Z: tools/list -> 200')
  const revoke = await fetch(`${V.publicApp}/oauth/revoke`, { method: 'POST', headers: { 'content-type': 'application/x-www-form-urlencoded' }, body: new URLSearchParams({ token: z.tokens.refresh_token, client_id: z.clientId }) })
  c.ok(revoke.status === 200, 'POST /oauth/revoke with the refresh token -> 200', revoke.status)
  c.ok((await list(z.tokens.access_token, 'A4 after revoke')) === 401, "after /oauth/revoke, client Z's access token -> 401")
  save(F, 'oauth-calls.json', calls)
  await ctx.close()
  c.finish()
} catch (error) {
  save(F, 'oauth-calls.json', calls)
  c.finish(error)
} finally {
  await browser?.close()
}

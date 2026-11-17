import assert from 'node:assert/strict'
import { createHash, randomBytes } from 'node:crypto'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { after, before, describe, test } from 'node:test'
import { ArtifactStore } from '../src/artifact-store.ts'
import type { Config } from '../src/config.ts'
import { validRedirectUri } from '../src/oauth.ts'
import { OwnerAuth, SESSION_COOKIE } from '../src/owner-auth.ts'
import { createPublicAppServer } from '../src/public-server.ts'

const APP = 'https://pub-app.test'
const MCP = `${APP}/mcp`
const REDIRECT = 'http://127.0.0.1:33418/callback'
let dir: string
let store: ArtifactStore
let auth: OwnerAuth
let app: ReturnType<typeof createPublicAppServer>
let cookie: string

before(() => {
  dir = mkdtempSync(join(tmpdir(), 'artifacts-mcp-'))
  store = new ArtifactStore(dir)
  const config: Config = {
    dataDir: dir,
    appPort: 0,
    contentPort: 0,
    appOrigin: 'http://app.test',
    contentOrigin: 'http://content.test',
    tokens: ['test-token-0123456789'],
    public: { appPort: 0, contentPort: 0, appOrigin: APP, contentOrigin: 'https://pub-content.test' }
  }
  auth = new OwnerAuth(dir, APP)
  app = createPublicAppServer(store, config, auth)
  cookie = `${SESSION_COOKIE}=${auth.createSession()}`
})

after(() => {
  store.close()
  auth.close()
  rmSync(dir, { recursive: true, force: true })
})

const form = (fields: Record<string, string>) => ({ method: 'POST', headers: { 'content-type': 'application/x-www-form-urlencoded' }, body: new URLSearchParams(fields).toString() })

async function register(redirectUris = [REDIRECT]) {
  const response = await app.request(`${APP}/oauth/register`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ client_name: 'Test Agent', redirect_uris: redirectUris }) })
  assert.equal(response.status, 201)
  return ((await response.json()) as { client_id: string }).client_id
}

function pkce() {
  const verifier = randomBytes(32).toString('base64url')
  return { verifier, challenge: createHash('sha256').update(verifier).digest('base64url') }
}

/** Through the consent page as the signed-in owner; returns the redirect the client receives. */
async function authorize(clientId: string, challenge: string, extra: Record<string, string> = {}, allow = true) {
  const query = new URLSearchParams({ response_type: 'code', client_id: clientId, redirect_uri: REDIRECT, code_challenge: challenge, code_challenge_method: 'S256', state: 'st', resource: MCP, scope: 'artifacts', ...extra })
  const page = await app.request(`${APP}/oauth/authorize?${query}`, { headers: { cookie } })
  assert.equal(page.status, 200, await page.clone().text())
  const html = await page.text()
  assert.match(html, /Test Agent/)
  assert.match(html, /127\.0\.0\.1:33418/)
  const request = /data-request="([^"]+)"/.exec(html)![1]
  const decision = await app.request(`${APP}/oauth/authorize/decision`, { method: 'POST', headers: { cookie, 'content-type': 'application/json', 'x-artifact-client': 'viewer' }, body: JSON.stringify({ request, allow }) })
  assert.equal(decision.status, 200)
  return new URL(((await decision.json()) as { redirect: string }).redirect)
}

async function signIn() {
  const clientId = await register()
  const { verifier, challenge } = pkce()
  const back = await authorize(clientId, challenge)
  const tokens = await app.request(`${APP}/oauth/token`, form({ grant_type: 'authorization_code', code: back.searchParams.get('code')!, redirect_uri: REDIRECT, client_id: clientId, code_verifier: verifier, resource: MCP }))
  assert.equal(tokens.status, 200)
  return { clientId, ...((await tokens.json()) as { access_token: string; refresh_token: string; expires_in: number }) }
}

let rpcId = 0
async function rpc(token: string | null, method: string, params: Record<string, unknown> = {}) {
  const response = await app.request(MCP, {
    method: 'POST',
    headers: { 'content-type': 'application/json', accept: 'application/json, text/event-stream', 'mcp-protocol-version': '2025-11-25', ...(token ? { authorization: `Bearer ${token}` } : {}) },
    body: JSON.stringify({ jsonrpc: '2.0', id: ++rpcId, method, params })
  })
  return { response, body: response.status === 200 ? ((await response.json()) as { result?: any; error?: any }) : null }
}

async function call(token: string, name: string, args: Record<string, unknown>) {
  const { body } = await rpc(token, 'tools/call', { name, arguments: args })
  const result = body!.result as { isError?: boolean; content: { text: string }[] }
  return { isError: result.isError === true, value: JSON.parse(result.content[0].text) }
}

describe('discovery', () => {
  test('the MCP answers 401 with where to sign in, and both metadata documents agree', async () => {
    const { response } = await rpc(null, 'tools/list')
    assert.equal(response.status, 401)
    assert.equal(response.headers.get('www-authenticate'), `Bearer resource_metadata="${APP}/.well-known/oauth-protected-resource/mcp", scope="artifacts"`)
    const resource = (await (await app.request(`${APP}/.well-known/oauth-protected-resource/mcp`)).json()) as { resource: string; authorization_servers: string[] }
    assert.deepEqual([resource.resource, resource.authorization_servers], [MCP, [APP]])
    const server = (await (await app.request(`${APP}/.well-known/oauth-authorization-server`)).json()) as Record<string, unknown>
    assert.equal(server.issuer, APP)
    assert.deepEqual(server.code_challenge_methods_supported, ['S256'])
    assert.equal(server.authorization_response_iss_parameter_supported, true)
    assert.equal(server.registration_endpoint, `${APP}/oauth/register`)
  })
})

describe('authorization', () => {
  test('signed out, the authorize URL goes through the passkey sign-in and back', async () => {
    const clientId = await register()
    const query = `response_type=code&client_id=${clientId}&redirect_uri=${encodeURIComponent(REDIRECT)}&code_challenge=${pkce().challenge}&code_challenge_method=S256`
    const response = await app.request(`${APP}/oauth/authorize?${query}`)
    assert.equal(response.status, 302)
    assert.equal(response.headers.get('location'), `/login?next=${encodeURIComponent(`/oauth/authorize?${query}`)}`)
  })

  test('the code comes back with state and iss; it works once, only with its verifier, client and redirect', async () => {
    const clientId = await register()
    const { verifier, challenge } = pkce()
    const back = await authorize(clientId, challenge)
    assert.equal(back.origin + back.pathname, REDIRECT)
    assert.equal(back.searchParams.get('state'), 'st')
    assert.equal(back.searchParams.get('iss'), APP)
    const code = back.searchParams.get('code')!
    const exchange = (fields: Record<string, string>) => app.request(`${APP}/oauth/token`, form({ grant_type: 'authorization_code', code, redirect_uri: REDIRECT, client_id: clientId, code_verifier: verifier, ...fields }))
    assert.equal((await exchange({ code_verifier: pkce().verifier })).status, 400)
    // The failed try spent the code: nobody may keep guessing verifiers.
    assert.equal((await exchange({})).status, 400)

    const second = await authorize(clientId, challenge)
    const good = await app.request(`${APP}/oauth/token`, form({ grant_type: 'authorization_code', code: second.searchParams.get('code')!, redirect_uri: REDIRECT, client_id: clientId, code_verifier: verifier }))
    assert.equal(good.status, 200)
    assert.equal(good.headers.get('cache-control'), 'no-store')
    const again = await app.request(`${APP}/oauth/token`, form({ grant_type: 'authorization_code', code: second.searchParams.get('code')!, redirect_uri: REDIRECT, client_id: clientId, code_verifier: verifier }))
    assert.equal(again.status, 400)
  })

  test('a loopback redirect may change port, nothing else; unregistered redirects get no redirect at all', async () => {
    const clientId = await register()
    const ok = await app.request(`${APP}/oauth/authorize?response_type=code&client_id=${clientId}&redirect_uri=${encodeURIComponent('http://127.0.0.1:5555/callback')}&code_challenge=${pkce().challenge}&code_challenge_method=S256`, { headers: { cookie } })
    assert.equal(ok.status, 200)
    for (const uri of ['http://127.0.0.1:5555/other', 'https://evil.test/callback', 'http://localhost:33418/callback']) {
      const bad = await app.request(`${APP}/oauth/authorize?response_type=code&client_id=${clientId}&redirect_uri=${encodeURIComponent(uri)}&code_challenge=${pkce().challenge}&code_challenge_method=S256`, { headers: { cookie } })
      assert.equal(bad.status, 400, uri)
      assert.equal(bad.headers.get('location'), null)
    }
  })

  test('CORS covers only the OAuth endpoints and metadata; /mcp answers its own preflight; the owner app stays closed', async () => {
    const preflight = await app.request(MCP, { method: 'OPTIONS', headers: { origin: 'https://inspector.test', 'access-control-request-method': 'POST', 'access-control-request-headers': 'mcp-method' } })
    assert.match(preflight.headers.get('access-control-allow-headers') ?? '', /mcp-method/)
    const account = await app.request(`${APP}/api/account`, { method: 'OPTIONS', headers: { origin: 'https://evil.test', 'access-control-request-method': 'GET' } })
    assert.equal(account.headers.get('access-control-allow-origin'), null)
    assert.notEqual(account.status, 204)
    const page = await app.request(`${APP}/api/account`, { headers: { cookie, origin: 'https://evil.test' } })
    assert.equal(page.headers.get('access-control-allow-origin'), null)
    assert.equal((await app.request(`${APP}/.well-known/oauth-authorization-server`)).headers.get('access-control-allow-origin'), '*')
  })

  test('a client registered long ago and never used again can still sign in', async () => {
    const clientId = await register()
    const old = new Date(Date.now() - 40 * 86400_000).toISOString()
    auth.db.prepare('UPDATE oauth_clients SET created_at = ?, last_used_at = ? WHERE id = ?').run(old, old, clientId)
    await register(['https://other.test/cb'])
    const { verifier, challenge } = pkce()
    const back = await authorize(clientId, challenge)
    const tokens = await app.request(`${APP}/oauth/token`, form({ grant_type: 'authorization_code', code: back.searchParams.get('code')!, redirect_uri: REDIRECT, client_id: clientId, code_verifier: verifier }))
    assert.equal(tokens.status, 200)
  })

  test('a callback that is not the registered one, or not a URL, is a 400 page, never a redirect or a 500', async () => {
    const clientId = await register([REDIRECT, 'https://app.test/cb'])
    for (const uri of ['https://127.0.0.1:33418/callback', 'notaurl', 'http://127.0.0.1:33418/callback#x']) {
      const response = await app.request(`${APP}/oauth/authorize?response_type=code&client_id=${clientId}&redirect_uri=${encodeURIComponent(uri)}&code_challenge=${pkce().challenge}&code_challenge_method=S256`, { headers: { cookie } })
      assert.equal(response.status, 400, uri)
    }
  })

  test('no PKCE, another resource or another scope is refused on this site, not by a redirect to the client', async () => {
    const clientId = await register()
    const base = { response_type: 'code', client_id: clientId, redirect_uri: REDIRECT, code_challenge_method: 'S256', state: 'x' }
    for (const [params, error] of [
      [{ ...base }, 'invalid_request'],
      [{ ...base, code_challenge: pkce().challenge, resource: 'https://other.test/mcp' }, 'invalid_target'],
      [{ ...base, code_challenge: pkce().challenge, scope: 'admin' }, 'invalid_scope'],
      [{ ...base, code_challenge: pkce().challenge, code_challenge_method: 'plain' }, 'invalid_request']
    ] as const) {
      const response = await app.request(`${APP}/oauth/authorize?${new URLSearchParams(params as Record<string, string>)}`, { headers: { cookie } })
      assert.equal(response.status, 400)
      assert.equal(response.headers.get('location'), null)
      assert.match(await response.text(), new RegExp(error))
    }
  })

  test('declining tells the client access_denied; the decision needs the session and the viewer header', async () => {
    const clientId = await register()
    const back = await authorize(clientId, pkce().challenge, {}, false)
    assert.equal(back.searchParams.get('error'), 'access_denied')
    assert.equal(back.searchParams.get('code'), null)
    const attempts: Record<string, string>[] = [{ 'content-type': 'application/json' }, { 'content-type': 'application/json', 'x-artifact-client': 'viewer' }, { cookie, 'content-type': 'application/json' }]
    for (const headers of attempts) {
      const response = await app.request(`${APP}/oauth/authorize/decision`, { method: 'POST', headers, body: JSON.stringify({ request: 'x', allow: true }) })
      assert.equal(response.status, 403)
    }
  })

  test('registration takes only public clients with safe redirects', async () => {
    assert.ok(validRedirectUri('https://claude.ai/api/mcp/auth_callback'))
    assert.ok(validRedirectUri('http://localhost:8080/cb'))
    assert.ok(validRedirectUri('cursor.app:/callback'))
    for (const bad of ['http://evil.test/cb', 'javascript:alert(1)', 'https://x.test/cb#frag', 'data:text/html,x', 'notaurl']) assert.ok(!validRedirectUri(bad), bad)
    const secret = await app.request(`${APP}/oauth/register`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ redirect_uris: [REDIRECT], token_endpoint_auth_method: 'client_secret_basic' }) })
    assert.equal(secret.status, 400)
  })
})

describe('tokens', () => {
  test('refresh tokens rotate; using an old one again ends the whole sign-in', async () => {
    const first = await signIn()
    const refresh = (token: string) => app.request(`${APP}/oauth/token`, form({ grant_type: 'refresh_token', refresh_token: token, client_id: first.clientId }))
    const next = await refresh(first.refresh_token)
    assert.equal(next.status, 200)
    const second = (await next.json()) as { access_token: string; refresh_token: string }
    assert.notEqual(second.refresh_token, first.refresh_token)
    assert.equal((await rpc(second.access_token, 'tools/list')).response.status, 200)
    // Replay of the rotated-out token after the grace period: everything of this sign-in stops working.
    auth.db.prepare('UPDATE oauth_tokens SET used_at = used_at - 120000 WHERE used_at IS NOT NULL').run()
    assert.equal((await refresh(first.refresh_token)).status, 400)
    assert.equal((await rpc(second.access_token, 'tools/list')).response.status, 401)
    assert.equal((await refresh(second.refresh_token)).status, 400)
  })

  test('a refresh retried within seconds (a lost response, two tabs) still works; a late replay ends the sign-in', async () => {
    const first = await signIn()
    const refresh = (token: string) => app.request(`${APP}/oauth/token`, form({ grant_type: 'refresh_token', refresh_token: token, client_id: first.clientId }))
    const a = await refresh(first.refresh_token)
    const b = await refresh(first.refresh_token)
    assert.equal(a.status, 200)
    assert.equal(b.status, 200)
    const winner = (await a.json()) as { access_token: string }
    assert.equal((await rpc(winner.access_token, 'tools/list')).response.status, 200)
    // Past the grace period, the same old token is a replay.
    auth.db.prepare('UPDATE oauth_tokens SET used_at = used_at - 120000 WHERE used_at IS NOT NULL').run()
    assert.equal((await refresh(first.refresh_token)).status, 400)
    assert.equal((await rpc(winner.access_token, 'tools/list')).response.status, 401)
  })

  test('revoking, or revoking the client from the account page, stops its tokens', async () => {
    const one = await signIn()
    await app.request(`${APP}/oauth/revoke`, form({ token: one.refresh_token }))
    assert.equal((await rpc(one.access_token, 'tools/list')).response.status, 401)

    const two = await signIn()
    const account = (await (await app.request(`${APP}/api/account`, { headers: { cookie } })).json()) as { clients: { id: string; name: string }[] }
    assert.ok(account.clients.some((client) => client.id === two.clientId && client.name === 'Test Agent'))
    const revoke = await app.request(`${APP}/api/account/clients/${two.clientId}`, { method: 'DELETE', headers: { cookie, 'x-artifact-client': 'viewer' } })
    assert.equal(revoke.status, 204)
    const response = await rpc(two.access_token, 'tools/list')
    assert.equal(response.response.status, 401)
    assert.match(response.response.headers.get('www-authenticate')!, /error="invalid_token"/)
  })

  test('a session cookie is not a way into the MCP', async () => {
    const response = await app.request(MCP, { method: 'POST', headers: { cookie, 'content-type': 'application/json', accept: 'application/json, text/event-stream' }, body: JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'tools/list' }) })
    assert.equal(response.status, 401)
  })
})

describe('tools', () => {
  test('publish, read, update with baseVersion, conflict with changedFiles, visibility, comments, delete', async () => {
    const { access_token: token } = await signIn()
    const init = await rpc(token, 'initialize', { protocolVersion: '2025-11-25', capabilities: {}, clientInfo: { name: 'test', version: '1' } })
    assert.equal(init.body!.result.serverInfo.name, 'artifacts')
    const listed = await rpc(token, 'tools/list')
    assert.deepEqual(
      listed.body!.result.tools.map((tool: { name: string }) => tool.name).sort(),
      ['comments', 'delete', 'errors', 'get', 'list', 'publish', 'read', 'reply', 'set_visibility', 'update']
    )

    const published = await call(token, 'publish', {
      title: 'MCP 報告',
      files: [
        { path: 'index.md', content: '# 第一版\n\n![圖](pic.png)\n' },
        { path: 'pic.png', content: 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', encoding: 'base64' }
      ]
    })
    assert.equal(published.isError, false)
    const id = published.value.id as string
    assert.equal(published.value.visibility, 'private')
    assert.equal(published.value.url, `${APP}/a/${id}`)
    assert.equal(published.value.publicUrl, null)
    assert.equal(store.requireSummary(id).currentVersion, 1)

    const read = await call(token, 'read', { id })
    assert.equal(read.value.content, '# 第一版\n\n![圖](pic.png)\n')
    assert.equal((await call(token, 'read', { id, path: 'pic.png' })).value.encoding, 'base64')

    const noBase = await call(token, 'update', { id, files: [{ path: 'index.md', content: '# x' }] })
    assert.equal(noBase.isError, true)
    assert.equal(noBase.value.error.code, 'base_version_required')

    const v2 = await call(token, 'update', { id, baseVersion: 1, files: [{ path: 'index.md', content: '# 第二版\n' }, { path: 'pic.png', delete: true }] })
    assert.equal(v2.value.currentVersion, 2)
    const conflict = await call(token, 'update', { id, baseVersion: 1, files: [{ path: 'index.md', content: '# 舊草稿\n' }] })
    assert.equal(conflict.isError, true)
    assert.equal(conflict.value.error.code, 'version_conflict')
    assert.deepEqual([...conflict.value.error.details.changedFiles].sort(), ['index.md', 'pic.png'])
    assert.equal(store.requireSummary(id).currentVersion, 2)

    const shared = await call(token, 'set_visibility', { id, visibility: 'public' })
    assert.equal(shared.value.publicUrl, `${APP}/a/${id}`)

    store.addComment(id, { author: '讀者', body: '第二段有錯字' })
    const comments = await call(token, 'comments', { id, open_only: true })
    assert.equal(comments.value.comments.length, 1)
    await call(token, 'reply', { id, commentId: comments.value.comments[0].id, body: '已修正', resolve: true })
    assert.equal((await call(token, 'comments', { id, open_only: true })).value.comments.length, 0)

    assert.equal((await call(token, 'delete', { id })).value.deleted, id)
    assert.equal((await call(token, 'get', { id })).isError, true)
  })

  test('a file of several megabytes publishes (the request limit matches the API, not the SDK default)', async () => {
    const { access_token: token } = await signIn()
    const picture = Buffer.alloc(5 * 1024 * 1024, 7).toString('base64')
    const result = await call(token, 'publish', { title: '大附件', files: [{ path: 'index.md', content: '# x' }, { path: 'big.bin', content: picture, encoding: 'base64' }] })
    assert.equal(result.isError, false, JSON.stringify(result.value).slice(0, 200))
  })

  test('capabilities and public do not mix through the MCP either', async () => {
    const { access_token: token } = await signIn()
    const result = await call(token, 'publish', { title: 'db 頁', capabilities: ['db'], visibility: 'public', files: [{ path: 'index.html', content: '<p>x</p>' }] })
    assert.equal(result.isError, true)
  })
})

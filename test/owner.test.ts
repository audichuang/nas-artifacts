import assert from 'node:assert/strict'
import { createHash, generateKeyPairSync, sign } from 'node:crypto'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { after, before, describe, test } from 'node:test'
import { createAppServer } from '../src/app-server.ts'
import { ArtifactStore } from '../src/artifact-store.ts'
import type { Config } from '../src/config.ts'
import { createContentServer } from '../src/content-server.ts'
import { CONTENT_TOKEN_HOURS, OwnerAuth, SESSION_COOKIE } from '../src/owner-auth.ts'
import { createPublicAppServer, publicContentConfig, safeNext } from '../src/public-server.ts'

const TOKEN = 'test-token-0123456789'
const PUBLIC_APP = 'https://pub-app.test'
const PUBLIC_CONTENT = 'https://pub-content.test'
let dir: string
let store: ArtifactStore
let auth: OwnerAuth
let app: ReturnType<typeof createAppServer>
let publicApp: ReturnType<typeof createPublicAppServer>
let publicContent: ReturnType<typeof createContentServer>
let cookie: string

before(() => {
  dir = mkdtempSync(join(tmpdir(), 'artifacts-owner-'))
  store = new ArtifactStore(dir)
  const config: Config = {
    dataDir: dir,
    appPort: 0,
    contentPort: 0,
    appOrigin: 'http://app.test',
    contentOrigin: 'http://content.test',
    tokens: [TOKEN],
    public: { appPort: 0, contentPort: 0, appOrigin: PUBLIC_APP, contentOrigin: PUBLIC_CONTENT }
  }
  auth = new OwnerAuth(dir, PUBLIC_APP)
  app = createAppServer(store, config, { auth })
  publicApp = createPublicAppServer(store, config, auth)
  publicContent = createContentServer(store, publicContentConfig(config), { publicOnly: true, contentTokens: auth })
  cookie = `${SESSION_COOKIE}=${auth.createSession()}`
})

after(() => {
  store.close()
  auth.close()
  rmSync(dir, { recursive: true, force: true })
})

describe('content tokens', () => {
  test('open only the artifact they were issued for, until they expire, and resist tampering', () => {
    const { token, expiresAt } = auth.contentToken('abc')
    assert.ok(auth.verifyContentToken(token, 'abc'))
    assert.ok(!auth.verifyContentToken(token, 'abd'))
    assert.ok(expiresAt > Date.now() + (CONTENT_TOKEN_HOURS - 1) * 3600_000)
    assert.ok(!auth.verifyContentToken(token, 'abc', expiresAt))
    const [expires, signature] = token.split('.')
    // A later expiry with the old signature does not pass.
    assert.ok(!auth.verifyContentToken(`${(parseInt(expires, 36) + 60_000).toString(36)}.${signature}`, 'abc'))
    assert.ok(!auth.verifyContentToken(`${expires}.${signature.slice(0, -1)}A`, 'abc') || signature.endsWith('A'))
    assert.ok(!auth.verifyContentToken('', 'abc'))
  })

  test('an encoded path segment does not get an anonymous visitor past the current-version gate', async () => {
    const { id } = store.publish({ title: 'Old', entry: 'docs/index.html', files: { 'docs/index.html': { content: Buffer.from('<p>v1</p>') }, '1/secret.txt': { content: Buffer.from('OLD SECRET') } } })
    store.update(id, { files: { '1/secret.txt': null }, baseVersion: 1 })
    store.setVisibility(id, 'public')
    for (const path of [`/a/${id}/%76/1/secret.txt`, `/a/${id}/%76/1`, `/a/${id}/%76/1/`, `/%61/${id}/v/1/1/secret.txt`, `/a/${id}/v/%31/1/secret.txt`]) {
      const response = await publicContent.request(`${PUBLIC_CONTENT}${path}`)
      assert.equal(response.status, 404, `${path} → ${response.status} ${response.headers.get('location') ?? ''}`)
    }
    assert.equal((await publicContent.request(`${PUBLIC_CONTENT}/a/${id}/v/2/docs/index.html`)).status, 200)
  })

  test('a valid token opens any version of a private artifact; the public gate is unchanged without one', async () => {
    const { id } = store.publish({ title: 'Draft', files: { 'index.md': { content: Buffer.from('# v1\n') } } })
    store.update(id, { files: { 'index.md': { content: Buffer.from('# v2\n') } }, baseVersion: 1 })
    const { token } = auth.contentToken(id)
    assert.equal((await publicContent.request(`${PUBLIC_CONTENT}/t/${token}/a/${id}/v/1/`)).status, 200)
    assert.equal((await publicContent.request(`${PUBLIC_CONTENT}/t/${token}/a/${id}/v/2/`)).status, 200)
    assert.equal((await publicContent.request(`${PUBLIC_CONTENT}/a/${id}/v/2/`)).status, 404)
    // An encoded id is the same id to the gate and to the route.
    assert.equal((await publicContent.request(`${PUBLIC_CONTENT}/t/${token}/a/${encodeURIComponent(id)}/v/2/`)).status, 200)
    const other = store.publish({ title: 'Other', files: { 'index.md': { content: Buffer.from('# o\n') } } })
    assert.equal((await publicContent.request(`${PUBLIC_CONTENT}/t/${token}/a/${other.id}/v/1/`)).status, 404)
    // Nothing outside /a/<id>/ is reachable under a token.
    assert.equal((await publicContent.request(`${PUBLIC_CONTENT}/t/${token}/_/bridge.js`)).status, 404)
    const page = await publicContent.request(`${PUBLIC_CONTENT}/t/${token}/a/${id}/v/2/`)
    const csp = page.headers.get('content-security-policy')!
    assert.match(csp, new RegExp(`img-src ${PUBLIC_CONTENT}/t/${token.replace('.', '\\.')}/a/${id}/ `))
    assert.match(csp, new RegExp(`frame-ancestors ${PUBLIC_APP}`))
    assert.equal(page.headers.get('cache-control'), 'private, no-cache')
    assert.equal(page.headers.get('referrer-policy'), 'no-referrer')
  })

  test('a nested entry redirects within the token path', async () => {
    const { id } = store.publish({ title: 'Nested', entry: 'docs/a.html', files: { 'docs/a.html': { content: Buffer.from('<p>a</p>') } } })
    const { token } = auth.contentToken(id)
    const response = await publicContent.request(`${PUBLIC_CONTENT}/t/${token}/a/${id}/v/1/`)
    assert.equal(response.status, 302)
    assert.equal(response.headers.get('location'), `/t/${token}/a/${id}/v/1/docs/a.html`)
  })
})

describe('the public app gate', () => {
  test('the not-found page and the OAuth error page may load their own stylesheet and icon, and still no script', async () => {
    for (const path of ['/a/nothing-here', '/oauth/authorize?client_id=nope']) {
      const response = await publicApp.request(`${PUBLIC_APP}${path}`, { headers: { cookie } })
      if (path.startsWith('/a/')) {
        // Signed in, an unknown id goes through the owner app; signed out it is the sign-in 404 page.
        const anonymous = await publicApp.request(`${PUBLIC_APP}${path}`)
        const csp = anonymous.headers.get('content-security-policy')!
        assert.match(csp, /style-src 'self'/)
        assert.match(csp, /img-src 'self'/)
        assert.doesNotMatch(csp, /script-src/)
      } else {
        assert.equal(response.status, 400)
        assert.match(response.headers.get('content-security-policy')!, /style-src 'self'/)
      }
    }
  })

  test('signed out: no gallery, list or writes; signed in: the full app', async () => {
    assert.equal((await publicApp.request(`${PUBLIC_APP}/api/artifacts`)).status, 404)
    const gallery = await publicApp.request(`${PUBLIC_APP}/`)
    assert.equal(gallery.status, 404)
    assert.match(await gallery.text(), /href="\/login\?next=%2F"/)
    const list = await publicApp.request(`${PUBLIC_APP}/api/artifacts`, { headers: { cookie } })
    assert.equal(list.status, 200)
    assert.equal(list.headers.get('cache-control'), 'private, no-cache')
    // A forged or ended session is no session.
    assert.equal((await publicApp.request(`${PUBLIC_APP}/api/artifacts`, { headers: { cookie: `${SESSION_COOKIE}=nope` } })).status, 404)
  })

  test('the session never publishes, deletes or changes visibility, and the publish token does nothing here', async () => {
    const { id } = store.publish({ title: 'Mine', files: { 'index.md': { content: Buffer.from('# m\n') } } })
    const headers = { cookie, authorization: `Bearer ${TOKEN}`, 'content-type': 'application/json', 'x-artifact-client': 'viewer' }
    const body = JSON.stringify({ title: 'x', files: { 'index.md': { content: '# x' } } })
    assert.equal((await publicApp.request(`${PUBLIC_APP}/api/artifacts`, { method: 'POST', headers, body })).status, 401)
    assert.equal((await publicApp.request(`${PUBLIC_APP}/api/artifacts/${id}`, { method: 'DELETE', headers })).status, 401)
    assert.equal((await publicApp.request(`${PUBLIC_APP}/api/artifacts/${id}`, { method: 'PATCH', headers, body: JSON.stringify({ visibility: 'public' }) })).status, 401)
    assert.equal(store.requireSummary(id).visibility, 'private')
    // Without a session, the token is not even looked at.
    const { cookie: _, ...tokenOnly } = headers
    assert.equal((await publicApp.request(`${PUBLIC_APP}/api/artifacts/${id}`, { method: 'DELETE', headers: tokenOnly })).status, 404)
  })

  test('signed-in writes still need the viewer header', async () => {
    const { id } = store.publish({ title: 'Comment', files: { 'index.md': { content: Buffer.from('# c\n') } } })
    const post = (extra: Record<string, string>) =>
      publicApp.request(`${PUBLIC_APP}/api/artifacts/${id}/comments`, { method: 'POST', headers: { cookie, 'content-type': 'application/json', ...extra }, body: JSON.stringify({ author: 'me', body: 'hi' }) })
    assert.equal((await post({})).status, 401)
    assert.equal((await post({ 'x-artifact-client': 'viewer' })).status, 201)
  })

  test('the viewer frames the token path and links assets under it', async () => {
    const { id } = store.publish({ title: 'Viewer', files: { 'index.md': { content: Buffer.from('# v\n') } } })
    const viewer = await publicApp.request(`${PUBLIC_APP}/a/${id}`, { headers: { cookie } })
    assert.equal(viewer.status, 200)
    assert.equal(viewer.headers.get('content-security-policy'), `frame-src ${PUBLIC_CONTENT}/t/`)
    const html = await viewer.text()
    assert.match(html, new RegExp(`"contentBase":"${PUBLIC_CONTENT}/t/[0-9a-z]+\\.[A-Za-z0-9_-]+"`))
    assert.doesNotMatch(html, /"officeEnabled":true/)
    // The API hands out no tokens: one stolen session must not harvest a token for every artifact.
    const detail = (await (await publicApp.request(`${PUBLIC_APP}/api/artifacts/${id}`, { headers: { cookie } })).json()) as { contentUrl: string | null }
    assert.equal(detail.contentUrl, null)
    const list = await (await publicApp.request(`${PUBLIC_APP}/api/artifacts`, { headers: { cookie } })).text()
    assert.doesNotMatch(list, /\/t\//)
  })

  test('sign-in calls need the viewer header; enrolment needs a live link', async () => {
    const call = (path: string, headers: Record<string, string> = {}, body = {}) =>
      publicApp.request(`${PUBLIC_APP}${path}`, { method: 'POST', headers: { 'content-type': 'application/json', ...headers }, body: JSON.stringify(body) })
    assert.equal((await call('/auth/login/options')).status, 403)
    assert.equal((await call('/auth/login/options', { 'x-artifact-client': 'viewer' })).status, 200)
    assert.equal((await call('/auth/register/options', { 'x-artifact-client': 'viewer' })).status, 403)
    const { code } = auth.createLoginLink()
    const options = await call('/auth/register/options', { 'x-artifact-client': 'viewer' }, { code })
    assert.equal(options.status, 200)
    const { options: creation } = (await options.json()) as { options: { rp: { id: string }; authenticatorSelection: { userVerification: string } } }
    assert.equal(creation.rp.id, 'pub-app.test')
    assert.equal(creation.authenticatorSelection.userVerification, 'required')
    // A made-up ceremony is refused, not crashed.
    const verify = await call('/auth/login/verify', { 'x-artifact-client': 'viewer' }, { flow: 'nope', response: {} })
    assert.equal(verify.status, 400)
    assert.equal(verify.headers.get('set-cookie'), null)
  })

  test('logout ends the session', async () => {
    const raw = auth.createSession()
    const own = `${SESSION_COOKIE}=${raw}`
    const out = await publicApp.request(`${PUBLIC_APP}/auth/logout`, { method: 'POST', headers: { cookie: own, 'x-artifact-client': 'viewer' } })
    assert.match(out.headers.get('set-cookie')!, /Max-Age=0/)
    assert.equal((await publicApp.request(`${PUBLIC_APP}/api/artifacts`, { headers: { cookie: own } })).status, 404)
  })
})

describe('enrolment while signed in', () => {
  test('a link opened on a signed-in device is still spent by the passkey it enrols', async () => {
    const { code } = auth.createLoginLink()
    const options = await publicApp.request(`${PUBLIC_APP}/auth/register/options`, { method: 'POST', headers: { cookie, 'content-type': 'application/json', 'x-artifact-client': 'viewer' }, body: JSON.stringify({ code }) })
    assert.equal(options.status, 200)
    // The ceremony is tied to the link: it carries the link, so finishing it would spend it.
    const { flow } = (await options.json()) as { flow: string }
    assert.equal((auth as unknown as { challenges: Map<string, { linkHash?: string }> }).challenges.get(flow)?.linkHash !== undefined, true)
  })
})

/** A software passkey: a P-256 key, its COSE public key, and assertions signed the way an authenticator does. */
function softwarePasskey(rpID: string, origin: string) {
  const { privateKey, publicKey } = generateKeyPairSync('ec', { namedCurve: 'P-256' })
  const jwk = publicKey.export({ format: 'jwk' }) as { x: string; y: string }
  const cose = Buffer.concat([Buffer.from([0xa5, 0x01, 0x02, 0x03, 0x26, 0x20, 0x01, 0x21, 0x58, 0x20]), Buffer.from(jwk.x, 'base64url'), Buffer.from([0x22, 0x58, 0x20]), Buffer.from(jwk.y, 'base64url')])
  const id = Buffer.from('software-passkey').toString('base64url')
  return {
    id,
    cose,
    assert(challenge: string, counter: number) {
      const clientDataJSON = Buffer.from(JSON.stringify({ type: 'webauthn.get', challenge, origin }))
      const authenticatorData = Buffer.concat([createHash('sha256').update(rpID).digest(), Buffer.from([0x05]), Buffer.from([0, 0, 0, counter])])
      const signature = sign('sha256', Buffer.concat([authenticatorData, createHash('sha256').update(clientDataJSON).digest()]), privateKey)
      return {
        id,
        rawId: id,
        type: 'public-key',
        clientExtensionResults: {},
        response: { clientDataJSON: clientDataJSON.toString('base64url'), authenticatorData: authenticatorData.toString('base64url'), signature: signature.toString('base64url') }
      }
    }
  }
}

describe('passkey sign-in', () => {
  test('signs in with a registered passkey, but not with one removed while it was being verified', async () => {
    const key = softwarePasskey('pub-app.test', PUBLIC_APP)
    const insert = () => auth.db.prepare('INSERT INTO passkeys (id, public_key, counter, name, created_at) VALUES (?, ?, 0, ?, ?)').run(key.id, key.cose, 'soft', '2026-01-01')
    insert()
    const first = await auth.loginOptions()
    await auth.login(first.flow, key.assert(first.options.challenge, 1))
    // Removed mid-verification: the verification is async, so the removal lands before it finishes.
    const second = await auth.loginOptions()
    const pending = auth.login(second.flow, key.assert(second.options.challenge, 2))
    await new Promise((resolve) => setImmediate(resolve))
    auth.removePasskey(key.id)
    await assert.rejects(pending, /已被移除/)
    cookie = `${SESSION_COOKIE}=${auth.createSession()}`
  })
})

describe('ending sessions', () => {
  test('logging out everywhere or removing a passkey also ends every content token', async () => {
    const { id } = store.publish({ title: 'Revoke', files: { 'index.md': { content: Buffer.from('# r\n') } } })
    const open = (token: string) => publicContent.request(`${PUBLIC_CONTENT}/t/${token}/a/${id}/v/1/`)
    const before = auth.contentToken(id).token
    assert.equal((await open(before)).status, 200)
    auth.endAllSessions()
    assert.equal((await open(before)).status, 404)
    cookie = `${SESSION_COOKIE}=${auth.createSession()}`
    const second = auth.contentToken(id).token
    auth.db.prepare("INSERT INTO passkeys (id, public_key, counter, name, created_at) VALUES ('k1', x'00', 0, 'old', '2026-01-01')").run()
    auth.removePasskey('k1')
    assert.equal((await open(second)).status, 404)
    assert.equal((await open(auth.contentToken(id).token)).status, 200)
  })
})

describe('enrolment links', () => {
  test('only the publisher on the tailnet issues them, as a public URL with the code in the fragment', async () => {
    assert.equal((await app.request('http://app.test/api/login-links', { method: 'POST' })).status, 401)
    const response = await app.request('http://app.test/api/login-links', { method: 'POST', headers: { authorization: `Bearer ${TOKEN}` } })
    assert.equal(response.status, 201)
    const { url } = (await response.json()) as { url: string }
    assert.match(url, /^https:\/\/pub-app\.test\/enroll#[A-Za-z0-9_-]{43}$/)
    assert.ok(auth.isLoginLink(url.split('#')[1]))
  })
})

describe('sign-in redirects', () => {
  test('go only to a path on this app', () => {
    assert.equal(safeNext('/a/x?v=2'), '/a/x?v=2')
    assert.equal(safeNext('//evil.test/'), '/')
    assert.equal(safeNext('/\\evil.test'), '/')
    assert.equal(safeNext('https://evil.test/'), '/')
    assert.equal(safeNext(undefined), '/')
    // Browsers drop tabs and newlines while parsing a URL: "/\t/evil" would become "//evil".
    for (const sneaky of ['/\t/evil.test', '/\n/evil.test', '/\r\\evil.test', '/\t\\evil.test']) assert.equal(safeNext(sneaky), '/', JSON.stringify(sneaky))
  })
})

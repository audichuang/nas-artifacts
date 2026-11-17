import assert from 'node:assert/strict'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { after, before, describe, test } from 'node:test'
import { createAppServer } from '../src/app-server.ts'
import { ArtifactStore } from '../src/artifact-store.ts'
import type { Config } from '../src/config.ts'
import { createContentServer } from '../src/content-server.ts'
import { createPublicAppServer, publicContentConfig } from '../src/public-server.ts'

const TOKEN = 'test-token-0123456789'
const PUBLIC_APP = 'https://pub-app.test'
const PUBLIC_CONTENT = 'https://pub-content.test'
let dir: string
let store: ArtifactStore
let app: ReturnType<typeof createAppServer>
let publicApp: ReturnType<typeof createPublicAppServer>
let publicContent: ReturnType<typeof createContentServer>

const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64')

before(() => {
  dir = mkdtempSync(join(tmpdir(), 'artifacts-public-'))
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
  app = createAppServer(store, config)
  publicApp = createPublicAppServer(store, config)
  publicContent = createContentServer(store, publicContentConfig(config), { publicOnly: true })
})

after(() => {
  store.close()
  rmSync(dir, { recursive: true, force: true })
})

const setVisibility = (id: string, visibility: string, token: string | null = TOKEN) =>
  app.request(`http://app.test/api/artifacts/${id}`, {
    method: 'PATCH',
    headers: { 'content-type': 'application/json', ...(token ? { authorization: `Bearer ${token}` } : {}) },
    body: JSON.stringify({ visibility })
  })

const page = (id: string, path = '') => publicContent.request(`${PUBLIC_CONTENT}/a/${id}/v/1/${path}`)

describe('visibility', () => {
  test('artifacts are private until the publisher makes them public, and only the publisher can', async () => {
    const { id, visibility } = store.publish({ title: 'Report', files: { 'index.md': { content: Buffer.from('# 報告\n\n![圖](pic.png)\n') }, 'pic.png': { content: PNG } } })
    assert.equal(visibility, 'private')
    assert.equal((await page(id)).status, 404)
    assert.equal((await page(id, 'pic.png')).status, 404)
    assert.equal((await publicApp.request(`${PUBLIC_APP}/a/${id}`)).status, 404)

    assert.equal((await setVisibility(id, 'public', null)).status, 401)
    const viewerTry = await app.request(`http://app.test/api/artifacts/${id}`, {
      method: 'PATCH',
      headers: { 'content-type': 'application/json', 'x-artifact-client': 'viewer' },
      body: JSON.stringify({ visibility: 'public' })
    })
    assert.equal(viewerTry.status, 401)
    assert.equal((await setVisibility(id, 'everyone')).status, 400)
    const made = await setVisibility(id, 'public')
    assert.equal(made.status, 200)
    assert.equal(((await made.json()) as { visibility: string }).visibility, 'public')

    const shown = await page(id)
    assert.equal(shown.status, 200)
    assert.match(await shown.text(), /報告/)
    assert.equal((await page(id, 'pic.png')).status, 200)
  })

  test('a public page is served for the public origins and never cached as public', async () => {
    const { id } = store.publish({ title: 'Cached', files: { 'index.html': { content: Buffer.from('<!doctype html><img src="pic.png">') }, 'pic.png': { content: PNG } } })
    await setVisibility(id, 'public')
    for (const path of ['', 'pic.png']) {
      const response = await page(id, path)
      assert.equal(response.status, 200, path)
      assert.equal(response.headers.get('cache-control'), 'private, no-cache', path)
      assert.match(response.headers.get('x-robots-tag') ?? '', /noindex/, path)
      assert.match(response.headers.get('content-security-policy') ?? '', new RegExp(`frame-ancestors ${PUBLIC_APP}(;|$)`), path)
      assert.match(response.headers.get('content-security-policy') ?? '', new RegExp(`img-src ${PUBLIC_CONTENT}/a/${id}/ `), path)
    }
    // The bridge talks to the public viewer, not the tailnet one.
    const bridge = await (await publicContent.request(`${PUBLIC_CONTENT}/_/bridge.js`)).text()
    assert.match(bridge, new RegExp(JSON.stringify(PUBLIC_APP).replace(/[.]/g, '\\.')))
    assert.doesNotMatch(bridge, /http:\/\/app\.test/)
    // Opened on its own, a page goes back into the public viewer.
    const top = await publicContent.request(`${PUBLIC_CONTENT}/a/${id}/v/1/`, { headers: { 'sec-fetch-dest': 'document' } })
    assert.equal(top.status, 302)
    assert.equal(top.headers.get('location'), `${PUBLIC_APP}/a/${id}?v=1`)
  })

  test('making a page private again closes its public URLs at once', async () => {
    const { id } = store.publish({ title: 'Flip', files: { 'index.md': { content: Buffer.from('# x\n') }, 'pic.png': { content: PNG } } })
    await setVisibility(id, 'public')
    assert.equal((await page(id, 'pic.png')).status, 200)
    assert.equal((await setVisibility(id, 'private')).status, 200)
    assert.equal((await page(id, 'pic.png')).status, 404)
    assert.equal((await page(id)).status, 404)
    assert.equal((await publicApp.request(`${PUBLIC_APP}/a/${id}`)).status, 404)
    assert.equal((await publicApp.request(`${PUBLIC_APP}/api/artifacts/${id}/files/index.md`)).status, 404)
  })

  test('a page that declares capabilities cannot be public, and a public page cannot gain them', async () => {
    const withCaps = store.publish({ title: 'Caps', capabilities: ['db'], files: { 'index.html': { content: Buffer.from('<p>x</p>') } } })
    const refused = await setVisibility(withCaps.id, 'public')
    assert.equal(refused.status, 409)
    assert.equal(((await refused.json()) as { error: { code: string } }).error.code, 'has_capabilities')

    const plain = store.publish({ title: 'Plain', files: { 'index.html': { content: Buffer.from('<p>x</p>') } } })
    await setVisibility(plain.id, 'public')
    const gain = await app.request(`http://app.test/api/artifacts/${plain.id}`, {
      method: 'PUT',
      headers: { 'content-type': 'application/json', authorization: `Bearer ${TOKEN}` },
      body: JSON.stringify({ capabilities: ['db'], files: {}, baseVersion: 1 })
    })
    assert.equal(gain.status, 409)
    assert.equal(store.requireSummary(plain.id).capabilities.length, 0)
  })
})

describe('public content gate', () => {
  test('a percent-encoded id does not slip past the check on a private artifact', async () => {
    const { id } = store.publish({ title: 'Secret', files: { 'index.md': { content: Buffer.from('# secret\n') }, 'sub/page.html': { content: Buffer.from('<p>secret</p>') } } })
    const asset = store.addAsset(id, 'secret.png', PNG, 'publisher')
    const encoded = `%${id.charCodeAt(0).toString(16).toUpperCase()}${id.slice(1)}`
    for (const path of [`/a/${encoded}/v/1`, `/a/${encoded}/v/1/`, `/a/${encoded}/v/1/sub/page.html`, `/a/${encoded}/assets/${asset.id}/secret.png`, `/a/${id}/assets/${asset.id}/secret.png`]) {
      const response = await publicContent.request(`${PUBLIC_CONTENT}${path}`)
      assert.equal(response.status, 404, path)
    }
  })

  test('only the current version of a public artifact is public', async () => {
    const { id } = store.publish({ title: 'Draft', files: { 'index.md': { content: Buffer.from('# internal draft\n') } } })
    store.update(id, { files: { 'index.md': { content: Buffer.from('# shared\n') } }, baseVersion: 1 })
    await setVisibility(id, 'public')
    assert.equal((await publicContent.request(`${PUBLIC_CONTENT}/a/${id}/v/1/`)).status, 404)
    assert.match(await (await publicContent.request(`${PUBLIC_CONTENT}/a/${id}/v/2/`)).text(), /shared/)
    assert.equal((await publicApp.request(`${PUBLIC_APP}/a/${id}?v=1`)).status, 404)
    assert.equal((await publicApp.request(`${PUBLIC_APP}/api/artifacts/${id}/files/index.md?v=1`)).status, 404)
    const html = await (await publicApp.request(`${PUBLIC_APP}/a/${id}`)).text()
    assert.doesNotMatch(html, /v1|\?v=1/)
    assert.match(html, new RegExp(`/a/${id}/v/2/`))
  })
})

describe('public app', () => {
  test('the public viewer shows the page and nothing about where it came from', async () => {
    const { id } = store.publish({
      title: '公開報告',
      label: 'internal note: fixed /Users/me/secret',
      source: { host: 'Audi的Mac mini', path: '/Users/me/secret-project' },
      files: { 'index.md': { content: Buffer.from('# 公開報告\n') } }
    })
    await setVisibility(id, 'public')
    const response = await publicApp.request(`${PUBLIC_APP}/a/${id}`)
    assert.equal(response.status, 200)
    const html = await response.text()
    assert.match(html, /公開報告/)
    assert.match(html, new RegExp(`<iframe[^>]+src="${PUBLIC_CONTENT}/a/${id}/v/1/"`))
    assert.doesNotMatch(html, /Mac mini|secret|internal note/)
    assert.doesNotMatch(html, /<script/)
    const csp = response.headers.get('content-security-policy') ?? ''
    assert.match(csp, new RegExp(`frame-src ${PUBLIC_CONTENT}/a/${id}/(;|$)`))
    assert.match(csp, /script-src 'none'/)
    assert.match(csp, /frame-ancestors 'none'/)
    assert.equal(response.headers.get('cache-control'), 'private, no-cache')
    assert.match(response.headers.get('x-robots-tag') ?? '', /noindex/)

    // A file of the version, by query, as the content origin's redirect sends it; one it lacks is ignored.
    const other = await publicApp.request(`${PUBLIC_APP}/a/${id}?v=1&path=nope.html`)
    assert.equal(other.status, 200)
    assert.doesNotMatch(await other.text(), /nope\.html/)

    const download = await publicApp.request(`${PUBLIC_APP}/api/artifacts/${id}/files/index.md`)
    assert.equal(download.status, 200)
    assert.match(download.headers.get('content-disposition') ?? '', /^attachment/)
  })

  test('everything but viewing is absent from the public app', async () => {
    const { id } = store.publish({ title: 'Closed', files: { 'index.md': { content: Buffer.from('# x\n') } } })
    await setVisibility(id, 'public')
    const closed: [string, string][] = [
      ['GET', '/'],
      ['GET', '/api/artifacts'],
      ['GET', `/api/artifacts/${id}/comments`],
      ['POST', `/api/artifacts/${id}/comments`],
      ['GET', `/api/artifacts/${id}/errors`],
      ['POST', `/api/artifacts/${id}/errors`],
      ['GET', `/api/artifacts/${id}/diff`],
      ['GET', `/api/artifacts/${id}/events`],
      ['GET', `/api/artifacts/${id}/export.pptx`],
      ['PATCH', `/api/artifacts/${id}`],
      ['PUT', `/api/artifacts/${id}`],
      ['DELETE', `/api/artifacts/${id}`],
      ['POST', '/api/artifacts'],
      ['GET', `/api/artifacts/${id}/db/c`]
    ]
    for (const [method, path] of closed) {
      const response = await publicApp.request(`${PUBLIC_APP}${path}`, { method, headers: { authorization: `Bearer ${TOKEN}` } })
      assert.equal(response.status, 404, `${method} ${path}`)
    }
    assert.equal(store.getSummary(id)?.title, 'Closed')
  })
})

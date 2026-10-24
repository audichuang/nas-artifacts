import assert from 'node:assert/strict'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { after, before, describe, test } from 'node:test'
import { createAppServer } from '../src/app-server.ts'
import { ArtifactStore } from '../src/artifact-store.ts'
import type { Config } from '../src/config.ts'
import { contentSecurityPolicy, createContentServer } from '../src/content-server.ts'

const TOKEN = 'test-token-0123456789'
let dir: string
let store: ArtifactStore
let app: ReturnType<typeof createAppServer>
let content: ReturnType<typeof createContentServer>

const config = (dataDir: string): Config => ({
  dataDir,
  appPort: 0,
  contentPort: 0,
  appOrigin: 'http://app.test',
  contentOrigin: 'http://content.test',
  tokens: [TOKEN]
})

function call(path: string, init: { method?: string; json?: unknown; token?: string; viewer?: boolean; body?: BodyInit } = {}) {
  const headers: Record<string, string> = {}
  if (init.token) headers.authorization = `Bearer ${init.token}`
  if (init.viewer) headers['x-artifact-client'] = 'viewer'
  if (init.json !== undefined) headers['content-type'] = 'application/json'
  return app.request(`http://app.test${path}`, {
    method: init.method ?? 'GET',
    headers,
    body: init.json !== undefined ? JSON.stringify(init.json) : init.body
  })
}

async function publish(files: Record<string, string | null>, extra: Record<string, unknown> = {}) {
  const response = await call('/api/artifacts', {
    method: 'POST',
    token: TOKEN,
    json: { title: 'T', files: Object.fromEntries(Object.entries(files).map(([path, text]) => [path, text === null ? null : { content: text }])), ...extra }
  })
  assert.equal(response.status, 201, await response.clone().text())
  return (await response.json()) as { id: string; currentVersion: number }
}

before(() => {
  dir = mkdtempSync(join(tmpdir(), 'artifacts-test-'))
  store = new ArtifactStore(dir)
  app = createAppServer(store, config(dir))
  content = createContentServer(store, config(dir))
})

after(() => {
  store.close()
  rmSync(dir, { recursive: true, force: true })
})

describe('publishing', () => {
  test('republishing keeps the URL, versions are immutable, and files left out are kept', async () => {
    const { id } = await publish({ 'index.html': '<p>one</p>', 'a.css': 'a{}' })
    const update = await call(`/api/artifacts/${id}`, {
      method: 'PUT',
      token: TOKEN,
      json: { files: { 'index.html': { content: '<p>two</p>' } }, baseVersion: 1 }
    })
    assert.equal(update.status, 200)
    assert.equal(((await update.json()) as { currentVersion: number }).currentVersion, 2)

    const v1 = await content.request(`http://content.test/a/${id}/v/1/`)
    const v2 = await content.request(`http://content.test/a/${id}/v/2/`)
    assert.match(await v1.text(), /one/)
    assert.match(await v2.text(), /two/)
    assert.equal((await content.request(`http://content.test/a/${id}/v/2/a.css`)).status, 200)
  })

  test('a stale baseVersion is a conflict unless forced', async () => {
    const { id } = await publish({ 'index.html': 'x' })
    await call(`/api/artifacts/${id}`, { method: 'PUT', token: TOKEN, json: { files: { 'index.html': { content: 'y' } } } })
    const stale = await call(`/api/artifacts/${id}`, { method: 'PUT', token: TOKEN, json: { files: {}, baseVersion: 1 } })
    assert.equal(stale.status, 409)
    const forced = await call(`/api/artifacts/${id}`, { method: 'PUT', token: TOKEN, json: { files: {}, baseVersion: 1, force: true } })
    assert.equal(forced.status, 200)
  })

  test('null removes a file, and removing the entry needs a new one', async () => {
    const { id } = await publish({ 'index.html': 'x', 'b.html': 'b' })
    const response = await call(`/api/artifacts/${id}`, { method: 'PUT', token: TOKEN, json: { files: { 'index.html': null } } })
    assert.equal(response.status, 200)
    assert.equal(((await response.json()) as { entry: string }).entry, 'b.html')
  })

  test('unsafe paths are rejected', async () => {
    for (const path of ['../x.html', '/abs.html', 'a//b.html', 'a\\b.html', './x.html']) {
      const response = await call('/api/artifacts', { method: 'POST', token: TOKEN, json: { title: 'T', files: { [path]: { content: 'x' } } } })
      assert.equal(response.status, 400, path)
    }
  })

  test('writes need the token; reads do not', async () => {
    const denied = await call('/api/artifacts', { method: 'POST', json: { title: 'T', files: { 'index.html': { content: 'x' } } } })
    assert.equal(denied.status, 401)
    const wrong = await call('/api/artifacts', { method: 'POST', token: 'nope-nope-nope-nope', json: { title: 'T', files: {} } })
    assert.equal(wrong.status, 401)
    const { id } = await publish({ 'index.html': 'x' })
    assert.equal((await call(`/api/artifacts/${id}`, { method: 'DELETE', viewer: true })).status, 401)
    assert.equal((await call(`/api/artifacts/${id}`)).status, 200)
  })

  test('deleting drops blobs nothing else references', async () => {
    const shared = 'shared-body-' + Math.random()
    const keep = await publish({ 'index.html': shared })
    const gone = await publish({ 'index.html': shared, 'only.txt': 'only-' + Math.random() })
    const before = store.blobs.list().length
    assert.equal((await call(`/api/artifacts/${gone.id}`, { method: 'DELETE', token: TOKEN })).status, 204)
    assert.equal(store.blobs.list().length, before - 1)
    assert.equal((await content.request(`http://content.test/a/${keep.id}/v/1/`)).status, 200)
  })
})

describe('content origin', () => {
  test('pages are sandboxed, framed only by the app, and get the bridge', async () => {
    const { id } = await publish({ 'index.html': '<html><head><title>x</title></head><body>hi</body></html>' })
    const response = await content.request(`http://content.test/a/${id}/v/1/`)
    assert.equal(response.headers.get('content-security-policy'), contentSecurityPolicy('http://app.test'))
    assert.match(response.headers.get('content-security-policy') ?? '', /^sandbox allow-scripts/)
    assert.doesNotMatch(response.headers.get('content-security-policy') ?? '', /allow-same-origin/)
    assert.match(await response.text(), /<head><script src="\/_\/bridge\.js"><\/script><title>/)
  })

  test('markdown entries render to HTML', async () => {
    const { id } = await publish({ 'index.md': '# Hi\n\n*there*' })
    const html = await (await content.request(`http://content.test/a/${id}/v/1/`)).text()
    assert.match(html, /<h1>Hi<\/h1>/)
    assert.match(html, /<em>there<\/em>/)
  })
})

describe('page capabilities', () => {
  test('the viewer may use only what the artifact declared', async () => {
    const plain = await publish({ 'index.html': 'x' })
    const denied = await call(`/api/artifacts/${plain.id}/db/c/d`, { method: 'PUT', viewer: true, json: { data: 1 } })
    assert.equal(denied.status, 403)
    const noHeader = await call(`/api/artifacts/${plain.id}/comments`, { method: 'POST', json: { author: 'a', body: 'b' } })
    assert.equal(noHeader.status, 401)

    const withDb = await publish({ 'index.html': 'x' }, { capabilities: ['db'] })
    const set = await call(`/api/artifacts/${withDb.id}/db/c/d`, { method: 'PUT', viewer: true, json: { data: { a: 1 } } })
    assert.equal(set.status, 200)
    const upload = await call(`/api/artifacts/${withDb.id}/assets?name=x.png`, { method: 'POST', viewer: true, body: 'png' })
    assert.equal(upload.status, 403)
  })

  test('db is scoped per artifact, merges on update, and queries by field', async () => {
    const one = await publish({ 'index.html': 'x' }, { capabilities: ['db'] })
    const two = await publish({ 'index.html': 'x' }, { capabilities: ['db'] })
    await call(`/api/artifacts/${one.id}/db/tasks/t1`, { method: 'PUT', viewer: true, json: { data: { title: 'a', done: false } } })
    await call(`/api/artifacts/${one.id}/db/tasks/t1`, { method: 'PATCH', viewer: true, json: { data: { done: true } } })
    await call(`/api/artifacts/${one.id}/db/tasks/t2`, { method: 'PUT', viewer: true, json: { data: { title: 'b', done: false } } })

    const doc = (await (await call(`/api/artifacts/${one.id}/db/tasks/t1`, { viewer: true })).json()) as { data: unknown }
    assert.deepEqual(doc.data, { title: 'a', done: true })
    const done = (await (await call(`/api/artifacts/${one.id}/db/tasks/_query`, { method: 'POST', viewer: true, json: { where: { done: true } } })).json()) as { docs: { id: string }[] }
    assert.deepEqual(done.docs.map((d) => d.id), ['t1'])
    assert.equal((await call(`/api/artifacts/${two.id}/db/tasks/t1`, { viewer: true })).status, 404)
  })

  test('a failing batch writes nothing', async () => {
    const { id } = await publish({ 'index.html': 'x' }, { capabilities: ['db'] })
    const response = await call(`/api/artifacts/${id}/db/_batch`, {
      method: 'POST',
      viewer: true,
      json: { ops: [{ op: 'set', collection: 'c', id: 'ok', data: 1 }, { op: 'bogus', collection: 'c', id: 'x' }] }
    })
    assert.equal(response.status, 400)
    assert.equal((await call(`/api/artifacts/${id}/db/c/ok`, { viewer: true })).status, 404)
  })

  test('comments thread under their root and resolve on the root only', async () => {
    const { id } = await publish({ 'index.html': 'x' })
    const root = (await (await call(`/api/artifacts/${id}/comments`, { method: 'POST', viewer: true, json: { author: 'a', body: 'one' } })).json()) as { id: string }
    const reply = (await (await call(`/api/artifacts/${id}/comments`, { method: 'POST', viewer: true, json: { author: 'b', body: 'two', parentId: root.id } })).json()) as { id: string; parentId: string }
    const nested = (await (await call(`/api/artifacts/${id}/comments`, { method: 'POST', viewer: true, json: { author: 'c', body: 'three', parentId: reply.id } })).json()) as { parentId: string }
    assert.equal(reply.parentId, root.id)
    assert.equal(nested.parentId, root.id)
    assert.equal((await call(`/api/artifacts/${id}/comments/${reply.id}`, { method: 'PATCH', viewer: true, json: { resolved: true } })).status, 404)
    assert.equal((await call(`/api/artifacts/${id}/comments/${root.id}`, { method: 'PATCH', viewer: true, json: { resolved: true } })).status, 204)
  })
})

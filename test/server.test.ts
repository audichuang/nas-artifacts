import assert from 'node:assert/strict'
import { mkdtempSync, rmSync } from 'node:fs'
import { DatabaseSync } from 'node:sqlite'
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

  test('PUT without baseVersion is rejected unless forced', async () => {
    const { id } = await publish({ 'index.html': 'x' })
    const missing = await call(`/api/artifacts/${id}`, { method: 'PUT', token: TOKEN, json: { files: { 'index.html': { content: 'y' } } } })
    assert.equal(missing.status, 400)
    assert.equal(((await missing.json()) as { error: { code: string } }).error.code, 'base_version_required')
    const forced = await call(`/api/artifacts/${id}`, { method: 'PUT', token: TOKEN, json: { files: { 'index.html': { content: 'y' } }, force: true } })
    assert.equal(forced.status, 200)
  })

  test('a stale baseVersion is a conflict that names the files changed since, unless forced', async () => {
    const { id } = await publish({ 'index.html': 'x', 'keep.txt': 'k', 'gone.txt': 'g', 'edit.txt': 'e' })
    const other = await call(`/api/artifacts/${id}`, {
      method: 'PUT',
      token: TOKEN,
      json: { files: { 'edit.txt': { content: 'e2' }, 'new.txt': { content: 'n' }, 'gone.txt': null }, baseVersion: 1 }
    })
    assert.equal(other.status, 200)
    const stale = await call(`/api/artifacts/${id}`, { method: 'PUT', token: TOKEN, json: { files: {}, baseVersion: 1 } })
    assert.equal(stale.status, 409)
    const { error } = (await stale.json()) as { error: { code: string; details: { currentVersion: number; baseVersion: number; changedFiles: string[] } } }
    assert.equal(error.code, 'version_conflict')
    assert.equal(error.details.currentVersion, 2)
    assert.equal(error.details.baseVersion, 1)
    assert.deepEqual([...error.details.changedFiles].sort(), ['edit.txt', 'gone.txt', 'new.txt'])
    const forced = await call(`/api/artifacts/${id}`, { method: 'PUT', token: TOKEN, json: { files: {}, baseVersion: 1, force: true } })
    assert.equal(forced.status, 200)
  })

  test('null removes a file, and removing the entry needs a new one', async () => {
    const { id } = await publish({ 'index.html': 'x', 'b.html': 'b' })
    const response = await call(`/api/artifacts/${id}`, { method: 'PUT', token: TOKEN, json: { files: { 'index.html': null }, baseVersion: 1 } })
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

describe('store hardening', () => {
  const unique = (label: string) => `${label}-${Math.random()}`

  test('a rejected publish or update leaves no blob behind', async () => {
    const before = store.blobs.list().length
    const badEntry = await call('/api/artifacts', { method: 'POST', token: TOKEN, json: { title: 'T', entry: 'missing.html', files: { 'index.html': { content: unique('entry') } } } })
    assert.equal(badEntry.status, 400)
    const files = Object.fromEntries(Array.from({ length: 512 }, (_, i) => [`f${i}.txt`, { content: unique('many') }]))
    const tooMany = await call('/api/artifacts', { method: 'POST', token: TOKEN, json: { title: 'T', files: { 'index.html': { content: 'x' }, ...files } } })
    assert.equal(tooMany.status, 413)
    assert.equal(store.blobs.list().length, before)

    const { id } = await publish({ 'index.html': unique('base') })
    const kept = store.blobs.list().length
    const rejected = await call(`/api/artifacts/${id}`, { method: 'PUT', token: TOKEN, json: { entry: 'nope.html', files: { 'extra.txt': { content: unique('upd') } }, baseVersion: 1 } })
    assert.equal(rejected.status, 400)
    assert.equal(store.blobs.list().length, kept)
  })

  test('non-integer limit and offset are clamped, not a server error', async () => {
    for (const query of ['limit=abc', 'limit=1.5', 'offset=abc', 'offset=1.5', 'offset=1e30', 'limit=Infinity']) {
      assert.equal((await call(`/api/artifacts?${query}`)).status, 200, query)
    }
    const { id } = await publish({ 'index.html': 'x' }, { capabilities: ['db'] })
    await call(`/api/artifacts/${id}/db/c/d`, { method: 'PUT', viewer: true, json: { data: 1 } })
    assert.equal((await call(`/api/artifacts/${id}/db/c?limit=1.5`, { viewer: true })).status, 200)
    assert.equal((await call(`/api/artifacts/${id}/db/c?limit=abc`, { viewer: true })).status, 200)
    const query = await call(`/api/artifacts/${id}/db/c/_query`, { method: 'POST', viewer: true, json: { limit: 1.5 } })
    assert.equal(query.status, 200)
  })

  test('a malformed file path or a missing static file is a 404', async () => {
    const { id } = await publish({ 'index.html': 'x' })
    const bad = await call(`/api/artifacts/${id}/files/%E0%A4%A`)
    assert.equal(bad.status, 404)
    assert.equal(((await bad.json()) as { error: { code: string } }).error.code, 'file_not_found')
    assert.equal((await call('/static/missing.js')).status, 404)
  })

  test('deleting an asset keeps blobs other artifacts or versions still reference', async () => {
    const shared = unique('asset-shared')
    const keep = await publish({ 'index.html': shared })
    const other = await publish({ 'index.html': 'x' })
    const solo = unique('asset-solo')
    const sharedAsset = store.addAsset(other.id, 'a.txt', Buffer.from(shared), 'publisher')
    const soloAsset = store.addAsset(other.id, 'b.txt', Buffer.from(solo), 'publisher')
    const count = store.blobs.list().length
    store.deleteAsset(other.id, sharedAsset.id)
    assert.equal(store.blobs.list().length, count)
    store.deleteAsset(other.id, soloAsset.id)
    assert.equal(store.blobs.list().length, count - 1)
    assert.equal((await content.request(`http://content.test/a/${keep.id}/v/1/`)).status, 200)
  })
})

describe('content origin', () => {
  test('pages are sandboxed, framed only by the app, and get the bridge', async () => {
    const { id } = await publish({ 'index.html': '<html><head><title>x</title></head><body>hi</body></html>' })
    const response = await content.request(`http://content.test/a/${id}/v/1/`)
    assert.equal(response.headers.get('content-security-policy'), contentSecurityPolicy('http://app.test', 'http://content.test', id))
    assert.match(response.headers.get('content-security-policy') ?? '', /^sandbox allow-scripts/)
    assert.doesNotMatch(response.headers.get('content-security-policy') ?? '', /allow-same-origin/)
    assert.match(await response.text(), /<head><script src="\/_\/bridge\.js"><\/script><title>/)
  })

  test('markdown entries render to HTML', async () => {
    const { id } = await publish({ 'index.md': '# Hi\n\n*there*' })
    const html = await (await content.request(`http://content.test/a/${id}/v/1/`)).text()
    assert.match(html, /<h1 id="hi">.*Hi<\/h1>/)
    assert.match(html, /<em>there<\/em>/)
    assert.doesNotMatch(html, /class="toc"|mermaid/)
  })

  test('markdown headings get unique ids that keep CJK, and three or more sections get a table of contents', async () => {
    const md = '# 報告\n\n## 背景說明\n\n## 背景說明\n\n### API **設計**\n\n## Q&A: 問題?\n'
    const { id } = await publish({ 'index.md': md })
    const html = await (await content.request(`http://content.test/a/${id}/v/1/`)).text()
    assert.match(html, /<h2 id="背景說明">/)
    assert.match(html, /<h2 id="背景說明-1">/)
    assert.match(html, /<h3 id="api-設計">/)
    assert.match(html, /<h2 id="qa-問題">/)
    assert.match(html, /<nav class="toc"[^>]*>[\s\S]*href="#背景說明-1"[\s\S]*<\/nav>/)
    assert.doesNotMatch(/<nav class="toc"[\s\S]*<\/nav>/.exec(html)?.[0] ?? '', /#報告/)
  })

  test('markdown code blocks are highlighted on the server; unknown languages are only escaped', async () => {
    const md = '```js\nconst a = "<b>"\n```\n\n```nosuchlang\n<b>raw</b>\n```\n'
    const { id } = await publish({ 'index.md': md })
    const html = await (await content.request(`http://content.test/a/${id}/v/1/`)).text()
    assert.match(html, /<code class="hljs language-js"><span class="hljs-keyword">const<\/span>/)
    assert.match(html, /&#60;b&#62;raw&#60;\/b&#62;/)
    assert.doesNotMatch(html, /<b>raw<\/b>/)
  })

  test('mermaid is drawn on the server as SVG themed by CSS, with no script; a broken diagram keeps its source', async () => {
    const md = '```mermaid\nflowchart LR\n  A[寫檔] --> B[發布]\n```\n\n```mermaid\nflowchart LR\n  A -->\n```\n'
    const { id } = await publish({ 'index.md': md })
    const response = await content.request(`http://content.test/a/${id}/v/1/`)
    const html = await response.text()
    assert.match(html, /<figure class="diagram"><svg [^>]*class="pele pele-flowchart"[^>]*>[\s\S]*寫檔[\s\S]*<\/svg><\/figure>/)
    assert.match(html, /var\(--pele-fg/)
    assert.match(html, /<div class="diagram-failed"><p>Mermaid: [^<]*line 3/)
    assert.match(html, /<pre><code>flowchart LR\n {2}A --&#62;\n<\/code><\/pre>/)
    assert.equal(html.match(/<script>/g)?.length, 1)
    assert.doesNotMatch(response.headers.get('content-security-policy') ?? '', /mermaid|pele/)
  })

  test('a renderer change revalidates markdown pages browsers already hold', async () => {
    const { id } = await publish({ 'index.md': '# x' })
    const etag = async (server: ReturnType<typeof createContentServer>) => (await server.request(`http://content.test/a/${id}/v/1/`)).headers.get('etag')
    const bumped = createContentServer(store, config(dir), { renderVersion: 'test-bump' })
    assert.notEqual(await etag(content), await etag(bumped))
  })
})

describe('content origin paths and caching', () => {
  test('a nested entry redirects so relative URLs resolve against its directory', async () => {
    const { id } = await publish({ 'pages/my report.html': '<html><head></head><body>r</body></html>', 'pages/app.js': 'x' })
    for (const path of [`/a/${id}/v/1/`, `/a/${id}/v/1`]) {
      const response = await content.request(`http://content.test${path}`)
      assert.equal(response.status, 302, path)
      assert.equal(response.headers.get('location'), `/a/${id}/v/1/pages/my%20report.html`)
    }
    const target = await content.request(`http://content.test/a/${id}/v/1/pages/my%20report.html`)
    assert.equal(target.status, 200)
  })

  test('a flat entry is served at the version root and redirected without the slash', async () => {
    const { id } = await publish({ 'index.html': 'x' })
    assert.equal((await content.request(`http://content.test/a/${id}/v/1/`)).status, 200)
    const bare = await content.request(`http://content.test/a/${id}/v/1`)
    assert.equal(bare.status, 302)
    assert.equal(bare.headers.get('location'), `/a/${id}/v/1/`)
  })

  test('malformed percent-encoding is a 404, not a 500', async () => {
    const { id } = await publish({ 'index.html': 'x' })
    const response = await content.request(`http://content.test/a/${id}/v/1/%E0%A4%A`)
    assert.equal(response.status, 404)
  })

  test('each artifact may load only its own files', async () => {
    const policy = contentSecurityPolicy('http://app.test', 'http://content.test', 'abc')
    assert.match(policy, /script-src http:\/\/content\.test\/a\/abc\/ http:\/\/content\.test\/_\/bridge\.js /)
    assert.doesNotMatch(policy, /'self'|allow-popups/)
  })

  test('If-None-Match answers 304 without a body; markdown revalidates when the title changes', async () => {
    const { id } = await publish({ 'index.html': '<p>x</p>', 'style.css': 'a{}' })
    for (const path of ['', 'style.css']) {
      const first = await content.request(`http://content.test/a/${id}/v/1/${path}`)
      const etag = first.headers.get('etag')
      assert.ok(etag)
      const again = await content.request(`http://content.test/a/${id}/v/1/${path}`, { headers: { 'if-none-match': etag } })
      assert.equal(again.status, 304)
      assert.equal(await again.text(), '')
      assert.match(again.headers.get('content-security-policy') ?? '', /sandbox/)
    }
    const md = await publish({ 'index.md': '# a' })
    const mdTag = (await content.request(`http://content.test/a/${md.id}/v/1/`)).headers.get('etag')
    const other = await publish({ 'index.md': '# a' }, { title: 'Different' })
    const otherTag = (await content.request(`http://content.test/a/${other.id}/v/1/`)).headers.get('etag')
    assert.notEqual(mdTag, otherTag)
  })
})

describe('rendered pages opened as a top-level document', () => {
  const get = (path: string, dest?: string) =>
    content.request(`http://content.test${path}`, { headers: dest ? { 'sec-fetch-dest': dest } : {} })

  test('a document navigation to a rendered page goes back to the viewer, with version and path', async () => {
    const { id } = await publish({ 'index.html': '<p>x</p>', 'docs/other.md': '# o', 'a.css': 'a{}' })
    const entry = await get(`/a/${id}/v/1/`, 'document')
    assert.equal(entry.status, 302)
    assert.equal(entry.headers.get('location'), `http://app.test/a/${id}?v=1`)
    assert.match(entry.headers.get('vary') ?? '', /sec-fetch-dest/i)

    const nested = await get(`/a/${id}/v/1/docs/other.md`, 'document')
    assert.equal(nested.status, 302)
    assert.equal(nested.headers.get('location'), `http://app.test/a/${id}?v=1&path=docs%2Fother.md`)
  })

  test('the viewer frame and header-less clients still get the page', async () => {
    const { id } = await publish({ 'index.html': '<p>x</p>' })
    for (const dest of ['iframe', undefined]) {
      const response = await get(`/a/${id}/v/1/`, dest)
      assert.equal(response.status, 200)
      assert.match(response.headers.get('vary') ?? '', /sec-fetch-dest/i)
    }
  })

  test('non-rendered files are served to documents unchanged', async () => {
    const { id } = await publish({ 'index.html': 'x', 'a.css': 'a{}' })
    assert.equal((await get(`/a/${id}/v/1/a.css`, 'document')).status, 200)
  })

  test('the viewer loads a requested file only when it exists in that version', async () => {
    const { id } = await publish({ 'index.html': 'x', 'docs/other.md': '# o' })
    const pathOf = async (query: string) => {
      const html = await (await call(`/a/${id}?${query}`)).text()
      const json = /<script id="page-data" type="application\/json">(.*?)<\/script>/s.exec(html)?.[1] ?? '{}'
      return (JSON.parse(json) as { path?: string }).path
    }
    assert.equal(await pathOf('path=docs/other.md'), 'docs/other.md')
    assert.equal(await pathOf('path=../x'), undefined)
    assert.equal(await pathOf('path=missing.html'), undefined)
  })
})

describe('app origin framing', () => {
  test('the viewer may frame only its own artifact; the gallery frames nothing', async () => {
    const { id } = await publish({ 'index.html': 'x' })
    const viewer = await call(`/a/${id}`)
    assert.equal(viewer.headers.get('content-security-policy'), `frame-src http://content.test/a/${id}/`)
    assert.doesNotMatch(await viewer.text(), /allow-popups/)
    assert.equal((await call('/')).headers.get('content-security-policy'), "frame-src 'none'")
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

describe('page errors', () => {
  const report = (id: string, body: unknown, viewer = true) => call(`/api/artifacts/${id}/errors`, { method: 'POST', viewer, json: body })
  const list = async (id: string, query = '') =>
    ((await (await call(`/api/artifacts/${id}/errors${query}`)).json()) as { errors: { kind: string; message: string; path: string; count: number; version: number }[] }).errors

  test('the viewer records them per version, counting repeats instead of storing them again', async () => {
    const { id } = await publish({ 'index.html': 'x', 'b.html': 'y' })
    const one = { kind: 'csp', message: 'connect-src blocked https://x.test/', source: 'index.html:3' }
    assert.equal((await report(id, { version: 1, path: 'index.html', errors: [one] }, false)).status, 401)
    assert.equal((await report(id, { version: 1, path: 'index.html', errors: [one, one] })).status, 204)
    assert.equal((await report(id, { version: 1, path: 'b.html', errors: [{ kind: 'error', message: 'boom', source: null }] })).status, 204)
    const errors = await list(id)
    assert.deepEqual(errors.map((e) => [e.kind, e.path, e.count]).sort(), [['csp', 'index.html', 2], ['error', 'b.html', 1]])

    await call(`/api/artifacts/${id}`, { method: 'PUT', token: TOKEN, json: { baseVersion: 1, files: { 'index.html': { content: 'z' } } } })
    assert.deepEqual(await list(id), [])
    assert.equal((await list(id, '?v=1')).length, 2)
  })

  test('reports are bounded and validated', async () => {
    const { id } = await publish({ 'index.html': 'x' })
    assert.equal((await report(id, { version: 9, path: 'index.html', errors: [] })).status, 404)
    assert.equal((await report(id, { version: 1, path: 'index.html', errors: [{ kind: 'nope', message: 'm' }] })).status, 400)
    const many = Array.from({ length: 21 }, (_, i) => ({ kind: 'error', message: `e${i}` }))
    assert.equal((await report(id, { version: 1, path: 'index.html', errors: many })).status, 400)
    for (let batch = 0; batch < 8; batch++) {
      const errors = Array.from({ length: 20 }, (_, i) => ({ kind: 'error', message: `${batch}-${i}${'m'.repeat(2000)}` }))
      assert.equal((await report(id, { version: 1, path: 'index.html', errors })).status, 204)
    }
    const stored = await list(id)
    assert.equal(stored.length, 100)
    assert.ok(stored.every((e) => e.message.length <= 500))
  })
})

describe('source', () => {
  test('each version records where it was published from, and the artifact shows the latest', async () => {
    const source = { host: 'mac-mini', path: '~/Project/report' }
    const created = await publish({ 'index.html': 'x' }, { source })
    let detail = (await (await call(`/api/artifacts/${created.id}`)).json()) as { source: unknown; versions: { version: number; source: unknown }[] }
    assert.deepEqual(detail.source, source)
    await call(`/api/artifacts/${created.id}`, { method: 'PUT', token: TOKEN, json: { baseVersion: 1, files: { 'index.html': { content: 'y' } } } })
    const other = { host: 'macbook-pro', path: '~/work/report' }
    await call(`/api/artifacts/${created.id}`, { method: 'PUT', token: TOKEN, json: { baseVersion: 2, source: other, files: { 'index.html': { content: 'z' } } } })
    detail = (await (await call(`/api/artifacts/${created.id}`)).json()) as typeof detail
    assert.deepEqual(detail.source, other)
    assert.deepEqual(detail.versions.map((v) => v.source), [other, null, source])
    const bad = await call('/api/artifacts', { method: 'POST', token: TOKEN, json: { title: 'T', source: { host: 1 }, files: { 'index.html': { content: 'x' } } } })
    assert.equal(bad.status, 400)
  })
})

describe('pruning old versions', () => {
  const put = (id: string, baseVersion: number, files: Record<string, string>) =>
    call(`/api/artifacts/${id}`, { method: 'PUT', token: TOKEN, json: { baseVersion, files: Object.fromEntries(Object.entries(files).map(([p, t]) => [p, { content: t }])) } })

  test('keeps the newest N versions, frees blobs only nothing else references, and needs the token', async () => {
    const { id } = await publish({ 'index.html': 'shared', 'a.txt': 'only-v1' })
    await put(id, 1, { 'a.txt': 'only-v2' })
    await put(id, 2, { 'a.txt': 'v3' })
    await put(id, 3, { 'a.txt': 'v4' })
    const other = await publish({ 'index.html': 'only-v2' })
    assert.equal((await call(`/api/artifacts/${id}/prune`, { method: 'POST', viewer: true, json: { keep: 1 } })).status, 401)
    assert.equal((await call(`/api/artifacts/${id}/prune`, { method: 'POST', token: TOKEN, json: { keep: 0 } })).status, 400)
    store.recordErrors(id, { version: 1, path: 'index.html', errors: [{ kind: 'error', message: 'old' }] })

    const response = await call(`/api/artifacts/${id}/prune`, { method: 'POST', token: TOKEN, json: { keep: 2 } })
    assert.equal(response.status, 200)
    const result = (await response.json()) as { removed: number[]; freedBytes: number }
    assert.deepEqual(result.removed, [2, 1])
    assert.equal(result.freedBytes, 'only-v1'.length)
    const detail = (await (await call(`/api/artifacts/${id}`)).json()) as { versions: { version: number }[]; currentVersion: number }
    assert.deepEqual(detail.versions.map((v) => v.version), [4, 3])
    assert.equal(detail.currentVersion, 4)
    // Still referenced: by the kept versions (index.html) or by another artifact (only-v2).
    assert.equal((await content.request(`http://content.test/a/${id}/v/4/`)).status, 200)
    assert.equal((await content.request(`http://content.test/a/${other.id}/v/1/`)).status, 200)
    assert.equal(store.collectGarbage(), 0)
    assert.equal((await call(`/api/artifacts/${id}/errors?v=1`)).status, 404)

    const again = (await (await call(`/api/artifacts/${id}/prune`, { method: 'POST', token: TOKEN, json: { keep: 5 } })).json()) as { removed: number[] }
    assert.deepEqual(again.removed, [])
  })

  test('a publish based on a pruned version is still a conflict naming every file, not a missing version', async () => {
    const { id } = await publish({ 'index.html': 'a', 'b.txt': 'b' })
    await put(id, 1, { 'index.html': 'a2' })
    await put(id, 2, { 'index.html': 'a3' })
    await call(`/api/artifacts/${id}/prune`, { method: 'POST', token: TOKEN, json: { keep: 1 } })
    const response = await put(id, 1, { 'index.html': 'mine' })
    assert.equal(response.status, 409)
    const body = (await response.json()) as { error: { code: string; details: { changedFiles: string[] } } }
    assert.equal(body.error.code, 'version_conflict')
    assert.deepEqual(body.error.details.changedFiles, ['b.txt', 'index.html'])
  })

  test('the viewer asked for a pruned version opens the current one', async () => {
    const { id } = await publish({ 'index.html': 'a' })
    await put(id, 1, { 'index.html': 'b' })
    await call(`/api/artifacts/${id}/prune`, { method: 'POST', token: TOKEN, json: { keep: 1 } })
    const page = await (await call(`/a/${id}?v=1`)).text()
    assert.match(page, /"version":2/)
  })
})

describe('search', () => {
  type Found = { artifacts: { id: string; snippet?: string | null }[] }
  const search = async (q: string) => (await (await call(`/api/artifacts?q=${encodeURIComponent(q)}`)).json()) as Found

  test('finds words in the page text, including two-character CJK words, with a snippet around the hit', async () => {
    const md = `# 週報\n\n${'前言很長。'.repeat(30)}這週完成了倉庫驗貨流程的自動化。\n`
    const { id } = await publish({ 'index.md': md, 'b.html': '<html><head><style>.zebra{}</style><script>var giraffe = 1</script></head><body><p>Hello &amp; welcome</p></body></html>' })
    const hit = (await search('驗貨')).artifacts.find((a) => a.id === id)
    assert.ok(hit, 'found by body text')
    assert.match(hit.snippet ?? '', /倉庫驗貨流程/)
    assert.ok((hit.snippet ?? '').length < 120)
    assert.ok((await search('welcome')).artifacts.some((a) => a.id === id))
    const words = await publish({ 'index.md': '- 用 nas-artifacts 的 search_text 表\n\n**粗體**與 `code_name`\n' })
    for (const q of ['nas-artifacts', 'search_text', 'code_name', '粗體與']) {
      assert.ok((await search(q)).artifacts.some((a) => a.id === words.id), q)
    }
    assert.ok((await search('Hello & welcome')).artifacts.some((a) => a.id === id))
    assert.ok(!(await search('giraffe')).artifacts.some((a) => a.id === id), 'script text is not page text')
    assert.ok(!(await search('zebra')).artifacts.some((a) => a.id === id), 'style text is not page text')
  })

  test('searches the current version only, and the source machine and folder', async () => {
    const { id } = await publish({ 'index.md': '舊的說法:麒麟' }, { source: { host: 'studio-mac', path: '~/Project/kirin' } })
    await call(`/api/artifacts/${id}`, { method: 'PUT', token: TOKEN, json: { baseVersion: 1, files: { 'index.md': { content: '新的說法:鳳凰' } } } })
    assert.ok(!(await search('麒麟')).artifacts.some((a) => a.id === id))
    assert.ok((await search('鳳凰')).artifacts.some((a) => a.id === id))
    assert.ok((await search('studio-mac')).artifacts.some((a) => a.id === id))
    assert.ok((await search('Project/kirin')).artifacts.some((a) => a.id === id))
  })

  test('artifacts published before search existed are indexed when the server starts', async () => {
    const own = mkdtempSync(join(tmpdir(), 'artifacts-backfill-'))
    try {
      const before = new ArtifactStore(own)
      const { id } = before.publish({ title: 'Old', files: { 'index.md': { content: Buffer.from('早就發佈的玄武') } } })
      before.close()
      const raw = new DatabaseSync(join(own, 'artifacts.db'))
      raw.exec('DELETE FROM search_text')
      raw.close()
      const after = new ArtifactStore(own)
      assert.deepEqual(after.list({ query: '玄武' }).artifacts.map((a) => a.id), [id])
      after.close()
    } finally {
      rmSync(own, { recursive: true, force: true })
    }
  })
})

describe('version diff', () => {
  test('lists changed files with line changes, defaulting to the current version against the one before', async () => {
    const md = (middle: string) => `# 報告\n\n${Array.from({ length: 20 }, (_, i) => `第 ${i} 段`).join('\n')}\n${middle}\n結尾\n`
    const { id } = await publish({ 'index.md': md('舊結論'), 'logo.png': 'PNGDATA', 'gone.txt': 'bye' })
    await call(`/api/artifacts/${id}`, {
      method: 'PUT',
      token: TOKEN,
      json: { baseVersion: 1, label: '改結論', source: { host: 'mbp', path: '~/r' }, files: { 'index.md': { content: md('新結論') }, 'logo.png': { content: 'PNG2' }, 'gone.txt': null, 'new.txt': { content: 'hi' } } }
    })
    const diff = (await (await call(`/api/artifacts/${id}/diff`)).json()) as { from: number; to: number; label: string; source: { host: string }; files: { path: string; status: string; note?: string; lines?: { op: string; text?: string; skipped?: number }[] }[] }
    assert.equal(diff.from, 1)
    assert.equal(diff.to, 2)
    assert.equal(diff.label, '改結論')
    assert.equal(diff.source.host, 'mbp')
    assert.deepEqual(diff.files.map((f) => [f.path, f.status, f.note ?? null]), [
      ['gone.txt', 'removed', null],
      ['index.md', 'modified', null],
      ['logo.png', 'modified', 'binary'],
      ['new.txt', 'added', null]
    ])
    const lines = diff.files[1].lines ?? []
    assert.deepEqual(lines.filter((l) => l.op === '-' || l.op === '+'), [{ op: '-', text: '舊結論' }, { op: '+', text: '新結論' }])
    assert.equal(lines[0].op, '…')
    const first = (await (await call(`/api/artifacts/${id}/diff?to=1`)).json()) as { from: number | null; files: { status: string }[] }
    assert.equal(first.from, null)
    assert.ok(first.files.every((f) => f.status === 'added'))
  })
})

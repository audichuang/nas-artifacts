import assert from 'node:assert/strict'
import { createHmac } from 'node:crypto'
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createServer, type Server } from 'node:http'
import { after, before, test } from 'node:test'
import { ArtifactStore } from '../src/artifact-store.ts'
import { createAppServer } from '../src/app-server.ts'
import { zip } from './helpers/zip.ts'

const secret = 'office-editor-test-secret-0123456789'
const pptx = (text: string) => zip({
  '[Content_Types].xml': '<Types/>',
  'ppt/presentation.xml': `<p:presentation xmlns:p="p"><p:sldSz cx="9144000" cy="6858000"/>${text}</p:presentation>`
})
const sign = (value: unknown) => {
  const body = `${Buffer.from('{"alg":"HS256","typ":"JWT"}').toString('base64url')}.${Buffer.from(JSON.stringify(value)).toString('base64url')}`
  return `${body}.${createHmac('sha256', secret).update(body).digest('base64url')}`
}
let server: Server
let docsUrl: string
let result = pptx('changed')
let downloads = 0
before(async () => {
  server = createServer((req, res) => {
    downloads++
    if (req.url === '/redirect') { res.writeHead(302, { location: 'http://other.test/secret' }); res.end(); return }
    if (req.method === 'POST') { res.setHeader('content-type', 'application/json'); res.end('{"error":4}'); return }
    res.end(result)
  })
  await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve))
  docsUrl = `http://127.0.0.1:${(server.address() as { port: number }).port}`
})
after(() => server.close())

function fixture() {
  const dir = mkdtempSync(join(tmpdir(), 'office-editor-'))
  const store = new ArtifactStore(dir)
  const config = { dataDir: dir, appPort: 0, contentPort: 0, appOrigin: 'http://app.test', contentOrigin: 'http://content.test', tokens: ['test-token'],
    office: { url: docsUrl, internalUrl: docsUrl, callbackOrigin: 'http://app-internal.test', secret } }
  const app = createAppServer(store, config)
  const artifact = store.publish({ title: '測試簡報', files: { 'deck.pptx': { content: pptx('original') } } })
  const root = `/api/artifacts/${artifact.id}/office-edits`
  const start = async (headers = { 'x-artifact-client': 'viewer' }) => app.request(root, { method: 'POST', headers: { ...headers, 'content-type': 'application/json' }, body: '{"baseVersion":1}' })
  const callback = (sid: string, payload: Record<string, unknown>, body: unknown = payload) => app.request(`/api/office/callback/${sid}`, {
    method: 'POST', headers: { 'content-type': 'application/json', authorization: `Bearer ${sign({ payload, operation: 'callback' })}` }, body: JSON.stringify(body)
  })
  return { app, store, config, artifact, root, start, callback, close: () => { store.close(); rmSync(dir, { recursive: true, force: true }) } }
}

test('Office editing requires the viewer header and is restricted to current capability-free PPTX', async () => {
  const f = fixture()
  try {
    assert.equal((await f.start({} as { 'x-artifact-client': string })).status, 401)
    const response = await f.start()
    assert.equal(response.status, 201)
    const session = await response.json()
    assert.match(session.editorUrl, new RegExp(`/a/${f.artifact.id}/edit/`))
    const cfg = await (await f.app.request(`${f.root}/${session.sessionId}/config`)).json()
    assert.equal(cfg.documentType, 'slide')
    assert.equal(cfg.editorConfig.lang, 'zh-TW')
    assert.equal(cfg.document.permissions.edit, true)
    assert.equal(cfg.editorConfig.customization.forcesave, true)
    const source = await f.app.request(cfg.document.url)
    assert.ok(Buffer.from(await source.arrayBuffer()).equals(pptx('original')))
    f.store.update(f.artifact.id, { capabilities: ['db'], baseVersion: 1 })
    assert.equal((await f.start()).status, 403)
  } finally { f.close() }
})

test('signed callbacks save native PPTX as new versions, survive restarts and are idempotent', async () => {
  const f = fixture()
  try {
    const { sessionId } = await (await f.start()).json()
    // Recreating the HTTP application must not lose the edit session.
    const restarted = createAppServer(f.store, f.config)
    assert.equal((await restarted.request(`${f.root}/${sessionId}/config`)).status, 200)
    const payload = { key: sessionId, status: 6, url: `${docsUrl}/result.pptx`, filetype: 'pptx' }
    assert.equal((await f.callback(sessionId, payload)).status, 200)
    assert.equal(f.store.requireSummary(f.artifact.id).currentVersion, 2)
    assert.deepEqual(await (await f.callback(sessionId, { ...payload, status: 2 })).json(), { error: 0 })
    assert.equal(f.store.requireSummary(f.artifact.id).currentVersion, 2)
    const state = await (await f.app.request(`${f.root}/${sessionId}`)).json()
    assert.equal(state.status, 'saved')
    assert.equal(state.savedVersion, 2)
    assert.equal(state.closed, true)
    assert.ok(f.store.blobs.read(f.store.requireVersion(f.artifact.id, 2).files['deck.pptx'].sha256).equals(result))
    assert.ok(f.store.blobs.read(f.store.requireVersion(f.artifact.id, 1).files['deck.pptx'].sha256).equals(pptx('original')))
  } finally { f.close() }
})

test('callbacks cannot use unsigned fields, fetch another origin or overwrite concurrent publishes', async () => {
  const f = fixture()
  try {
    const { sessionId } = await (await f.start()).json()
    const bad = await f.app.request(`/api/office/callback/${sessionId}`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ key: sessionId, status: 2, url: `${docsUrl}/result` }) })
    assert.equal(bad.status, 401)
    const count = downloads
    assert.equal((await f.callback(sessionId, { key: sessionId, status: 2, url: 'http://other.test/secret' })).status, 403)
    assert.equal(downloads, count)
    f.store.update(f.artifact.id, { baseVersion: 1, files: { 'other.txt': { content: Buffer.from('newer') } } })
    const signed = { key: sessionId, status: 2, url: `${docsUrl}/result.pptx` }
    assert.deepEqual(await (await f.callback(sessionId, signed, { ...signed, url: 'http://other.test/secret' })).json(), { error: 0 })
    assert.equal(f.store.requireSummary(f.artifact.id).currentVersion, 2)
    const state = await (await f.app.request(`${f.root}/${sessionId}`)).json()
    assert.equal(state.status, 'conflict')
    const download = await f.app.request(state.downloadUrl)
    assert.ok(Buffer.from(await download.arrayBuffer()).equals(result))
    await f.callback(sessionId, signed)
    assert.equal((await (await f.app.request(`${f.root}/${sessionId}`)).json()).status, 'conflict')
  } finally { f.close() }
})

test('invalid PPTX and redirected downloads cannot publish and report a recoverable save error', async () => {
  const f = fixture()
  try {
    const { sessionId } = await (await f.start()).json()
    assert.equal((await f.callback(sessionId, { key: sessionId, status: 2, url: `${docsUrl}/redirect` })).status, 400)
    const previous = result
    result = Buffer.from('<html>not a PowerPoint file</html>')
    assert.equal((await f.callback(sessionId, { key: sessionId, status: 2, url: `${docsUrl}/result.pptx` })).status, 400)
    result = previous
    assert.equal(f.store.requireSummary(f.artifact.id).currentVersion, 1)
    assert.equal((await (await f.app.request(`${f.root}/${sessionId}`)).json()).status, 'error')
    await f.callback(sessionId, { key: sessionId, status: 2, url: `${docsUrl}/result.pptx` })
    assert.equal(f.store.requireSummary(f.artifact.id).currentVersion, 2)
  } finally { f.close() }
})

test('capabilities granted during an edit prevent callback writes and preserve the result', async () => {
  const f = fixture()
  try {
    const { sessionId } = await (await f.start()).json()
    f.store.update(f.artifact.id, { baseVersion: 1, capabilities: ['uploads'] })
    await f.callback(sessionId, { key: sessionId, status: 2, url: `${docsUrl}/result.pptx` })
    assert.equal(f.store.requireSummary(f.artifact.id).currentVersion, 2)
    const state = await (await f.app.request(`${f.root}/${sessionId}`)).json()
    assert.equal(state.status, 'conflict')
    assert.ok(state.downloadUrl)
  } finally { f.close() }
})

test('a no-change disconnect followed by reconnect resumes the edit session', async () => {
  const f = fixture()
  try {
    const { sessionId } = await (await f.start()).json()
    await f.callback(sessionId, { key: sessionId, status: 4 })
    await f.callback(sessionId, { key: sessionId, status: 1 })
    const state = await (await f.app.request(`${f.root}/${sessionId}`)).json()
    assert.equal(state.closed, false)
    assert.equal((await f.app.request(`${f.root}/${sessionId}/config`)).status, 200)
  } finally { f.close() }
})

test('an Office save error preserves the supplied recovery file without publishing it', async () => {
  const f = fixture()
  try {
    const { sessionId } = await (await f.start()).json()
    for (const status of [3, 7]) {
      await f.callback(sessionId, { key: sessionId, status, url: `${docsUrl}/recovery.pptx` })
      const state = await (await f.app.request(`${f.root}/${sessionId}`)).json()
      assert.equal(state.status, 'error')
      assert.ok(state.downloadUrl)
      assert.ok(Buffer.from(await (await f.app.request(state.downloadUrl)).arrayBuffer()).equals(result))
      assert.equal(f.store.requireSummary(f.artifact.id).currentVersion, 1)
    }
    await f.callback(sessionId, { key: sessionId, status: 6, url: `${docsUrl}/recovery.pptx` })
    assert.equal(f.store.requireSummary(f.artifact.id).currentVersion, 2)
  } finally { f.close() }
})

test('leftovers from a crashed start or stray files do not block new edits', async () => {
  const f = fixture()
  try {
    const edits = join(f.config.dataDir, 'office-edits')
    const leftover = join(edits, 'A'.repeat(43))
    // A start interrupted while copying the source leaves a folder without session.json.
    mkdirSync(leftover)
    writeFileSync(join(leftover, 'source.pptx.tmp'), 'partial')
    writeFileSync(join(edits, '.DS_Store'), '')
    assert.equal((await f.start()).status, 201)
    assert.equal(existsSync(leftover), false)
  } finally { f.close() }
})

test('sessions Office never opened are cleaned after a day; opened ones are kept', async () => {
  const f = fixture()
  try {
    const age = (sid: string) => {
      const file = join(f.config.dataDir, 'office-edits', sid, 'session.json')
      writeFileSync(file, JSON.stringify({ ...JSON.parse(readFileSync(file, 'utf8')), createdAt: Date.now() - 2 * 86_400_000 }))
    }
    const abandoned = (await (await f.start()).json()).sessionId
    const opened = (await (await f.start()).json()).sessionId
    const cfg = await (await f.app.request(`${f.root}/${opened}/config`)).json()
    assert.equal((await f.app.request(cfg.document.url)).status, 200)
    age(abandoned)
    age(opened)
    await f.start()
    assert.equal((await f.app.request(`${f.root}/${abandoned}`)).status, 404)
    assert.equal((await f.app.request(`${f.root}/${opened}`)).status, 200)
  } finally { f.close() }
})

test('a malformed callback body is rejected as unauthorized, not a server error', async () => {
  const f = fixture()
  try {
    const { sessionId } = await (await f.start()).json()
    const response = await f.app.request(`/api/office/callback/${sessionId}`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: '{not json' })
    assert.equal(response.status, 401)
  } finally { f.close() }
})

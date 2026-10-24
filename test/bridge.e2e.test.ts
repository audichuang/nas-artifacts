import assert from 'node:assert/strict'
import { existsSync, mkdtempSync, rmSync } from 'node:fs'
import { homedir, tmpdir } from 'node:os'
import { join } from 'node:path'
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

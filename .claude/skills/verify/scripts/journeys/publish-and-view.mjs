#!/usr/bin/env node
// Journey: publish an HTML site, a Markdown page and a Markdown deck with the CLI, then read them in the viewer.
// Usage: node .claude/skills/verify/scripts/journeys/publish-and-view.mjs   (instance from launch.sh must be up)
import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { V, checks, cli, contentFrame, launch, log, save, scratch, shot } from '../browser.mjs'

const F = 'publish-and-view'
const c = checks(F)
// Unique per run, so re-running on the same data dir still finds exactly one card.
const needle = `needle-${Date.now().toString(36)}`
const title = `Verify page ${needle}`
const site = scratch('pv-site')
writeFileSync(join(site, 'index.html'), `<!doctype html><html><head><title>Verify page</title></head><body>
<h1 id="t">Verify page</h1><p>${needle}</p><p id="leak">pending</p>
<a id="next" href="two.html">second page</a>
<script>
fetch('${V.app}/api/artifacts').then(() => { document.getElementById('leak').textContent = 'leaked' }, () => { document.getElementById('leak').textContent = 'blocked' })
</script></body></html>`)
writeFileSync(join(site, 'two.html'), '<!doctype html><title>two</title><h1 id="two">Second page</h1>')
const notes = join(scratch('pv-md'), 'notes.md')
writeFileSync(notes, '# Verify notes\n\nIntro.\n\n## Alpha\n\na\n\n## Beta\n\nb\n\n## Gamma\n\nc\n')
const deck = join(scratch('pv-deck'), 'talk.slides.md')
writeFileSync(deck, '# Slide one\n\nfirst\n\n---\n\n# Slide two\n\nsecond\n')

let browser
try {
  const checked = cli(F, ['check', site])
  c.ok(checked.code === 0 || /connect|fetch/i.test(checked.stdout), 'artifact check ran (warnings name what the sandbox blocks)', `${checked.code} ${checked.stdout.trim()}`)
  const pub = cli(F, ['publish', site, '--new', '--title', title, '--icon', '🧪', '--description', 'verify journey', '--json'])
  const art = JSON.parse(pub.stdout)
  c.ok(pub.code === 0 && art.currentVersion === 1 && art.url === `${V.app}/a/${art.id}` && art.contentUrl?.startsWith(`${V.content}/a/${art.id}/`), 'publish -> v1, viewer url on app origin, contentUrl on content origin', `${art.url} ${art.contentUrl}`)
  const md = JSON.parse(cli(F, ['publish', notes, '--new', '--json']).stdout)
  const dk = JSON.parse(cli(F, ['publish', deck, '--new', '--json']).stdout)
  const list = cli(F, ['list'])
  c.ok(list.code === 0 && list.stdout.includes(art.id) && list.stdout.includes(md.id) && list.stdout.includes(dk.id), 'artifact list shows all three')

  browser = await launch()
  const page = await browser.newPage({ viewport: { width: 1280, height: 860 } })
  await page.goto(`${V.app}/`)
  await page.locator('.card', { hasText: title }).waitFor({ timeout: 5000 })
  await shot(page, F, '01-gallery')
  await page.locator('#search').fill(needle)
  await page.waitForFunction(() => document.querySelectorAll('.card').length === 1, undefined, { timeout: 5000 }).catch(() => {})
  c.ok((await page.locator('.card', { hasText: title }).count()) === 1 && (await page.locator('.card').count()) === 1, 'gallery search for page text finds exactly that card', await page.locator('.card').count())
  await shot(page, F, '02-gallery-search')

  const frameResponses = []
  page.on('response', (r) => { if (r.url().startsWith(`${V.content}/a/${art.id}/`)) frameResponses.push(r) })
  const [viewerResponse] = await Promise.all([page.waitForNavigation(), page.locator('.card', { hasText: title }).click()])
  c.ok(page.url().startsWith(`${V.app}/a/${art.id}`), 'clicking the card opens the viewer', page.url())
  await page.frameLocator('#frame').locator('#t').waitFor({ timeout: 5000 })
  const frameSrc = (viewerResponse.headers()['content-security-policy'] ?? '').match(/frame-src [^;]*/)?.[0] ?? ''
  c.ok(frameSrc === `frame-src ${V.content}/a/${art.id}/`, 'viewer CSP frame-src is scoped to <content>/a/<id>/', frameSrc)
  const frame = contentFrame(page, V.content)
  c.ok((await frame.evaluate(() => origin)) === 'null', 'page frame runs with an opaque origin (sandbox without allow-same-origin)')
  await page.waitForTimeout(1500)
  const leak = await frame.locator('#leak').textContent()
  c.ok(leak === 'blocked', 'in-page fetch to the app API was rejected (marker reads blocked)', leak)
  const doc = frameResponses.find((r) => /\/v\/1\/(index\.html)?$/.test(r.url()))
  const csp = doc?.headers()['content-security-policy'] ?? ''
  save(F, 'headers.json', { viewer: viewerResponse.headers(), frameDocument: doc?.headers() })
  c.ok(/sandbox [^;]*allow-scripts/.test(csp) && !/allow-same-origin/.test(csp) && /connect-src 'none'/.test(csp), "frame CSP: sandbox without allow-same-origin, connect-src 'none'", csp)
  await shot(page, F, '03-viewer')
  await page.frameLocator('#frame').locator('#next').click()
  await page.frameLocator('#frame').locator('#two').waitFor({ timeout: 5000 })
  await page.waitForFunction(() => location.search.includes('path='), undefined, { timeout: 5000 })
  c.ok(new URL(page.url()).searchParams.get('path') === 'two.html', 'relative link moves the viewer to ?path=two.html', page.url())
  await shot(page, F, '04-second-page')

  const top = await browser.newPage()
  await top.goto(`${V.content}/a/${art.id}/v/1/index.html`)
  await top.waitForURL((u) => u.origin === V.app, { timeout: 5000 }).catch(() => {})
  c.ok(top.url().startsWith(`${V.app}/a/${art.id}`), 'content URL opened top-level is redirected into the viewer', top.url())
  await top.close()

  const readBack = cli(F, ['read', art.id], { quiet: true })
  c.ok(readBack.stdout === readFileSync(join(site, 'index.html'), 'utf8'), 'artifact read returns the published bytes')

  await page.goto(md.url)
  await page.frameLocator('#frame').locator('h1', { hasText: 'Verify notes' }).waitFor({ timeout: 5000 })
  c.ok((await page.frameLocator('#frame').locator('nav.toc li').count()) === 3, 'Markdown renders with a 3-entry table of contents')
  await shot(page, F, '05-markdown')

  await page.goto(dk.url)
  const slides = page.frameLocator('#frame')
  await slides.locator('.slide.active', { hasText: 'Slide one' }).waitFor({ timeout: 5000 })
  await shot(page, F, '06-deck-slide1')
  await page.locator('#title').click()
  await page.keyboard.press('ArrowRight')
  await slides.locator('.slide.active', { hasText: 'Slide two' }).waitFor({ timeout: 5000 })
  c.ok(true, 'ArrowRight in the viewer turns the deck to slide two')
  await shot(page, F, '07-deck-slide2')
  save(F, 'ids.json', { html: art.id, markdown: md.id, deck: dk.id })
  c.finish()
} catch (error) {
  c.finish(error)
} finally {
  await browser?.close()
}

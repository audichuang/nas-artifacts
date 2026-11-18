#!/usr/bin/env node
// Journey: a page with db + uploads capabilities, live across two viewer tabs and the CLI; a page without them is refused.
// Usage: node .claude/skills/verify/scripts/journeys/capabilities-live-page.mjs
import { writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { V, checks, cli, contentFrame, launch, log, save, scratch, shot } from '../browser.mjs'

const F = 'capabilities-live-page'
const c = checks(F)
const nonce = Date.now().toString(36)
const COUNTER = `<!doctype html><html><head><title>Counter</title></head><body>
<h1>Counter ${nonce}</h1><p>count: <b id="count">?</b></p><p id="err"></p><p id="upload"></p>
<button id="bump">+1</button> <button id="up">upload note.txt</button>
<script>
const show = (e) => { document.getElementById('err').textContent = e.code || e.message }
async function refresh() {
  try { const doc = await artifact.db.get('c', 'main'); document.getElementById('count').textContent = String(doc?.data?.n ?? 0) } catch (e) { show(e) }
}
try { artifact.db.subscribe(refresh) } catch (e) { show(e) }
refresh()
document.getElementById('bump').onclick = async () => {
  try { const doc = await artifact.db.get('c', 'main'); await artifact.db.set('c', 'main', { n: (doc?.data?.n ?? 0) + 1 }) } catch (e) { show(e) }
}
document.getElementById('up').onclick = async () => {
  try { const asset = await artifact.uploads.put(new Blob(['hello from the page']), 'note.txt'); document.getElementById('upload').textContent = asset.url } catch (e) { show(e) }
}
</script></body></html>`

let browser
try {
  const withCaps = scratch(`cap-${nonce}`)
  const without = scratch(`nocap-${nonce}`)
  writeFileSync(join(withCaps, 'index.html'), COUNTER)
  writeFileSync(join(without, 'index.html'), COUNTER)
  const art = JSON.parse(cli(F, ['publish', withCaps, '--new', '--cap', 'db,uploads', '--title', `Counter ${nonce}`, '--json']).stdout)
  c.ok(art.capabilities?.join(',') === 'db,uploads', 'published with capabilities db,uploads', art.capabilities)
  const bare = JSON.parse(cli(F, ['publish', without, '--new', '--title', `No caps ${nonce}`, '--json']).stdout)

  browser = await launch()
  const tab1 = await browser.newPage({ viewport: { width: 900, height: 600 } })
  const tab2 = await browser.newPage({ viewport: { width: 900, height: 600 } })
  await tab1.goto(art.url)
  await tab2.goto(art.url)
  const count = (tab) => tab.frameLocator('#frame').locator('#count')
  await count(tab1).filter({ hasText: /^0$/ }).waitFor({ timeout: 5000 })
  await count(tab2).filter({ hasText: /^0$/ }).waitFor({ timeout: 5000 })
  c.ok((await contentFrame(tab1, V.content).evaluate(() => origin)) === 'null', 'page frame origin is null: capabilities come only through the bridge')
  await shot(tab2, F, '01-tab2-before')
  await tab1.frameLocator('#frame').locator('#bump').click()
  await count(tab2).filter({ hasText: /^1$/ }).waitFor({ timeout: 5000 })
  c.ok(true, '+1 clicked in tab 1 shows 1 in tab 2 without a reload')
  await shot(tab2, F, '02-tab2-after-bump')
  const got = JSON.parse(cli(F, ['db', 'get', art.id, 'c', 'main']).stdout)
  c.ok(got.data?.n === 1, 'CLI db get agrees with the page (n = 1)', JSON.stringify(got))

  const set = cli(F, ['db', 'set', art.id, 'c', 'main', '{"n":41}'])
  c.ok(set.code === 0, 'CLI db set n = 41', set.stdout.trim())
  await count(tab1).filter({ hasText: /^41$/ }).waitFor({ timeout: 5000 })
  await count(tab2).filter({ hasText: /^41$/ }).waitFor({ timeout: 5000 })
  c.ok(true, 'both tabs show 41 live after the CLI write')
  await shot(tab1, F, '03-tab1-after-cli-set')

  await tab1.frameLocator('#frame').locator('#up').click()
  await tab1.frameLocator('#frame').locator('#upload').filter({ hasText: `/a/${art.id}/assets/` }).waitFor({ timeout: 5000 })
  const uploadedUrl = (await tab1.frameLocator('#frame').locator('#upload').textContent()).trim()
  const uploaded = await fetch(uploadedUrl)
  const uploadedBody = await uploaded.text()
  c.ok(uploaded.status === 200 && uploadedBody === 'hello from the page', 'the URL the page got back serves the uploaded bytes (200, "hello from the page")', `${uploaded.status} ${JSON.stringify(uploadedBody.slice(0, 60))} ${uploadedUrl}`)
  const assets = cli(F, ['asset', 'list', art.id])
  const assetLine = assets.stdout.split('\n').find((line) => line.includes(uploadedUrl)) ?? ''
  c.ok(assets.code === 0 && /\bpage\b/.test(assetLine) && assetLine.trim().endsWith(uploadedUrl), 'asset list shows the same URL with source page', assets.stdout.trim())
  await shot(tab1, F, '04-uploaded')

  const tab3 = await browser.newPage({ viewport: { width: 900, height: 600 } })
  await tab3.goto(bare.url)
  await tab3.frameLocator('#frame').locator('#err').filter({ hasText: 'capability_not_declared' }).waitFor({ timeout: 5000 })
  c.ok(true, 'page without declared capabilities gets capability_not_declared from the bridge')
  await shot(tab3, F, '05-no-capability')
  // The server checks again: a viewer-style call straight to the API (no token) is refused too.
  const direct = await fetch(`${V.app}/api/artifacts/${bare.id}/db/c/main`, { method: 'PUT', headers: { 'content-type': 'application/json', 'x-artifact-client': 'viewer' }, body: JSON.stringify({ data: { n: 1 } }) })
  const directBody = await direct.json()
  save(F, 'direct-api-call.json', { request: `PUT ${V.app}/api/artifacts/${bare.id}/db/c/main (x-artifact-client: viewer, no token)`, status: direct.status, body: directBody })
  c.ok(direct.status === 403 && directBody.error?.code === 'capability_not_declared', 'server refuses a viewer db write on the undeclared page with 403 capability_not_declared', direct.status)

  const pub = cli(F, ['visibility', art.id, 'public'])
  c.ok(pub.code !== 0 && /capabilit/i.test(pub.stderr), 'visibility public on a page with capabilities fails', pub.stderr.trim())
  const still = JSON.parse(cli(F, ['get', art.id, '--json']).stdout)
  c.ok((still.visibility ?? 'private') === 'private', 'it stays private', still.visibility)
  save(F, 'ids.json', { withCapabilities: art.id, without: bare.id })
  c.finish()
} catch (error) {
  c.finish(error)
} finally {
  await browser?.close()
}

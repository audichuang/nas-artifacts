#!/usr/bin/env node
// Journey: a reader pins an artifact from the viewer bar (☆ → ★). The pin is kept on the server: it survives a
// reload, another browser (device) sees the gallery list it first with the ★ mark, and CLI get reports it.
// The gallery itself has no pin control; it only shows the mark and sorts pinned first.
// Usage: node .claude/skills/verify/scripts/journeys/pins.mjs
import { writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { V, checks, cli, launch, save, scratch, shot } from '../browser.mjs'

const F = 'pins'
const c = checks(F)
const nonce = Date.now().toString(36)

function publish(name) {
  const dir = scratch(`pins-${name}-${nonce}`)
  writeFileSync(join(dir, 'index.html'), `<!doctype html><h1 id="t">${name} ${nonce}</h1>`)
  return JSON.parse(cli(F, ['publish', dir, '--new', '--title', `Pin ${name} ${nonce}`, '--json']).stdout)
}
/** Gallery cards for this run's nonce, in order: [{ title, pinned }]. Filtering by nonce: ordering spans the data dir. */
async function galleryOrder(page) {
  await page.goto(`${V.app}/`)
  await page.locator('#search').fill(nonce)
  await page.locator('#status', { hasText: '共 2 個' }).waitFor({ timeout: 5000 })
  return page.locator('#list .card').evaluateAll((cards) => cards.map((card) => ({ title: card.querySelector('.card-title').textContent, pinned: Boolean(card.querySelector('.pinned-mark')) })))
}
const show = (order) => order.map((x) => `${x.title}${x.pinned ? ' ★' : ''}`).join(' | ')
const pinState = async (page) => ({ text: (await page.locator('#pin').textContent()).trim(), pressed: await page.locator('#pin').getAttribute('aria-pressed') })

let browser
try {
  const a = publish('A')
  await new Promise((resolve) => setTimeout(resolve, 1100))
  const b = publish('B')
  browser = await launch()
  const device1 = await browser.newContext({ viewport: { width: 1200, height: 800 } })
  const gallery1 = await device1.newPage()
  const before = await galleryOrder(gallery1)
  c.ok(before.map((x) => x.title).join() === `Pin B ${nonce},Pin A ${nonce}` && before.every((x) => !x.pinned), 'gallery before pinning: newer B first, then A, no ★ mark', show(before))
  await shot(gallery1, F, '01-gallery-before')

  const viewer = await device1.newPage()
  await viewer.goto(a.url)
  await viewer.frameLocator('#frame').locator('#t').waitFor({ timeout: 5000 })
  const off = await pinState(viewer)
  c.ok(off.text === '☆' && off.pressed === 'false', 'viewer bar #pin reads ☆, aria-pressed false', JSON.stringify(off))
  await viewer.locator('#pin').click()
  await viewer.locator('#pin[aria-pressed="true"]').waitFor({ timeout: 5000 })
  const on = await pinState(viewer)
  c.ok(on.text === '★' && on.pressed === 'true', 'clicking #pin turns it ★, aria-pressed true', JSON.stringify(on))
  await viewer.reload()
  await viewer.frameLocator('#frame').locator('#t').waitFor({ timeout: 5000 })
  const reloaded = await pinState(viewer)
  c.ok(reloaded.text === '★' && reloaded.pressed === 'true', 'after a reload #pin is still ★', JSON.stringify(reloaded))
  await shot(viewer, F, '02-viewer-pinned')

  const device2 = await browser.newContext({ viewport: { width: 1200, height: 800 } })
  const gallery2 = await device2.newPage()
  const pinnedOrder = await galleryOrder(gallery2)
  c.ok(pinnedOrder.map((x) => x.title).join() === `Pin A ${nonce},Pin B ${nonce}`, 'another browser: the gallery lists pinned A first although B is newer', show(pinnedOrder))
  c.ok(pinnedOrder[0]?.pinned === true && pinnedOrder[1]?.pinned === false, "another browser: A's card has the ★ mark (.pinned-mark), B's does not", show(pinnedOrder))
  await shot(gallery2, F, '03-other-browser-gallery')
  c.ok(JSON.parse(cli(F, ['get', a.id, '--json']).stdout).pinned === true, 'CLI get A --json: pinned true')

  await viewer.locator('#pin').click()
  await viewer.locator('#pin[aria-pressed="false"]').waitFor({ timeout: 5000 })
  await viewer.reload()
  await viewer.frameLocator('#frame').locator('#t').waitFor({ timeout: 5000 })
  const unpinned = await pinState(viewer)
  c.ok(unpinned.text === '☆' && unpinned.pressed === 'false', 'unpinned: after a reload #pin reads ☆ again', JSON.stringify(unpinned))
  const after = await galleryOrder(gallery2)
  c.ok(after.map((x) => x.title).join() === `Pin B ${nonce},Pin A ${nonce}` && after.every((x) => !x.pinned), 'another browser: order back to B, A with no ★ mark', show(after))
  c.ok(JSON.parse(cli(F, ['get', a.id, '--json']).stdout).pinned === false, 'CLI get A --json: pinned false')
  await shot(gallery2, F, '04-gallery-after-unpin')
  save(F, 'ids.json', { a: a.id, b: b.id })
  for (const id of [a.id, b.id]) cli(F, ['delete', id, '--yes'], { quiet: true })
  c.finish()
} catch (error) {
  c.finish(error)
} finally {
  await browser?.close()
}

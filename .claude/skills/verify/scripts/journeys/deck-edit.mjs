#!/usr/bin/env node
// Journey: a reader edits a capability-free Markdown deck in the tailnet viewer (編輯 → type → 儲存). The typing
// previews live without a version; saving makes a version labelled 瀏覽器編輯 that a second tab, CLI get and CLI
// checkout all see. A stale draft never overwrites a newer version, and a deck with capabilities offers no edit.
// Usage: node .claude/skills/verify/scripts/journeys/deck-edit.mjs
import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { V, checks, cli, launch, log, save, scratch, shot } from '../browser.mjs'

const F = 'deck-edit'
const c = checks(F)
const nonce = Date.now().toString(36)
const SOURCE = `# Deck ${nonce}\n\nfirst slide\n\n---\n\n# Slide two\n\nsecond\n`
const EDITED = `# Edited ${nonce}\n\nfirst slide, changed in the browser\n\n---\n\n# Slide two\n\nsecond\n`
const CLI_V3 = `# From the CLI ${nonce}\n\nv3\n\n---\n\n# Slide two\n\nsecond\n`
const getJson = (id) => JSON.parse(cli(F, ['get', id, '--json']).stdout)

let browser
try {
  const dir = scratch(`deck-${nonce}`)
  writeFileSync(join(dir, 'talk.slides.md'), SOURCE)
  const art = JSON.parse(cli(F, ['publish', join(dir, 'talk.slides.md'), '--new', '--title', `Deck ${nonce}`, '--json']).stdout)
  c.ok(art.entry === 'talk.slides.md' && art.currentVersion === 1, 'published a capability-free deck, entry talk.slides.md, v1', `${art.entry} v${art.currentVersion}`)

  browser = await launch()
  const tab1 = await browser.newPage({ viewport: { width: 1300, height: 800 } })
  const tab2 = await browser.newPage({ viewport: { width: 1300, height: 800 } })
  await tab1.goto(art.url)
  await tab2.goto(art.url)
  const active = (tab) => tab.frameLocator('#frame').locator('.slide.active')
  await active(tab1).filter({ hasText: `Deck ${nonce}` }).waitFor({ timeout: 5000 })
  await active(tab2).filter({ hasText: `Deck ${nonce}` }).waitFor({ timeout: 5000 })

  await tab1.locator('#edit').waitFor({ state: 'visible', timeout: 5000 })
  c.ok(await tab1.locator('#edit').isVisible(), '編輯 (#edit) is shown on a capability-free deck')
  await tab1.locator('#edit').click()
  await tab1.locator('#editor').waitFor({ state: 'visible', timeout: 5000 })
  const loaded = await tab1.locator('#editor-source').inputValue()
  c.ok(loaded === SOURCE, 'the editor holds the published source byte for byte', JSON.stringify(loaded.slice(0, 60)))
  const status1 = (await tab1.locator('#editor-status').textContent()).trim()
  c.ok(status1 === '根據 v1' && (await tab1.locator('#save-deck').isDisabled()), '#editor-status 根據 v1 and 儲存 disabled before typing', status1)
  await shot(tab1, F, '01-editor-open')

  await tab1.locator('#editor-source').fill(EDITED)
  await tab1.locator('#editor-status', { hasText: '根據 v1 · 未儲存' }).waitFor({ timeout: 3000 })
  c.ok(await tab1.locator('#save-deck').isEnabled(), 'typing: status 根據 v1 · 未儲存 and 儲存 enabled')
  // fill() leaves the cursor at the end, so the preview shows the last slide; the edited one is in the deck, not on show.
  await tab1.frameLocator('#frame').locator('.slide', { hasText: `Edited ${nonce}` }).first().waitFor({ state: 'attached', timeout: 5000 })
  c.ok(true, 'live preview: the frame deck now holds the typed heading before saving')
  c.ok(getJson(art.id).currentVersion === 1, 'the preview made no version (CLI get still v1)')
  await shot(tab1, F, '02-live-preview')

  await tab1.locator('#save-deck').click()
  await tab1.locator('#editor-status', { hasText: '已存成 v2' }).waitFor({ timeout: 5000 })
  c.ok(true, '儲存: #editor-status 已存成 v2')
  await active(tab2).filter({ hasText: `Edited ${nonce}` }).waitFor({ timeout: 5000 })
  c.ok(true, 'tab 2 shows the edited slide without a reload')
  await shot(tab2, F, '03-tab2-after-save')
  const got = getJson(art.id)
  const v2 = got.versions.find((v) => v.version === 2)
  c.ok(got.currentVersion === 2 && v2?.label === '瀏覽器編輯', 'CLI get: current v2, labelled exactly 瀏覽器編輯 (no viewer name in a fresh browser)', got.versions.map((v) => `v${v.version}:${v.label}`).join(' '))
  const optionText = await tab1.locator('#version option[value="2"]').textContent()
  c.ok(optionText.includes('v2 · 瀏覽器編輯'), 'the version picker names v2 · 瀏覽器編輯', optionText)
  const out = scratch(`deck-checkout-${nonce}`)
  const co = cli(F, ['checkout', art.id, join(out, 'co'), '--force'])
  const checkedOut = readFileSync(join(out, 'co', 'talk.slides.md'), 'utf8')
  c.ok(co.code === 0 && checkedOut === EDITED, 'CLI checkout writes talk.slides.md byte-equal to what was typed', `${co.code} ${JSON.stringify(checkedOut.slice(0, 60))}`)

  // A draft based on v2, while the CLI publishes v3 from the checkout: saving must not overwrite v3.
  await tab1.locator('#editor-source').fill(`${EDITED}\n---\n\n# Stale draft ${nonce}\n`)
  await tab1.locator('#editor-status', { hasText: '根據 v2 · 未儲存' }).waitFor({ timeout: 3000 })
  writeFileSync(join(out, 'co', 'talk.slides.md'), CLI_V3)
  const v3 = cli(F, ['publish', join(out, 'co')])
  c.ok(v3.code === 0 && v3.stdout.includes('v3'), 'CLI publishes v3 from the checkout', v3.stdout.trim())
  await tab1.waitForTimeout(500)
  tab1.once('dialog', (dialog) => { log(F, `conflict dialog: ${dialog.message().split('\n')[0]}`); dialog.dismiss() })
  await tab1.locator('#save-deck').click()
  await tab1.locator('#editor-status', { hasText: '別人已存了新版本' }).waitFor({ timeout: 5000 })
  c.ok(true, 'saving the stale draft: #editor-status 別人已存了新版本 (kept the draft)')
  await shot(tab1, F, '04-conflict')
  const after = getJson(art.id)
  const read = cli(F, ['read', art.id, 'talk.slides.md'])
  c.ok(after.currentVersion === 3 && read.stdout === CLI_V3, "CLI get still v3 and read returns the CLI's text (nothing overwritten)", `v${after.currentVersion} ${JSON.stringify(read.stdout.slice(0, 40))}`)

  // The same deck with a capability: no edit in the viewer, and the server refuses a direct save.
  const capDir = scratch(`deck-cap-${nonce}`)
  writeFileSync(join(capDir, 'talk.slides.md'), SOURCE)
  const cap = JSON.parse(cli(F, ['publish', join(capDir, 'talk.slides.md'), '--new', '--cap', 'db', '--title', `Cap deck ${nonce}`, '--json']).stdout)
  const tab3 = await browser.newPage({ viewport: { width: 1300, height: 800 } })
  await tab3.goto(cap.url)
  await active(tab3).filter({ hasText: `Deck ${nonce}` }).waitFor({ timeout: 5000 })
  await tab3.waitForTimeout(300)
  c.ok(await tab3.locator('#edit').isHidden(), 'a deck with --cap db shows no 編輯', await tab3.locator('#edit').getAttribute('hidden'))
  await shot(tab3, F, '05-cap-deck-no-edit')
  const put = await fetch(`${V.app}/api/artifacts/${cap.id}/slides`, { method: 'PUT', headers: { 'content-type': 'application/json', 'x-artifact-client': 'viewer' }, body: JSON.stringify({ markdown: EDITED, baseVersion: 1 }) })
  const putBody = await put.json()
  save(F, 'direct-slides-put.json', { request: `PUT ${V.app}/api/artifacts/${cap.id}/slides (x-artifact-client: viewer, no token)`, status: put.status, body: putBody })
  c.ok(put.status === 403 && putBody.error?.code === 'deck_not_editable', 'server refuses a viewer save of the capabilities deck: 403 deck_not_editable', `${put.status} ${putBody.error?.code}`)
  c.ok(getJson(cap.id).currentVersion === 1, 'CLI get: the capabilities deck is still v1')
  save(F, 'ids.json', { deck: art.id, capDeck: cap.id })
  c.finish()
} catch (error) {
  c.finish(error)
} finally {
  await browser?.close()
}

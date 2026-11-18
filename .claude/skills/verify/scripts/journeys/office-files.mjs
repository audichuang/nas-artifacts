#!/usr/bin/env node
// Journey: publish .pptx, .docx and .xlsx with the CLI; the viewer draws them (deck, document, sheet tabs);
// download, `read --out` and `export` return the published bytes. VERIFY_OFFICE_FILE=<path> adds a real file.
// Usage: node .claude/skills/verify/scripts/journeys/office-files.mjs
import { createHash } from 'node:crypto'
import { existsSync, readFileSync } from 'node:fs'
import { basename, join } from 'node:path'
import { V, checks, cli, contentFrame, launch, log, save, scratch, shot } from '../browser.mjs'
import { makeOffice } from '../make-office.mjs'

const F = 'office-files'
const c = checks(F)
const nonce = Date.now().toString(36)
const work = scratch(`office-${nonce}`)
const sha = (path) => createHash('sha256').update(readFileSync(path)).digest('hex')

let browser
try {
  const files = await makeOffice(work)
  const ids = {}
  for (const [kind, path] of Object.entries(files)) {
    const out = cli(F, ['publish', path, '--new', '--title', `Office ${kind} ${nonce}`, '--json'])
    ids[kind] = JSON.parse(out.stdout).id
    c.ok(out.code === 0, `publish ${basename(path)}`, ids[kind])
  }
  browser = await launch()
  const page = await browser.newPage({ viewport: { width: 1280, height: 860 }, acceptDownloads: true })
  const frame = () => page.frameLocator('#frame')

  // Whether the file in the frame was drawn, checked from what src/pptx.ts, src/docx.ts and src/xlsx.ts emit.
  // It first waits for a render root or a failure page (a timeout is a FAIL, never a count of 0), then counts
  // the whole-file failure boxes and the per-part "cannot draw" markers of all three formats, and reads the
  // viewer's page-error badge (#errors), which the renderers' failure script trips.
  const rendered = async (kind) => {
    const roots = '.slide, .dx-doc, .xl-tabs, .xl-fail, .xl-empty'
    const drawn = await frame().locator(roots).first().waitFor({ state: 'attached', timeout: 15_000 }).then(() => true, () => false)
    c.ok(drawn, `${kind}: the frame shows a rendered document or a failure page within 15 s`)
    if (!drawn) return
    const fail = await frame().locator('.dx-fail, .xl-fail, .xl-empty, .pp-missing[role=alert]').allInnerTexts()
    c.ok(fail.length === 0, `${kind}: no whole-file failure page (.dx-fail, .xl-fail, .xl-empty, alert .pp-missing)`, fail.join(' | ').replace(/\s+/g, ' '))
    // .xl-note and .dx-notice also carry "not shown / unsupported / cut" notes; an empty sheet's note is not a failure.
    const markers = '.pp-missing, .dx-unsup, .dx-cut, .dx-notice, .xl-cellmiss, .xl-note, .xl-shape-note'
    const missing = (await frame().locator(markers).allInnerTexts()).filter((text) => text.trim() !== '這個工作表是空的。')
    c.ok(missing.length === 0, `${kind}: no "cannot draw" markers (${markers})`, `${missing.length}: ${missing.slice(0, 5).join(' | ')}`)
    // Secondary signal only: a 2 s window for the error report to arrive, so its absence proves less than the markers.
    await page.locator('#errors').waitFor({ state: 'visible', timeout: 2000 }).catch(() => {})
    const errors = (await page.locator('#errors').isVisible()) ? await page.locator('#error-count').textContent() : '0'
    c.ok(errors === '0', `${kind}: viewer reports no page errors (#errors badge hidden)`, `${errors} page errors`)
  }

  const download = async (kind, path) => {
    const button = page.locator('#export-pptx')
    await button.waitFor({ state: 'visible', timeout: 5000 })
    const label = (await button.textContent()).trim()
    const [file] = await Promise.all([page.waitForEvent('download'), button.click()])
    const saved = join(work, `downloaded-${kind}-${file.suggestedFilename()}`)
    await file.saveAs(saved)
    c.ok(sha(saved) === sha(path), `${kind}: viewer download button (${label}) returns identical bytes (sha256)`, sha(saved).slice(0, 16))
    const out = join(work, `read-${kind}-${basename(path)}`)
    const read = cli(F, ['read', ids[kind], '--out', out])
    c.ok(read.code === 0 && sha(out) === sha(path), `${kind}: artifact read --out returns identical bytes (sha256)`)
    c.ok(await page.locator('#edit').isHidden(), `${kind}: no 編輯 button (ONLYOFFICE not configured)`)
  }

  // pptx: a deck drawn from the file; slide two's picture comes out of the file under /~office/…/ppt/.
  await page.goto(`${V.app}/a/${ids.pptx}`)
  await frame().locator('.slide.active', { hasText: 'Verify deck slide one' }).waitFor({ timeout: 8000 })
  await rendered('pptx')
  await shot(page, F, '01-pptx-slide1')
  await page.locator('#title').click()
  await page.keyboard.press('ArrowRight')
  const slide2 = frame().locator('.slide.active', { hasText: 'Slide two with a picture' })
  await slide2.waitFor({ timeout: 5000 })
  const img = slide2.locator('img').first()
  await img.evaluate((el) => (el.complete && el.naturalWidth ? null : new Promise((done) => el.addEventListener('load', done, { once: true }))))
  const picture = await img.evaluate((el) => ({ naturalWidth: el.naturalWidth, src: el.currentSrc }))
  c.ok(picture.naturalWidth > 0 && /\/~office\/[^/]+\/ppt\//.test(picture.src), 'pptx: slide-two picture loads from /~office/…/ppt/', `${picture.naturalWidth}px ${picture.src}`)
  c.ok((await frame().locator('.slide-notes', { hasText: 'speaker notes for slide two' }).count()) === 1, 'pptx: speaker notes are carried in the slide')
  c.ok((await contentFrame(page, V.content).evaluate(() => origin)) === 'null', 'pptx: frame origin null')
  await shot(page, F, '02-pptx-slide2-picture')
  await download('pptx', files.pptx)
  const exported = join(work, 'exported.pptx')
  const exp = cli(F, ['export', ids.pptx, '--out', exported])
  c.ok(exp.code === 0 && sha(exported) === sha(files.pptx), 'pptx: artifact export returns the published .pptx unchanged')

  // docx: real text in the DOM, not a picture of it.
  await page.goto(`${V.app}/a/${ids.docx}`)
  await frame().getByText('A paragraph of selectable verify text.').waitFor({ timeout: 8000 })
  await rendered('docx')
  const docText = await frame().locator('body').innerText()
  c.ok(docText.includes('Verify document heading') && docText.includes('A paragraph of selectable verify text.'), 'docx: heading and paragraph are DOM text')
  c.ok((await contentFrame(page, V.content).evaluate(() => origin)) === 'null', 'docx: frame origin null')
  await shot(page, F, '03-docx')
  await download('docx', files.docx)

  // xlsx: a tab per sheet; #sheet-2 is the hash the second tab links to.
  await page.goto(`${V.app}/a/${ids.xlsx}`)
  await frame().getByText('verify-first-sheet').waitFor({ timeout: 8000 })
  await rendered('xlsx')
  c.ok((await frame().locator('a.xl-tab').count()) === 2, 'xlsx: two sheet tabs', await frame().locator('a.xl-tab').allTextContents())
  await shot(page, F, '04-xlsx-sheet1')
  await frame().locator('a.xl-tab[href="#sheet-2"]').click()
  await frame().getByText('verify-second-sheet').waitFor({ state: 'visible', timeout: 5000 })
  c.ok(await frame().getByText('verify-first-sheet').isHidden(), 'xlsx: clicking the second tab shows sheet 2 and hides sheet 1')
  await shot(page, F, '05-xlsx-sheet2')
  await download('xlsx', files.xlsx)

  const real = process.env.VERIFY_OFFICE_FILE
  if (real) {
    if (!existsSync(real)) throw new Error(`VERIFY_OFFICE_FILE ${real} does not exist`)
    const out = cli(F, ['publish', real, '--new', '--json'])
    ids.real = out.code === 0 ? JSON.parse(out.stdout).id : undefined
    c.ok(out.code === 0, `real file ${basename(real)}: published`, ids.real ?? out.stderr.trim())
    await page.goto(`${V.app}/a/${ids.real}`)
    await rendered(`real file ${basename(real)}`)
    await shot(page, F, '06-real-file')
    await download('real', real)
  } else {
    log(F, 'VERIFY_OFFICE_FILE not set: only the generated files were checked')
  }
  save(F, 'ids.json', ids)
  c.finish()
} catch (error) {
  c.finish(error)
} finally {
  await browser?.close()
}

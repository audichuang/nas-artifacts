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
  const docText = await frame().locator('body').innerText()
  c.ok(docText.includes('Verify document heading') && docText.includes('A paragraph of selectable verify text.'), 'docx: heading and paragraph are DOM text')
  c.ok((await contentFrame(page, V.content).evaluate(() => origin)) === 'null', 'docx: frame origin null')
  await shot(page, F, '03-docx')
  await download('docx', files.docx)

  // xlsx: a tab per sheet; #sheet-2 is the hash the second tab links to.
  await page.goto(`${V.app}/a/${ids.xlsx}`)
  await frame().getByText('verify-first-sheet').waitFor({ timeout: 8000 })
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
    ids.real = JSON.parse(out.stdout).id
    await page.goto(`${V.app}/a/${ids.real}`)
    await page.waitForTimeout(3000)
    const missing = await frame().locator('.pp-missing').count()
    c.ok(out.code === 0 && missing === 0, `real file ${basename(real)}: published, no "cannot draw" boxes`, `${missing} missing boxes`)
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

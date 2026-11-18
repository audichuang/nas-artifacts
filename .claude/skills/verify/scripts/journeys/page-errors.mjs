#!/usr/bin/env node
// Journey: what breaks in a page reaches the viewer's ⚠ and the CLI. One page throws inline, makes a fetch the
// sandbox blocks and leaves a Promise rejection unhandled; a second page has a missing picture. Each kind is listed
// with its page file, repeats bump a count instead of adding rows, and the log is per version.
// Usage: node .claude/skills/verify/scripts/journeys/page-errors.mjs
import { writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { V, checks, cli, launch, log, save, scratch, shot } from '../browser.mjs'

const F = 'page-errors'
const c = checks(F)
const nonce = Date.now().toString(36)
const FETCH_URL = `${V.publicApp}/.well-known/oauth-protected-resource/mcp`
// Separate inline scripts, so one throwing does not stop the others. The fetch target answers CORS, so only the
// sandbox's connect-src 'none' stops it; its .catch keeps it from adding a rejection row of its own.
const INDEX = `<!doctype html>
<title>errors</title>
<h1 id="t">Errors ${nonce}</h1>
<script>throw new Error('verify boom ${nonce}')</script>
<script>fetch('${FETCH_URL}').catch(() => {})</script>
<script>Promise.reject(new Error('verify reject ${nonce}'))</script>
<a id="next" href="two.html">two</a>
`
const TWO = `<!doctype html><title>two</title><h1 id="two">Two ${nonce}</h1><img src="missing.png" alt="missing">\n`
const LABELS = { error: '腳本錯誤', rejection: '未處理的 Promise', resource: '載入失敗', csp: '被沙箱擋下' }

const cliErrors = (id, extra = []) => JSON.parse(cli(F, ['errors', id, '--json', ...extra]).stdout).errors
async function pollErrors(id, ready) {
  let rows = []
  for (let i = 0; i < 30; i++) {
    rows = cliErrors(id)
    if (ready(rows)) break
    await new Promise((resolve) => setTimeout(resolve, 300))
  }
  return rows
}
const brief = (rows) => rows.map((e) => `[${e.kind}] ${e.path} ×${e.count} ${e.message} (${e.source})`).join(' | ')

let browser
try {
  const dir = scratch(`errors-${nonce}`)
  writeFileSync(join(dir, 'index.html'), INDEX)
  writeFileSync(join(dir, 'two.html'), TWO)
  const art = JSON.parse(cli(F, ['publish', dir, '--new', '--title', `Errors ${nonce}`, '--json']).stdout)

  browser = await launch()
  const page = await browser.newPage({ viewport: { width: 1200, height: 800 } })
  await page.goto(art.url)
  await page.frameLocator('#frame').locator('#t').waitFor({ timeout: 5000 })
  await page.locator('#errors:not([hidden])').waitFor({ timeout: 5000 })
  await page.locator('#error-count', { hasText: /^3$/ }).waitFor({ timeout: 5000 }).catch(() => {})
  const count1 = (await page.locator('#error-count').textContent()).trim()
  c.ok(count1 === '3', 'index.html: the viewer shows ⚠ 3 (throw, blocked fetch, rejection)', count1)
  await shot(page, F, '01-errors-index')

  await page.frameLocator('#frame').locator('#next').click()
  await page.frameLocator('#frame').locator('#two').waitFor({ timeout: 5000 })
  await page.locator('#error-count', { hasText: /^4$/ }).waitFor({ timeout: 5000 }).catch(() => {})
  const count2 = (await page.locator('#error-count').textContent()).trim()
  c.ok(count2 === '4', 'after clicking through to two.html (missing picture): ⚠ 4', count2)
  await page.locator('#errors').click()
  await page.locator('#error-dialog[open]').waitFor({ timeout: 5000 })
  const dialogRows = await page.locator('#error-list li.thread').evaluateAll((items) => items.map((li) => ({ meta: [...li.querySelectorAll('.comment-meta span')].map((s) => s.textContent), message: li.querySelector('.comment-body')?.textContent, source: li.querySelector('.error-source')?.textContent ?? null })))
  save(F, 'dialog-rows.json', dialogRows)
  const byLabel = Object.fromEntries(dialogRows.map((row) => [row.meta[0], row]))
  c.ok(dialogRows.length === 4 && Object.values(LABELS).every((label) => byLabel[label]), 'the ⚠ dialog lists four rows, one per kind label (腳本錯誤, 未處理的 Promise, 載入失敗, 被沙箱擋下)', dialogRows.map((r) => r.meta.join(' ')).join(' | '))
  c.ok(byLabel[LABELS.resource]?.meta[1] === 'two.html' && [LABELS.error, LABELS.rejection, LABELS.csp].every((label) => byLabel[label]?.meta[1] === 'index.html'), 'the dialog names the page file: 載入失敗 on two.html, the other three on index.html', dialogRows.map((r) => `${r.meta[0]}@${r.meta[1]}`).join(' '))
  await shot(page, F, '02-error-dialog')
  await page.locator('#close-errors').click()

  const rows = await pollErrors(art.id, (r) => r.length >= 4)
  save(F, 'errors-v1.json', rows)
  log(F, `CLI errors --json: ${brief(rows)}`)
  const kind = (k) => rows.find((e) => e.kind === k)
  c.ok(rows.length === 4, 'CLI errors --json: exactly four rows', rows.length)
  c.ok(kind('error')?.path === 'index.html' && kind('error').message.includes(`verify boom ${nonce}`) && /^index\.html:4(:\d+)?$/.test(kind('error').source ?? ''), 'error row: index.html, the thrown message, source index.html:4[:col]', JSON.stringify(kind('error')))
  c.ok(kind('csp')?.path === 'index.html' && kind('csp').message === `connect-src blocked ${FETCH_URL}` && /^index\.html:5(:\d+)?$/.test(kind('csp').source ?? ''), 'csp row: index.html, "connect-src blocked <the fetch URL>", source index.html:5[:col]', JSON.stringify(kind('csp')))
  c.ok(kind('rejection')?.path === 'index.html' && kind('rejection').message === `Error: verify reject ${nonce}`, 'rejection row: index.html, "Error: verify reject …"', JSON.stringify(kind('rejection')))
  c.ok(kind('resource')?.path === 'two.html' && kind('resource').message === '<img> failed to load' && kind('resource').source === 'missing.png', 'resource row: two.html, "<img> failed to load", source missing.png', JSON.stringify(kind('resource')))
  const plain = cli(F, ['errors', art.id])
  const lines = plain.stdout.trim().split('\n')
  c.ok(plain.code === 0 && lines.length === 4 && ['error', 'csp', 'rejection', 'resource'].every((k) => lines.some((l) => new RegExp(`^\\[${k}\\] (index|two)\\.html ×1  `).test(l))), 'CLI errors (plain): one "[kind] <page file> ×1  message" line per row', plain.stdout.trim())

  // A second visit: the same errors bump their count; no new rows.
  await page.goto(art.url) // not reload(): the address bar names ?path=two.html now
  await page.frameLocator('#frame').locator('#t').waitFor({ timeout: 5000 })
  const again = await pollErrors(art.id, (r) => r.filter((e) => e.path === 'index.html' && e.count === 2).length === 3)
  save(F, 'errors-v1-second-visit.json', again)
  const indexCounts = again.filter((e) => e.path === 'index.html').map((e) => e.count)
  c.ok(again.length === 4 && indexCounts.length === 3 && indexCounts.every((n) => n === 2) && again.find((e) => e.path === 'two.html')?.count === 1, 'second visit: the three index.html rows read ×2, two.html stays ×1, still four rows', brief(again))
  await page.locator('#error-count', { hasText: /^4$/ }).waitFor({ timeout: 5000 }).catch(() => {})
  c.ok((await page.locator('#error-count').textContent()).trim() === '4', '⚠ still counts 4 rows (distinct errors, not the sum)', await page.locator('#error-count').textContent())
  await shot(page, F, '03-second-visit')

  // Per version: a fixed v2 has no errors; v1 keeps its four.
  writeFileSync(join(dir, 'index.html'), `<!doctype html><title>fixed</title><h1 id="t">Fixed ${nonce}</h1>\n`)
  writeFileSync(join(dir, 'two.html'), `<!doctype html><h1 id="two">Two fixed ${nonce}</h1>\n`)
  const v2 = cli(F, ['publish', dir])
  c.ok(v2.code === 0 && v2.stdout.includes('v2'), 'published a fixed v2', v2.stdout.trim().split('\n')[0])
  const fresh = await browser.newPage({ viewport: { width: 1200, height: 800 } })
  await fresh.goto(art.url)
  await fresh.frameLocator('#frame').locator('#t', { hasText: `Fixed ${nonce}` }).waitFor({ timeout: 5000 })
  await fresh.waitForTimeout(1200)
  c.ok(await fresh.locator('#errors').isHidden(), 'v2 in the viewer: no ⚠')
  await shot(fresh, F, '04-fixed-v2')
  const empty = cli(F, ['errors', art.id])
  c.ok(empty.stdout.trim() === '(no errors reported for this version)', 'CLI errors for v2: (no errors reported for this version)', empty.stdout.trim())
  const old = cliErrors(art.id, ['--v', '1'])
  c.ok(old.length === 4, 'CLI errors --v 1 still lists the four', brief(old))
  save(F, 'ids.json', { artifact: art.id })
  cli(F, ['delete', art.id, '--yes'], { quiet: true })
  c.finish()
} catch (error) {
  c.finish(error)
} finally {
  await browser?.close()
}

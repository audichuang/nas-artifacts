#!/usr/bin/env node
// Journey: a private page is invisible on the public listeners; made public it is readable anonymously, read-only,
// current version only; made private again it is gone. Anonymous = a fresh browser context with no session.
// Usage: node .claude/skills/verify/scripts/journeys/public-sharing.mjs
import { createHash } from 'node:crypto'
import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { V, checks, cli, contentFrame, launch, save, scratch, shot } from '../browser.mjs'

const F = 'public-sharing'
const c = checks(F)
const nonce = Date.now().toString(36)
const dir = scratch(`pub-${nonce}`)
writeFileSync(join(dir, 'index.md'), `# Public report ${nonce}\n\nFirst edition.\n\n[next page](two.md)\n`)
writeFileSync(join(dir, 'two.md'), '# Second page\n')
const sha = (bytes) => createHash('sha256').update(bytes).digest('hex')

let browser
try {
  const art = JSON.parse(cli(F, ['publish', dir, '--new', '--title', `Public report ${nonce}`, '--json']).stdout)
  const publicUrl = `${V.publicApp}/a/${art.id}`
  browser = await launch()
  const anon = await browser.newContext({ viewport: { width: 1100, height: 760 } })
  const page = await anon.newPage()

  const closed = await page.goto(publicUrl)
  const closedText = (await closed.text()).replaceAll(art.id, 'ID')
  const missing = await page.request.get(`${V.publicApp}/a/no-such-id`)
  const missingText = (await missing.text()).replaceAll('no-such-id', 'ID')
  c.ok(closed.status() === 404 && missing.status() === 404 && closedText === missingText, 'private page on the public app: 404, same body as a missing id', `${closed.status()} ${missing.status()}`)
  await shot(page, F, '01-private-is-404')
  const closedFile = await page.request.get(`${V.publicContent}/a/${art.id}/v/1/index.md`)
  c.ok(closedFile.status() === 404, 'private file on the public content listener: 404', closedFile.status())

  const made = cli(F, ['visibility', art.id, 'public'])
  c.ok(made.code === 0 && made.stdout.includes(`is public: ${publicUrl}`), 'artifact visibility public prints the public URL', made.stdout.trim())
  c.ok(cli(F, ['list']).stdout.split('\n').some((l) => l.startsWith(art.id) && l.includes('[public]')), 'artifact list marks it [public]')

  const open = await page.goto(publicUrl)
  const headers = open.headers()
  save(F, 'public-viewer-headers.json', { url: publicUrl, status: open.status(), headers })
  c.ok(headers['cache-control'] === 'private, no-cache' && /noindex/.test(headers['x-robots-tag'] ?? ''), 'public viewer: Cache-Control private, no-cache and X-Robots-Tag noindex', `${headers['cache-control']} | ${headers['x-robots-tag']}`)
  const frame = page.frameLocator('iframe.public-frame')
  await frame.locator('h1', { hasText: `Public report ${nonce}` }).waitFor({ timeout: 5000 })
  c.ok((await contentFrame(page, V.publicContent).evaluate(() => origin)) === 'null', 'public page frame is sandboxed (origin null)')
  c.ok((await page.locator('#comments, #pin, #menu-delete, #toggle-comments').count()) === 0, 'no write controls on the public viewer')
  await shot(page, F, '02-public-viewer')
  const fileHeaders = (await page.request.get(`${V.publicContent}/a/${art.id}/v/1/index.md`)).headers()
  c.ok(fileHeaders['cache-control'] === 'private, no-cache' && /noindex/.test(fileHeaders['x-robots-tag'] ?? ''), 'public content response carries the same cache and robots headers', `${fileHeaders['cache-control']} | ${fileHeaders['x-robots-tag']}`)

  const [download] = await Promise.all([page.waitForEvent('download'), page.locator('a.button[download]').click()])
  const downloaded = readFileSync(await download.path())
  c.ok(sha(downloaded) === sha(readFileSync(join(dir, 'index.md'))), 'MD download button returns the published file (sha256 equal)', download.suggestedFilename())
  await frame.locator('a', { hasText: 'next page' }).click()
  await frame.locator('h1', { hasText: 'Second page' }).waitFor({ timeout: 5000 })
  c.ok(true, 'relative link works inside the public frame')
  await shot(page, F, '03-public-second-page')
  const root = await page.request.get(`${V.publicApp}/`)
  const api = await page.request.get(`${V.publicApp}/api/artifacts`)
  c.ok(root.status() === 404 && api.status() === 404, 'anonymous / and /api/artifacts on the public app are 404', `${root.status()} ${api.status()}`)

  const tailnet = await browser.newPage({ viewport: { width: 1100, height: 760 } })
  await tailnet.goto(art.url)
  await tailnet.locator('#public-link').waitFor({ state: 'visible', timeout: 5000 })
  c.ok((await tailnet.locator('#public-link').getAttribute('href')) === publicUrl, 'tailnet viewer shows #public-link to the public URL')
  await shot(tailnet, F, '04-tailnet-public-link')

  writeFileSync(join(dir, 'index.md'), `# Public report ${nonce}\n\nSecond edition.\n\n[next page](two.md)\n`)
  const v2 = cli(F, ['publish', dir])
  c.ok(v2.code === 0 && v2.stdout.includes(`public: ${publicUrl}`), 'republish keeps it public (v2)', v2.stdout.trim())
  await page.goto(publicUrl)
  await frame.locator('p', { hasText: 'Second edition.' }).waitFor({ timeout: 5000 })
  const old = await page.request.get(`${publicUrl}?v=1`)
  const oldFile = await page.request.get(`${V.publicContent}/a/${art.id}/v/1/index.md`)
  c.ok(old.status() === 404 && oldFile.status() === 404, 'public listeners serve only the current version: ?v=1 and /v/1/ are 404', `${old.status()} ${oldFile.status()}`)

  const priv = cli(F, ['visibility', art.id, 'private'])
  c.ok(priv.code === 0 && priv.stdout.includes('is private'), 'artifact visibility private', priv.stdout.trim())
  const gone = await page.goto(publicUrl)
  const goneFile = await page.request.get(`${V.publicContent}/a/${art.id}/v/2/index.md`)
  c.ok(gone.status() === 404 && goneFile.status() === 404, 'back to private: public app and public content are 404 again', `${gone.status()} ${goneFile.status()}`)
  await shot(page, F, '05-private-again')
  c.ok(!cli(F, ['list']).stdout.split('\n').some((l) => l.startsWith(art.id) && l.includes('[public]')), 'artifact list no longer marks it [public]')
  save(F, 'ids.json', { id: art.id })
  c.finish()
} catch (error) {
  c.finish(error)
} finally {
  await browser?.close()
}

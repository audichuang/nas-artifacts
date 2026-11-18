#!/usr/bin/env node
// Journey: a viewer comment read and answered by the CLI, live in the viewer; `comments --wait` wakes on a new one.
// Usage: node .claude/skills/verify/scripts/journeys/comments.mjs
import { writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { checks, cli, cliBackground, launch, log, save, scratch, shot } from '../browser.mjs'

const F = 'comments'
const c = checks(F)
const nonce = Date.now().toString(36)
const dir = scratch(`cm-${nonce}`)
writeFileSync(join(dir, 'page.md'), `# Comment target ${nonce}\n\nSomething to talk about.\n`)

let browser
let waiter
try {
  const art = JSON.parse(cli(F, ['publish', join(dir, 'page.md'), '--new', '--json']).stdout)
  browser = await launch()
  const viewer = await browser.newPage({ viewport: { width: 1280, height: 860 } })
  await viewer.goto(art.url)
  await viewer.frameLocator('#frame').locator('h1').waitFor({ timeout: 5000 })
  await viewer.evaluate(() => { window.__verifyNoReload = true })
  await viewer.locator('#toggle-comments').click()
  await viewer.locator('#comments').waitFor({ state: 'visible' })
  viewer.once('dialog', (dialog) => dialog.accept('Verifier'))
  await viewer.locator('#comment-body').fill(`first comment ${nonce}`)
  await viewer.locator('#comment-form button[type="submit"]').click()
  await viewer.locator('#threads', { hasText: `first comment ${nonce}` }).waitFor({ timeout: 5000 })
  await shot(viewer, F, '01-comment-posted')

  const listed = JSON.parse(cli(F, ['comments', art.id, '--json']).stdout).comments
  const root = listed.find((x) => x.body === `first comment ${nonce}`)
  c.ok(root && root.author === 'Verifier' && !root.resolved, 'CLI comments --json has the viewer comment (author Verifier, open)', JSON.stringify(root))

  const reply = cli(F, ['comment', 'reply', art.id, root.id, `agent reply ${nonce}`, '--resolve', '--author', 'Agent'])
  c.ok(reply.code === 0 && reply.stdout.includes('and resolved it'), 'CLI reply --resolve', reply.stdout.trim())
  // Resolved threads leave the default list; the viewer re-renders live, no reload.
  await viewer.locator('#threads', { hasText: `first comment ${nonce}` }).waitFor({ state: 'detached', timeout: 5000 })
  await shot(viewer, F, '02-resolved-thread-hidden')
  await viewer.locator('#show-resolved').check()
  const thread = viewer.locator('.thread.resolved', { hasText: `agent reply ${nonce}` })
  await thread.waitFor({ timeout: 5000 })
  c.ok((await thread.textContent()).includes(`first comment ${nonce}`) && (await viewer.evaluate(() => window.__verifyNoReload === true)), 'viewer shows the reply in a resolved thread without a reload')
  await shot(viewer, F, '03-resolved-thread-with-reply')
  const open = cli(F, ['comments', art.id, '--open'])
  c.ok(open.code === 0 && !open.stdout.includes(`first comment ${nonce}`), 'comments --open omits the resolved thread', open.stdout.trim())

  const reopen = cli(F, ['comment', 'reopen', art.id, root.id])
  c.ok(reopen.code === 0 && cli(F, ['comments', art.id, '--open']).stdout.includes(`first comment ${nonce}`), 'comment reopen puts it back in --open')
  await viewer.locator('#show-resolved').uncheck()
  await viewer.locator('#threads', { hasText: `first comment ${nonce}` }).waitFor({ timeout: 5000 })

  waiter = cliBackground(F, ['comments', art.id, '--wait', '--timeout', '20'])
  const deadline = Date.now() + 10_000
  while (!waiter.stderr().includes('waiting up to') && Date.now() < deadline) await new Promise((r) => setTimeout(r, 100))
  log(F, `background comments --wait pid ${waiter.child.pid} is listening: ${waiter.stderr().includes('waiting up to')}`)
  viewer.once('dialog', (dialog) => dialog.accept('Verifier'))
  await viewer.locator('#comment-body').fill(`second comment ${nonce}`)
  await viewer.locator('#comment-form button[type="submit"]').click()
  const woke = await waiter.done
  c.ok(woke.code === 0 && woke.stdout.includes(`second comment ${nonce}`) && !woke.stdout.includes(`first comment ${nonce}`), 'comments --wait wakes with exit 0 and prints only the new comment', woke.stdout.trim())
  c.ok(waiter.child.exitCode !== null, 'the background CLI process has exited', `pid ${waiter.child.pid} exit ${waiter.child.exitCode}`)
  await shot(viewer, F, '04-second-comment')
  save(F, 'ids.json', { id: art.id, comment: root.id })
  c.finish()
} catch (error) {
  c.finish(error)
} finally {
  if (waiter && waiter.child.exitCode === null) waiter.child.kill()
  await browser?.close()
}

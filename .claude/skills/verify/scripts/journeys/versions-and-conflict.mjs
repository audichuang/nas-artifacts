#!/usr/bin/env node
// Journey: republish to the same URL (viewer follows live), diff, a second machine's conflicting publish, --force, prune.
// Usage: node .claude/skills/verify/scripts/journeys/versions-and-conflict.mjs
import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { V, checks, cli, contentFrame, launch, log, save, scratch, shot } from '../browser.mjs'

const F = 'versions-and-conflict'
const c = checks(F)
const nonce = Date.now().toString(36)
const a = scratch(`vc-a-${nonce}`)
const b = join(scratch(`vc-b-${nonce}`), 'copy')
const second = { VERIFY_CLI_DIR: scratch(`cli-second-${nonce}`) }
const page = (text) => `<!doctype html><title>versions</title><h1 id="v">${text}</h1>\n<p>shared line</p>\n`
const viewerFrameText = (p) => p.frameLocator('#frame').locator('#v')

let browser
try {
  writeFileSync(join(a, 'index.html'), page('version one'))
  const first = JSON.parse(cli(F, ['publish', a, '--new', '--title', `Versions ${nonce}`, '--json']).stdout)
  c.ok(first.currentVersion === 1, 'machine A publishes v1', first.url)

  browser = await launch()
  const viewer = await browser.newPage({ viewport: { width: 1200, height: 800 } })
  await viewer.goto(first.url)
  await viewerFrameText(viewer).filter({ hasText: 'version one' }).waitFor({ timeout: 5000 })
  await shot(viewer, F, '01-viewer-v1')
  // A reload would wipe this marker; it survives only if the viewer switched versions in place.
  await viewer.evaluate(() => { window.__verifyNoReload = true })

  writeFileSync(join(a, 'index.html'), page('version two'))
  const republish = cli(F, ['publish', a, '--label', 'second'])
  c.ok(republish.code === 0 && republish.stdout.includes('(v2, same URL)'), 'republish from the same folder prints "v2, same URL"', republish.stdout.trim())
  await viewerFrameText(viewer).filter({ hasText: 'version two' }).waitFor({ timeout: 8000 })
  c.ok((await viewer.evaluate(() => window.__verifyNoReload === true)) && viewer.url().startsWith(first.url), 'open viewer switched to v2 without a reload, same URL', viewer.url())
  c.ok(contentFrame(viewer, V.content)?.url().includes('/v/2/'), 'viewer frame now shows /v/2/', contentFrame(viewer, V.content)?.url())
  await shot(viewer, F, '02-viewer-v2-live')

  const diff = cli(F, ['diff', first.id])
  c.ok(diff.code === 0 && diff.stdout.includes('-<!doctype html><title>versions</title><h1 id="v">version one</h1>') && diff.stdout.includes('+<!doctype html><title>versions</title><h1 id="v">version two</h1>'), 'artifact diff lists the changed line', diff.stdout.split('\n')[0])

  const checkout = cli(F, ['checkout', first.id, b], { env: second })
  c.ok(checkout.code === 0 && checkout.stdout.includes('v2'), 'machine B checks out v2 into its own folder', checkout.stdout.trim())
  writeFileSync(join(a, 'index.html'), page('version three from A'))
  const third = cli(F, ['publish', a])
  c.ok(third.stdout.includes('(v3, same URL)'), 'machine A publishes v3', third.stdout.trim())

  writeFileSync(join(b, 'index.html'), page('edit from B'))
  const refused = cli(F, ['publish', b], { env: second })
  c.ok(refused.code !== 0 && refused.stderr.includes('changed remotely since v2: index.html') && refused.stderr.includes('Nothing was uploaded') && refused.stderr.includes('--force'), 'B is refused with changedFiles and both options', refused.stderr.trim())
  const getAfterRefusal = JSON.parse(cli(F, ['get', first.id, '--json']).stdout)
  c.ok(getAfterRefusal.currentVersion === 3, 'the refused publish created no version', getAfterRefusal.currentVersion)

  const forced = cli(F, ['publish', b, '--force'], { env: second })
  c.ok(forced.code === 0 && forced.stdout.includes('(v4, same URL)'), 'B publishes with --force -> v4', forced.stdout.trim())
  const afterForce = JSON.parse(cli(F, ['get', first.id, '--json']).stdout)
  c.ok(afterForce.versions.map((v) => v.version).join(',') === '1,2,3,4' || afterForce.versions.map((v) => v.version).sort().join(',') === '1,2,3,4', 'get lists v1..v4', afterForce.versions.map((v) => v.version).join(','))
  c.ok(cli(F, ['read', first.id], { quiet: true }).stdout === readFileSync(join(b, 'index.html'), 'utf8'), 'current content is B\'s file')

  const prune = cli(F, ['prune', first.id, '--keep', '1'])
  c.ok(prune.code === 0 && /dropped v/.test(prune.stdout), 'prune --keep 1 drops the old versions', prune.stdout.trim())
  const afterPrune = JSON.parse(cli(F, ['get', first.id, '--json']).stdout)
  c.ok(afterPrune.versions.length === 1 && afterPrune.versions[0].version === 4, 'get lists only v4 after prune', afterPrune.versions.map((v) => v.version).join(','))

  const old = await browser.newPage({ viewport: { width: 1200, height: 800 } })
  await old.goto(`${first.url}?v=1`)
  await viewerFrameText(old).filter({ hasText: 'edit from B' }).waitFor({ timeout: 5000 })
  c.ok(contentFrame(old, V.content)?.url().includes('/v/4/'), '?v=1 after prune shows the current version (v4)', `${old.url()} -> ${contentFrame(old, V.content)?.url()}`)
  await shot(old, F, '03-pruned-v1-shows-v4')
  save(F, 'ids.json', { id: first.id })
  c.finish()
} catch (error) {
  c.finish(error)
} finally {
  await browser?.close()
}

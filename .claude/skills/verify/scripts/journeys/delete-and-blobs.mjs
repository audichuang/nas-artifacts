#!/usr/bin/env node
// Journey: deleting an artifact and an asset. Two artifacts share an identical file (one content-addressed blob).
// Deleting one needs --yes, tells its open viewer, and makes its URLs 404, while the other still serves the shared
// bytes. An uploaded asset with the same bytes is deleted through the token-only API (the CLI has no command for
// it): 401 without the token, 204 with it, then its URL is 404 and the published file with those bytes survives.
// Usage: node .claude/skills/verify/scripts/journeys/delete-and-blobs.mjs
import { writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { V, checks, cli, launch, save, scratch, shot } from '../browser.mjs'

const F = 'delete-and-blobs'
const c = checks(F)
const nonce = Date.now().toString(36)
const SHARED = `shared bytes ${nonce}\n`.repeat(20)

async function get(url, headers = {}) {
  const response = await fetch(url, { redirect: 'manual', headers })
  return { status: response.status, body: await response.text() }
}
function publish(name) {
  const dir = scratch(`del-${name}-${nonce}`)
  writeFileSync(join(dir, 'index.html'), `<!doctype html><h1 id="t">${name} ${nonce}</h1><a href="shared.txt">shared</a>`)
  writeFileSync(join(dir, 'shared.txt'), SHARED)
  return JSON.parse(cli(F, ['publish', dir, '--new', '--title', `Delete ${name} ${nonce}`, '--json']).stdout)
}

let browser
try {
  const a = publish('A')
  const b = publish('B')
  const sharedA = `${V.content}/a/${a.id}/v/1/shared.txt`
  const sharedB = `${V.content}/a/${b.id}/v/1/shared.txt`
  const beforeA = await get(sharedA)
  const beforeB = await get(sharedB)
  c.ok(beforeA.status === 200 && beforeB.status === 200 && beforeA.body === SHARED && beforeB.body === SHARED, 'both artifacts serve the identical shared.txt (200, same bytes)', `${beforeA.status} ${beforeB.status}`)

  browser = await launch()
  const viewer = await browser.newPage({ viewport: { width: 1200, height: 800 } })
  await viewer.goto(a.url)
  await viewer.frameLocator('#frame').locator('#t').waitFor({ timeout: 5000 })
  await shot(viewer, F, '01-viewer-on-A')

  const refused = cli(F, ['delete', a.id])
  const stillListed = cli(F, ['list', '--json']).stdout.includes(a.id)
  c.ok(refused.code !== 0 && refused.stderr.includes('pass --yes') && stillListed, 'delete without --yes is refused ("pass --yes") and A is still listed', `${refused.code} ${refused.stderr.trim()}`)
  const deleted = cli(F, ['delete', a.id, '--yes'])
  c.ok(deleted.code === 0 && deleted.stdout.includes(`deleted ${a.id}`), 'delete --yes', deleted.stdout.trim())

  await viewer.locator('#toast', { hasText: '這個 artifact 已被刪除' }).waitFor({ timeout: 5000 })
  c.ok(true, 'the open viewer shows the toast 這個 artifact 已被刪除')
  await shot(viewer, F, '02-viewer-told')
  await viewer.waitForURL(`${V.app}/`, { timeout: 5000 })
  c.ok(true, 'and is sent to the gallery')

  const listed = cli(F, ['list', '--json']).stdout
  c.ok(!listed.includes(a.id) && listed.includes(b.id), 'CLI list: A is gone, B is still there')
  c.ok(cli(F, ['get', a.id]).code !== 0, 'CLI get A fails')
  const gone = await Promise.all([get(`${V.app}/a/${a.id}`), get(`${V.content}/a/${a.id}/v/1/`), get(sharedA)])
  c.ok(gone.every((r) => r.status === 404), "A's viewer URL, entry URL and shared.txt URL all 404", gone.map((r) => r.status).join(' '))
  await viewer.locator('#search').fill(nonce)
  await viewer.locator('#status', { hasText: '共 1 個' }).waitFor({ timeout: 5000 })
  const titles = await viewer.locator('#list .card-title').allTextContents()
  c.ok(titles.join() === `Delete B ${nonce}`, 'gallery search for this run shows only B', titles.join())
  await shot(viewer, F, '03-gallery-only-B')

  const afterB = await get(sharedB)
  c.ok(afterB.status === 200 && afterB.body === SHARED, "B's shared.txt still serves the same bytes (the shared blob survived)", afterB.status)
  const readB = cli(F, ['read', b.id, 'shared.txt'])
  c.ok(readB.code === 0 && readB.stdout === SHARED, 'CLI read B shared.txt returns the same bytes')
  const freshCtx = await browser.newContext({ viewport: { width: 1200, height: 800 } })
  const fresh = await freshCtx.newPage()
  await fresh.goto(b.url)
  await fresh.frameLocator('#frame').locator('#t', { hasText: `B ${nonce}` }).waitFor({ timeout: 5000 })
  c.ok(true, "B's page renders in a fresh browser")
  await freshCtx.close()

  // Asset delete: same bytes as shared.txt, so it shares that blob too.
  const assetFile = join(scratch(`del-asset-${nonce}`), 'asset.txt')
  writeFileSync(assetFile, SHARED)
  const up = cli(F, ['asset', 'upload', b.id, assetFile])
  const assetUrl = up.stdout.trim()
  const assetId = /\/assets\/([^/]+)\//.exec(assetUrl)?.[1]
  const served = await get(assetUrl)
  c.ok(up.code === 0 && assetId && served.status === 200 && served.body === SHARED, 'asset upload: the printed URL serves the same bytes', `${up.code} ${served.status} ${assetUrl}`)
  const listedAsset = cli(F, ['asset', 'list', b.id]).stdout
  c.ok(listedAsset.includes(assetId) && listedAsset.includes(assetUrl), 'asset list shows it', listedAsset.trim())
  const route = `${V.app}/api/artifacts/${b.id}/assets/${assetId}`
  const noToken = await fetch(route, { method: 'DELETE' })
  c.ok(noToken.status === 401 && (await get(assetUrl)).status === 200, 'DELETE asset without a token -> 401, the asset still serves', noToken.status)
  const viewerOnly = await fetch(route, { method: 'DELETE', headers: { 'x-artifact-client': 'viewer' } })
  c.ok(viewerOnly.status === 401 && (await get(assetUrl)).status === 200, 'DELETE asset with only x-artifact-client: viewer -> 401, the asset still serves', viewerOnly.status)
  const withToken = await fetch(route, { method: 'DELETE', headers: { authorization: `Bearer ${V.token}` } })
  c.ok(withToken.status === 204, 'DELETE asset with the publish token -> 204', withToken.status)
  const assetsAfter = cli(F, ['asset', 'list', b.id]).stdout.trim()
  const assetGone = await get(assetUrl)
  c.ok(assetsAfter === '(no assets)' && assetGone.status === 404, 'asset list is empty and the asset URL is 404', `${assetsAfter} ${assetGone.status}`)
  const stillB = await get(sharedB)
  c.ok(stillB.status === 200 && stillB.body === SHARED, "B's shared.txt (same bytes as the asset) still serves", stillB.status)
  save(F, 'ids.json', { deleted: a.id, kept: b.id, asset: assetId })
  cli(F, ['delete', b.id, '--yes'], { quiet: true })
  c.finish()
} catch (error) {
  c.finish(error)
} finally {
  await browser?.close()
}

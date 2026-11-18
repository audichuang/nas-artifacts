#!/usr/bin/env node
// Journey: the owner enrols a passkey from a login-link, reads a private page on the public origins, authorizes an
// MCP client (OAuth + PKCE + consent), drives every MCP tool, cross-checks with the CLI and viewer, then revokes it.
// Usage: node .claude/skills/verify/scripts/journeys/owner-signin-and-mcp.mjs
import { writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { V, checks, cli, contentFrame, enrolOwner, launch, log, redact, save, scratch, shot } from '../browser.mjs'
import { rpc, tool } from '../mcp-call.mjs'
import { MCP_URL, authorizeClient } from '../mcp-token.mjs'

const F = 'owner-signin-and-mcp'
const c = checks(F)
const nonce = Date.now().toString(36)
const calls = []
const record = (name, args, out) => calls.push({ call: name, args, status: out.status, isError: out.isError, result: out.value ?? out.body })

let browser
try {
  browser = await launch()
  const owner = await browser.newContext({ viewport: { width: 1200, height: 800 } })
  const page = await owner.newPage()
  await enrolOwner(page, F, `verify device ${nonce}`)
  await page.locator('a[href="/account"]').waitFor({ timeout: 5000 })
  const cookie = (await owner.cookies()).find((x) => x.name === '__Host-artifacts_session')
  c.ok(cookie && cookie.httpOnly && cookie.secure && cookie.sameSite === 'Lax' && cookie.path === '/', 'session cookie __Host-artifacts_session is HttpOnly, Secure, SameSite=Lax', cookie && `${cookie.httpOnly} ${cookie.secure} ${cookie.sameSite}`)
  await shot(page, F, '01-owner-gallery')
  await page.goto(`${V.publicApp}/account`)
  await page.locator('#passkeys li', { hasText: `verify device ${nonce}` }).waitFor({ timeout: 5000 })
  c.ok(true, '/account lists the new passkey')
  await shot(page, F, '02-account')

  const dir = scratch(`owner-${nonce}`)
  writeFileSync(join(dir, 'index.html'), `<!doctype html><title>private</title><h1 id="t">Private ${nonce}</h1><a id="next" href="two.html">two</a>`)
  writeFileSync(join(dir, 'two.html'), '<!doctype html><h1 id="two">Private two</h1>')
  const priv = JSON.parse(cli(F, ['publish', dir, '--new', '--title', `Private ${nonce}`, '--json']).stdout)
  await page.goto(`${V.publicApp}/a/${priv.id}`)
  await page.frameLocator('#frame').locator('#t').waitFor({ timeout: 5000 })
  const ownerFrame = contentFrame(page, V.publicContent)
  c.ok(new RegExp(`^${V.publicContent}/t/[^/]+/a/${priv.id}/v/1/`).test(ownerFrame.url()), 'owner reads the private page from the public content listener under /t/<token>/a/<id>/v/1/', redact(ownerFrame.url()))
  c.ok((await ownerFrame.evaluate(() => origin)) === 'null', 'owner frame is still sandboxed (origin null)')
  await page.frameLocator('#frame').locator('#next').click()
  await page.frameLocator('#frame').locator('#two').waitFor({ timeout: 5000 })
  c.ok(true, 'relative link keeps working under the content token')
  await shot(page, F, '03-owner-private-page')
  const anon = await browser.newContext()
  const stranger = await anon.newPage()
  const anonResponse = await stranger.goto(`${V.publicApp}/a/${priv.id}`)
  c.ok(anonResponse.status() === 404, 'the same URL in an anonymous context is 404', anonResponse.status())
  await anon.close()

  const bare = await rpc('tools/list', {}, null)
  c.ok(bare.status === 401 && (bare.headers['www-authenticate'] ?? '').startsWith('Bearer resource_metadata='), 'POST /mcp without a token: 401 with WWW-Authenticate: Bearer resource_metadata=…', bare.headers['www-authenticate'])
  const meta = await (await fetch(`${V.publicApp}/.well-known/oauth-protected-resource/mcp`)).json()
  c.ok(meta.resource === MCP_URL && meta.authorization_servers?.[0] === V.publicApp, 'protected-resource metadata names /mcp and the public app as authorization server', JSON.stringify(meta))

  const clientName = `verify client ${nonce}`
  const auth = await authorizeClient(page, F, clientName)
  c.ok(auth.consent.includes(clientName), 'consent page names the client')
  c.ok(auth.iss === V.publicApp, 'callback iss equals the public app origin', auth.iss)

  const init = await rpc('initialize', { protocolVersion: '2025-11-25', capabilities: {}, clientInfo: { name: 'verify', version: '1' } })
  c.ok(init.status === 200 && init.body?.result?.serverInfo, 'initialize', JSON.stringify(init.body?.result?.serverInfo))
  const listed = await rpc('tools/list')
  const names = (listed.body?.result?.tools ?? []).map((t) => t.name).sort()
  c.ok(names.join(',') === 'comments,delete,errors,get,list,publish,read,reply,set_visibility,update', 'tools/list returns the 10 tools', names.join(','))

  const step = async (name, args) => {
    const out = await tool(name, args)
    record(name, args, out)
    return out
  }
  const published = await step('publish', { title: `From MCP ${nonce}`, files: [{ path: 'index.html', content: `<!doctype html><h1 id="m">mcp v1 ${nonce}</h1>` }] })
  const id = published.value.id
  c.ok(!published.isError && published.value.currentVersion === 1, 'tools/call publish -> v1', id)
  const got = await step('get', { id })
  c.ok(!got.isError && got.value.currentVersion === 1, 'tools/call get')

  const noBase = await step('update', { id, files: [{ path: 'index.html', content: `<!doctype html><h1 id="m">mcp v2 ${nonce}</h1>` }] })
  c.ok(noBase.isError && noBase.value.error?.code === 'base_version_required', 'update without baseVersion -> base_version_required', noBase.value.error?.code)
  // Another machine moves it to v2 with the CLI: a full snapshot on top of the current version.
  const other = scratch(`owner-cli-${nonce}`)
  writeFileSync(join(other, 'index.html'), `<!doctype html><h1 id="m">cli v2 ${nonce}</h1>`)
  const cliV2 = cli(F, ['publish', other, '--id', id])
  c.ok(cliV2.code === 0 && cliV2.stdout.includes('(v2, same URL)'), 'CLI publishes v2 of the MCP page', cliV2.stdout.trim())
  const stale = await step('update', { id, baseVersion: 1, files: [{ path: 'index.html', content: `<!doctype html><h1 id="m">mcp stale ${nonce}</h1>` }] })
  c.ok(stale.isError && stale.value.error?.code === 'version_conflict' && stale.value.error?.details?.changedFiles?.includes('index.html'), 'update with stale baseVersion 1 -> version_conflict with changedFiles', JSON.stringify(stale.value.error?.details))
  const good = await step('update', { id, baseVersion: 2, files: [{ path: 'index.html', content: `<!doctype html><h1 id="m">mcp v3 ${nonce}</h1>` }] })
  c.ok(!good.isError && good.value.currentVersion === 3, 'update with baseVersion 2 -> v3', good.value.currentVersion)

  const replied = await step('reply', { id, body: `note from mcp ${nonce}`, author: 'Verify agent' })
  c.ok(!replied.isError, 'tools/call reply starts a thread')
  const comments = await step('comments', { id })
  c.ok(JSON.stringify(comments.value).includes(`note from mcp ${nonce}`), 'tools/call comments lists it')
  const read = await step('read', { id })
  c.ok(JSON.stringify(read.value).includes(`mcp v3 ${nonce}`), 'tools/call read returns v3 content')
  const errors = await step('errors', { id })
  c.ok(!errors.isError, 'tools/call errors')
  const listedArtifacts = await step('list', { query: nonce })
  c.ok(JSON.stringify(listedArtifacts.value).includes(id), 'tools/call list finds it')
  const shared = await step('set_visibility', { id, visibility: 'public' })
  c.ok(!shared.isError && shared.value.visibility === 'public', 'set_visibility public')
  c.ok((await fetch(`${V.publicApp}/a/${id}`)).status === 200, 'anonymous public URL now 200')
  const unshared = await step('set_visibility', { id, visibility: 'private' })
  c.ok(!unshared.isError && (await fetch(`${V.publicApp}/a/${id}`)).status === 404, 'set_visibility private -> anonymous 404 again')

  const cliGet = JSON.parse(cli(F, ['get', id, '--json']).stdout)
  c.ok(cliGet.currentVersion === 3 && cliGet.versions.some((v) => v.label === 'MCP'), 'CLI get sees the MCP-written v3 (label MCP)', cliGet.versions.map((v) => `v${v.version}:${v.label}`).join(' '))
  c.ok(cli(F, ['comments', id]).stdout.includes(`note from mcp ${nonce}`), 'CLI comments sees the MCP comment')
  const viewer = await browser.newPage({ viewport: { width: 1200, height: 800 } })
  await viewer.goto(`${V.app}/a/${id}`)
  await viewer.frameLocator('#frame').locator('#m', { hasText: `mcp v3 ${nonce}` }).waitFor({ timeout: 5000 })
  c.ok(true, 'tailnet viewer renders the MCP page')
  await shot(viewer, F, '04-tailnet-viewer-mcp-page')

  const deleted = await step('delete', { id })
  c.ok(!deleted.isError && cli(F, ['get', id]).code !== 0, 'tools/call delete; CLI get then fails', JSON.stringify(deleted.value))

  await page.goto(`${V.publicApp}/account`)
  const row = page.locator('#clients li', { hasText: clientName })
  await row.waitFor({ timeout: 5000 })
  await shot(page, F, '05-account-client-listed')
  page.once('dialog', (dialog) => dialog.accept())
  await row.locator('button', { hasText: '撤銷' }).click()
  await page.locator('#status', { hasText: '已撤銷' }).waitFor({ timeout: 5000 })
  await row.waitFor({ state: 'detached', timeout: 5000 })
  await shot(page, F, '06-account-client-revoked')
  const after = await rpc('tools/list')
  c.ok(after.status === 401, 'the same access token gets 401 after revocation', after.status)
  save(F, 'mcp-calls.json', calls)
  c.finish()
} catch (error) {
  save(F, 'mcp-calls.json', calls)
  c.finish(error)
} finally {
  await browser?.close()
}

#!/usr/bin/env node
// Journey: owner sessions on the public app. Sign in again with the existing passkey at /login, sign out, sign out
// everywhere, remove a passkey; every old session cookie and every captured /t/<content token>/ URL is replayed from
// Node afterwards to prove the server ended it (not just the browser dropping a cookie).
// Usage: node .claude/skills/verify/scripts/journeys/owner-sessions.mjs
import { writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { SESSION_COOKIE, V, checks, cli, contentFrame, enrolOwner, launch, log, redact, save, scratch, secret, shot, withCookie } from '../browser.mjs'

const F = 'owner-sessions'
const c = checks(F)
const nonce = Date.now().toString(36)
const DATA = `data for ${nonce}\n`

async function sessionCookie(context) {
  const value = (await context.cookies()).find((x) => x.name === SESSION_COOKIE)?.value
  return secret(value)
}
/** Opens the private page as the owner and returns the URL of data.txt under the frame's content token. */
async function tokenUrlFor(page, id) {
  await page.goto(`${V.publicApp}/a/${id}`)
  await page.frameLocator('#frame').locator('#t').waitFor({ timeout: 5000 })
  const frameUrl = contentFrame(page, V.publicContent).url()
  const url = new URL('data.txt', frameUrl).href
  secret(/\/t\/([^/]+)\//.exec(url)?.[1])
  return url
}
async function signIn(page, label) {
  await page.goto(`${V.publicApp}/login`)
  await page.locator('#go').click()
  await page.waitForURL(`${V.publicApp}/`, { timeout: 10_000 }).catch(() => {})
  const ok = page.url() === `${V.publicApp}/`
  const status = ok ? '' : await page.locator('#status').textContent().catch(() => '')
  log(F, `${label}: /login -> #go -> ${page.url()}${status ? `  (#status ${status})` : ''}`)
  return ok
}

let browser
try {
  browser = await launch()
  const ctxA = await browser.newContext({ viewport: { width: 1200, height: 800 } })
  const pageA = await ctxA.newPage()
  await enrolOwner(pageA, F, `device A ${nonce}`)

  const dir = scratch(`sessions-${nonce}`)
  writeFileSync(join(dir, 'index.html'), `<!doctype html><title>private</title><h1 id="t">Private ${nonce}</h1>`)
  writeFileSync(join(dir, 'data.txt'), DATA)
  const art = JSON.parse(cli(F, ['publish', dir, '--new', '--title', `Sessions ${nonce}`, '--json']).stdout)
  const pageUrl = `${V.publicApp}/a/${art.id}`
  const tokenUrl1 = await tokenUrlFor(pageA, art.id)
  const cookieA = await sessionCookie(ctxA)
  c.ok(Boolean(cookieA), 'enrolment left a session cookie in context A')

  // Controls: the replay technique works while the session is alive.
  const liveA = await withCookie(pageUrl, cookieA)
  c.ok(liveA.status === 200, 'control: Node GET /a/<id> with cookie A (explicit cookie header) -> 200', liveA.status)
  const anon = await withCookie(pageUrl, null)
  c.ok(anon.status === 404, 'control: the same GET without a cookie -> 404', anon.status)
  const tok1 = await withCookie(tokenUrl1, null)
  c.ok(tok1.status === 200 && tok1.body === DATA, 'control: Node GET of data.txt under the content token -> 200, same bytes', `${tok1.status} ${JSON.stringify(tok1.body)}`)
  await shot(pageA, F, '01-owner-private-page')

  // 1. Sign out on this device.
  await pageA.goto(`${V.publicApp}/account`)
  await pageA.locator('#logout').click()
  await pageA.waitForURL(`${V.publicApp}/login`, { timeout: 5000 })
  c.ok(true, '#logout lands on /login')
  const afterLogoutPage = await withCookie(pageUrl, cookieA)
  const afterLogoutApi = await withCookie(`${V.publicApp}/api/account`, cookieA)
  c.ok(afterLogoutPage.status === 404 && afterLogoutApi.status === 404, 'after #logout, replaying cookie A: /a/<id> and /api/account -> 404 (session ended on the server)', `${afterLogoutPage.status} ${afterLogoutApi.status}`)
  const goneInBrowser = await pageA.goto(pageUrl)
  c.ok(goneInBrowser.status() === 404 && (await pageA.locator('body').innerText()).includes('擁有者登入'), 'the browser now gets the 404 page with 擁有者登入', goneInBrowser.status())
  const tokAfterLogout = await withCookie(tokenUrl1, null)
  log(F, `observation: data.txt under the content token after a single #logout -> ${tokAfterLogout.status} (one logout does not rotate the content key; only ending all sessions does)`)
  await shot(pageA, F, '02-after-logout')

  // 2. Sign in again with the passkey the authenticator kept.
  c.ok(await signIn(pageA, 'sign in again'), 'the existing passkey signs in again at /login (#go) and lands on the gallery', pageA.url())
  await pageA.goto(pageUrl)
  await pageA.frameLocator('#frame').locator('#t', { hasText: `Private ${nonce}` }).waitFor({ timeout: 5000 })
  c.ok(true, 'signed in again, the private page renders')
  const cookieB = await sessionCookie(ctxA)
  c.ok(cookieB && cookieB !== cookieA, 'a new session cookie B, different from A')
  await shot(pageA, F, '03-signed-in-again')

  // 3. A second live session (the browser forgets cookie B, the server keeps it), then sign out everywhere.
  await ctxA.clearCookies()
  c.ok(await signIn(pageA, 'second session'), 'signed in a second time after clearing cookies')
  const cookieC = await sessionCookie(ctxA)
  const tokenUrl2 = await tokenUrlFor(pageA, art.id)
  await pageA.goto(`${V.publicApp}/account`)
  await pageA.locator('#sessions').filter({ hasText: /\d/ }).waitFor({ timeout: 5000 })
  const sessionsBefore = Number(await pageA.locator('#sessions').textContent())
  c.ok(sessionsBefore >= 2, '/account counts at least 2 sessions (B and C)', sessionsBefore)
  const liveB = await withCookie(pageUrl, cookieB)
  const liveC = await withCookie(pageUrl, cookieC)
  const tok2 = await withCookie(tokenUrl2, null)
  c.ok(liveB.status === 200 && liveC.status === 200 && tok2.status === 200, 'before #logout-all: cookies B and C and the second content-token URL all -> 200', `${liveB.status} ${liveC.status} ${tok2.status}`)
  log(F, `observation: the second content-token URL ${tokenUrl2 === tokenUrl1 ? 'equals' : 'differs from'} the first (same key, the token covers id and expiry)`)
  await shot(pageA, F, '04-account-two-sessions')
  pageA.once('dialog', (dialog) => dialog.accept())
  await pageA.locator('#logout-all').click()
  await pageA.waitForURL(`${V.publicApp}/login`, { timeout: 5000 })
  const deadB = await withCookie(pageUrl, cookieB)
  const deadC = await withCookie(pageUrl, cookieC)
  c.ok(deadB.status === 404 && deadC.status === 404, 'after #logout-all: replaying cookies B and C -> 404', `${deadB.status} ${deadC.status}`)
  const deadTok1 = await withCookie(tokenUrl1, null)
  const deadTok2 = await withCookie(tokenUrl2, null)
  c.ok(deadTok1.status === 404 && deadTok2.status === 404, 'after #logout-all: both captured content-token URLs -> 404 (content key rotated)', `${deadTok1.status} ${deadTok2.status}`)

  // 4. The passkey survives signing out everywhere.
  c.ok(await signIn(pageA, 'after logout-all'), 'the passkey still signs in after #logout-all')
  const tokenUrl3 = await tokenUrlFor(pageA, art.id)
  c.ok(tokenUrl3 !== tokenUrl1 && tokenUrl3 !== tokenUrl2, 'the frame now uses a different content token', redact(tokenUrl3))
  const tok3 = await withCookie(tokenUrl3, null)
  c.ok(tok3.status === 200 && tok3.body === DATA, 'the new content-token URL serves data.txt', tok3.status)
  await pageA.goto(`${V.publicApp}/account`)
  await pageA.locator('#sessions').filter({ hasText: /\d/ }).waitFor({ timeout: 5000 })
  const sessionsAfter = (await pageA.locator('#sessions').textContent()).trim()
  c.ok(sessionsAfter === '1', '/account counts exactly 1 session after #logout-all and one sign-in', sessionsAfter)
  await shot(pageA, F, '05-account-after-logout-all')

  // 5. A second device enrols, then device A removes its passkey.
  const ctxB = await browser.newContext({ viewport: { width: 1200, height: 800 } })
  const pageB = await ctxB.newPage()
  await enrolOwner(pageB, F, `device B ${nonce}`)
  const cookieDevB = await sessionCookie(ctxB)
  const tokenUrlB = await tokenUrlFor(pageB, art.id)
  const cookieA2 = await sessionCookie(ctxA)
  const liveDevB = await withCookie(pageUrl, cookieDevB)
  const liveTokB = await withCookie(tokenUrlB, null)
  c.ok(liveDevB.status === 200 && liveTokB.status === 200, 'device B is signed in: its cookie and content-token URL -> 200', `${liveDevB.status} ${liveTokB.status}`)
  await pageA.reload()
  const rowB = pageA.locator('#passkeys li', { hasText: `device B ${nonce}` })
  const rowA = pageA.locator('#passkeys li', { hasText: `device A ${nonce}` })
  await rowB.waitFor({ timeout: 5000 })
  const rows = await pageA.locator('#passkeys li').count()
  const disabledAll = await pageA.locator('#passkeys li button').evaluateAll((buttons) => buttons.map((b) => b.disabled))
  c.ok(rows >= 2 && disabledAll.every((d) => d === false), `/account lists ${rows} passkeys (data dir is shared with other journeys), each 刪除 enabled`, JSON.stringify(disabledAll))
  await shot(pageA, F, '06-account-two-devices')
  pageA.once('dialog', (dialog) => dialog.accept())
  await rowB.locator('button', { hasText: '刪除' }).click()
  await pageA.waitForURL(`${V.publicApp}/login?next=/account`, { timeout: 5000 })
  c.ok(true, 'removing device B lands device A on /login?next=/account (removal ends every session)')
  const deadDevB = await withCookie(pageUrl, cookieDevB)
  const deadA2 = await withCookie(pageUrl, cookieA2)
  const deadTokB = await withCookie(tokenUrlB, null)
  c.ok(deadDevB.status === 404 && deadA2.status === 404, 'after the removal: device B cookie and device A pre-removal cookie -> 404', `${deadDevB.status} ${deadA2.status}`)
  c.ok(deadTokB.status === 404, "after the removal: device B's content-token URL -> 404 (content key rotated)", deadTokB.status)
  await pageB.goto(`${V.publicApp}/login`)
  await pageB.locator('#go').click()
  await pageB.locator('#status', { hasText: '這把 passkey 沒有註冊在這裡' }).waitFor({ timeout: 5000 }).catch(() => {})
  const statusB = (await pageB.locator('#status').textContent()) ?? ''
  c.ok(statusB.includes('這把 passkey 沒有註冊在這裡') && pageB.url() === `${V.publicApp}/login`, 'device B can no longer sign in: #status 這把 passkey 沒有註冊在這裡, stays on /login', `${statusB} @ ${pageB.url()}`)
  await shot(pageB, F, '07-device-b-refused')
  await pageA.locator('#go').click()
  await pageA.waitForURL(`${V.publicApp}/account`, { timeout: 10_000 })
  await rowA.waitFor({ timeout: 5000 })
  const rowsAfter = await pageA.locator('#passkeys li').count()
  const goneB = await pageA.locator('#passkeys li', { hasText: `device B ${nonce}` }).count()
  const disabledAfter = await pageA.locator('#passkeys li button').evaluateAll((buttons) => buttons.map((b) => b.disabled))
  c.ok(goneB === 0 && rowsAfter === rows - 1, 'device A signs back in at /account; device B is gone, device A remains', `${rows} -> ${rowsAfter}`)
  c.ok(disabledAfter.every((d) => d === (rowsAfter <= 1)), '刪除 is disabled exactly when one passkey is left', `${rowsAfter} rows ${JSON.stringify(disabledAfter)}`)
  await shot(pageA, F, '08-account-after-removal')
  save(F, 'ids.json', { artifact: art.id })
  cli(F, ['delete', art.id, '--yes'], { quiet: true })
  await ctxB.close()
  await ctxA.close()
  c.finish()
} catch (error) {
  c.finish(error)
} finally {
  await browser?.close()
}

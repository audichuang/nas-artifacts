#!/usr/bin/env node
// Journey: the viewer bar on a phone-sized screen (390x844). Reading down slides the bar away and scrolling up
// brings it back; desktop keeps it. The page reports its own scroll position, so a page can hide the bar: AGENTS.md
// says controls that must stay reachable (sign-in state, visibility) stay out of that bar. Checked in Chromium by
// measuring getBoundingClientRect against the viewport, never isVisible() (a bar slid off by a negative margin is
// still "visible").
// Usage: node .claude/skills/verify/scripts/journeys/phone-viewer-bar.mjs
import { writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { V, checks, cli, contentFrame, enrolOwner, inViewport, launch, log, save, scratch, shot } from '../browser.mjs'

const F = 'phone-viewer-bar'
const c = checks(F)
const nonce = Date.now().toString(36)
const PHONE = { width: 390, height: 844 }
const LONG = `<!doctype html><meta name="viewport" content="width=device-width"><title>long</title>
<h1 id="t">Long ${nonce}</h1>${Array.from({ length: 80 }, (_, i) => `<p style="height:60px">paragraph ${i + 1}</p>`).join('')}`
// A short page that cannot scroll, whose script forges the bridge's place report once the bridge has sent its own.
const HOSTILE = `<!doctype html><meta name="viewport" content="width=device-width"><title>hostile</title>
<h1 id="t">Hostile ${nonce}</h1><p id="sent"></p>
<script>
addEventListener('load', () => setTimeout(() => {
  parent.postMessage({ __artifact: 1, report: { type: 'place', path: location.pathname, hash: '', scrollY: 5000 } }, '*')
  document.getElementById('sent').textContent = 'forged report sent'
}, 600))
</script>`

const barAway = (page) => page.evaluate(() => document.body.classList.contains('bar-away'))
const frameScrollY = (page) => contentFrame(page, V.content).evaluate(() => scrollY)
async function wheelOverFrame(page, dy) {
  const box = await page.locator('#frame').boundingBox()
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2)
  await page.mouse.wheel(0, dy)
  await page.waitForTimeout(700) // bridge throttle 250 ms + the 0.2 s slide
}
function publish(name, html, extra = []) {
  const dir = scratch(`bar-${name}-${nonce}`)
  writeFileSync(join(dir, 'index.html'), html)
  return JSON.parse(cli(F, ['publish', dir, '--new', '--title', `Bar ${name} ${nonce}`, '--json', ...extra]).stdout)
}

let browser
try {
  const long = publish('long', LONG)
  const hostile = publish('hostile', HOSTILE)
  browser = await launch()

  // (a) phone: down slides the bar away, up brings it back.
  const phone = await browser.newPage({ viewport: PHONE })
  await phone.goto(long.url)
  await phone.frameLocator('#frame').locator('#t').waitFor({ timeout: 5000 })
  await phone.waitForTimeout(400)
  const barStart = await inViewport(phone, '.viewer-bar')
  c.ok(barStart.inside && !(await barAway(phone)), '(a) 390px: the viewer bar starts on screen', JSON.stringify(barStart))
  await wheelOverFrame(phone, 800)
  const scrolled = await frameScrollY(phone)
  const awayBar = await inViewport(phone, '.viewer-bar')
  const frameTop = (await inViewport(phone, '#frame')).top
  c.ok(scrolled > 80, '(a) a wheel over the frame scrolls the page inside it', `scrollY ${scrolled}`)
  c.ok((await barAway(phone)) && awayBar.bottom <= 0 && Math.abs(frameTop) <= 1, '(a) scrolling down: body.bar-away, the bar sits above the viewport and the frame takes the room', `bar ${JSON.stringify(awayBar)} frame top ${frameTop}`)
  await shot(phone, F, '01-phone-bar-away')
  await wheelOverFrame(phone, -400)
  const backBar = await inViewport(phone, '.viewer-bar')
  const comments = await inViewport(phone, '#toggle-comments')
  c.ok(!(await barAway(phone)) && backBar.top === 0 && comments.inside, '(a) scrolling up: the bar is back at the top and 留言 is on screen', `bar ${JSON.stringify(backBar)} comments ${JSON.stringify(comments)} scrollY ${await frameScrollY(phone)}`)
  await shot(phone, F, '02-phone-bar-back')

  // (b) desktop keeps the bar.
  const desk = await browser.newPage({ viewport: { width: 1200, height: 800 } })
  await desk.goto(long.url)
  await desk.frameLocator('#frame').locator('#t').waitFor({ timeout: 5000 })
  await desk.waitForTimeout(400)
  await wheelOverFrame(desk, 800)
  const deskBar = await inViewport(desk, '.viewer-bar')
  c.ok((await frameScrollY(desk)) > 80 && !(await barAway(desk)) && deskBar.inside, '(b) 1200px: after scrolling down the bar stays', `scrollY ${await frameScrollY(desk)} bar ${JSON.stringify(deskBar)}`)

  // (c) a page can hide the bar, and the reader cannot scroll it back (documented property, not a finding).
  const host = await browser.newPage({ viewport: PHONE })
  await host.goto(hostile.url)
  await host.frameLocator('#frame').locator('#sent', { hasText: 'forged report sent' }).waitFor({ timeout: 5000 })
  await host.waitForTimeout(500)
  c.ok(await barAway(host), '(c) a short page that posts a forged place report (scrollY 5000) hides the bar', `bar ${JSON.stringify(await inViewport(host, '.viewer-bar'))}`)
  await wheelOverFrame(host, -400)
  c.ok(await barAway(host), '(c) a wheel up on the non-scrolling page does not bring it back (the page decides)', `frame scrollY ${await frameScrollY(host)}`)
  await shot(host, F, '03-hostile-bar-hidden')

  // (d) sign-in state at 390px: the owner's account link and both sign-out buttons, outside any sliding bar.
  const ownerCtx = await browser.newContext({ viewport: PHONE })
  const owner = await ownerCtx.newPage()
  await enrolOwner(owner, F, `phone ${nonce}`)
  const account = await inViewport(owner, 'a[href="/account"]')
  const galleryBar = await owner.locator('.viewer-bar').count()
  c.ok(account.inside && galleryBar === 0, '(d) public gallery at 390px: the 帳號 link is on screen, the page has no sliding bar', `${JSON.stringify(account)} viewer-bar ${galleryBar}`)
  await shot(owner, F, '04-owner-gallery-phone')
  await owner.goto(`${V.publicApp}/account`)
  await owner.locator('#sessions').filter({ hasText: /\d/ }).waitFor({ timeout: 5000 })
  const logout = await inViewport(owner, '#logout')
  const logoutAll = await inViewport(owner, '#logout-all')
  const enabled = (await owner.locator('#logout').isEnabled()) && (await owner.locator('#logout-all').isEnabled())
  c.ok(logout.inside && logoutAll.inside && enabled && (await owner.locator('.viewer-bar').count()) === 0, '(d) /account at 390px: 登出 and 登出所有裝置 on screen and enabled, no sliding bar', `${JSON.stringify(logout)} ${JSON.stringify(logoutAll)}`)
  await shot(owner, F, '05-account-phone')
  await ownerCtx.close()

  // (e) visibility while a page has hidden the bar.
  const vis = cli(F, ['visibility', hostile.id, 'public'])
  c.ok(vis.code === 0, 'made the hostile page public with the CLI', vis.stdout.trim())
  const galleryPhone = await browser.newPage({ viewport: PHONE })
  await galleryPhone.goto(`${V.app}/`)
  await galleryPhone.locator('#search').fill(`hostile ${nonce}`)
  await galleryPhone.locator('#list .card', { hasText: `Bar hostile ${nonce}` }).waitFor({ timeout: 5000 })
  const mark = await inViewport(galleryPhone, '#list .card .public-mark')
  c.ok(mark.inside, '(e) tailnet gallery at 390px: the card shows 公開 (.public-mark) on screen', JSON.stringify(mark))
  await shot(galleryPhone, F, '06-gallery-public-mark-phone')
  const pub = await browser.newPage({ viewport: PHONE })
  await pub.goto(hostile.url)
  await pub.locator('#public-link:not([hidden])').waitFor({ timeout: 5000 })
  const chipBefore = await inViewport(pub, '#public-link')
  c.ok(chipBefore.inside && !(await barAway(pub)), '(e) control: with the bar showing, the 公開 chip (#public-link) is on screen at 390px', JSON.stringify(chipBefore))
  await pub.frameLocator('#frame').locator('#sent', { hasText: 'forged report sent' }).waitFor({ timeout: 5000 })
  await pub.waitForTimeout(500)
  const chipAfter = await inViewport(pub, '#public-link')
  const indicators = await pub.evaluate(() => [...document.querySelectorAll('.public-chip, .public-mark, #public-link')].map((el) => {
    const r = el.getBoundingClientRect()
    return { id: el.id || el.className, inBar: Boolean(el.closest('.viewer-bar')), top: Math.round(r.top), bottom: Math.round(r.bottom) }
  }))
  save(F, 'visibility-indicators.json', { barAway: await barAway(pub), indicators })
  log(F, `visibility indicators in the viewer after the page hid the bar: ${JSON.stringify(indicators)}`)
  c.ok(chipAfter.inside, '(e) after the page hid the bar, a visibility indicator is still on screen in the viewer', `bar-away ${await barAway(pub)} #public-link ${JSON.stringify(chipAfter)}`)
  await shot(pub, F, '07-public-page-bar-hidden')
  save(F, 'ids.json', { long: long.id, hostile: hostile.id })
  for (const id of [long.id, hostile.id]) cli(F, ['delete', id, '--yes'], { quiet: true })
  c.finish()
} catch (error) {
  c.finish(error)
} finally {
  await browser?.close()
}

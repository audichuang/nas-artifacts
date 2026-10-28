import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { runInNewContext } from 'node:vm'
import { test } from 'node:test'

type Element = { textContent: string; disabled: boolean; hidden: boolean; href: string; onclick?: () => void; onload?: () => void; addEventListener: (name: string, fn: () => void) => void }
async function client() {
  const nodes = new Map<string, Element>()
  const element = (): Element => ({ textContent: '', disabled: false, hidden: false, href: '', addEventListener(name, fn) { if (name === 'click') this.onclick = fn } })
  const get = (id: string) => { if (!nodes.has(id)) nodes.set(id, element()); return nodes.get(id)! }
  get('page-data').textContent = JSON.stringify({ artifactId: 'a', sessionId: 's', officeUrl: 'http://office.test' })
  let events: Record<string, (event?: unknown) => void> = {}
  const calls: string[] = []
  const timers: (() => void)[] = []
  let state: Record<string, unknown> = { status: 'editing', closed: false }
  const redirects: string[] = []
  const requests: string[] = []
  await runInNewContext(`(async()=>{${readFileSync(new URL('../public/office-editor.js', import.meta.url), 'utf8')}\n})()`, {
    document: { getElementById: get, createElement: element, head: { append(script: Element) { queueMicrotask(() => script.onload?.()) } } },
    fetch: async (url: string) => { requests.push(url); return { ok: true, json: async () => url.endsWith('/config') ? {} : url.endsWith('/save') ? { queued: true } : state } },
    window: { DocsAPI: { DocEditor: class {
      constructor(_id: string, config: { events: typeof events }) { events = config.events }
      requestClose() { calls.push('requestClose') }
      destroyEditor() { calls.push('destroyEditor') }
    } } },
    location: { assign(url: string) { redirects.push(url) } },
    setTimeout: (fn: () => void) => { timers.push(fn) }, addEventListener() {}, queueMicrotask
  })
  await new Promise<void>((resolve) => setImmediate(resolve))
  const tick = async () => { timers.shift()?.(); await new Promise<void>((resolve) => setImmediate(resolve)) }
  return { get, calls, events, tick, redirects, requests, state(value: Record<string, unknown>) { state = value } }
}

test('return is disabled during loading and uses the SDK close guard even with unsynchronized edits', async () => {
  const c = await client()
  assert.equal(c.get('office-close').disabled, true)
  c.events.onDocumentReady()
  assert.equal(c.get('office-close').disabled, false)
  c.events.onDocumentStateChange({ data: true })
  c.get('office-close').onclick!()
  assert.deepEqual(c.calls, ['requestClose'])
  assert.equal(c.get('office-close').disabled, false)
  c.events.onRequestClose()
  assert.deepEqual(c.calls, ['requestClose', 'destroyEditor'])
  c.state({ status: 'saved', closed: true, savedVersion: 2 })
  await c.tick()
  assert.deepEqual(c.redirects, ['/a/a?v=2'])
})

test('cancelling an SDK close prompt leaves return available in manual-save mode', async () => {
  const c = await client()
  c.events.onDocumentReady()
  c.events.onDocumentStateChange({ data: true })
  c.get('office-close').onclick!()
  await c.tick()
  c.get('office-close').onclick!()
  assert.deepEqual(c.calls, ['requestClose', 'requestClose'])
  assert.equal(c.get('office-close').disabled, false)
})

test('polling survives a temporary closed status and resumes after reconnection', async () => {
  const c = await client()
  c.events.onDocumentReady()
  c.state({ status: 'closed', closed: true })
  await c.tick()
  c.state({ status: 'editing', closed: false })
  await c.tick()
  assert.equal(c.get('office-close').disabled, false)
  c.get('office-close').onclick!()
  assert.deepEqual(c.calls, ['requestClose'])
  assert.deepEqual(c.redirects, [])
})

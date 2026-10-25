import assert from 'node:assert/strict'
import { describe, test } from 'node:test'
import { parseXml } from '../src/xml.ts'

describe('reading Office XML', () => {
  test('attributes read as before: prefixes, either quote, entities, spaces around =, stray tokens skipped', () => {
    const node = parseXml(`<a:x r:embed="r1" b = 'two &amp; three' stray c="&lt;4&gt;" d='"q"' e="it's"/>`)
    assert.deepEqual(node.attrs, { 'r:embed': 'r1', b: 'two & three', c: '<4>', d: '"q"', e: "it's" })
    assert.deepEqual(parseXml('<a x="1"y="2"></a>').attrs, { x: '1', y: '2' })
  })

  test('a long tag with no = parses in linear time', () => {
    const started = performance.now()
    const node = parseXml(`<a ${'x'.repeat(200_000)} b="1"/>`)
    assert.equal(node.attrs.b, '1')
    assert.ok(performance.now() - started < 1000, `took ${Math.round(performance.now() - started)} ms`)
  })
})

import assert from 'node:assert/strict'
import { describe, test } from 'node:test'
import { MAX_DEPTH, parseXml, tagAttrs, XmlError } from '../src/xml.ts'

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

  test('the attributes of a raw start tag follow the same rules', () => {
    assert.deepEqual(tagAttrs(`<row r='3' ht = "20" hidden>`), { r: '3', ht: '20' })
    assert.deepEqual(tagAttrs('<x:row\nr="9"/>'), { r: '9' })
    assert.deepEqual(tagAttrs('<row>'), {})
  })

  test('a document read only in part closes what is still open; a whole one must close everything', () => {
    const cut = parseXml('<worksheet><sheetViews><sheetView a="1"/></sheetViews><cols><col min="1"', undefined, true)
    assert.equal(cut.name, 'worksheet')
    assert.equal(cut.children[0].children[0].attrs.a, '1')
    assert.throws(() => parseXml('<worksheet><cols>'), /XML 的 <cols> 沒有結束/)
  })

  test('deep nesting is refused with a message, not a stack overflow later', () => {
    const deep = (n: number) => '<g>'.repeat(n) + '</g>'.repeat(n)
    assert.equal(parseXml(deep(MAX_DEPTH)).name, 'g')
    assert.throws(() => parseXml(deep(MAX_DEPTH + 1)), (error: Error) => error instanceof XmlError && /巢狀超過 1000 層/.test(error.message))
  })

  test('every error a reader may see is in Traditional Chinese', () => {
    for (const [source, message] of [
      ['<!DOCTYPE x><x/>', /不接受含 DTD 的 XML/],
      ['<a><!-- open', /XML 標記沒有結束/],
      ['<a><![CDATA[x', /CDATA 區段沒有結束/],
      ['<a></b>', /XML 多出了 <\/b>/],
      ['<a b="1"', /XML 標籤沒有結束/],
      ['', /XML 是空的/]
    ] as const) {
      assert.throws(() => parseXml(source), (error: Error) => error instanceof XmlError && message.test(error.message), source)
    }
    assert.throws(() => parseXml('<a><b/><b/></a>', 2), /XML 內容太多/)
  })
})

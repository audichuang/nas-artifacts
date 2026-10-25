import assert from 'node:assert/strict'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { after, before, describe, test } from 'node:test'
import pptxgen from 'pptxgenjs'
import { createAppServer } from '../src/app-server.ts'
import { ArtifactStore } from '../src/artifact-store.ts'
import { createContentServer } from '../src/content-server.ts'
import { readPptx } from '../src/pptx.ts'
import { parseXml } from '../src/xml.ts'
import { ZipReader } from '../src/zip.ts'
import { zip } from './helpers/zip.ts'

const PptxGenJS = ((pptxgen as unknown as { default?: typeof pptxgen.default }).default ?? pptxgen) as typeof pptxgen.default

const PNG_DATA = 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='

/** A deck built the way other tools build them: its own master, placeholders, a table, charts, shapes, a hidden slide. */
async function sampleDeck(): Promise<Buffer> {
  const pptx = new PptxGenJS()
  pptx.layout = 'LAYOUT_4x3'
  pptx.defineSlideMaster({
    title: 'M',
    background: { color: '0B3D2E' },
    objects: [
      { rect: { x: 0, y: 6.8, w: 10, h: 0.7, fill: { color: 'F2C14E' } } },
      { placeholder: { options: { name: 'body', type: 'body', x: 0.6, y: 1.6, w: 8.8, h: 4.6, color: 'FFFFFF', fontSize: 20 }, text: '' } },
      { placeholder: { options: { name: 'title', type: 'title', x: 0.6, y: 0.4, w: 8.8, h: 1, color: 'F2C14E', fontSize: 34 }, text: '' } }
    ]
  })
  let slide = pptx.addSlide({ masterName: 'M' })
  slide.addText('母片的標題', { placeholder: 'title' })
  slide.addText(
    [
      { text: '第一點', options: { bullet: { type: 'number' } } },
      { text: '第二點', options: { bullet: { type: 'number' } } }
    ],
    { placeholder: 'body' }
  )
  slide.addNotes('講稿 <一>')
  slide = pptx.addSlide()
  slide.addTable([[{ text: '項目', options: { fill: { color: '1F4E79' }, color: 'FFFFFF' } }, { text: '數量' }], [{ text: '蘋果' }, { text: '3' }]], { x: 0.5, y: 0.5, w: 5 })
  slide.addChart(pptx.ChartType.bar, [{ name: '銷售', labels: ['一月', '二月'], values: [3, 7] }], { x: 5.6, y: 0.5, w: 4, h: 3, showTitle: true, title: '季度' })
  slide.addChart(pptx.ChartType.line, [{ name: '趨勢', labels: ['春', '夏'], values: [1, 2] }, { name: '目標', labels: ['春', '夏'], values: [2, 2] }], { x: 5.6, y: 3.6, w: 4, h: 2.5, showTitle: true, title: '走勢' })
  slide.addShape(pptx.ShapeType.rightArrow, { x: 0.5, y: 3, w: 2, h: 1, fill: { color: 'E74C3C' }, rotate: 20 })
  slide.addImage({ data: PNG_DATA, x: 3, y: 3, w: 1, h: 1 })
  slide = pptx.addSlide()
  slide.hidden = true
  slide.addText('隱藏頁', { x: 1, y: 1, w: 5, h: 1 })
  slide = pptx.addSlide()
  slide.addText('回第一頁', { x: 1, y: 1, w: 5, h: 1, hyperlink: { slide: 1 } })
  return (await pptx.write({ outputType: 'nodebuffer' })) as Buffer
}

const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
const rels = (list: [string, string, string][]) =>
  `<Relationships>${list.map(([id, type, target]) => `<Relationship Id="${id}" Type="http://x/${type}" Target="${target}"/>`).join('')}</Relationships>`
const textShape = (text: string) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="2" name="t"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`

/** Two slides; the second names no layout, which PowerPoint would never write. */
function handBuiltDeck(): Buffer {
  return zip({
    '_rels/.rels': rels([['r1', 'officeDocument', 'ppt/presentation.xml']]),
    'ppt/presentation.xml': `<p:presentation ${NS}><p:sldIdLst><p:sldId id="256" r:id="s1"/><p:sldId id="257" r:id="s2"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>`,
    'ppt/_rels/presentation.xml.rels': rels([['s1', 'slide', 'slides/slide1.xml'], ['s2', 'slide', 'slides/slide2.xml']]),
    'ppt/slides/slide1.xml': `<p:sld ${NS}><p:cSld><p:spTree>${textShape('第一頁 &amp; &lt;b&gt;')}</p:spTree></p:cSld></p:sld>`,
    'ppt/slides/_rels/slide1.xml.rels': rels([['l', 'slideLayout', '../slideLayouts/slideLayout1.xml']]),
    'ppt/slideLayouts/slideLayout1.xml': `<p:sldLayout ${NS}><p:cSld><p:spTree/></p:cSld></p:sldLayout>`,
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels': rels([['m', 'slideMaster', '../slideMasters/slideMaster1.xml']]),
    'ppt/slideMasters/slideMaster1.xml': `<p:sldMaster ${NS}><p:cSld><p:spTree/></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"/></p:sldMaster>`,
    'ppt/slides/slide2.xml': `<p:sld ${NS}><p:cSld><p:spTree>${textShape('第二頁')}</p:spTree></p:cSld></p:sld>`
  })
}

/** PowerPoint's own defaults: dark text in the master's otherStyle, a shape whose style says white text, a table style with a white header. */
function themedDeck(): Buffer {
  const theme = `<a:theme ${NS}><a:themeElements><a:clrScheme name="c"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="111111"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1></a:clrScheme><a:fontScheme name="f"><a:majorFont><a:latin typeface="Calibri"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme><a:fmtScheme name="s"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`
  const shape = `<p:sp><p:nvSpPr><p:cNvPr id="2" name="s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr><p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style><p:txBody><a:bodyPr/><a:p><a:r><a:t>白字形狀</a:t></a:r></a:p></p:txBody></p:sp>`
  const table = `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="t"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="1828800"/><a:ext cx="1828800" cy="740000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{S}</a:tableStyleId></a:tblPr><a:tblGrid><a:gridCol w="1828800"/></a:tblGrid><a:tr h="370000"><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>標題列</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr><a:tr h="370000"><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>內容列</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>`
  const tableStyles = `<a:tblStyleLst ${NS} def="{S}"><a:tblStyle styleId="{S}" styleName="m"><a:wholeTbl><a:tcTxStyle><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="dk1"/></a:tcTxStyle><a:tcStyle><a:fill><a:solidFill><a:schemeClr val="accent1"><a:tint val="20000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:wholeTbl><a:firstRow><a:tcTxStyle b="on"><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:firstRow></a:tblStyle></a:tblStyleLst>`
  return zip({
    '_rels/.rels': rels([['r1', 'officeDocument', 'ppt/presentation.xml']]),
    'ppt/presentation.xml': `<p:presentation ${NS}><p:sldIdLst><p:sldId id="256" r:id="s1"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>`,
    'ppt/_rels/presentation.xml.rels': rels([['s1', 'slide', 'slides/slide1.xml'], ['ts', 'tableStyles', 'tableStyles.xml']]),
    'ppt/tableStyles.xml': tableStyles,
    'ppt/slides/slide1.xml': `<p:sld ${NS}><p:cSld><p:spTree>${shape}${table}</p:spTree></p:cSld></p:sld>`,
    'ppt/slides/_rels/slide1.xml.rels': rels([['l', 'slideLayout', '../slideLayouts/slideLayout1.xml']]),
    'ppt/slideLayouts/slideLayout1.xml': `<p:sldLayout ${NS}><p:cSld><p:spTree/></p:cSld></p:sldLayout>`,
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels': rels([['m', 'slideMaster', '../slideMasters/slideMaster1.xml']]),
    'ppt/slideMasters/slideMaster1.xml': `<p:sldMaster ${NS}><p:cSld><p:spTree/></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"/><p:txStyles><p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:defRPr></a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>`,
    'ppt/slideMasters/_rels/slideMaster1.xml.rels': rels([['t', 'theme', '../theme/theme1.xml']]),
    'ppt/theme/theme1.xml': theme
  })
}

describe('reading PowerPoint files', () => {
  test('any deck draws from its own master, layouts and theme: placeholders, numbering, tables, charts, notes', async () => {
    const deck = readPptx(await sampleDeck(), 'm/')
    // 4:3, and the hidden slide is skipped as a slide show skips it.
    assert.equal(deck.width / deck.height, 4 / 3)
    assert.equal(deck.slides.length, 3)
    const [first, second, third] = deck.slides
    // The title placeholder takes its colour from the master it was defined on.
    assert.match(first.html, /color:rgb\(242,193,78\)[^>]*>母片的標題</)
    assert.match(first.html, /class="pp-bu"[^>]*>1\.<\/span>[\s\S]*class="pp-bu"[^>]*>2\.<\/span>/)
    assert.match(first.html, /<aside class="slide-notes" hidden><p>講稿 &#60;一&#62;<\/p><\/aside>/)
    assert.match(first.html, /^<section class="slide layout-pptx"[^>]*style="background:rgb\(11,61,46\)"/)
    assert.match(second.html, /<table class="pp-tbl">[\s\S]*項目[\s\S]*蘋果/)
    assert.match(second.html, /<svg class="pp-geom"[^>]*aria-label="季度"[\s\S]*>一月</)
    // Line charts keep their categories in a multi-level cache.
    assert.match(second.html, /aria-label="走勢"[\s\S]*>春<\/text>/)
    assert.doesNotMatch(second.html, /NaN/)
    assert.match(second.html, /transform:rotate\(20deg\)/)
    assert.match(second.html, /<img src="m\/ppt\/media\/[^"]+\.png"/)
    assert.match(third.html, /<a href="#1"/)
    assert.doesNotMatch(deck.slides.map((s) => s.html).join(''), /隱藏頁/)
  })

  test('a slide that cannot be read says so on its own face; the others still draw, text escaped', () => {
    const deck = readPptx(handBuiltDeck(), '')
    assert.equal(deck.slides.length, 2)
    assert.match(deck.slides[0].html, />第一頁 &#38; &#60;b&#62;</)
    assert.match(deck.slides[1].html, /class="pp-missing">第 2 頁無法顯示/)
  })

  test("a shape's style and a table style colour text over the master's default text colour", () => {
    const [slide] = readPptx(themedDeck(), '').slides
    assert.match(slide.html, /color:rgb\(255,255,255\)[^>]*>白字形狀</)
    assert.match(slide.html, /font-weight:700;color:rgb\(255,255,255\)[^>]*>標題列</)
    assert.match(slide.html, /color:rgb\(0,0,0\)[^>]*>內容列</)
  })

  test('Office parts never expand a DTD; a file that is not a zip is refused', () => {
    assert.throws(() => parseXml('<!DOCTYPE x [<!ENTITY a "aaaa">]><x>&a;</x>'), /DTD/)
    assert.throws(() => parseXml('<a><b></a>'), /unexpected/)
    assert.throws(() => new ZipReader(Buffer.from('not a zip')), /not a zip/)
    assert.equal(parseXml('<a t="1 &gt; 0"><b>x &amp; y</b></a>').children[0].text, 'x & y')
  })
})

describe('serving PowerPoint files', () => {
  let dir: string
  let store: ArtifactStore
  let content: ReturnType<typeof createContentServer>
  let app: ReturnType<typeof createAppServer>
  let bytes: Buffer
  before(async () => {
    dir = mkdtempSync(join(tmpdir(), 'pptx-test-'))
    store = new ArtifactStore(dir)
    const config = { dataDir: dir, appPort: 0, contentPort: 0, appOrigin: 'http://app.test', contentOrigin: 'http://content.test', tokens: ['t'] }
    content = createContentServer(store, config)
    app = createAppServer(store, config)
    bytes = await sampleDeck()
  })
  after(() => {
    store.close()
    rmSync(dir, { recursive: true, force: true })
  })

  test('a published .pptx is the entry, renders as a deck at its own size, and is searchable by its text', async () => {
    const { id, entry } = store.publish({ title: 'Q3 報告', files: { '季報 (final).pptx': { content: bytes } } })
    assert.equal(entry, '季報 (final).pptx')
    const response = await content.request(`http://content.test/a/${id}/v/1/`)
    assert.equal(response.status, 200)
    assert.equal(response.headers.get('cache-control'), 'no-cache')
    const html = await response.text()
    assert.match(html, /<body class="deck">/)
    assert.equal(html.match(/<section class="slide layout-pptx"/g)?.length, 3)
    assert.match(html, /--aspect:1\.333333/)
    // Picture URLs stay relative, inside the version's own path, and cannot close CSS url('…').
    const src = /<img src="([^"]+)"/.exec(html)![1]
    assert.match(src, /^~office\/%E5%AD%A3%E5%A0%B1%20%28final%29\.pptx\/ppt\/media\//)
    const picture = await content.request(`http://content.test/a/${id}/v/1/${src}`)
    assert.equal(picture.status, 200)
    assert.equal(picture.headers.get('content-type'), 'image/png')
    assert.match(picture.headers.get('content-security-policy') ?? '', /sandbox allow-scripts/)
    assert.deepEqual(store.list({ query: '母片的標題' }).artifacts.map((a) => a.id), [id])
  })

  test('the picture route serves only pictures inside a .pptx of that version', async () => {
    const { id } = store.publish({ title: 'D', files: { 'deck.pptx': { content: bytes }, 'a.txt': { content: Buffer.from('x') } } })
    for (const path of ['~office/deck.pptx/ppt/presentation.xml', '~office/deck.pptx/ppt/slides/slide1.xml', '~office/a.txt/ppt/media/x.png', '~office/other.pptx/ppt/media/image-2-1.png', '~office/deck.pptx/ppt/media/missing.png']) {
      assert.equal((await content.request(`http://content.test/a/${id}/v/1/${path}`)).status, 404, path)
    }
  })

  test('a deck in a folder finds its pictures beside it', async () => {
    const { id } = store.publish({ title: 'D', entry: 'talks/q3.pptx', files: { 'talks/q3.pptx': { content: bytes }, 'index.html': { content: Buffer.from('<p>x') } } })
    const html = await (await content.request(`http://content.test/a/${id}/v/1/talks/q3.pptx`)).text()
    const src = /<img src="([^"]+)"/.exec(html)![1]
    assert.equal((await content.request(`http://content.test/a/${id}/v/1/talks/${src}`)).status, 200)
    assert.equal((await content.request(`http://content.test/a/${id}/v/1/${src}`)).status, 404)
  })

  test('a file that is not a PowerPoint deck still publishes; its page says why and throws for the error log', async () => {
    const { id } = store.publish({ title: 'Broken', files: { 'broken.pptx': { content: Buffer.from('PK nope') } } })
    const html = await (await content.request(`http://content.test/a/${id}/v/1/`)).text()
    assert.match(html, /class="pp-missing">PowerPoint: not a zip file/)
    assert.match(html, /<script>throw new Error\("PowerPoint: not a zip file"\)<\/script>/)
  })

  test('a .pptx opened as a top-level document goes back into the viewer', async () => {
    const { id } = store.publish({ title: 'D', files: { 'deck.pptx': { content: bytes } } })
    const response = await content.request(`http://content.test/a/${id}/v/1/deck.pptx`, { headers: { 'sec-fetch-dest': 'document' } })
    assert.equal(response.status, 302)
    assert.equal(response.headers.get('location'), `http://app.test/a/${id}?v=1`)
  })

  test('PPTX export hands back the published file unchanged; the viewer cannot edit it as Markdown', async () => {
    const { id } = store.publish({ title: 'D', files: { 'deck.pptx': { content: bytes } } })
    const exported = await app.request(`http://app.test/api/artifacts/${id}/export.pptx`)
    assert.equal(exported.status, 200)
    assert.match(exported.headers.get('content-disposition') ?? '', /^attachment;/)
    assert.ok(Buffer.from(await exported.arrayBuffer()).equals(bytes))
    const edit = await app.request(`http://app.test/api/artifacts/${id}/slides`, {
      method: 'PUT',
      headers: { 'content-type': 'application/json', 'x-artifact-client': 'viewer' },
      body: JSON.stringify({ markdown: '# x', baseVersion: 1 })
    })
    assert.equal(edit.status, 409)
  })
})

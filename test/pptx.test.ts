import assert from 'node:assert/strict'
import { execFileSync } from 'node:child_process'
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { after, before, describe, test } from 'node:test'
import pptxgen from 'pptxgenjs'
import { createAppServer } from '../src/app-server.ts'
import { ArtifactStore } from '../src/artifact-store.ts'
import { createContentServer } from '../src/content-server.ts'
import { readPptx, renderPptxDeck } from '../src/pptx.ts'
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

const THEME = `<a:theme ${NS}><a:themeElements><a:clrScheme name="c"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:accent1><a:srgbClr val="4472C4"/></a:accent1></a:clrScheme><a:fontScheme name="f"><a:majorFont><a:latin typeface="Calibri"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme><a:fmtScheme name="s"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"/></a:gs></a:gsLst><a:lin ang="5400000"/></a:gradFill><a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"/></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="50000"/></a:schemeClr></a:gs></a:gsLst></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`

type Rels = [string, string, string][]

/** A deck of hand-written slides on one plain layout and master; `parts` adds or replaces any part. */
function deckOf(slides: string[], parts: Record<string, string> = {}, slideRels: Record<number, Rels> = {}): Buffer {
  const files: Record<string, string> = {
    '_rels/.rels': rels([['r1', 'officeDocument', 'ppt/presentation.xml']]),
    'ppt/presentation.xml': `<p:presentation ${NS}><p:sldIdLst>${slides.map((_, at) => `<p:sldId id="${256 + at}" r:id="s${at + 1}"/>`).join('')}</p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>`,
    'ppt/_rels/presentation.xml.rels': rels(slides.map((_, at) => [`s${at + 1}`, 'slide', `slides/slide${at + 1}.xml`])),
    'ppt/slideLayouts/slideLayout1.xml': `<p:sldLayout ${NS}><p:cSld><p:spTree/></p:cSld></p:sldLayout>`,
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels': rels([['m', 'slideMaster', '../slideMasters/slideMaster1.xml']]),
    'ppt/slideMasters/slideMaster1.xml': `<p:sldMaster ${NS}><p:cSld><p:spTree/></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"/></p:sldMaster>`,
    'ppt/slideMasters/_rels/slideMaster1.xml.rels': rels([['t', 'theme', '../theme/theme1.xml']]),
    'ppt/theme/theme1.xml': THEME
  }
  slides.forEach((slide, at) => {
    files[`ppt/slides/slide${at + 1}.xml`] = slide.startsWith('<p:sld') || slide.startsWith('<!') ? slide : `<p:sld ${NS}><p:cSld><p:spTree>${slide}</p:spTree></p:cSld></p:sld>`
    files[`ppt/slides/_rels/slide${at + 1}.xml.rels`] = rels([['l', 'slideLayout', '../slideLayouts/slideLayout1.xml'], ...(slideRels[at + 1] ?? [])])
  })
  return zip({ ...files, ...parts })
}

const xfrm = (x = 0, y = 0, cx = 914400, cy = 914400, extra = '') => `<a:xfrm${extra}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`
const shape = (spPr: string, body = '', style = '') =>
  `<p:sp><p:nvSpPr><p:cNvPr id="2" name="s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${spPr}</p:spPr>${style}${body ? `<p:txBody><a:bodyPr/>${body}</p:txBody>` : ''}</p:sp>`
const frame = (kind: string, inner: string) =>
  `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="g"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="5486400" cy="3657600"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/${kind}">${inner}</a:graphicData></a:graphic></p:graphicFrame>`
const chartXml = (group: string, grouping: string, series: [string, number[]][]) =>
  `<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:${group}>${group === 'barChart' ? '<c:barDir val="col"/>' : ''}<c:grouping val="${grouping}"/>${series
    .map(([color, values]) => `<c:ser><c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr><c:cat><c:strRef><c:strCache><c:ptCount val="${values.length}"/>${values.map((_, at) => `<c:pt idx="${at}"><c:v>k${at}</c:v></c:pt>`).join('')}</c:strCache></c:strRef></c:cat><c:val><c:numRef><c:numCache><c:ptCount val="${values.length}"/>${values.map((v, at) => `<c:pt idx="${at}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numCache></c:numRef></c:val></c:ser>`)
    .join('')}</c:${group}></c:plotArea></c:chart></c:chartSpace>`
const chartDeck = (chart: string) =>
  deckOf([frame('chart', '<c:chart xmlns:c="c" r:id="c1"/>')], { 'ppt/charts/chart1.xml': chart }, { 1: [['c1', 'chart', '../charts/chart1.xml']] })

/** The y of the gridline an axis labels with `label`. */
const tickY = (html: string, label: string) => Number(new RegExp(`<line [^>]*y1="([-\\d.]+)"[^>]*/><text[^>]*>${label.replace('%', '%')}</text>`).exec(html)?.[1])

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

describe('crafted and unusual PowerPoint files', () => {
  test('a chart value near the largest number ends; the process does not run out of memory', () => {
    const dir = mkdtempSync(join(tmpdir(), 'pptx-huge-'))
    try {
      const file = join(dir, 'huge.pptx')
      writeFileSync(file, chartDeck(chartXml('barChart', 'clustered', [['FF0000', [1.7e308]]])))
      const script = `import { readFileSync } from 'node:fs'; const { readPptx } = await import(${JSON.stringify(new URL('../src/pptx.ts', import.meta.url).href)}); process.stdout.write(String(readPptx(readFileSync(process.argv[1]), '').slides[0].html.length))`
      const length = Number(execFileSync(process.execPath, ['--max-old-space-size=128', '--input-type=module', '-e', script, file], { timeout: 15_000, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] }))
      assert.ok(length > 0 && length < 100_000, String(length))
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test('a numbered list cannot start at a number that makes its label huge', () => {
    const body = `<a:p><a:pPr><a:buAutoNum type="romanUcPeriod" startAt="100000000"/></a:pPr><a:r><a:t>第一點</a:t></a:r></a:p>`
    const [slide] = readPptx(deckOf([shape(xfrm(), body)]), '').slides
    assert.ok(slide.html.length < 20_000, String(slide.html.length))
    assert.match(slide.html, /class="pp-bu"[^>]*>[^<]{1,16}<\/span>[^]*第一點/)
  })

  test('a SmartArt drawing that leads back to itself draws once, as a labelled box inside', () => {
    const smartArt = frame('diagram', '<dgm:relIds xmlns:dgm="d" r:dm="dm" r:lo="lo" r:qs="qs" r:cs="cs"/>')
    const parts = {
      'ppt/diagrams/data1.xml': `<dgm:dataModel xmlns:dgm="d"><dgm:extLst><a:ext ${NS}><dsp:dataModelExt xmlns:dsp="s" relId="dr"/></a:ext></dgm:extLst></dgm:dataModel>`,
      'ppt/diagrams/drawing1.xml': `<dsp:drawing xmlns:dsp="s" ${NS}><dsp:spTree>${smartArt}${textShape('圖形文字')}</dsp:spTree></dsp:drawing>`,
      'ppt/diagrams/_rels/drawing1.xml.rels': rels([['dm', 'diagramData', 'data1.xml'], ['dr', 'diagramDrawing', 'drawing1.xml']])
    }
    const started = performance.now()
    const [slide] = readPptx(deckOf([smartArt], parts, { 1: [['dm', 'diagramData', '../diagrams/data1.xml'], ['dr', 'diagramDrawing', '../diagrams/drawing1.xml']] }), '').slides
    assert.ok(performance.now() - started < 2000)
    assert.doesNotMatch(slide.html, /call stack|Maximum/)
    assert.equal(slide.html.match(/圖形文字/g)?.length, 1)
    assert.match(slide.html, /class="pp-missing">SmartArt 圖形/)
  })

  test('a table larger than PowerPoint allows draws its first 75 × 75 cells and says the rest is in the file', () => {
    const row = `<a:tr h="100">${'<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>x</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>'.repeat(100)}</a:tr>`
    const table = frame('table', `<a:tbl><a:tblPr/><a:tblGrid>${'<a:gridCol w="1000"/>'.repeat(100)}</a:tblGrid>${row.repeat(100)}</a:tbl>`)
    const [slide] = readPptx(deckOf([table]), '').slides
    assert.equal(slide.html.match(/<td/g)?.length, 75 * 75 + 1)
    assert.equal(slide.html.match(/<col /g)?.length, 75)
    assert.match(slide.html, /表格太大,只顯示前 75 列、75 欄/)
  })

  test('a slide with more content than a page can hold stops at a budget and says so', () => {
    const body = '<a:p><a:r><a:t>x</a:t></a:r></a:p>'.repeat(150_000)
    const deck = readPptx(deckOf([shape(xfrm(), body), textShape('第二頁')]), '')
    const html = deck.slides.map((slide) => slide.html).join('')
    assert.ok(html.length < 25_000_000, String(html.length))
    assert.match(deck.slides[0].html, /內容太多,其餘部分請下載 PPTX 查看/)
    assert.match(deck.slides[1].html, /第二頁/)
  })

  test('one paragraph with more runs than a page can hold stops at the budget too', () => {
    const body = `<a:p>${'<a:r><a:t>x</a:t></a:r>'.repeat(300_000)}</a:p>`
    const [slide] = readPptx(deckOf([shape(xfrm(), body)]), '').slides
    assert.ok(slide.html.length < 5_000_000, String(slide.html.length))
    assert.match(slide.html, /內容太多,其餘部分請下載 PPTX 查看/)
  })

  test('a slide, notes, theme or table style that cannot be read leaves the rest of the deck drawing', () => {
    const deck = readPptx(
      deckOf([textShape('一'), `<p:sld ${NS}><p:cSld><p:spTree></p:cSld></p:sld>`, '<!DOCTYPE x><p:sld/>', textShape('四')], {
        'ppt/notesSlides/notesSlide1.xml': '<p:notes><broken>',
        'ppt/tableStyles.xml': '<a:tblStyleLst><x></a:tblStyleLst>',
        'ppt/_rels/presentation.xml.rels': rels([...[1, 2, 3, 4].map((n): [string, string, string] => [`s${n}`, 'slide', `slides/slide${n}.xml`]), ['ts', 'tableStyles', 'tableStyles.xml']]),
        'docProps/core.xml': '<cp:coreProperties><dc:title>t</cp:coreProperties>'
      }, { 1: [['n', 'notesSlide', '../notesSlides/notesSlide1.xml']] }),
      ''
    )
    assert.equal(deck.slides.length, 4)
    assert.match(deck.slides[0].html, />一</)
    assert.match(deck.slides[1].html, /class="pp-missing">第 2 頁無法顯示/)
    assert.match(deck.slides[2].html, /class="pp-missing">第 3 頁無法顯示/)
    assert.match(deck.slides[3].html, />四</)
    assert.match(deck.slides[3].html, /data-number="4"/)
  })

  test('a password-protected or old .ppt file names the cause in Chinese and points to the download', () => {
    const cfb = Buffer.concat([Buffer.from('d0cf11e0a1b11ae1', 'hex'), Buffer.alloc(504)])
    assert.throws(() => readPptx(cfb, ''), /舊版 PowerPoint \(\.ppt\) 或受密碼保護/)
    const page = renderPptxDeck(cfb, 'x.pptx', 'x', { css: '', js: '' })
    assert.match(page, /這份 PowerPoint 檔無法顯示/)
    assert.match(page, /舊版 PowerPoint \(\.ppt\) 或受密碼保護/)
    assert.match(page, /可以下載原始檔,用 PowerPoint 開啟查看/)
    const truncated = renderPptxDeck(Buffer.from('PK\x03\x04 truncated'), 'x.pptx', 'x', { css: '', js: '' })
    assert.match(truncated, /檔案不是有效的 \.pptx/)
    assert.match(truncated, /<script>throw new Error\("PowerPoint: not a zip file"\)<\/script>/)
  })

  test('a line or connector with no height or no width keeps its stroke', () => {
    const ln = '<a:ln w="50800"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>'
    const flat = shape(`${xfrm(0, 0, 4000000, 0)}<a:prstGeom prst="line"/>${ln}`)
    const upright = shape(`${xfrm(0, 0, 0, 3000000)}<a:prstGeom prst="straightConnector1"/>${ln}`)
    const [slide] = readPptx(deckOf([flat + upright]), '').slides
    assert.doesNotMatch(slide.html, /viewBox="[^"]*0\.01/)
    // The 4 pt stroke gets a 4 pt box, centred on the line, with the same scale on both axes.
    assert.match(slide.html, /viewBox="0 -2 314\.961 4"[^>]*style="top:calc\(-2\*var\(--pt\)\);height:calc\(4\*var\(--pt\)\)/)
    assert.match(slide.html, /viewBox="-2 0 4 236\.22"[^>]*style="left:calc\(-2\*var\(--pt\)\);width:calc\(4\*var\(--pt\)\)/)
  })

  test("a path shape filled from the theme's gradient style is filled, not hollow", () => {
    const style = '<p:style><a:lnRef idx="1"><a:srgbClr val="00FF00"/></a:lnRef><a:fillRef idx="3"><a:srgbClr val="FF0000"/></a:fillRef><a:effectRef idx="0"><a:srgbClr val="FF0000"/></a:effectRef><a:fontRef idx="minor"/></p:style>'
    const [slide] = readPptx(deckOf([shape(`${xfrm()}<a:prstGeom prst="triangle"/>`, '', style)]), '').slides
    assert.match(slide.html, /<path d="M[^"]*" fill="rgb\((255,0,0|128,0,0)\)"/)
  })

  test('stacked charts stack: negatives below zero, line and area on top of each other, 100% stacked up to 100%', () => {
    const bar = readPptx(chartDeck(chartXml('barChart', 'stacked', [['FF0000', [5]], ['00FF00', [-2]]])), '').slides[0].html
    const negative = /<rect x="[^"]*" y="([-\d.]+)"[^>]*fill="rgb\(0,255,0\)"/.exec(bar)
    assert.equal(Number(negative?.[1]), tickY(bar, '0'))
    const line = readPptx(chartDeck(chartXml('lineChart', 'stacked', [['FF0000', [1, 1]], ['00FF00', [2, 2]]])), '').slides[0].html
    const second = /<polyline points="[-\d.]+,([-\d.]+)[^"]*" fill="none" stroke="rgb\(0,255,0\)"/.exec(line)
    assert.equal(Number(second?.[1]), tickY(line, '3'))
    const area = readPptx(chartDeck(chartXml('areaChart', 'stacked', [['FF0000', [1, 1]], ['00FF00', [2, 2]]])), '').slides[0].html
    const top = /<path d="M[-\d.]+,([-\d.]+)[^"]*" fill="rgb\(0,255,0\)"/.exec(area)
    assert.equal(Number(top?.[1]), tickY(area, '3'))
    const percent = readPptx(chartDeck(chartXml('barChart', 'percentStacked', [['FF0000', [1]], ['00FF00', [3]]])), '').slides[0].html
    assert.ok(Number.isFinite(tickY(percent, '100%')))
    assert.doesNotMatch(percent, />(200|300|400)%?</)
    const full = /<rect x="[^"]*" y="([-\d.]+)"[^>]*fill="rgb\(0,255,0\)"/.exec(percent)
    assert.equal(Number(full?.[1]), tickY(percent, '100%'))
  })

  test('text in a flipped group stays readable; the group still mirrors its shapes', () => {
    const group = (flip: string) =>
      `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="5" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>${xfrm(0, 0, 914400, 914400, ` ${flip}="1"`).replace('</a:xfrm>', '<a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/></a:xfrm>')}</p:grpSpPr>${textShape('群組文字')}</p:grpSp>`
    for (const [flip, scale] of [['flipV', 'scaleY(-1)'], ['flipH', 'scaleX(-1)']]) {
      const [slide] = readPptx(deckOf([group(flip)]), '').slides
      assert.match(slide.html, new RegExp(`class="pp-grp" style="[^"]*transform:${scale.replace(/[()]/g, '\\$&')}"`))
      assert.match(slide.html, new RegExp(`transform:${scale.replace(/[()]/g, '\\$&')}"><div class="pp-tx"[^>]*><p[^>]*><span[^>]*>群組文字`))
    }
  })

  test('outlined text with no fill draws its outline', () => {
    const body = '<a:p><a:r><a:rPr sz="2400"><a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln><a:noFill/></a:rPr><a:t>空心字</a:t></a:r></a:p>'
    const [slide] = readPptx(deckOf([shape(xfrm(), body)]), '').slides
    assert.match(slide.html, /color:transparent;[^"]*-webkit-text-stroke:calc\(1\*var\(--pt\)\) rgb\(255,0,0\)[^"]*">空心字/)
  })

  test('a video with a poster frame shows the poster with a label saying it does not play here', () => {
    const video = `<p:pic><p:nvPicPr><p:cNvPr id="4" name="v.mp4"/><p:cNvPicPr/><p:nvPr><a:videoFile r:link="v"/></p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="img"/></p:blipFill><p:spPr>${xfrm()}</p:spPr></p:pic>`
    const deck = deckOf([video], { 'ppt/media/image1.png': 'png', 'ppt/media/media1.mp4': 'mp4' }, { 1: [['img', 'image', '../media/image1.png'], ['v', 'video', '../media/media1.mp4']] })
    const [slide] = readPptx(deck, 'm/').slides
    assert.match(slide.html, /<img src="m\/ppt\/media\/image1\.png"/)
    assert.match(slide.html, /class="pp-missing"[^>]*>影音/)
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
    assert.match(html, /class="pp-missing" role="alert"[^>]*><strong[^>]*>這份 PowerPoint 檔無法顯示<\/strong><span>檔案不是有效的 \.pptx[^<]*<\/span><code>PowerPoint: not a zip file<\/code><span>可以下載原始檔/)
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

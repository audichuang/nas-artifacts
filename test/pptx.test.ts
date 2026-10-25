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
import { parseXml, XmlError } from '../src/xml.ts'
import { ZipError, ZipReader } from '../src/zip.ts'
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
    assert.match(second.html, /<div class="pp-geom pp-chart"><svg [^>]*aria-label="季度"[\s\S]*>一月</)
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
    assert.throws(() => parseXml('<!DOCTYPE x [<!ENTITY a "aaaa">]><x>&a;</x>'), XmlError)
    assert.throws(() => parseXml('<a><b></a>'), XmlError)
    assert.throws(() => new ZipReader(Buffer.from('not a zip')), ZipError)
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

  test('W-4 (PowerPoint): Chinese and circled auto-numbers keep counting past twenty, as Word lists do', () => {
    const list = (type: string, start: number) => `<a:p><a:pPr><a:buAutoNum type="${type}" startAt="${start}"/></a:pPr><a:r><a:t>項目</a:t></a:r></a:p>`
    const [slide] = readPptx(deckOf([shape(xfrm(), list('ea1ChtPeriod', 25) + list('circleNumDbPlain', 30) + list('ea1JpnChsDbPeriod', 101))]), '').slides
    const labels = [...slide.html.matchAll(/class="pp-bu"[^>]*>([^<]*)<\/span>/g)].map((m) => m[1].trim())
    assert.deepEqual(labels, ['二十五.', '㉚', '一百〇一.'])
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
    // The reason is said once: no second copy of the same sentence under it.
    assert.equal(page.split('網頁版無法開啟').length - 1 - (page.match(/<script>throw[^<]*網頁版無法開啟/g)?.length ?? 0), 1)
    assert.match(page, /可以下載原始檔,用 PowerPoint 開啟查看/)
    const truncated = renderPptxDeck(Buffer.from('PK\x03\x04 truncated'), 'x.pptx', 'x', { css: '', js: '' })
    assert.match(truncated, /檔案不是有效的 \.pptx/)
    assert.match(truncated, /<script>throw new Error\("PowerPoint: [^"]+"\)<\/script>/)
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
    assert.match(slide.html, /<linearGradient id="([\w-]+)"[^>]*><stop offset="0%" stop-color="rgb\(255,0,0\)"\/>[\s\S]*<path d="M[^"]*" fill="url\(#\1\)"/)
  })

  test('charts draw through the shared chart renderer, coloured by the file; a combo chart draws every group', () => {
    const bar = readPptx(chartDeck(chartXml('barChart', 'stacked', [['FF0000', [5]], ['00FF00', [-2]]])), '').slides[0].html
    assert.match(bar, /<rect [^>]*fill="rgb\(255,0,0\)"/)
    assert.match(bar, /<rect [^>]*fill="rgb\(0,255,0\)"/)
    const ser = (color: string, values: number[], line = false) =>
      `<c:ser><c:idx val="0"/><c:order val="0"/><c:spPr>${line ? `<a:ln w="28575"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:ln>` : `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`}</c:spPr><c:cat><c:strRef><c:strCache><c:ptCount val="${values.length}"/>${values.map((_, at) => `<c:pt idx="${at}"><c:v>k${at}</c:v></c:pt>`).join('')}</c:strCache></c:strRef></c:cat><c:val><c:numRef><c:numCache><c:ptCount val="${values.length}"/>${values.map((v, at) => `<c:pt idx="${at}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numCache></c:numRef></c:val></c:ser>`
    const ax = (id: number, cross: number, kind: string, pos: string) => `<c:${kind}><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${pos}"/><c:crossAx val="${cross}"/></c:${kind}>`
    const combo = `<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser('FF0000', [100, 140, 90])}<c:axId val="1"/><c:axId val="2"/></c:barChart><c:lineChart><c:grouping val="standard"/>${ser('0000FF', [0.2, 0.35, 0.1], true)}<c:axId val="3"/><c:axId val="4"/></c:lineChart>${ax(1, 2, 'catAx', 'b')}${ax(2, 1, 'valAx', 'l')}${ax(3, 4, 'catAx', 'b')}${ax(4, 3, 'valAx', 'r')}</c:plotArea></c:chart></c:chartSpace>`
    const html = readPptx(chartDeck(combo), '').slides[0].html
    assert.equal((html.match(/<rect [^>]*fill="rgb\(255,0,0\)"/g) ?? []).length, 3)
    assert.match(html, /<path d="M[^"]*"[^>]*stroke="rgb\(0,0,255\)"/)
    assert.doesNotMatch(html, /pp-missing/)
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
    assert.match(html, /class="pp-missing" role="alert"[^>]*><strong[^>]*>這份 PowerPoint 檔無法顯示<\/strong><span>檔案不是有效的 \.pptx[^<]*<\/span><code>[^<]+<\/code><span>可以下載原始檔/)
    assert.match(html, /<script>throw new Error\("PowerPoint: [^"]+"\)<\/script>/)
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

const picture = (blip: string, spPr = xfrm()) =>
  `<p:pic><p:nvPicPr><p:cNvPr id="4" name="p"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill>${blip}<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${spPr}</p:spPr></p:pic>`

/** Words only an English sentence has: a face may still name PowerPoint, XML or a part, but not explain itself in English. */
const ENGLISH = /\b(?:the|is|not|of|exceeded|unexpected|unclosed|accepted|too|does|file|error)\b/i

describe('what a reader is told, in Traditional Chinese', () => {
  test('a missing picture says it is missing; a linked one says the web page does not fetch it; an unknown format says so', () => {
    const deck = deckOf(
      [picture('<a:blip r:embed="gone"/>'), picture('<a:blip r:embed="emf"/>'), picture('<a:blip/>')],
      { 'ppt/media/image2.emf': 'emf' },
      { 1: [['gone', 'image', '../media/image1.png']], 2: [['emf', 'image', '../media/image2.emf']] }
    )
    const slides = readPptx(deck, '').slides
    // Worded as the Word and Excel pages word them.
    assert.match(slides[0].html, /class="pp-missing">圖片遺失\(檔案裡找不到這張圖片\)</)
    assert.match(slides[1].html, /class="pp-missing">EMF 圖片\(網頁版未支援,請下載原始檔查看\)</)
    assert.match(slides[2].html, /class="pp-missing">圖片遺失\(檔案裡找不到這張圖片\)</)
    const linked = deckOf([picture('<a:blip r:link="web"/>')], {
      'ppt/slides/_rels/slide1.xml.rels': `<Relationships><Relationship Id="l" Type="http://x/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="web" Type="http://x/image" Target="https://example.com/x.png" TargetMode="External"/></Relationships>`
    })
    const [shown] = readPptx(linked, '').slides
    assert.match(shown.html, /class="pp-missing">外部連結的圖片\(圖片在文件之外,網頁版不載入外部連結\)</)
    assert.doesNotMatch(shown.html, /example\.com/)
  })

  test('a slide that cannot be read, or groups nested past any real deck, are explained in Chinese', () => {
    let nested = textShape('最裡面')
    for (let i = 0; i < 600; i++) nested = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="5" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>${xfrm()}</p:grpSpPr>${nested}</p:grpSp>`
    const bytes = deckOf([`<p:sld ${NS}><p:cSld><p:spTree></p:cSld></p:sld>`, '<!DOCTYPE x><p:sld/>', nested])
    const deck = readPptx(bytes, '')
    // The cause, as thrown, still reaches the viewer's error log.
    assert.match(renderPptxDeck(bytes, 'x.pptx', 'x', { css: '', js: '' }), /<script>throw new Error\("PowerPoint: 第 1 頁:[^"]+; 第 2 頁:[^"]+"\)<\/script>/)
    const faces = deck.slides.map((slide) => [...slide.html.matchAll(/class="pp-missing"[^>]*>([^<]*)</g)].map((m) => m[1]).join(' '))
    for (const face of faces) {
      assert.match(face, /[一-鿿]/, face)
      assert.doesNotMatch(face, ENGLISH, face)
    }
    assert.match(faces[0], /第 1 頁無法顯示:/)
    assert.match(faces[2], /群組/)
  })

  test('a file that is not a deck explains itself in Chinese on its face; the raw cause goes to the error log', () => {
    const page = renderPptxDeck(Buffer.from('PK\x03\x04 truncated'), 'x.pptx', 'x', { css: '', js: '' })
    const face = /role="alert"[^>]*>([\s\S]*?)<\/div>/.exec(page)![1].replace(/<[^>]+>/g, ' ')
    assert.doesNotMatch(face, ENGLISH, face)
    assert.match(page, /<script>throw new Error\("PowerPoint: [^"]+"\)<\/script>/)
  })
})

const tableCell = (text: string) => `<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="zh-TW"/><a:t>${text}</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>`
/** A 3 × 3 table whose cells say where they are: 頭A … 尾C. */
const styledTable = (tblPr: string) =>
  frame('table', `<a:tbl>${tblPr}<a:tblGrid>${'<a:gridCol w="914400"/>'.repeat(3)}</a:tblGrid>${['頭', '中', '尾'].map((row) => `<a:tr h="370000">${['A', 'B', 'C'].map((col) => tableCell(row + col)).join('')}</a:tr>`).join('')}</a:tbl>`)
const tablePr = (flags: string, id?: string) => `<a:tblPr ${flags}>${id ? `<a:tableStyleId>${id}</a:tableStyleId>` : ''}</a:tblPr>`
/** The style of the cell that says `text`. */
const cellStyle = (html: string, text: string) => new RegExp(`<td[^>]* style="([^"]*)"><div class="pp-tx"[^>]*><p[^>]*><span style="([^"]*)">${text}<`).exec(html)?.slice(1).join(' || ') ?? ''
const tableDeck = (tables: string[], def = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}') =>
  deckOf(tables, {
    'ppt/_rels/presentation.xml.rels': rels([...tables.map((_, at): [string, string, string] => [`s${at + 1}`, 'slide', `slides/slide${at + 1}.xml`]), ['ts', 'tableStyles', 'tableStyles.xml']]),
    // What python-pptx and pptxgenjs write: a default, and no definition of any built-in style.
    'ppt/tableStyles.xml': `<a:tblStyleLst ${NS} def="${def}"/>`
  })

describe('table styles PowerPoint has built in', () => {
  test('a table that names a built-in style by its id alone is drawn in that style, colours as PowerPoint mixes them', () => {
    const [slide] = readPptx(tableDeck([styledTable(tablePr('firstRow="1" bandRow="1"', '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}'))]), '').slides
    // Medium Style 2 - Accent 1: an accent header with white bold text over a thick white rule, banded rows.
    assert.match(cellStyle(slide.html, '頭A'), /background:rgb\(68,114,196\).*border-bottom:calc\(3\*var\(--pt\)\) solid rgb\(255,255,255\).*\|\|.*font-weight:700;color:rgb\(255,255,255\)/)
    assert.match(cellStyle(slide.html, '中A'), /background:rgb\(207,213,234\)/)
    assert.match(cellStyle(slide.html, '尾A'), /background:rgb\(233,235,245\)/)
    assert.match(cellStyle(slide.html, '中B'), /border-left:calc\(1\*var\(--pt\)\) solid rgb\(255,255,255\)/)
  })

  test("a table with no style, or one PowerPoint does not know, takes the presentation's default", () => {
    const deck = readPptx(tableDeck([styledTable('<a:tblPr/>'), styledTable(tablePr('', '{00000000-0000-0000-0000-000000000000}'))], '{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}'), '')
    // Medium Style 2: a light grey body, and with no header flag no header.
    for (const slide of deck.slides) {
      assert.match(cellStyle(slide.html, '頭A'), /background:rgb\(231,231,231\)/)
      assert.doesNotMatch(cellStyle(slide.html, '頭A'), /font-weight/)
    }
  })

  test("built-in styles keep what sets them apart: the table's own background, double rules, corner cells, column bands after the first column", () => {
    const [themed, light, medium, banded] = readPptx(tableDeck([
      styledTable(tablePr('firstRow="1"', '{D113A9D2-9D6B-4929-AA2D-F23B5EE8CBE7}')),
      styledTable(tablePr('lastRow="1"', '{69012ECD-51FC-41F1-AA8D-1B2483CD663E}')),
      styledTable(tablePr('firstRow="1" lastRow="1" firstCol="1" lastCol="1"', '{6E25E649-3F16-4E02-A733-19D2CDBF48F0}')),
      styledTable(tablePr('firstCol="1" bandCol="1"', '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}'))
    ]), '').slides
    // Themed Style 2: the theme's third fill, in the accent, behind the whole table.
    assert.match(themed.html, /<table class="pp-tbl" style="background:linear-gradient\([^"]*rgb\(68,114,196\)/)
    // Light Style 2: a double rule above the total row.
    assert.match(cellStyle(light.html, '尾A'), /border-top:calc\(4\*var\(--pt\)\) double rgb\(68,114,196\)/)
    // Medium Style 3: the last column is white on the accent, but its bottom corner is dark text again.
    assert.match(cellStyle(medium.html, '中C'), /background:rgb\(68,114,196\).*color:rgb\(255,255,255\)/)
    assert.match(cellStyle(medium.html, '尾C'), /\|\|.*font-weight:700;color:rgb\(0,0,0\)/)
    // Column bands start after the first column, as row bands start after the header.
    assert.match(cellStyle(banded.html, '中B'), /background:rgb\(207,213,234\)/)
    assert.match(cellStyle(banded.html, '中C'), /background:rgb\(233,235,245\)/)
  })
})

describe('fonts for Chinese, Japanese and Latin text', () => {
  const scripted = THEME.replace('<a:minorFont><a:latin typeface="Calibri"/>', '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/>')
  const run = (rPr: string, text: string) => `<a:p><a:r><a:rPr sz="2800" ${rPr}</a:rPr><a:t>${text}</a:t></a:r></a:p>`
  const fontOf = (html: string, text: string) => new RegExp(`font-family:([^"]*)">${text}<`).exec(html)?.[1] ?? ''
  const [slide] = readPptx(deckOf([shape(xfrm(), [
    run('lang="zh-TW"><a:latin typeface="Times New Roman"/><a:ea typeface="微軟正黑體"/>', '襯線英文黑體中文'),
    run('lang="zh-TW">', '主題的中文'),
    run('lang="en-US" altLang="ja-JP">', '日本語'),
    run('lang="zh-TW"><a:latin typeface="Calibri"/><a:ea typeface="標楷體"/>', '楷書'),
    run('lang="zh-TW"><a:latin typeface="Courier New"/><a:ea typeface="微軟正黑體"/>', '等寬'),
    run('lang="zh-CN"><a:latin typeface="Impact"/><a:ea typeface="SimSun"/>', '宋体英文'),
    run('lang="zh-CN"><a:latin typeface="Impact"/><a:ea typeface="NSimSun"/>', '新宋体英文'),
    run('lang="ko-KR"><a:latin typeface="Impact"/><a:ea typeface="Gungsuh"/>', '궁서체')
  ].join(''))], { 'ppt/theme/theme1.xml': scripted }), '').slides

  test('Latin and East Asian letters each fall back within their own family', () => {
    const serifLatin = fontOf(slide.html, '襯線英文黑體中文')
    // Times for the Latin letters, then the Chinese font and other sans-serif Chinese fonts for the rest.
    assert.match(serifLatin, /^'Times New Roman',[^]*'微軟正黑體','PingFang TC'[^]*,sans-serif$/)
    assert.doesNotMatch(serifLatin, /PMingLiU|Songti/)
    assert.match(fontOf(slide.html, '楷書'), /^'Calibri',[^]*'標楷體','DFKai-SB'[^]*,serif$/)
    const mono = fontOf(slide.html, '等寬')
    assert.match(mono, /^'Courier New',[^]*Menlo[^]*'微軟正黑體','PingFang TC'/)
  })

  test('Song faces named in Latin letters (SimSun, NSimSun) and Gungsuh fall back to serif faces', () => {
    assert.match(fontOf(slide.html, '宋体英文'), /^'Impact',[^]*'SimSun','Songti SC'[^]*,serif$/)
    assert.match(fontOf(slide.html, '新宋体英文'), /'NSimSun','SimSun','Songti SC'[^]*,serif$/)
    assert.match(fontOf(slide.html, '궁서체'), /'Gungsuh','AppleMyungjo'[^]*,serif$/)
  })

  test("text with no East Asian font of its own takes the theme's font for its language", () => {
    assert.match(fontOf(slide.html, '主題的中文'), /^'Calibri',[^]*'新細明體','PMingLiU'[^]*,serif$/)
    // An English run whose second language is Japanese draws its Japanese in the theme's Japanese font.
    assert.match(fontOf(slide.html, '日本語'), /'ＭＳ Ｐゴシック','Hiragino Sans'/)
  })
})

describe('picture, pattern and gradient fills', () => {
  const PNG = Buffer.from(PNG_DATA.split(',')[1], 'base64').toString('latin1')
  const blip = '<a:blipFill><a:blip r:embed="img"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>'
  const pattern = (prst: string) => `<a:pattFill prst="${prst}"><a:fgClr><a:srgbClr val="C00000"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill>`
  const withPicture = (slides: string[]) => deckOf(slides, { 'ppt/media/image1.png': PNG }, Object.fromEntries(slides.map((_, at) => [at + 1, [['img', 'image', '../media/image1.png']] as Rels])))
  /** The paint server a path names, and the path's fill. */
  const paintOf = (html: string) => {
    const fill = /<path d="[^"]*" fill="url\(#([\w-]+)\)"/.exec(html)
    return fill ? new RegExp(`<(pattern|linearGradient|radialGradient) id="${fill[1]}"[^>]*>([\\s\\S]*?)</\\1>`).exec(html) : null
  }

  test('a picture fills any shape, not only a rectangle', () => {
    const [hexagon, star] = readPptx(withPicture([shape(`${xfrm()}<a:prstGeom prst="hexagon"/>${blip}`), shape(`${xfrm()}<a:prstGeom prst="star5"/>${blip}`)]), 'm/').slides
    for (const slide of [hexagon, star]) {
      const paint = paintOf(slide.html)
      assert.equal(paint?.[1], 'pattern', slide.html)
      assert.match(paint![2], /<image href="m\/ppt\/media\/image1\.png"[^>]* preserveAspectRatio="none"/)
    }
  })

  test('a pattern fill draws its pattern in both colours, on boxes, paths and slide backgrounds', () => {
    const [box, path, background] = readPptx(deckOf([
      shape(`${xfrm()}<a:prstGeom prst="rect"/>${pattern('dkDnDiag')}`),
      shape(`${xfrm()}<a:prstGeom prst="triangle"/>${pattern('smCheck')}`),
      `<p:sld ${NS}><p:cSld><p:bg><p:bgPr>${pattern('horz')}<a:effectLst/></p:bgPr></p:bg><p:spTree/></p:cSld></p:sld>`
    ]), '').slides
    for (const html of [box.html, background.html]) {
      const url = /background:url\('data:image\/svg\+xml,([^']+)'\)[^;"]*repeat/.exec(html)
      assert.ok(url, html)
      const svg = decodeURIComponent(url[1])
      assert.match(svg, /fill="rgb\(255,255,255\)"/)
      assert.match(svg, /fill="rgb\(192,0,0\)"/)
    }
    // The slide itself sits outside the canvas that defines --pt: its tiles are sized in percentages.
    assert.match(background.html, /<section [^>]*style="background:url\('data:[^']+'\) 0% 0%\/0\.8333% 1\.1111% repeat"/)
    const paint = paintOf(path.html)
    assert.equal(paint?.[1], 'pattern')
    assert.match(paint![2], /rgb\(192,0,0\)[\s\S]*rgb\(255,255,255\)|rgb\(255,255,255\)[\s\S]*rgb\(192,0,0\)/)
  })

  test('a gradient fills a path shape as a gradient', () => {
    const grad = '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="5400000"/></a:gradFill>'
    const [slide] = readPptx(deckOf([shape(`${xfrm()}<a:prstGeom prst="triangle"/>${grad}`)]), '').slides
    const paint = paintOf(slide.html)
    assert.equal(paint?.[1], 'linearGradient')
    assert.match(paint![0], /x1="0\.5" y1="0" x2="0\.5" y2="1"/)
    assert.match(paint![2], /<stop offset="0%" stop-color="rgb\(255,0,0\)"\/><stop offset="100%" stop-color="rgb\(0,0,255\)"\/>/)
  })
})

describe('shape geometry (ECMA-376 presets and custom shapes)', () => {
  const fill = '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>'
  const ln = '<a:ln w="12700"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln>'
  const slideOf = (spPr: string, body = '') => readPptx(deckOf([shape(spPr, body)]), '').slides[0].html
  const pathsOf = (html: string) => [...html.matchAll(/<path d="([^"]*)" fill="([^"]*)"([^>]*)\/>/g)].map((m) => ({ d: m[1], fill: m[2], rest: m[3] }))
  const coords = (d: string) => (d.match(/-?[\d.]+/g) ?? []).map(Number)

  test('non-rectangular presets draw their own outline: heart curves, a callout tail leaves the box, a cloud is arcs', () => {
    const heart = pathsOf(slideOf(`${xfrm()}<a:prstGeom prst="heart"><a:avLst/></a:prstGeom>${fill}`))
    assert.ok(heart.length >= 1)
    assert.match(heart[0].d, /C/)
    assert.equal(heart[0].fill, 'rgb(255,0,0)')
    // wedgeRectCallout's default tail points below the box (adj2 = 62500 of the height past the centre).
    const callout = pathsOf(slideOf(`${xfrm()}<a:prstGeom prst="wedgeRectCallout"><a:avLst/></a:prstGeom>${fill}`))
    assert.ok(Math.max(...coords(callout[0].d)) > 72, callout[0].d)
    const cloud = slideOf(`${xfrm()}<a:prstGeom prst="cloud"><a:avLst/></a:prstGeom>${fill}${ln}`)
    assert.ok((pathsOf(cloud)[0].d.match(/A/g) ?? []).length >= 10)
    assert.doesNotMatch(cloud, /border-radius/)
  })

  test('fills come first, then strokes; shaded paths get an overlay; lines take no fill and carry the arrowheads', () => {
    const can = pathsOf(slideOf(`${xfrm()}<a:prstGeom prst="can"/>${fill}${ln}`))
    const firstStroke = can.findIndex((p) => p.rest.includes('stroke='))
    assert.ok(firstStroke > 0 && can.slice(firstStroke).every((p) => p.rest.includes('stroke=')), JSON.stringify(can))
    // The can's lid is lightened: drawn in the fill colour, then white at 40%.
    assert.ok(can.some((p) => p.fill === '#fff' && /fill-opacity="0\.4"/.test(p.rest)))
    const arrow = slideOf(`${xfrm()}<a:prstGeom prst="bentConnector3"/>${ln.replace('</a:ln>', '<a:tailEnd type="triangle"/></a:ln>')}`)
    const lines = pathsOf(arrow)
    assert.ok(lines.every((p) => p.fill === 'none'))
    assert.match(lines[0].rest, /marker-end="url\(#pp-m\d+t\)"/)
    // A line with only a fill is never a filled box: its one segment encloses nothing.
    const filledLine = slideOf(`${xfrm()}<a:prstGeom prst="line"/>${fill}<a:ln><a:noFill/></a:ln>`)
    assert.doesNotMatch(filledLine, /background:rgb\(255,0,0\)/)
    assert.deepEqual(pathsOf(filledLine).map((p) => p.d), ['M0,0 L72,72'])
    // smileyFace's eyes are open arcs the standard fills (darkened); its mouth is a line and is not filled.
    const smiley = pathsOf(slideOf(`${xfrm()}<a:prstGeom prst="smileyFace"/>${fill}${ln}`))
    assert.ok(smiley.some((p) => p.fill === '#000' && /fill-opacity="0\.2"/.test(p.rest)))
    assert.ok(smiley.some((p) => /Q/.test(p.d) && p.fill === 'none'))
    assert.ok(!smiley.some((p) => /Q/.test(p.d) && p.fill !== 'none'))
  })

  test("text sits in the geometry's text rectangle, as PowerPoint lays it out", () => {
    const html = slideOf(`${xfrm()}<a:prstGeom prst="ellipse"/>${fill}`, '<a:p><a:r><a:t>橢圓裡的字</a:t></a:r></a:p>')
    // An ellipse's text box is the square inscribed at 45°: (1 - cos 45°) / 2 = 14.645% in from each side.
    assert.match(html, /<div class="pp-grp" style="left:14\.6\d+%;top:14\.6\d+%;width:70\.7\d+%;height:70\.7\d+%"><div class="pp-tx"[\s\S]*橢圓裡的字/)
    // A rectangle keeps the whole box.
    assert.doesNotMatch(slideOf(`${xfrm()}<a:prstGeom prst="rect"/>${fill}`, '<a:p><a:r><a:t>方</a:t></a:r></a:p>'), /class="pp-grp"/)
  })

  test('a custom geometry is evaluated with its own guides; an unknown preset is a labelled box', () => {
    const cust = `<a:custGeom><a:avLst/><a:gdLst><a:gd name="mid" fmla="*/ w 1 2"/></a:gdLst><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path><a:moveTo><a:pt x="mid" y="t"/></a:moveTo><a:lnTo><a:pt x="r" y="b"/></a:lnTo><a:lnTo><a:pt x="l" y="b"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>`
    const [path] = pathsOf(slideOf(`${xfrm()}${cust}${fill}`))
    assert.match(path.d, /^M36,0 L72,72 L0,72 Z$/)
    const unknown = slideOf(`${xfrm()}<a:prstGeom prst="notAShape"/>${fill}`)
    assert.match(unknown, /background:rgb\(255,0,0\)/)
    assert.match(unknown, /class="pp-missing"[^>]*>未支援的圖形:notAShape</)
  })

  test("a picture is clipped to its shape's path, scaled with the slide", () => {
    const PNG = Buffer.from(PNG_DATA.split(',')[1], 'base64').toString('latin1')
    const pic = `<p:pic><p:nvPicPr><p:cNvPr id="4" name="p"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="img"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${xfrm()}<a:prstGeom prst="ellipse"/>${ln}</p:spPr></p:pic>`
    const html = readPptx(deckOf([pic], { 'ppt/media/image1.png': PNG }, { 1: [['img', 'image', '../media/image1.png']] }), 'm/').slides[0].html
    const id = /class="pp-pic" style="clip-path:url\(#([\w-]+)\)/.exec(html)?.[1]
    assert.ok(id, html)
    assert.match(html, new RegExp(`<clipPath id="${id}" clipPathUnits="objectBoundingBox"><path transform="scale\\([\\d.]+ [\\d.]+\\)" d="M[^"]*A`))
    // The outline follows the same path instead of a square border.
    assert.match(html, /<svg class="pp-geom" viewBox="0 0 72 72"[^>]*><path d="M[^"]*A[^"]*" fill="none" stroke="rgb\(0,0,255\)"/)
  })
})

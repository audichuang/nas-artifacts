import assert from 'node:assert/strict'
import { describe, test } from 'node:test'
import { chartTitle, formatChartNumber, renderChart, type ChartOptions } from '../src/office-chart.ts'
import { child, parseXml, type XmlNode } from '../src/xml.ts'

const NS = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
const ACCENTS = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47']

/** A chart part from a plot area, the way a tool would write one. */
function chart(plot: string, extra = '', outer = ''): XmlNode {
  // Series numbered in order, as Office writes them (automatic colours follow c:idx).
  let n = 0
  plot = plot.replace(/<c:idx val="0"\/><c:order val="0"\/>/g, () => `<c:idx val="${n}"/><c:order val="${n++}"/>`)
  return parseXml(`<?xml version="1.0"?><c:chartSpace ${NS}>${outer}<c:chart>${extra}<c:plotArea><c:layout/>${plot}</c:plotArea></c:chart></c:chartSpace>`)
}

const strs = (values: string[]) => `<c:strRef><c:f>S!$A$1</c:f><c:strCache><c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:strCache></c:strRef>`
const nums = (values: (number | string | null)[], code = 'General') =>
  `<c:numRef><c:f>S!$B$1</c:f><c:numCache><c:formatCode>${code}</c:formatCode><c:ptCount val="${values.length}"/>${values.map((v, i) => (v === null ? '' : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)).join('')}</c:numCache></c:numRef>`
const fill = (hex: string) => `<c:spPr><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill></c:spPr>`

function ser(name: string, values: (number | string | null)[], cats: string[] | null = ['A', 'B', 'C'], more = '', code = 'General'): string {
  return `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx>${strs([name])}</c:tx>${more}${cats ? `<c:cat>${strs(cats)}</c:cat>` : ''}<c:val>${nums(values, code)}</c:val></c:ser>`
}

const catAx = (id: number, cross: number, more = '') => `<c:catAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>${more}<c:crossAx val="${cross}"/></c:catAx>`
const valAx = (id: number, cross: number, more = '', pos = 'l') => `<c:valAx><c:axId val="${id}"/>${more.includes('<c:scaling>') ? '' : '<c:scaling><c:orientation val="minMax"/></c:scaling>'}<c:delete val="0"/><c:axPos val="${pos}"/><c:majorGridlines/>${more}<c:crossAx val="${cross}"/></c:valAx>`
const axes = (more = '') => `${catAx(1, 2)}${valAx(2, 1, more)}`
const ids = (a = 1, b = 2) => `<c:axId val="${a}"/><c:axId val="${b}"/>`

/** The colour resolver a renderer would pass: srgbClr only, enough for these charts. */
function resolveColor(holder: XmlNode | null): string | null {
  const srgb = child(holder, 'srgbClr')
  const val = srgb?.attrs.val
  return val && /^[0-9A-Fa-f]{6}$/.test(val) ? `#${val.toUpperCase()}` : null
}

const OPTIONS: ChartOptions = { widthPt: 480, heightPt: 300, resolveColor, themeColor: (i) => `#${ACCENTS[i % 6]}`, fontFamily: "'Calibri',sans-serif", textColor: '#595959' }

function draw(node: XmlNode, options: Partial<ChartOptions> = {}): { svg: string; text: string } {
  const drawn = renderChart(node, { ...OPTIONS, ...options })
  assert.ok(drawn, 'the chart should draw')
  return drawn
}

const count = (svg: string, pattern: RegExp) => (svg.match(pattern) ?? []).length
const texts = (svg: string) => [...svg.matchAll(/<text[^>]*>([^<]*)<\/text>/g)].map((m) => m[1])
const rects = (svg: string) =>
  [...svg.matchAll(/<rect x="([-\d.]+)" y="([-\d.]+)" width="([\d.]+)" height="([\d.]+)" fill="([^"]+)"[^>]*data-s="(\d+)"/g)].map((m) => ({ x: +m[1], y: +m[2], w: +m[3], h: +m[4], fill: m[5], s: +m[6] }))

describe('number formats', () => {
  const cases: [string, number, string][] = [
    ['General', 1234.5, '1234.5'],
    ['General', 0.1 + 0.2, '0.3'],
    ['General', 12345678901234, '1.23457E+13'],
    ['0', 2.5, '3'],
    ['0.0', -1.25, '-1.3'],
    ['#,##0', 1234567.4, '1,234,567'],
    ['#,##0.00', 1234.5, '1,234.50'],
    ['0%', 0.256, '26%'],
    ['0.0%', 0.256, '25.6%'],
    ['"NT$"#,##0', 1500, 'NT$1,500'],
    ['[$$-409]#,##0.00', 3, '$3.00'],
    ['$#,##0;[Red]-$#,##0', -42, '-$42'],
    ['#,##0;(#,##0)', -42, '(42)'],
    ['0.00E+00', 12345, '1.23E+04'],
    ['#,##0,"K"', 25000, '25K'],
    ['yyyy/m/d', 45000, '2023/3/15'],
    ['@', 7, '7'],
    ['[>100]0;0.0', 7, '7']
  ]
  for (const [code, value, want] of cases) test(`${code} ${value}`, () => assert.equal(formatChartNumber(code, value), want))
  test('non-finite values format as nothing', () => {
    assert.equal(formatChartNumber('0', Number.NaN), '')
    assert.equal(formatChartNumber('General', Infinity), '')
  })
})

describe('bar and column charts', () => {
  test('a clustered column chart draws one bar per value, coloured by the series fill, with category labels', () => {
    const node = chart(`<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser('Sales', [10, 20, 30], undefined, fill('FF0000'))}${ser('Cost', [5, 15, 25])}<c:gapWidth val="150"/>${ids()}</c:barChart>${axes()}`, '', '')
    const { svg, text } = draw(node)
    const bars = rects(svg)
    assert.equal(bars.length, 6)
    assert.deepEqual([...new Set(bars.filter((b) => b.s === 0).map((b) => b.fill))], ['#FF0000'])
    assert.deepEqual([...new Set(bars.filter((b) => b.s === 1).map((b) => b.fill))], ['#ED7D31'])
    // Taller value, taller bar; vertical columns.
    const first = bars.filter((b) => b.s === 0)
    assert.ok(first[2].h > first[1].h && first[1].h > first[0].h)
    for (const label of ['A', 'B', 'C']) assert.ok(texts(svg).includes(label))
    assert.match(text, /Sales/)
    assert.match(svg, /^<svg [^>]*viewBox="0 0 480 300"/)
  })

  test('gap width and overlap set the bar widths', () => {
    const wide = rects(draw(chart(`<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser('S', [1, 2, 3])}<c:gapWidth val="0"/>${ids()}</c:barChart>${axes()}`)).svg)
    const narrow = rects(draw(chart(`<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser('S', [1, 2, 3])}<c:gapWidth val="300"/>${ids()}</c:barChart>${axes()}`)).svg)
    assert.ok(wide[0].w > narrow[0].w * 3)
    const overlapped = rects(draw(chart(`<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser('S', [1, 2, 3])}${ser('T', [1, 2, 3])}<c:gapWidth val="100"/><c:overlap val="100"/>${ids()}</c:barChart>${axes()}`)).svg)
    // Fully overlapped series sit on the same x.
    assert.equal(overlapped.filter((b) => b.s === 0)[0].x, overlapped.filter((b) => b.s === 1)[0].x)
  })

  test('a horizontal stacked bar chart stacks along x', () => {
    const node = chart(`<c:barChart><c:barDir val="bar"/><c:grouping val="stacked"/>${ser('S', [10, 20, 30])}${ser('T', [10, 10, 10])}<c:overlap val="100"/>${ids()}</c:barChart>${catAx(1, 2).replace('"b"', '"l"')}${valAx(2, 1, '', 'b')}`)
    const bars = rects(draw(node).svg)
    const s = bars.filter((b) => b.s === 0)
    const t = bars.filter((b) => b.s === 1)
    // Horizontal: the bars grow in width; T starts where S ends.
    assert.ok(s[2].w > s[0].w)
    assert.ok(Math.abs(t[0].x - (s[0].x + s[0].w)) < 0.6)
    assert.ok(Math.abs(t[0].y - s[0].y) < 0.6)
  })

  test('a 100% stacked column chart fills every category to the top and labels the axis in percent', () => {
    const node = chart(`<c:barChart><c:barDir val="col"/><c:grouping val="percentStacked"/>${ser('S', [1, 3, 5])}${ser('T', [3, 1, 5])}<c:overlap val="100"/>${ids()}</c:barChart>${axes()}`)
    const { svg } = draw(node)
    const bars = rects(svg)
    const heights = [0, 1, 2].map((at) => bars.filter((b) => b.x === bars[at].x).reduce((sum, b) => sum + b.h, 0))
    assert.ok(Math.abs(heights[0] - heights[1]) < 1 && Math.abs(heights[1] - heights[2]) < 1)
    assert.ok(texts(svg).includes('100%'))
  })

  test('per-point dPt colours and varyColors', () => {
    const dPt = `<c:dPt><c:idx val="1"/>${fill('00FF00')}</c:dPt>`
    const bars = rects(draw(chart(`<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${ser('S', [1, 2, 3], undefined, fill('112233') + dPt)}${ids()}</c:barChart>${axes()}`)).svg)
    assert.deepEqual(bars.map((b) => b.fill), ['#112233', '#00FF00', '#112233'])
    const varied = rects(draw(chart(`<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="1"/>${ser('S', [1, 2, 3])}${ids()}</c:barChart>${axes()}`)).svg)
    assert.deepEqual(varied.map((b) => b.fill), ['#4472C4', '#ED7D31', '#A5A5A5'])
  })
})

describe('axes', () => {
  test('axis min, max and major unit from the file', () => {
    const scaling = '<c:scaling><c:orientation val="minMax"/><c:max val="15"/><c:min val="-5"/></c:scaling><c:majorUnit val="5"/>'
    const { svg } = draw(chart(`<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser('S', [1, 2, 3])}${ids()}</c:barChart>${axes(scaling)}`))
    const labels = texts(svg)
    for (const want of ['-5', '0', '5', '10', '15']) assert.ok(labels.includes(want), `${want} in ${labels}`)
    assert.ok(!labels.includes('20') && !labels.includes('-10'))
  })

  test('a logarithmic axis ticks at powers of its base', () => {
    const scaling = '<c:scaling><c:logBase val="10"/><c:orientation val="minMax"/></c:scaling>'
    const labels = texts(draw(chart(`<c:lineChart><c:grouping val="standard"/>${ser('S', [1, 100, 10000])}<c:marker val="1"/>${ids()}</c:lineChart>${axes(scaling)}`)).svg)
    for (const want of ['1', '10', '100', '1000', '10000']) assert.ok(labels.includes(want), `${want} in ${labels}`)
  })

  test('a reversed value axis draws bars down from the top', () => {
    const normal = rects(draw(chart(`<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser('S', [1, 2, 3])}${ids()}</c:barChart>${axes()}`)).svg)
    const scaling = '<c:scaling><c:orientation val="maxMin"/></c:scaling>'
    const reversed = rects(draw(chart(`<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser('S', [1, 2, 3])}${ids()}</c:barChart>${axes(scaling)}`)).svg)
    assert.ok(reversed[2].y < normal[2].y)
  })

  test('the value axis number format, or the source format when linked', () => {
    const own = texts(draw(chart(`<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser('S', [0.1, 0.2, 0.3])}${ids()}</c:barChart>${axes('<c:numFmt formatCode="0%" sourceLinked="0"/>')}`)).svg)
    assert.ok(own.includes('30%'), own.join('|'))
    const linked = texts(draw(chart(`<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser('S', [1000, 2000, 3000], undefined, '', '"NT$"#,##0')}${ids()}</c:barChart>${axes('<c:numFmt formatCode="General" sourceLinked="1"/>')}`)).svg)
    assert.ok(linked.includes('NT$3,000'), linked.join('|'))
  })

  test('many long category labels are rotated or thinned, never overlapping', () => {
    const cats = Array.from({ length: 60 }, (_, i) => `Category number ${i}`)
    const { svg } = draw(chart(`<c:lineChart><c:grouping val="standard"/>${ser('S', cats.map((_, i) => i), cats)}${ids()}</c:lineChart>${axes()}`))
    const shown = texts(svg).filter((t) => t.startsWith('Category'))
    assert.ok(shown.length < 60 && shown.length > 3, `${shown.length} labels`)
    assert.match(svg, /rotate\(-45/)
  })

  test('a deleted axis draws no labels', () => {
    const { svg } = draw(chart(`<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser('S', [1, 2, 3])}${ids()}</c:barChart>${catAx(1, 2).replace('<c:delete val="0"/>', '<c:delete val="1"/>')}${valAx(2, 1)}`))
    assert.ok(!texts(svg).includes('A'))
  })
})

describe('data labels', () => {
  test('showVal labels each bar in the series number format', () => {
    const dLbls = '<c:dLbls><c:numFmt formatCode="0.0" sourceLinked="0"/><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/></c:dLbls>'
    const labels = texts(draw(chart(`<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser('S', [1, 2, 3])}${dLbls}${ids()}</c:barChart>${axes()}`)).svg)
    for (const want of ['1.0', '2.0', '3.0']) assert.ok(labels.includes(want), labels.join('|'))
  })

  test('a series dLbls delete overrides the group', () => {
    const group = '<c:dLbls><c:showVal val="1"/></c:dLbls>'
    const labels = texts(draw(chart(`<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser('S', [11, 22, 33], undefined, '<c:dLbls><c:delete val="1"/></c:dLbls>')}${group}${ids()}</c:barChart>${axes()}`)).svg)
    assert.ok(!labels.includes('33'))
  })

  test('pie labels show percent and category name', () => {
    const dLbls = '<c:dLbls><c:showVal val="0"/><c:showCatName val="1"/><c:showPercent val="1"/></c:dLbls>'
    const labels = texts(draw(chart(`<c:pieChart><c:varyColors val="1"/>${ser('S', [1, 1, 2], ['X', 'Y', 'Z'])}${dLbls}</c:pieChart>`)).svg)
    assert.ok(labels.some((l) => l.includes('Z') && l.includes('50%')), labels.join('|'))
  })
})

describe('line, area, scatter, radar, bubble', () => {
  test('a line chart with markers; a gap leaves the line out', () => {
    const { svg } = draw(chart(`<c:lineChart><c:grouping val="standard"/>${ser('S', [1, 2, null, 3], ['A', 'B', 'C', 'D'], '<c:marker><c:symbol val="circle"/><c:size val="7"/></c:marker>')}<c:marker val="1"/>${ids()}</c:lineChart>${axes()}`))
    assert.equal(count(svg, /<circle /g), 3)
    // One segment, A to B: C is a gap and D stands alone.
    assert.equal(count(svg, / L/g), 1)
    assert.match(svg, /<path d="M[^"]*"[^>]*fill="none" stroke="#4472C4"/)
  })

  test('symbol none hides markers; ln noFill hides the line', () => {
    const { svg } = draw(chart(`<c:lineChart><c:grouping val="standard"/>${ser('S', [1, 2, 3], undefined, '<c:spPr><a:ln><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val="none"/></c:marker>')}<c:marker val="1"/>${ids()}</c:lineChart>${axes()}`))
    assert.ok(!/stroke="#4472C4"/.test(svg))
  })

  test('stacked lines add up', () => {
    const { svg } = draw(chart(`<c:lineChart><c:grouping val="stacked"/>${ser('S', [10, 10, 10])}${ser('T', [10, 10, 10])}${ids()}</c:lineChart>${axes()}`))
    assert.ok(texts(svg).includes('20'))
  })

  test('a stacked area chart fills between the series', () => {
    const { svg } = draw(chart(`<c:areaChart><c:grouping val="stacked"/>${ser('S', [1, 2, 3])}${ser('T', [1, 2, 3])}${ids()}</c:areaChart>${axes()}`))
    assert.equal(count(svg, /<path d="M[^"]*Z" fill="#(4472C4|ED7D31)"/g), 2)
  })

  test('scatter plots real numeric x on a value axis', () => {
    const xy = (x: number[], y: number[]) => `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx>${strs(['P'])}</c:tx><c:xVal>${nums(x)}</c:xVal><c:yVal>${nums(y)}</c:yVal></c:ser>`
    const node = chart(`<c:scatterChart><c:scatterStyle val="lineMarker"/>${xy([1, 2, 10], [1, 2, 3])}${ids()}</c:scatterChart>${valAx(1, 2, '', 'b')}${valAx(2, 1)}`)
    const { svg } = draw(node)
    const dots = [...svg.matchAll(/<path d="M([-\d.]+),([-\d.]+)[^"]*"[^>]*data-marker/g)].map((m) => +m[1])
    assert.equal(dots.length, 3)
    // x=10 sits much further from x=2 than x=2 from x=1.
    assert.ok(dots[2] - dots[1] > (dots[1] - dots[0]) * 4, dots.join(','))
    assert.ok(texts(svg).includes('10'))
  })

  test('a radar chart draws a closed polygon per series around the categories', () => {
    const { svg } = draw(chart(`<c:radarChart><c:radarStyle val="marker"/>${ser('S', [1, 2, 3, 4, 5], ['a', 'b', 'c', 'd', 'e'])}${ids()}</c:radarChart>${axes()}`))
    assert.match(svg, /<path d="M[^"]*Z" fill="none" stroke="#4472C4"/)
    for (const label of ['a', 'b', 'c', 'd', 'e']) assert.ok(texts(svg).includes(label))
  })

  test('bubble sizes scale with the size values', () => {
    const b = `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx>${strs(['B'])}</c:tx><c:xVal>${nums([1, 2])}</c:xVal><c:yVal>${nums([1, 2])}</c:yVal><c:bubbleSize>${nums([1, 4])}</c:bubbleSize></c:ser>`
    const { svg } = draw(chart(`<c:bubbleChart><c:varyColors val="0"/>${b}<c:bubbleScale val="100"/>${ids()}</c:bubbleChart>${valAx(1, 2, '', 'b')}${valAx(2, 1)}`))
    const radii = [...svg.matchAll(/<circle [^>]*r="([\d.]+)"[^>]*data-bubble/g)].map((m) => +m[1])
    assert.equal(radii.length, 2)
    // Area follows size: 4× the size is 2× the radius.
    assert.ok(Math.abs(radii[1] / radii[0] - 2) < 0.05, radii.join(','))
  })
})

describe('line details', () => {
  test('a line that sets only its dash and width keeps the automatic colour', () => {
    const { svg } = draw(chart(`<c:lineChart><c:grouping val="standard"/>${ser('S', [1, 2, 3], undefined, '<c:spPr><a:ln w="38100"><a:prstDash val="dash"/></a:ln></c:spPr>')}${ids()}</c:lineChart>${axes()}`))
    assert.match(svg, /stroke="#4472C4" stroke-width="3"[^>]*stroke-dasharray=/)
  })

  test('blanks shown as zero or spanned', () => {
    const plot = (blanks: string) => chart(`<c:lineChart><c:grouping val="standard"/>${ser('S', [1, null, 3], undefined, '<c:marker><c:symbol val="none"/></c:marker>')}${ids()}</c:lineChart>${axes()}`, `<c:dispBlanksAs val="${blanks}"/>`)
    const segments = (svg: string) => count(/<path d="([^"]*)" fill="none"/.exec(svg)?.[1] ?? '', / L/g)
    assert.equal(segments(draw(plot('gap')).svg), 0)
    assert.equal(segments(draw(plot('span')).svg), 1)
    assert.equal(segments(draw(plot('zero')).svg), 2)
  })
})

describe('pie and doughnut', () => {
  test('a pie draws one slice per point in point colours, with a legend of categories', () => {
    const { svg } = draw(chart(`<c:pieChart><c:varyColors val="1"/>${ser('S', [1, 2, 3], ['X', 'Y', 'Z'])}<c:firstSliceAng val="0"/></c:pieChart>`, '<c:legend><c:legendPos val="r"/></c:legend>'))
    assert.equal(count(svg, /data-slice/g), 3)
    for (const name of ['X', 'Y', 'Z']) assert.ok(texts(svg).includes(name))
  })

  test('a doughnut leaves its hole transparent: no white disc drawn over the slices', () => {
    const { svg } = draw(chart(`<c:doughnutChart><c:varyColors val="1"/>${ser('S', [1, 2, 3], ['X', 'Y', 'Z'])}<c:holeSize val="50"/></c:doughnutChart>`))
    assert.ok(!/<circle[^>]*fill="#fff/i.test(svg))
    assert.equal(count(svg, /data-slice/g), 3)
    // Every slice is a ring segment: two arcs.
    for (const m of svg.matchAll(/<path d="([^"]*)"[^>]*data-slice/g)) assert.equal(count(m[1], /A/g), 2)
  })

  test("a doughnut's first series is its inner ring", () => {
    const { svg } = draw(chart(`<c:doughnutChart><c:varyColors val="0"/>${ser('In', [1, 1], ['X', 'Y'], fill('111111'))}${ser('Out', [1, 1], ['X', 'Y'], fill('222222'))}<c:holeSize val="50"/></c:doughnutChart>`))
    const radius = (hex: string) => Math.max(...[...svg.matchAll(new RegExp(`<path d="M[^"]*A([\\d.]+),[^"]*" fill="#${hex}"`, 'g'))].map((m) => +m[1]))
    assert.ok(radius('111111') < radius('222222'))
  })

  test('a single full slice still draws', () => {
    const { svg } = draw(chart(`<c:pieChart><c:varyColors val="1"/>${ser('S', [5, 0, 0], ['X', 'Y', 'Z'])}</c:pieChart>`))
    assert.ok(count(svg, /data-slice/g) >= 1)
  })
})

describe('combo charts', () => {
  test('every chart group in the plot area is drawn, the line on its own secondary axis', () => {
    const plot =
      `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser('Sales', [100, 140, 90])}${ids(1, 2)}</c:barChart>` +
      `<c:lineChart><c:grouping val="standard"/>${ser('Margin', [0.2, 0.35, 0.1], undefined, fill('C00000').replace('<a:solidFill>', '<a:ln><a:solidFill>').replace('</a:solidFill>', '</a:solidFill></a:ln>'))}<c:marker val="1"/>${ids(3, 4)}</c:lineChart>` +
      `${catAx(1, 2)}${valAx(2, 1)}${valAx(4, 3, '<c:numFmt formatCode="0%" sourceLinked="0"/><c:crosses val="max"/>', 'r')}${catAx(3, 4).replace('<c:delete val="0"/>', '<c:delete val="1"/>')}`
    const { svg, text } = draw(chart(plot, '<c:legend><c:legendPos val="b"/></c:legend>'))
    assert.equal(rects(svg).length, 3)
    assert.match(svg, /stroke="#C00000"/, 'the line group is drawn')
    const labels = texts(svg)
    // The secondary axis carries the line's percent scale; the primary axis the bars'.
    assert.ok(labels.includes('40%'), labels.join('|'))
    assert.ok(labels.includes('160'), labels.join('|'))
    assert.match(text, /Margin/)
    for (const name of ['Sales', 'Margin']) assert.ok(labels.includes(name), 'both series in the legend')
  })

  test('groups sharing one axis pair share its scale', () => {
    const plot =
      `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser('Sales', [100, 140, 90])}${ids()}</c:barChart>` +
      `<c:lineChart><c:grouping val="standard"/>${ser('Line', [0.2, 0.3, 0.1])}<c:marker val="1"/>${ids()}</c:lineChart>${axes()}`
    const { svg } = draw(chart(plot))
    assert.equal(rects(svg).length, 3)
    assert.equal(count(svg, /stroke="#ED7D31"/g) >= 1, true)
  })

  test('an unsupported group beside a supported one draws the rest and says a part is missing', () => {
    const plot = `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser('S', [1, 2, 3])}${ids()}</c:barChart><c:surfaceChart>${ser('T', [1, 2, 3])}${ids()}</c:surfaceChart>${axes()}`
    const { svg } = draw(chart(plot))
    assert.equal(rects(svg).length, 3)
    assert.match(svg, /未顯示/)
  })
})

describe('titles and legends', () => {
  const titled = (title: string) => `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>${title}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`

  test('a long title wraps inside the box', () => {
    const long = '24 年第三季各區域營收與毛利率比較(含海外子公司與新設據點之合併數字)'
    const { svg } = draw(chart(`<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser('S', [1, 2, 3])}${ids()}</c:barChart>${axes()}`, titled(long)), { widthPt: 240 })
    const lines = [...svg.matchAll(/<text[^>]*data-title[^>]*>([^<]*)<\/text>/g)].map((m) => m[1])
    assert.ok(lines.length >= 2, lines.join('|'))
    assert.equal(lines.join(''), long)
    assert.equal(chartTitle(chart('', titled(long))), long)
  })

  test('a single named series titles the chart unless the file turned the automatic title off', () => {
    const plot = `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser('Revenue', [1, 2, 3])}${ids()}</c:barChart>${axes()}`
    assert.match(draw(chart(plot)).svg, /data-title="">Revenue</)
    assert.ok(!/data-title/.test(draw(chart(plot, '<c:autoTitleDeleted val="1"/>')).svg))
  })

  test('the axis on the right is the secondary one even when its group comes first', () => {
    const plot =
      `<c:lineChart><c:grouping val="standard"/>${ser('Pct', [0.2, 0.3, 0.1])}${ids(3, 4)}</c:lineChart>` +
      `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser('Sales', [100, 140, 90])}${ids(1, 2)}</c:barChart>` +
      `${catAx(1, 2)}${valAx(2, 1)}${valAx(4, 3, '<c:numFmt formatCode="0%" sourceLinked="0"/>', 'r')}${catAx(3, 4).replace('<c:delete val="0"/>', '<c:delete val="1"/>')}`
    const { svg } = draw(chart(plot))
    const x = (label: string) => Number(new RegExp(`<text x="([\\d.]+)"[^>]*>${label}<`).exec(svg)?.[1])
    assert.ok(x('30%') > x('160'), `${x('30%')} vs ${x('160')}`)
    assert.ok(texts(svg).includes('B'), 'category labels from the primary pair')
  })

  test('text is escaped', () => {
    const { svg, text } = draw(chart(`<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser('&lt;img src=x&gt;', [1, 2, 3], ['&lt;b&gt;', 'B', 'C'])}${ids()}</c:barChart>${axes()}`, titled('&lt;script&gt;'), '') , {
      fontFamily: '"x" onload="alert(1)'
    })
    assert.ok(!/<script|<img|<b>/.test(svg))
    assert.ok(!/onload="/.test(svg))
    assert.match(text, /<script>/)
  })

  test('legend position right stacks entries on the right; bottom spreads them below', () => {
    const plot = `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser('Alpha', [1, 2, 3])}${ser('Beta', [1, 2, 3])}${ids()}</c:barChart>${axes()}`
    const legendX = (svg: string) => [...svg.matchAll(/<text x="([\d.]+)" y="([\d.]+)"[^>]*data-legend[^>]*>/g)].map((m) => [+m[1], +m[2]])
    const right = legendX(draw(chart(plot, '<c:legend><c:legendPos val="r"/></c:legend>')).svg)
    assert.equal(right.length, 2)
    assert.ok(right.every(([x]) => x > 360))
    assert.ok(right[0][1] !== right[1][1])
    const bottom = legendX(draw(chart(plot, '<c:legend><c:legendPos val="b"/></c:legend>')).svg)
    assert.ok(bottom.every(([, y]) => y > 270))
    assert.equal(bottom[0][1], bottom[1][1])
  })

  test('the chart text size from txPr', () => {
    const plot = `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser('S', [1, 2, 3])}${ids()}</c:barChart>${axes()}`
    const { svg } = draw(chart(plot, '', ''), {})
    assert.match(svg, /font-size="10"/)
    const big = parseXml(`<c:chartSpace ${NS}><c:chart><c:plotArea>${plot}</c:plotArea></c:chart><c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="1800"/></a:pPr></a:p></c:txPr></c:chartSpace>`)
    assert.match(draw(big).svg, /font-size="18"/)
  })
})

describe('what cannot be drawn', () => {
  test('no plot area, no groups, or only unsupported groups give null', () => {
    assert.equal(renderChart(parseXml(`<c:chartSpace ${NS}><c:chart/></c:chartSpace>`), OPTIONS), null)
    assert.equal(renderChart(chart(axes()), OPTIONS), null)
    assert.equal(renderChart(chart(`<c:surfaceChart>${ser('S', [1, 2, 3])}${ids()}</c:surfaceChart>${axes()}`), OPTIONS), null)
  })

  test('a group with no series, or series with no values, gives null', () => {
    assert.equal(renderChart(chart(`<c:barChart><c:barDir val="col"/>${ids()}</c:barChart>${axes()}`), OPTIONS), null)
  })

  test('non-finite and absurd values never hang or emit NaN', () => {
    const values = ['1e308', '1.7976931348623157e308', '-1e308', 'NaN', 'Infinity', 'x']
    for (const kind of ['barChart', 'lineChart', 'areaChart']) {
      const node = chart(`<c:${kind}><c:barDir val="col"/><c:grouping val="stacked"/>${ser('S', values, ['a', 'b', 'c', 'd', 'e', 'f'])}${ser('T', values, ['a', 'b', 'c', 'd', 'e', 'f'])}${ids()}</c:${kind}>${axes('<c:majorUnit val="1e-300"/>')}`)
      const drawn = renderChart(node, OPTIONS)
      if (drawn) assert.ok(!/NaN|Infinity/.test(drawn.svg), kind)
    }
    const tiny = renderChart(chart(`<c:barChart><c:barDir val="col"/>${ser('S', [1, 2, 3])}${ids()}</c:barChart>${axes('<c:scaling><c:orientation val="minMax"/><c:max val="1e300"/><c:min val="-1e300"/></c:scaling><c:majorUnit val="1e-10"/>')}`), OPTIONS)
    assert.ok(tiny && count(tiny.svg, /<line /g) < 200)
    const pie = renderChart(chart(`<c:pieChart>${ser('S', ['1e308', '1e308', 'NaN'])}</c:pieChart>`), OPTIONS)
    if (pie) assert.ok(!/NaN|Infinity/.test(pie.svg))
    const zero = renderChart(chart(`<c:pieChart>${ser('S', [0, 0, 0])}</c:pieChart>`), OPTIONS)
    if (zero) assert.ok(!/NaN/.test(zero.svg))
  })

  test('a huge chart is bounded', () => {
    const many = Array.from({ length: 5000 }, (_, i) => i)
    const plot = `<c:lineChart><c:grouping val="standard"/>${Array.from({ length: 300 }, (_, i) => ser(`S${i}`, many.slice(0, 200), null)).join('')}${ids()}</c:lineChart>${axes()}`
    const start = Date.now()
    const drawn = renderChart(chart(plot), OPTIONS)
    assert.ok(Date.now() - start < 3000)
    if (drawn) assert.ok(drawn.svg.length < 5_000_000)
  })

  test('values come from readCells when the part holds no cache', () => {
    const plain = `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>S!$B$1</c:f></c:strRef></c:tx><c:cat><c:strRef><c:f>S!$A$2:$A$4</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>S!$B$2:$B$4</c:f></c:numRef></c:val></c:ser>`
    const cells: Record<string, (string | number | null)[]> = { 'S!$B$1': ['Read'], 'S!$A$2:$A$4': ['p', 'q', 'r'], 'S!$B$2:$B$4': [1, '2', 3] }
    const { svg, text } = draw(chart(`<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${plain}${ids()}</c:barChart>${axes()}`), { readCells: (f) => cells[f] ?? [] })
    assert.equal(rects(svg).length, 3)
    assert.match(text, /Read/)
    assert.ok(texts(svg).includes('q'))
  })
})

describe('date systems', () => {
  // XlsxWriter's date_1904 workbook: the chart part says <c:date1904 val="1"/> and caches 1904 serials.
  const dated = (outer: string) =>
    chart(`<c:lineChart><c:grouping val="standard"/><c:ser><c:idx val="0"/><c:order val="0"/><c:cat>${nums([43830, 43831], 'yyyy-mm-dd')}</c:cat><c:val>${nums([1, 2])}</c:val></c:ser>${ids()}</c:lineChart><c:dateAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="yyyy-mm-dd" sourceLinked="1"/><c:crossAx val="2"/></c:dateAx>${valAx(2, 1)}`, '', outer)

  test('a chart from a 1904 workbook labels its dates by that system', () => {
    const labels = texts(draw(dated('<c:date1904 val="1"/>')).svg)
    assert.ok(labels.includes('2024-01-01') && labels.includes('2024-01-02'), labels.join(' | '))
  })

  test('the renderer can say the workbook uses 1904 when the chart part does not', () => {
    const labels = texts(draw(dated(''), { date1904: true }).svg)
    assert.ok(labels.includes('2024-01-01'), labels.join(' | '))
    assert.ok(texts(draw(dated('')).svg).includes('2019-12-31'), 'the 1900 system stays the default')
  })
})

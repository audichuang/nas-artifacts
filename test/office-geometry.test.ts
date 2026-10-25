import assert from 'node:assert/strict'
import { test } from 'node:test'
import { PRESET_NAMES, adjustValues, customGeometry, fillShade, presetGeometry, shapeGeometry, type Geometry } from '../src/office-geometry.ts'
import { parseXml } from '../src/xml.ts'

const EMU = 12700
const geom = (name: string, wPt: number, hPt: number, adjust: Record<string, number> = {}): Geometry => {
  const out = presetGeometry(name, wPt * EMU, hPt * EMU, adjust)
  assert.ok(out, `${name} is a preset`)
  return out
}

/** Every point a path visits: the end of each command, in order. */
function points(d: string): [number, number][] {
  const out: [number, number][] = []
  for (const [, cmd, body] of d.matchAll(/([MLACQZ])([^MLACQZ]*)/g)) {
    if (cmd === 'Z') continue
    const nums = body.trim().split(/[\s,]+/).filter(Boolean).map(Number)
    out.push([nums[nums.length - 2], nums[nums.length - 1]])
  }
  return out
}

const near = (actual: [number, number] | undefined, expected: [number, number], tolerance = 0.02) => {
  assert.ok(actual, `expected a point near ${expected}`)
  assert.ok(Math.abs(actual[0] - expected[0]) <= tolerance && Math.abs(actual[1] - expected[1]) <= tolerance, `${actual} is not near ${expected}`)
}
const hasPoint = (d: string, expected: [number, number], tolerance = 0.02) =>
  assert.ok(points(d).some(([x, y]) => Math.abs(x - expected[0]) <= tolerance && Math.abs(y - expected[1]) <= tolerance), `${d} does not visit ${expected}`)

const finite = (g: Geometry) => {
  for (const path of g.paths) {
    assert.match(path.d, /^[MLACQZ0-9 ,.\-]*$/, path.d)
    for (const n of path.d.match(/-?[\d.]+/g) ?? []) assert.ok(Number.isFinite(Number(n)), path.d)
  }
  for (const v of Object.values(g.textRect)) assert.ok(Number.isFinite(v))
}

test('the table carries every ECMA-376 preset shape', () => {
  assert.equal(PRESET_NAMES.length, 187)
  for (const name of ['rect', 'heart', 'cloud', 'wedgeRectCallout', 'actionButtonHome', 'chartPlus', 'gear9', 'swooshArrow']) assert.ok(PRESET_NAMES.includes(name), name)
  assert.equal(presetGeometry('notAShape', 1, 1, {}), null)
  assert.equal(presetGeometry('__proto__', 1, 1, {}), null)
  assert.equal(presetGeometry('toString', 1, 1, {}), null)
})

test('every preset evaluates to finite path data at a 2:1 box, zero size and huge size', () => {
  for (const name of PRESET_NAMES) {
    for (const [w, h] of [[200 * EMU, 100 * EMU], [0, 0], [0, 50 * EMU], [1e12, 1e12], [1e12, 3]]) {
      const g = presetGeometry(name, w, h, {})
      assert.ok(g, name)
      assert.ok(g.paths.length > 0, name)
      finite(g)
    }
  }
})

test('rect is the box, and its text rectangle is the whole box', () => {
  const g = geom('rect', 100, 50)
  assert.deepEqual(g.paths, [{ d: 'M0,0 L100,0 L100,50 L0,50 Z', fill: 'norm', stroke: true, closed: true }])
  assert.deepEqual(g.textRect, { l: 0, t: 0, r: 100, b: 50 })
})

test('roundRect corners follow adj (a share of the shorter side) and default to 16667', () => {
  const g = geom('roundRect', 200, 100, { adj: 25000 })
  assert.equal(g.paths[0].d, 'M0,25 A25,25 0 0 1 25,0 L175,0 A25,25 0 0 1 200,25 L200,75 A25,25 0 0 1 175,100 L25,100 A25,25 0 0 1 0,75 Z')
  near([g.textRect.l, g.textRect.t], [7.32, 7.32])
  assert.match(geom('roundRect', 200, 100).paths[0].d, /^M0,16\.67 A16\.67,16\.67 /)
  // An adjust value outside its range is pinned, as PowerPoint does.
  assert.match(geom('roundRect', 200, 100, { adj: 900000 }).paths[0].d, /^M0,50 A50,50 /)
  // A square corner (adj 0) draws no curve at all.
  assert.equal(geom('roundRect', 200, 100, { adj: 0 }).paths[0].d, 'M0,0 L0,0 L200,0 L200,0 L200,100 L200,100 L0,100 L0,100 Z')
})

test('ellipse is four quarter arcs through the side midpoints', () => {
  const { d } = geom('ellipse', 200, 100).paths[0]
  assert.equal(d, 'M0,50 A100,50 0 0 1 100,0 A100,50 0 0 1 200,50 A100,50 0 0 1 100,100 A100,50 0 0 1 0,50 Z')
})

test('arcTo angles are visual angles on the ellipse, not parametric ones', () => {
  // A chord from 45° to 135° on a 2:1 ellipse: both ends lie on the rays at those angles.
  const { d } = geom('chord', 200, 100, { adj1: 2700000, adj2: 8100000 }).paths[0]
  const [start, end] = points(d)
  near(start, [144.72, 94.72])
  near(end, [55.28, 94.72])
  // A parametric reading would have ended at (29.29, 85.36).
  // pie's arc from 0° to 270° lands on the top centre and is a large arc.
  const pie = geom('pie', 200, 100).paths[0].d
  assert.match(pie, /A100,50 0 1 1 100,0 L100,50 Z$/)
})

test('a whole turn in one arcTo is drawn, not dropped', () => {
  const xml = parseXml(
    '<a:custGeom xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:pathLst><a:path w="100" h="100">' +
      '<a:moveTo><a:pt x="100" y="50"/></a:moveTo><a:arcTo wR="50" hR="50" stAng="0" swAng="21600000"/><a:close/></a:path></a:pathLst></a:custGeom>'
  )
  const g = customGeometry(xml, 100 * EMU, 100 * EMU)
  assert.ok(g)
  assert.equal(g.paths[0].d, 'M100,50 A50,50 0 0 1 0,50 A50,50 0 0 1 100,50 Z')
})

test('rightArrow: shaft from adj1 of the height, head from adj2 of the shorter side', () => {
  const { d } = geom('rightArrow', 200, 100).paths[0]
  assert.equal(d, 'M0,25 L150,25 L150,0 L200,50 L150,100 L150,75 L0,75 Z')
  assert.equal(geom('rightArrow', 200, 100, { adj1: 20000, adj2: 100000 }).paths[0].d, 'M0,40 L100,40 L100,0 L200,50 L100,100 L100,60 L0,60 Z')
})

test('wedgeRectCallout points its tail at adj1/adj2 from the centre', () => {
  const { d } = geom('wedgeRectCallout', 200, 100).paths[0]
  hasPoint(d, [58.33, 112.5])
  hasPoint(d, [0, 0])
  hasPoint(d, [200, 100])
  hasPoint(geom('wedgeRectCallout', 200, 100, { adj1: 70000, adj2: 0 }).paths[0].d, [240, 50])
})

test('heart is drawn with curves meeting at the bottom point and the top notch', () => {
  const { d } = geom('heart', 200, 100).paths[0]
  assert.equal((d.match(/C/g) ?? []).length, 2)
  hasPoint(d, [100, 25])
  hasPoint(d, [100, 100])
})

test('star5 has ten corners with the top point centred', () => {
  const { d } = geom('star5', 200, 100).paths[0]
  const corners = points(d)
  assert.equal(corners.length, 10)
  assert.ok(corners.some(([x, y]) => Math.abs(x - 100) < 0.01 && Math.abs(y) < 0.01))
  // The inner radius follows adj (default 19098 of the outer radius × 50000).
  const fat = points(geom('star5', 200, 100, { adj: 40000 }).paths[0].d)
  const inner = (pts: [number, number][]) => Math.hypot(pts[1][0] - 100, pts[1][1] - 50)
  assert.ok(inner(fat) > inner(corners))
})

test('cloud: every bump arc chains into the next and the outline closes where it began', () => {
  const g = geom('cloud', 200, 100)
  const outline = g.paths[0]
  assert.equal((outline.d.match(/A/g) ?? []).length, 11)
  const pts = points(outline.d)
  near(pts[pts.length - 1], pts[0], 0.5) // the standard's literals close to within 0.3%
  // The small puffs are separate, unfilled, stroked paths.
  assert.ok(g.paths.slice(1).every((path) => path.fill === 'none'))
})

test('chevron and flowChartTerminator', () => {
  assert.equal(geom('chevron', 200, 100).paths[0].d, 'M0,0 L150,0 L200,50 L150,100 L0,100 L50,50 Z')
  assert.equal(geom('chevron', 200, 100, { adj: 20000 }).paths[0].d, 'M0,0 L180,0 L200,50 L180,100 L0,100 L20,50 Z')
  // Its path is laid out in a 21600 square, stretched to the box.
  assert.equal(geom('flowChartTerminator', 216, 108).paths[0].d, 'M34.75,0 L181.25,0 A34.75,54 0 0 1 181.25,108 L34.75,108 A34.75,54 0 0 1 34.75,0 Z')
})

test('path attributes come through: open lines, shaded faces, unstroked fills', () => {
  const line = geom('line', 100, 50).paths[0]
  assert.equal(line.fill, 'norm')
  assert.equal(line.closed, false)
  assert.equal(line.d, 'M0,0 L100,50')
  const can = geom('can', 100, 200).paths
  assert.deepEqual(can.map((p) => [p.fill, p.stroke]), [['norm', false], ['lighten', false], ['none', true]])
  assert.deepEqual(fillShade('darken'), { color: '#000', opacity: 0.4 })
  assert.equal(fillShade('norm'), null)
})

test('customGeometry evaluates its own adjust values and guides, including forward references', () => {
  const xml = parseXml(
    '<a:custGeom xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<a:avLst><a:gd name="inset" fmla="val 10000"/></a:avLst>' +
      '<a:gdLst><a:gd name="x2" fmla="+- r 0 x1"/><a:gd name="x1" fmla="*/ w inset 100000"/></a:gdLst>' +
      '<a:rect l="x1" t="t" r="x2" b="b"/>' +
      '<a:pathLst><a:path><a:moveTo><a:pt x="x1" y="t"/></a:moveTo><a:lnTo><a:pt x="x2" y="t"/></a:lnTo>' +
      '<a:quadBezTo><a:pt x="r" y="vc"/><a:pt x="x2" y="b"/></a:quadBezTo><a:lnTo><a:pt x="x1" y="b"/></a:lnTo><a:close/></a:path>' +
      '<a:path w="10" h="10" fill="none" stroke="0"><a:moveTo><a:pt x="0" y="5"/></a:moveTo><a:cubicBezTo><a:pt x="2" y="0"/><a:pt x="8" y="10"/><a:pt x="10" y="5"/></a:cubicBezTo></a:path>' +
      '</a:pathLst></a:custGeom>'
  )
  const g = customGeometry(xml, 200 * EMU, 100 * EMU)
  assert.ok(g)
  assert.deepEqual(g.paths, [
    { d: 'M20,0 L180,0 Q200,50 180,100 L20,100 Z', fill: 'norm', stroke: true, closed: true },
    { d: 'M0,50 C40,0 160,100 200,50', fill: 'none', stroke: false, closed: false }
  ])
  assert.deepEqual(g.textRect, { l: 20, t: 0, r: 180, b: 100 })
  // Guides run in order: a name defined again later (gear6 reuses a1) leaves earlier formulas their value.
  const again = parseXml('<custGeom><gdLst><gd name="a" fmla="val 5"/><gd name="b" fmla="+- a 0 0"/><gd name="a" fmla="val 7"/></gdLst><pathLst><path w="1" h="1"><moveTo><pt x="b" y="a"/></moveTo></path></pathLst></custGeom>')
  assert.equal(customGeometry(again, EMU, EMU)?.paths[0].d, 'M5,7')
  // A long chain of forward references is evaluated once per guide, not once per use.
  const chain = Array.from({ length: 40 }, (_, i) => `<gd name="g${i}" fmla="+- g${i + 1} g${i + 1} 0"/>`).join('') + '<gd name="g40" fmla="val 1"/>'
  const forward = parseXml(`<custGeom><gdLst>${chain}</gdLst><pathLst><path w="1" h="1"><moveTo><pt x="0" y="0"/></moveTo><lnTo><pt x="g35" y="0"/></lnTo></path></pathLst></custGeom>`)
  assert.equal(customGeometry(forward, EMU, EMU)?.paths[0].d, 'M0,0 L32,0')
  // A guide that refers to itself evaluates to 0 rather than looping.
  const loop = parseXml('<custGeom><gdLst><gd name="a" fmla="+- a 1 0"/></gdLst><pathLst><path><moveTo><pt x="a" y="0"/></moveTo></path></pathLst></custGeom>')
  assert.equal(customGeometry(loop, EMU, EMU)?.paths[0].d, 'M0,0')
  assert.equal(customGeometry(parseXml('<custGeom/>'), EMU, EMU), null)
})

test('shapeGeometry and adjustValues read a shape properties element', () => {
  const spPr = parseXml(
    '<p:spPr xmlns:a="a" xmlns:p="p"><a:xfrm/><a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 50000"/><a:gd name="bad" fmla="*/ w 1 2"/></a:avLst></a:prstGeom></p:spPr>'
  )
  assert.deepEqual(adjustValues(spPr.children[1]), { adj: 50000 })
  assert.match(shapeGeometry(spPr, 200 * EMU, 100 * EMU)?.paths[0].d ?? '', /^M0,50 A50,50/)
  assert.equal(shapeGeometry(parseXml('<spPr/>'), EMU, EMU), null)
  assert.equal(shapeGeometry(parseXml('<spPr><prstGeom prst="nope"/></spPr>'), EMU, EMU), null)
})

test('formula operators follow ECMA-376 20.1.9.11', () => {
  const ops: [string, number][] = [
    ['*/ 6 4 3', 8], ['+- 6 4 3', 7], ['+/ 6 4 2', 5], ['?: 1 4 3', 4], ['?: 0 4 3', 3], ['abs -7', 7],
    ['at2 1 1', 2700000], ['cat2 10 1 0', 10], ['sat2 10 0 1', 10], ['cos 10 0', 10], ['sin 10 5400000', 10],
    ['max 3 9', 9], ['min 3 9', 3], ['mod 3 4 12', 13], ['pin 0 -5 10', 0], ['pin 0 50 10', 10], ['sqrt 81', 9],
    ['tan 10 2700000', 10], ['val 42', 42], ['*/ 1 1 0', 0], ['cd2', 0]
  ]
  for (const [fmla, expected] of ops) {
    const xml = parseXml(`<custGeom><gdLst><gd name="g" fmla="${fmla}"/></gdLst><pathLst><path w="1" h="1"><moveTo><pt x="g" y="0"/></moveTo></path></pathLst></custGeom>`)
    const x = Number(customGeometry(xml, EMU, EMU)?.paths[0].d.slice(1).split(',')[0])
    assert.ok(Math.abs(x - expected) < 0.01, `${fmla} → ${x}, expected ${expected}`)
  }
})

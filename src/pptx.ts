import { posix } from 'node:path'
import { escapeHtml } from './content-render.ts'
import { failureScript, isDrawablePart, officeMediaBase, officeMediaUrl } from './office.ts'
import { deckPage, type DeckAssets } from './slides.ts'
import { attr, child, children, local, num, parseXml, path, relAttr, type XmlNode } from './xml.ts'
import { ZipReader } from './zip.ts'

/**
 * A PowerPoint file published as is (`*.pptx`), drawn on the server as slides the deck runtime pages
 * through: every shape is absolutely placed in slide-relative units, text stays text. Pictures load
 * from the file itself through the Office media route (office.ts), inside the artifact's CSP scope. What it does not draw (charts it cannot read, OLE, video) shows as a
 * labelled box; the original file is always one download away.
 */
const EMU_PER_PT = 12700

// ---------------------------------------------------------------------------------------------
// Package parts and relationships

type Rel = { type: string; target: string; external: boolean }

class Package {
  readonly zip: ZipReader
  private readonly parts = new Map<string, XmlNode | null>()
  private readonly relsCache = new Map<string, Map<string, Rel>>()

  constructor(data: Buffer) {
    this.zip = new ZipReader(data)
  }

  xml(name: string): XmlNode | null {
    if (!this.parts.has(name)) {
      const text = this.zip.text(name)
      this.parts.set(name, text === null ? null : parseXml(text))
    }
    return this.parts.get(name) ?? null
  }

  rels(part: string): Map<string, Rel> {
    let map = this.relsCache.get(part)
    if (map) return map
    map = new Map()
    const dir = posix.dirname(part)
    const doc = this.xml(posix.join(dir, '_rels', `${posix.basename(part)}.rels`))
    for (const rel of children(doc, 'Relationship')) {
      const id = attr(rel, 'Id')
      const target = attr(rel, 'Target')
      if (!id || target === undefined) continue
      const external = attr(rel, 'TargetMode') === 'External'
      let resolved = target
      if (!external) {
        resolved = target.startsWith('/') ? target.slice(1) : posix.normalize(posix.join(dir, target))
        if (resolved.startsWith('..')) continue
        try {
          resolved = decodeURIComponent(resolved)
        } catch {}
      }
      map.set(id, { type: attr(rel, 'Type') ?? '', target: resolved, external })
    }
    this.relsCache.set(part, map)
    return map
  }

  rel(part: string, id: string | undefined): Rel | undefined {
    return id ? this.rels(part).get(id) : undefined
  }

  relOfType(part: string, suffix: string): Rel | undefined {
    for (const rel of this.rels(part).values()) if (rel.type.endsWith(`/${suffix}`)) return rel
    return undefined
  }
}

// ---------------------------------------------------------------------------------------------
// Colours

type Rgba = { r: number; g: number; b: number; a: number }

type Theme = {
  colors: Record<string, string>
  major: { latin?: string; ea?: string }
  minor: { latin?: string; ea?: string }
  fills: XmlNode[]
  bgFills: XmlNode[]
  lines: XmlNode[]
}

const PRESET_COLORS: Record<string, string> = {
  black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', gray: '808080', grey: '808080',
  darkGray: 'A9A9A9', lightGray: 'D3D3D3', orange: 'FFA500', purple: '800080', navy: '000080', teal: '008080', silver: 'C0C0C0',
  maroon: '800000', olive: '808000', lime: '00FF00', aqua: '00FFFF', cyan: '00FFFF', fuchsia: 'FF00FF', magenta: 'FF00FF'
}

const COLOR_TAGS = new Set(['srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'scrgbClr', 'hslClr'])

const hexToRgba = (hex: string): Rgba | null => {
  if (!/^[0-9a-fA-F]{6}$/.test(hex)) return null
  return { r: parseInt(hex.slice(0, 2), 16), g: parseInt(hex.slice(2, 4), 16), b: parseInt(hex.slice(4, 6), 16), a: 1 }
}

function rgbToHsl({ r, g, b }: Rgba): [number, number, number] {
  const [R, G, B] = [r / 255, g / 255, b / 255]
  const max = Math.max(R, G, B)
  const min = Math.min(R, G, B)
  const l = (max + min) / 2
  if (max === min) return [0, 0, l]
  const d = max - min
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
  const h = max === R ? (G - B) / d + (G < B ? 6 : 0) : max === G ? (B - R) / d + 2 : (R - G) / d + 4
  return [h * 60, s, l]
}

function hslToRgb(h: number, s: number, l: number, a: number): Rgba {
  const hue = (p: number, q: number, t: number) => {
    t = ((t % 1) + 1) % 1
    if (t < 1 / 6) return p + (q - p) * 6 * t
    if (t < 1 / 2) return q
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6
    return p
  }
  if (s === 0) return { r: l * 255, g: l * 255, b: l * 255, a }
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s
  const p = 2 * l - q
  const H = h / 360
  return { r: hue(p, q, H + 1 / 3) * 255, g: hue(p, q, H) * 255, b: hue(p, q, H - 1 / 3) * 255, a }
}

const clamp01 = (value: number) => Math.min(1, Math.max(0, value))

function applyModifiers(color: Rgba, element: XmlNode): Rgba {
  let current = { ...color }
  for (const mod of element.children) {
    const value = (num(mod, 'val') ?? 0) / 100000
    const name = local(mod.name)
    if (name === 'alpha') current.a = clamp01(value)
    else if (name === 'alphaMod') current.a = clamp01(current.a * value)
    else if (name === 'alphaOff') current.a = clamp01(current.a + value)
    else if (name === 'tint') {
      current = { ...current, r: current.r + (255 - current.r) * (1 - value), g: current.g + (255 - current.g) * (1 - value), b: current.b + (255 - current.b) * (1 - value) }
    } else if (name === 'shade') {
      current = { ...current, r: current.r * value, g: current.g * value, b: current.b * value }
    } else if (['lumMod', 'lumOff', 'satMod', 'satOff', 'hueMod', 'hueOff'].includes(name)) {
      let [h, s, l] = rgbToHsl(current)
      if (name === 'lumMod') l = clamp01(l * value)
      else if (name === 'lumOff') l = clamp01(l + value)
      else if (name === 'satMod') s = clamp01(s * value)
      else if (name === 'satOff') s = clamp01(s + value)
      else if (name === 'hueMod') h = (h * value) % 360
      else h = (h + (num(mod, 'val') ?? 0) / 60000) % 360
      current = hslToRgb(h, s, l, current.a)
    } else if (name === 'comp') {
      const [h, s, l] = rgbToHsl(current)
      current = hslToRgb((h + 180) % 360, s, l, current.a)
    } else if (name === 'inv') current = { ...current, r: 255 - current.r, g: 255 - current.g, b: 255 - current.b }
    else if (name === 'gray') {
      const y = 0.3 * current.r + 0.59 * current.g + 0.11 * current.b
      current = { ...current, r: y, g: y, b: y }
    }
  }
  return current
}

const cssColor = ({ r, g, b, a }: Rgba): string => {
  const [R, G, B] = [r, g, b].map((v) => Math.round(Math.min(255, Math.max(0, v))))
  return a >= 1 ? `rgb(${R},${G},${B})` : `rgba(${R},${G},${B},${Math.round(a * 1000) / 1000})`
}

// ---------------------------------------------------------------------------------------------
// Rendering context

type Frame = { x: number; y: number; cx: number; cy: number }

type Ctx = {
  pkg: Package
  theme: Theme
  clrMap: Record<string, string>
  part: string
  slideNumber: number
  slideOf: (part: string) => number | undefined
  media: (part: string) => string
  master: Master
  layout: Layout
  phColor?: Rgba
  presentationStyle: XmlNode | null
  tableStyles: XmlNode | null
}

type Master = { part: string; xml: XmlNode; clrMap: Record<string, string>; theme: Theme }
type Layout = { part: string; xml: XmlNode; master: Master }

function resolveColor(element: XmlNode | null | undefined, ctx: Ctx): Rgba | null {
  if (!element) return null
  const name = local(element.name)
  let base: Rgba | null = null
  if (name === 'srgbClr') base = hexToRgba(attr(element, 'val') ?? '')
  else if (name === 'schemeClr') {
    const val = attr(element, 'val') ?? ''
    if (val === 'phClr') base = ctx.phColor ? { ...ctx.phColor } : null
    else base = hexToRgba(ctx.theme.colors[ctx.clrMap[val] ?? val] ?? '')
  } else if (name === 'sysClr') base = hexToRgba(attr(element, 'lastClr') ?? (attr(element, 'val') === 'window' ? 'FFFFFF' : '000000'))
  else if (name === 'prstClr') base = hexToRgba(PRESET_COLORS[attr(element, 'val') ?? ''] ?? '000000')
  else if (name === 'scrgbClr') {
    const linear = (v: number | undefined) => {
      const c = clamp01((v ?? 0) / 100000)
      return 255 * (c <= 0.0031308 ? c * 12.92 : 1.055 * c ** (1 / 2.4) - 0.055)
    }
    base = { r: linear(num(element, 'r')), g: linear(num(element, 'g')), b: linear(num(element, 'b')), a: 1 }
  } else if (name === 'hslClr') {
    base = hslToRgb((num(element, 'hue') ?? 0) / 60000, (num(element, 'sat') ?? 0) / 100000, (num(element, 'lum') ?? 0) / 100000, 1)
  }
  return base ? applyModifiers(base, element) : null
}

/** The colour held by a fill-like element (`solidFill`, `buClr`, `fontRef`…). */
const colorIn = (node: XmlNode | null | undefined, ctx: Ctx): Rgba | null =>
  resolveColor(node?.children.find((candidate) => COLOR_TAGS.has(local(candidate.name))), ctx)

const FILL_TAGS = new Set(['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'])

const firstFill = (node: XmlNode | null | undefined): XmlNode | null =>
  node?.children.find((candidate) => FILL_TAGS.has(local(candidate.name))) ?? null

/** CSS `background` for a fill element; '' for none. */
function fillCss(fill: XmlNode | null, ctx: Ctx): string {
  if (!fill) return ''
  const name = local(fill.name)
  if (name === 'solidFill') {
    const color = colorIn(fill, ctx)
    return color ? cssColor(color) : ''
  }
  if (name === 'gradFill') {
    const stops = children(child(fill, 'gsLst'), 'gs')
      .map((gs) => ({ pos: (num(gs, 'pos') ?? 0) / 1000, color: colorIn(gs, ctx) }))
      .filter((stop): stop is { pos: number; color: Rgba } => stop.color !== null)
      .sort((a, b) => a.pos - b.pos)
    if (stops.length === 0) return ''
    if (stops.length === 1) return cssColor(stops[0].color)
    const list = stops.map((stop) => `${cssColor(stop.color)} ${round(stop.pos)}%`).join(',')
    if (child(fill, 'path')) return `radial-gradient(ellipse at center,${list})`
    const angle = (num(child(fill, 'lin'), 'ang') ?? 0) / 60000
    return `linear-gradient(${round(angle + 90)}deg,${list})`
  }
  if (name === 'blipFill') {
    const url = blipUrl(fill, ctx)
    if (!url) return ''
    const tile = child(fill, 'tile')
    return tile ? `url('${url}') repeat` : `url('${url}') center/100% 100% no-repeat`
  }
  if (name === 'pattFill') {
    const color = colorIn(child(fill, 'fgClr'), ctx)
    return color ? cssColor(color) : ''
  }
  return ''
}

/** The URL of a picture part a `blip` names, or null when browsers cannot show it. */
function blipUrl(blipFill: XmlNode, ctx: Ctx): string | null {
  const blip = child(blipFill, 'blip')
  // An SVG picture carries its PNG fallback in r:embed and the vector in an extension.
  const svg = blip ? findDeep(blip, 'svgBlip') : null
  for (const id of [relAttr(svg, 'embed'), relAttr(blip, 'embed')]) {
    const rel = ctx.pkg.rel(ctx.part, id)
    if (rel && !rel.external && isDrawablePart(rel.target) && ctx.pkg.zip.has(rel.target)) return ctx.media(rel.target)
  }
  return null
}

function findDeep(node: XmlNode, name: string): XmlNode | null {
  for (const candidate of node.children) {
    if (local(candidate.name) === name) return candidate
    const found = findDeep(candidate, name)
    if (found) return found
  }
  return null
}

const round = (value: number, digits = 3) => {
  const factor = 10 ** digits
  return Math.round(value * factor) / factor
}

const pt = (value: number) => `calc(${round(value)}*var(--pt))`
const emuPt = (emu: number) => pt(emu / EMU_PER_PT)

// ---------------------------------------------------------------------------------------------
// Styles referenced from the theme (p:style)

function styleRefFill(style: XmlNode | null, ctx: Ctx): string {
  const ref = child(style, 'fillRef')
  const idx = num(ref, 'idx') ?? 0
  if (!ref || idx === 0) return ''
  const list = idx >= 1000 ? ctx.theme.bgFills : ctx.theme.fills
  const fill = list[(idx >= 1000 ? idx - 1000 : idx) - 1]
  return fill ? fillCss(fill, { ...ctx, phColor: colorIn(ref, ctx) ?? undefined }) : ''
}

type Line = { width: number; color: string; dash: string; head?: string; tail?: string } | null

function lineOf(ln: XmlNode | null, style: XmlNode | null, ctx: Ctx): Line {
  const ref = child(style, 'lnRef')
  const idx = num(ref, 'idx') ?? 0
  const themed = idx > 0 ? ctx.theme.lines[idx - 1] : undefined
  const refCtx = { ...ctx, phColor: colorIn(ref, ctx) ?? undefined }
  const sources: [XmlNode, Ctx][] = []
  if (ln) sources.push([ln, ctx])
  if (themed) sources.push([themed, refCtx])
  if (sources.length === 0) return null
  const fill = sources.map(([node, c]) => [firstFill(node), c] as const).find(([node]) => node)
  if (!fill || local(fill[0]!.name) === 'noFill') return null
  const color = local(fill[0]!.name) === 'solidFill' ? colorIn(fill[0], fill[1]) : local(fill[0]!.name) === 'gradFill' ? colorIn(child(child(fill[0], 'gsLst'), 'gs'), fill[1]) : null
  if (!color) return null
  const width = (sources.map(([node]) => num(node, 'w')).find((w) => w !== undefined) ?? 9525) / EMU_PER_PT
  const dash = sources.map(([node]) => attr(child(node, 'prstDash'), 'val')).find(Boolean) ?? 'solid'
  const end = (name: string) => {
    const type = sources.map(([node]) => attr(child(node, name), 'type')).find(Boolean)
    return type && type !== 'none' ? type : undefined
  }
  return { width, color: cssColor(color), dash, head: end('headEnd'), tail: end('tailEnd') }
}

const dashArray = (dash: string, width: number): string | null => {
  const w = Math.max(width, 0.75)
  const patterns: Record<string, number[]> = {
    dot: [1, 1], sysDot: [1, 1], dash: [4, 3], sysDash: [3, 1], lgDash: [8, 3], dashDot: [4, 3, 1, 3], sysDashDot: [3, 1, 1, 1],
    lgDashDot: [8, 3, 1, 3], lgDashDotDot: [8, 3, 1, 3, 1, 3], sysDashDotDot: [3, 1, 1, 1, 1, 1]
  }
  const pattern = patterns[dash]
  return pattern ? pattern.map((v) => round(v * w)).join(' ') : null
}

function shadowCss(spPr: XmlNode | null, style: XmlNode | null, ctx: Ctx): string {
  // Theme effect styles (effectRef) are left out: they are mostly subtle and often turned off in the theme.
  void style
  const shadow = child(child(spPr, 'effectLst'), 'outerShdw')
  if (!shadow) return ''
  const color = colorIn(shadow, ctx)
  if (!color) return ''
  const dist = (num(shadow, 'dist') ?? 0) / EMU_PER_PT
  const dir = ((num(shadow, 'dir') ?? 0) / 60000) * (Math.PI / 180)
  const blur = (num(shadow, 'blurRad') ?? 0) / EMU_PER_PT
  return `drop-shadow(${pt(Math.cos(dir) * dist)} ${pt(Math.sin(dir) * dist)} ${pt(blur / 2)} ${cssColor(color)})`
}

// ---------------------------------------------------------------------------------------------
// Geometry

const adjOf = (prstGeom: XmlNode | null): Record<string, number> => {
  const out: Record<string, number> = {}
  for (const gd of children(child(prstGeom, 'avLst'), 'gd')) {
    const match = /^val\s+(-?\d+)/.exec(attr(gd, 'fmla') ?? '')
    if (match) out[attr(gd, 'name') ?? ''] = Number(match[1])
  }
  return out
}

/** An SVG path for a preset shape in a w × h box, or null for the ones drawn as a CSS box. */
function presetPath(prst: string, w: number, h: number, adj: Record<string, number>): string | null {
  const ss = Math.min(w, h)
  const a = (name: string, fallback: number) => (adj[name] ?? fallback) / 100000
  const poly = (points: [number, number][]) => `M${points.map(([x, y]) => `${round(x, 2)},${round(y, 2)}`).join(' L')} Z`
  switch (prst) {
    case 'triangle': {
      const x = w * a('adj', 50000)
      return poly([[x, 0], [w, h], [0, h]])
    }
    case 'rtTriangle':
      return poly([[0, 0], [w, h], [0, h]])
    case 'diamond':
    case 'flowChartDecision':
      return poly([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]])
    case 'parallelogram':
    case 'flowChartInputOutput': {
      const x = ss * a('adj', 25000)
      return poly([[x, 0], [w, 0], [w - x, h], [0, h]])
    }
    case 'trapezoid': {
      const x = ss * a('adj', 25000)
      return poly([[x, 0], [w - x, 0], [w, h], [0, h]])
    }
    case 'pentagon':
      return poly([[w / 2, 0], [w, h * 0.38], [w * 0.81, h], [w * 0.19, h], [0, h * 0.38]])
    case 'homePlate': {
      const x = w - ss * a('adj', 50000)
      return poly([[0, 0], [x, 0], [w, h / 2], [x, h], [0, h]])
    }
    case 'chevron': {
      const x = ss * a('adj', 50000)
      return poly([[0, 0], [w - x, 0], [w, h / 2], [w - x, h], [0, h], [x, h / 2]])
    }
    case 'hexagon': {
      const x = ss * a('adj', 25000)
      return poly([[x, 0], [w - x, 0], [w, h / 2], [w - x, h], [x, h], [0, h / 2]])
    }
    case 'octagon': {
      const x = ss * a('adj', 29289)
      return poly([[x, 0], [w - x, 0], [w, x], [w, h - x], [w - x, h], [x, h], [0, h - x], [0, x]])
    }
    case 'rightArrow': {
      const shaft = h * a('adj1', 50000)
      const head = ss * a('adj2', 50000)
      const y1 = (h - shaft) / 2
      return poly([[0, y1], [w - head, y1], [w - head, 0], [w, h / 2], [w - head, h], [w - head, h - y1], [0, h - y1]])
    }
    case 'leftArrow': {
      const shaft = h * a('adj1', 50000)
      const head = ss * a('adj2', 50000)
      const y1 = (h - shaft) / 2
      return poly([[w, y1], [head, y1], [head, 0], [0, h / 2], [head, h], [head, h - y1], [w, h - y1]])
    }
    case 'upArrow': {
      const shaft = w * a('adj1', 50000)
      const head = ss * a('adj2', 50000)
      const x1 = (w - shaft) / 2
      return poly([[x1, h], [x1, head], [0, head], [w / 2, 0], [w, head], [w - x1, head], [w - x1, h]])
    }
    case 'downArrow': {
      const shaft = w * a('adj1', 50000)
      const head = ss * a('adj2', 50000)
      const x1 = (w - shaft) / 2
      return poly([[x1, 0], [x1, h - head], [0, h - head], [w / 2, h], [w, h - head], [w - x1, h - head], [w - x1, 0]])
    }
    case 'leftRightArrow': {
      const shaft = h * a('adj1', 50000)
      const head = ss * a('adj2', 50000)
      const y1 = (h - shaft) / 2
      return poly([[0, h / 2], [head, 0], [head, y1], [w - head, y1], [w - head, 0], [w, h / 2], [w - head, h], [w - head, h - y1], [head, h - y1], [head, h]])
    }
    case 'plus':
    case 'mathPlus': {
      const x = ss * a('adj', 25000)
      return poly([[x, 0], [w - x, 0], [w - x, x], [w, x], [w, h - x], [w - x, h - x], [w - x, h], [x, h], [x, h - x], [0, h - x], [0, x], [x, x]])
    }
    case 'star5': {
      const points: [number, number][] = []
      for (let i = 0; i < 10; i++) {
        const angle = -Math.PI / 2 + (i * Math.PI) / 5
        const r = i % 2 === 0 ? 0.5 : 0.19
        points.push([w / 2 + Math.cos(angle) * r * w, h / 2 + Math.sin(angle) * r * h * 1.05 + h * 0.025])
      }
      return poly(points)
    }
    case 'snip1Rect': {
      const x = ss * a('adj', 16667)
      return poly([[0, 0], [w - x, 0], [w, x], [w, h], [0, h]])
    }
    case 'round2SameRect': {
      const r = ss * a('adj1', 16667)
      return `M${r},0 L${w - r},0 Q${w},0 ${w},${r} L${w},${h} L0,${h} L0,${r} Q0,0 ${r},0 Z`
    }
    case 'flowChartTerminator': {
      const r = h / 2
      return `M${r},0 L${w - r},0 A${r},${r} 0 0 1 ${w - r},${h} L${r},${h} A${r},${r} 0 0 1 ${r},0 Z`
    }
    case 'line':
    case 'straightConnector1':
    case 'curvedConnector3':
      return `M0,0 L${w},${h}`
    case 'bentConnector2':
      return `M0,0 L${w},0 L${w},${h}`
    case 'bentConnector3': {
      const x = w * a('adj1', 50000)
      return `M0,0 L${x},0 L${x},${h} L${w},${h}`
    }
    default:
      return null
  }
}

/** An SVG path takes one colour: a gradient draws as its middle stop, a picture fill as none. */
function svgFill(fillNode: XmlNode | null, background: string, ctx: Ctx): string {
  if (!background || background.startsWith('url(')) return 'none'
  if (fillNode && local(fillNode.name) === 'gradFill') {
    const stops = children(child(fillNode, 'gsLst'), 'gs')
    const color = colorIn(stops[Math.floor(stops.length / 2)], ctx)
    return color ? cssColor(color) : 'none'
  }
  return background.includes('gradient') ? 'none' : background
}

const OPEN_PATHS = new Set(['line', 'straightConnector1', 'curvedConnector3', 'bentConnector2', 'bentConnector3'])

/** A custom geometry (`a:custGeom`) as SVG path data in a w × h box. */
function customPath(custGeom: XmlNode, w: number, h: number): { d: string; open: boolean } | null {
  const parts: string[] = []
  let open = true
  for (const pathNode of children(child(custGeom, 'pathLst'), 'path')) {
    const pw = num(pathNode, 'w') || w
    const ph = num(pathNode, 'h') || h
    const sx = w / pw
    const sy = h / ph
    let cx = 0
    let cy = 0
    const point = (node: XmlNode | null): [number, number] => [(num(node, 'x') ?? 0) * sx, (num(node, 'y') ?? 0) * sy]
    for (const cmd of pathNode.children) {
      const name = local(cmd.name)
      const pts = children(cmd, 'pt').map(point)
      if (name === 'moveTo' && pts[0]) {
        ;[cx, cy] = pts[0]
        parts.push(`M${round(cx, 2)},${round(cy, 2)}`)
      } else if (name === 'lnTo' && pts[0]) {
        ;[cx, cy] = pts[0]
        parts.push(`L${round(cx, 2)},${round(cy, 2)}`)
      } else if (name === 'cubicBezTo' && pts.length === 3) {
        parts.push(`C${pts.map(([x, y]) => `${round(x, 2)},${round(y, 2)}`).join(' ')}`)
        ;[cx, cy] = pts[2]
      } else if (name === 'quadBezTo' && pts.length === 2) {
        parts.push(`Q${pts.map(([x, y]) => `${round(x, 2)},${round(y, 2)}`).join(' ')}`)
        ;[cx, cy] = pts[1]
      } else if (name === 'arcTo') {
        const wr = (num(cmd, 'wR') ?? 0) * sx
        const hr = (num(cmd, 'hR') ?? 0) * sy
        const start = ((num(cmd, 'stAng') ?? 0) / 60000) * (Math.PI / 180)
        const sweep = ((num(cmd, 'swAng') ?? 0) / 60000) * (Math.PI / 180)
        const ox = cx - wr * Math.cos(start)
        const oy = cy - hr * Math.sin(start)
        const ex = ox + wr * Math.cos(start + sweep)
        const ey = oy + hr * Math.sin(start + sweep)
        parts.push(`A${round(wr, 2)},${round(hr, 2)} 0 ${Math.abs(sweep) > Math.PI ? 1 : 0} ${sweep > 0 ? 1 : 0} ${round(ex, 2)},${round(ey, 2)}`)
        ;[cx, cy] = [ex, ey]
      } else if (name === 'close') {
        parts.push('Z')
        open = false
      }
    }
  }
  return parts.length ? { d: parts.join(' '), open } : null
}

// ---------------------------------------------------------------------------------------------
// Text

const WINGDINGS: Record<string, string> = {
  'l': '●', 'n': '■', 'q': '❑', 'u': '◆', 'v': '❖', 'Ø': '➢', 'ü': '✓', 'þ': '☑', '§': '■', 'ª': '◆', 'Ÿ': '•', 'o': '○', 'p': '□', 'à': '→', 'è': '➔'
}

const CJK_SANS = "'PingFang TC','Heiti TC','Noto Sans TC','Noto Sans CJK TC','Microsoft JhengHei'"
const FONT_FALLBACKS: [RegExp, string][] = [
  [/細明|明體|MingLiU|PMingLiU|Song|宋|Serif|Times|Georgia|Cambria|Mincho|明朝/i, "'PMingLiU','Songti TC','Noto Serif TC','Noto Serif CJK TC',serif"],
  [/楷|Kai/i, "'DFKai-SB','BiauKai','Kaiti TC',serif"],
  [/Consolas|Courier|Mono|Menlo/i, 'ui-monospace,Menlo,Consolas,monospace']
]

const safeFont = (name: string | undefined) => (name ?? '').replace(/[^\p{L}\p{N} ._-]/gu, '').trim()

function fontStack(latin: string | undefined, ea: string | undefined, theme: Theme): string {
  const resolve = (name: string | undefined) => {
    if (!name) return undefined
    if (name === '+mj-lt') return theme.major.latin
    if (name === '+mn-lt') return theme.minor.latin
    if (name === '+mj-ea') return theme.major.ea
    if (name === '+mn-ea') return theme.minor.ea
    return name.startsWith('+') ? undefined : name
  }
  const names = [...new Set([resolve(latin) ?? theme.minor.latin, resolve(ea) ?? theme.minor.ea].map(safeFont).filter(Boolean))]
  const fallback = FONT_FALLBACKS.find(([pattern]) => names.some((name) => pattern.test(name)))?.[1] ?? `${CJK_SANS},sans-serif`
  return [...names.map((name) => `'${name}'`), fallback].join(',')
}

/** The text styles a shape's paragraphs inherit, nearest first. */
type TextStyles = { lists: XmlNode[]; bodyPrs: XmlNode[] }

const node = (name: string, attrs: Record<string, string> = {}, children: XmlNode[] = []): XmlNode => ({ name, attrs, children, text: '' })

/**
 * A list style that only sets the text colour (and boldness) at every level: how a shape's `p:style`
 * fontRef and a table style's tcTxStyle take part in the inheritance chain, above the master's text styles.
 */
function colorListStyle(color: XmlNode | null, bold = false): XmlNode | null {
  if (!color && !bold) return null
  const defRPr = node('a:defRPr', bold ? { b: '1' } : {}, color ? [node('a:solidFill', {}, [color])] : [])
  return node('a:lstStyle', {}, Array.from({ length: 9 }, (_, i) => node(`a:lvl${i + 1}pPr`, {}, [defRPr])))
}

const colorElement = (holder: XmlNode | null): XmlNode | null => holder?.children.find((candidate) => COLOR_TAGS.has(local(candidate.name))) ?? null

const truthy = (value: string | undefined) => value === '1' || value === 'true' || value === 'on'

function firstAttr(nodes: XmlNode[], name: string): string | undefined {
  for (const node of nodes) {
    const value = attr(node, name)
    if (value !== undefined) return value
  }
  return undefined
}

function firstChild(nodes: XmlNode[], names: string[]): XmlNode | null {
  for (const node of nodes) {
    const found = node.children.find((candidate) => names.includes(local(candidate.name)))
    if (found) return found
  }
  return null
}

type Paragraph = { html: string; text: string }

type Counter = { level: number; type: string; value: number }

function renderTextBody(txBody: XmlNode, styles: TextStyles, ctx: Ctx, options: { bodyPr: XmlNode[] }): { html: string; text: string } {
  const bodyPrs = [child(txBody, 'bodyPr'), ...options.bodyPr].filter((n): n is XmlNode => n !== null)
  const autofit = firstChild(bodyPrs, ['normAutofit', 'spAutoFit', 'noAutofit'])
  const fontScale = autofit && local(autofit.name) === 'normAutofit' ? (num(autofit, 'fontScale') ?? 100000) / 100000 : 1
  const spacingCut = autofit && local(autofit.name) === 'normAutofit' ? (num(autofit, 'lnSpcReduction') ?? 0) / 100000 : 0
  const lists = [child(txBody, 'lstStyle'), ...styles.lists].filter((n): n is XmlNode => n !== null)
  const paragraphs = children(txBody, 'p')
  const out: Paragraph[] = []
  const counters: Counter[] = []
  paragraphs.forEach((p, index) => {
    const pPr = child(p, 'pPr')
    const level = Math.min(9, Math.max(1, (num(pPr, 'lvl') ?? 0) + 1))
    const levelNodes = lists.map((list) => child(list, `lvl${level}pPr`)).filter((n): n is XmlNode => n !== null)
    const pNodes = [pPr, ...levelNodes].filter((n): n is XmlNode => n !== null)
    const defRPrs = pNodes.map((node) => child(node, 'defRPr')).filter((n): n is XmlNode => n !== null)
    const runNodes = (rPr: XmlNode | null) => [rPr, ...defRPrs].filter((n): n is XmlNode => n !== null)
    const runs = p.children.filter((node) => ['r', 'br', 'fld'].includes(local(node.name)))
    const endRPr = child(p, 'endParaRPr')
    const sizeOf = (nodes: XmlNode[]) => ((Number(firstAttr(nodes, 'sz')) || 1800) / 100) * fontScale
    const firstRunRPr = runs.length ? child(runs[0], 'rPr') : endRPr
    const firstSize = sizeOf(runNodes(firstRunRPr))
    let html = ''
    let text = ''
    for (const run of runs) {
      const rPr = child(run, 'rPr')
      const nodes = runNodes(rPr)
      if (local(run.name) === 'br') {
        html += '<br>'
        text += '\n'
        continue
      }
      let value = local(run.name) === 'fld' && attr(run, 'type') === 'slidenum' ? String(ctx.slideNumber) : (child(run, 't')?.text ?? '')
      if (!value) continue
      text += value
      const style: string[] = []
      const size = sizeOf(nodes)
      style.push(`font-size:${pt(size)}`)
      if (truthy(firstAttr(nodes, 'b'))) style.push('font-weight:700')
      if (truthy(firstAttr(nodes, 'i'))) style.push('font-style:italic')
      const decorations: string[] = []
      const underline = firstAttr(nodes, 'u')
      if (underline && underline !== 'none') decorations.push('underline')
      const strike = firstAttr(nodes, 'strike')
      if (strike && strike !== 'noStrike') decorations.push('line-through')
      const fillNode = firstChild(nodes, ['solidFill', 'noFill', 'gradFill'])
      let color: Rgba | null = null
      if (fillNode && local(fillNode.name) === 'solidFill') color = colorIn(fillNode, ctx)
      else if (fillNode && local(fillNode.name) === 'gradFill') color = colorIn(child(child(fillNode, 'gsLst'), 'gs'), ctx)
      if (fillNode && local(fillNode.name) === 'noFill') style.push('color:transparent')
      const link = child(rPr, 'hlinkClick')
      let href: string | null = null
      if (link) {
        const action = attr(link, 'action') ?? ''
        const rel = ctx.pkg.rel(ctx.part, relAttr(link, 'id'))
        if (action.includes('hlinksldjump') && rel) {
          const target = ctx.slideOf(rel.target)
          if (target) href = `#${target}`
        } else if (action.includes('jump=firstslide')) href = '#1'
        if (!action || action.startsWith('ppaction://hlinkfile') === false) {
          if (!color) color = hexToRgba(ctx.theme.colors[ctx.clrMap.hlink ?? 'hlink'] ?? '') ?? null
          if (!decorations.includes('underline')) decorations.push('underline')
        }
      }
      if (color) style.push(`color:${cssColor(color)}`)
      if (decorations.length) style.push(`text-decoration:${decorations.join(' ')}`)
      const latin = attr(firstChild(nodes, ['latin']), 'typeface')
      const ea = attr(firstChild(nodes, ['ea']), 'typeface')
      style.push(`font-family:${fontStack(latin, ea, ctx.theme)}`)
      const baseline = Number(firstAttr(nodes, 'baseline') ?? 0)
      if (baseline) style.push(`vertical-align:${round(baseline / 1000)}%`, `font-size:${pt(size * 0.66)}`)
      const cap = firstAttr(nodes, 'cap')
      if (cap === 'all') style.push('text-transform:uppercase')
      else if (cap === 'small') style.push('font-variant:small-caps')
      const spc = Number(firstAttr(nodes, 'spc') ?? 0)
      if (spc) style.push(`letter-spacing:${pt(spc / 100)}`)
      const highlight = colorIn(firstChild(nodes, ['highlight']), ctx)
      if (highlight) style.push(`background:${cssColor(highlight)}`)
      const tag = href ? 'a' : 'span'
      const hrefAttr = href ? ` href="${escapeHtml(href)}"` : ''
      const title = link && !href ? ctx.pkg.rel(ctx.part, relAttr(link, 'id')) : undefined
      const titleAttr = title?.external ? ` title="${escapeHtml(title.target)}"` : ''
      html += `<${tag}${hrefAttr}${titleAttr} style="${style.join(';')}">${escapeHtml(value)}</${tag}>`
    }

    // Paragraph box
    const pStyle: string[] = []
    const algn = firstAttr(pNodes, 'algn') ?? 'l'
    pStyle.push(`text-align:${{ l: 'left', ctr: 'center', r: 'right', just: 'justify', dist: 'justify', thaiDist: 'justify', justLow: 'justify' }[algn] ?? 'left'}`)
    const marL = Number(firstAttr(pNodes, 'marL') ?? 0) / EMU_PER_PT
    const indent = Number(firstAttr(pNodes, 'indent') ?? 0) / EMU_PER_PT
    if (marL) pStyle.push(`padding-left:${pt(marL)}`)
    const marR = Number(firstAttr(pNodes, 'marR') ?? 0) / EMU_PER_PT
    if (marR) pStyle.push(`padding-right:${pt(marR)}`)
    if (indent) pStyle.push(`text-indent:${pt(indent)}`)
    pStyle.push(`font-size:${pt(firstSize)}`)
    const lnSpc = firstChild(pNodes, ['lnSpc'])
    const lnPct = child(lnSpc, 'spcPct')
    const lnPts = child(lnSpc, 'spcPts')
    if (lnPts) pStyle.push(`line-height:${pt(((num(lnPts, 'val') ?? 0) / 100) * (1 - spacingCut))}`)
    else pStyle.push(`line-height:${round(1.2 * ((lnPct ? (num(lnPct, 'val') ?? 100000) : 100000) / 100000 - spacingCut), 3)}`)
    const space = (name: string) => {
      const node = firstChild(pNodes, [name])
      const pts = child(node, 'spcPts')
      if (pts) return pt((num(pts, 'val') ?? 0) / 100)
      const pct = child(node, 'spcPct')
      if (pct) return `${round(((num(pct, 'val') ?? 0) / 100000) * 1.2, 3)}em`
      return null
    }
    const before = index > 0 ? space('spcBef') : null
    const after = space('spcAft')
    if (before) pStyle.push(`margin-top:${before}`)
    if (after) pStyle.push(`margin-bottom:${after}`)

    // Bullet
    let bullet = ''
    const kind = firstChild(pNodes, ['buNone', 'buChar', 'buAutoNum', 'buBlip'])
    const hasText = text.trim().length > 0
    while (counters.length && counters[counters.length - 1].level > level) counters.pop()
    if (kind && hasText && local(kind.name) !== 'buNone') {
      const buStyle: string[] = [`min-width:${pt(Math.max(0, -indent))}`]
      const buFont = firstChild(pNodes, ['buFont', 'buFontTx'])
      const buClr = firstChild(pNodes, ['buClr', 'buClrTx'])
      const buSz = firstChild(pNodes, ['buSzPct', 'buSzPts', 'buSzTx'])
      const colorValue = buClr && local(buClr.name) === 'buClr' ? colorIn(buClr, ctx) : null
      if (colorValue) buStyle.push(`color:${cssColor(colorValue)}`)
      else {
        const firstColor = /color:(rgba?\([^)]*\))/.exec(html)?.[1]
        if (firstColor) buStyle.push(`color:${firstColor}`)
      }
      if (buSz && local(buSz.name) === 'buSzPct') buStyle.push(`font-size:${pt(firstSize * ((num(buSz, 'val') ?? 100000) / 100000))}`)
      else if (buSz && local(buSz.name) === 'buSzPts') buStyle.push(`font-size:${pt((num(buSz, 'val') ?? 1800) / 100)}`)
      else buStyle.push(`font-size:${pt(firstSize)}`)
      let symbol = ''
      if (local(kind.name) === 'buChar') {
        const char = attr(kind, 'char') ?? '•'
        const face = buFont && local(buFont.name) === 'buFont' ? attr(buFont, 'typeface') ?? '' : ''
        symbol = /wingdings|symbol/i.test(face) ? (WINGDINGS[char] ?? '•') : char
      } else if (local(kind.name) === 'buAutoNum') {
        const type = attr(kind, 'type') ?? 'arabicPeriod'
        const top = counters[counters.length - 1]
        let counter: Counter
        if (top && top.level === level && top.type === type) counter = top
        else {
          if (top && top.level === level) counters.pop()
          counter = { level, type, value: (num(kind, 'startAt') ?? 1) - 1 }
          counters.push(counter)
        }
        counter.value++
        symbol = autoNumber(type, counter.value)
      } else symbol = '•'
      bullet = `<span class="pp-bu" style="${buStyle.join(';')}">${escapeHtml(symbol)}</span>`
    } else if (hasText) {
      // A paragraph without a numbered bullet ends the numbering at its level.
      if (counters.length && counters[counters.length - 1].level === level && !(kind && local(kind.name) === 'buAutoNum')) counters.pop()
    }
    out.push({ html: `<p class="pp-p" style="${pStyle.join(';')}">${bullet}${html || '<br>'}</p>`, text })
  })

  // Body box: insets, anchor, wrapping, columns
  const insets = ['lIns', 'tIns', 'rIns', 'bIns'].map((name, at) => Number(firstAttr(bodyPrs, name) ?? (at % 2 === 0 ? 91440 : 45720)) / EMU_PER_PT)
  const anchor = firstAttr(bodyPrs, 'anchor') ?? 't'
  const style = [
    `padding:${pt(insets[1])} ${pt(insets[2])} ${pt(insets[3])} ${pt(insets[0])}`,
    `justify-content:${anchor === 'ctr' ? 'center' : anchor === 'b' ? 'flex-end' : 'flex-start'}`
  ]
  if (firstAttr(bodyPrs, 'wrap') === 'none') style.push('white-space:pre')
  const vert = firstAttr(bodyPrs, 'vert')
  if (vert && vert !== 'horz') style.push('writing-mode:vertical-rl')
  if (vert === 'vert270') style.push('transform:rotate(180deg)')
  const columns = Number(firstAttr(bodyPrs, 'numCol') ?? 1)
  const body = out.map((p) => p.html).join('')
  const inner = columns > 1 ? `<div style="column-count:${Math.min(16, columns)};column-gap:${emuPt(Number(firstAttr(bodyPrs, 'spcCol') ?? 0))}">${body}</div>` : body
  return { html: `<div class="pp-tx" style="${style.join(';')}">${inner}</div>`, text: out.map((p) => p.text).join('\n') }
}

function autoNumber(type: string, n: number): string {
  const roman = (value: number) => {
    const table: [number, string][] = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']]
    let out = ''
    for (const [size, letters] of table) while (value >= size) { out += letters; value -= size }
    return out
  }
  const alpha = (value: number) => {
    let out = ''
    while (value > 0) { out = String.fromCharCode(97 + ((value - 1) % 26)) + out; value = Math.floor((value - 1) / 26) }
    return out
  }
  const CJK = '〇一二三四五六七八九十'
  const cjk = (value: number) => (value <= 10 ? CJK[value] : value < 20 ? `十${CJK[value - 10]}` : String(value))
  const circled = (value: number) => (value >= 1 && value <= 20 ? String.fromCodePoint(0x2460 + value - 1) : String(value))
  const base = type.startsWith('alphaLc') ? alpha(n) : type.startsWith('alphaUc') ? alpha(n).toUpperCase() : type.startsWith('romanLc') ? roman(n)
    : type.startsWith('romanUc') ? roman(n).toUpperCase() : type.startsWith('circleNum') ? circled(n) : /ea1Cht|ea1Chs|ea1Jpn|taiwan|chinese/i.test(type) ? cjk(n) : String(n)
  if (type.endsWith('ParenBoth')) return `(${base})`
  if (type.endsWith('ParenR')) return `${base})`
  if (type.endsWith('Period')) return `${base}.`
  if (type.endsWith('Comma')) return `${base}、`
  return base
}

// ---------------------------------------------------------------------------------------------
// Placeholders and inheritance

type Placeholder = { type: string; idx?: string }

function placeholderOf(node: XmlNode): Placeholder | null {
  const nvPr = node.children.find((candidate) => /^nv\w+Pr$/.test(local(candidate.name)))
  const ph = path(nvPr, 'nvPr', 'ph')
  if (!ph) return null
  return { type: attr(ph, 'type') ?? 'obj', idx: attr(ph, 'idx') }
}

const normalizeType = (type: string) => (type === 'ctrTitle' ? 'title' : type === 'subTitle' || type === 'obj' ? 'body' : type)

function findPlaceholder(tree: XmlNode | null, ph: Placeholder, byIndex: boolean): XmlNode | null {
  if (!tree) return null
  const candidates: XmlNode[] = []
  const walk = (node: XmlNode) => {
    for (const item of node.children) {
      if (local(item.name) === 'grpSp') walk(item)
      else if (placeholderOf(item)) candidates.push(item)
    }
  }
  walk(tree)
  if (byIndex && ph.idx !== undefined) {
    const match = candidates.find((candidate) => placeholderOf(candidate)?.idx === ph.idx)
    if (match) return match
  }
  const type = normalizeType(ph.type)
  return candidates.find((candidate) => placeholderOf(candidate)!.type === ph.type) ??
    candidates.find((candidate) => normalizeType(placeholderOf(candidate)!.type) === type) ?? null
}

type Inherited = { nodes: XmlNode[]; lists: XmlNode[]; bodyPrs: XmlNode[]; master: XmlNode | null }

function inherited(ph: Placeholder | null, ctx: Ctx, level: 'slide' | 'layout' | 'master'): Inherited {
  if (!ph) return { nodes: [], lists: [], bodyPrs: [], master: null }
  const layoutTree = path(ctx.layout.xml, 'cSld', 'spTree')
  const masterTree = path(ctx.master.xml, 'cSld', 'spTree')
  const fromLayout = level === 'slide' ? findPlaceholder(layoutTree, ph, true) : null
  const fromMaster = level !== 'master' ? findPlaceholder(masterTree, fromLayout ? placeholderOf(fromLayout)! : ph, false) : null
  const nodes = [fromLayout, fromMaster].filter((n): n is XmlNode => n !== null)
  return {
    nodes,
    lists: nodes.map((node) => path(node, 'txBody', 'lstStyle')).filter((n): n is XmlNode => n !== null),
    bodyPrs: nodes.map((node) => path(node, 'txBody', 'bodyPr')).filter((n): n is XmlNode => n !== null),
    master: fromMaster
  }
}

function masterTextStyle(ph: Placeholder | null, ctx: Ctx): XmlNode | null {
  const styles = path(ctx.master.xml, 'txStyles')
  if (!ph) return child(styles, 'otherStyle')
  const type = normalizeType(ph.type)
  if (type === 'title') return child(styles, 'titleStyle')
  if (type === 'body') return child(styles, 'bodyStyle')
  return child(styles, 'otherStyle')
}

// ---------------------------------------------------------------------------------------------
// Shapes

type Xfrm = { x: number; y: number; cx: number; cy: number; rot: number; flipH: boolean; flipV: boolean; chX: number; chY: number; chCx: number; chCy: number }

function xfrmOf(node: XmlNode | null): Xfrm | null {
  if (!node) return null
  const off = child(node, 'off')
  const ext = child(node, 'ext')
  if (!off || !ext) return null
  const chOff = child(node, 'chOff')
  const chExt = child(node, 'chExt')
  return {
    x: num(off, 'x') ?? 0, y: num(off, 'y') ?? 0, cx: num(ext, 'cx') ?? 0, cy: num(ext, 'cy') ?? 0,
    rot: (num(node, 'rot') ?? 0) / 60000, flipH: truthy(attr(node, 'flipH')), flipV: truthy(attr(node, 'flipV')),
    chX: num(chOff, 'x') ?? 0, chY: num(chOff, 'y') ?? 0, chCx: num(chExt, 'cx') ?? 0, chCy: num(chExt, 'cy') ?? 0
  }
}

const pct = (value: number, of: number) => `${round(of ? (value / of) * 100 : 0, 4)}%`

function boxStyle(x: Xfrm, frame: Frame): string[] {
  const style = [
    `left:${pct(x.x - frame.x, frame.cx)}`, `top:${pct(x.y - frame.y, frame.cy)}`,
    `width:${pct(x.cx, frame.cx)}`, `height:${pct(x.cy, frame.cy)}`
  ]
  if (x.rot) style.push(`transform:rotate(${round(x.rot)}deg)`)
  return style
}

type Rendered = { html: string; text: string[] }

function renderTree(tree: XmlNode, frame: Frame, ctx: Ctx, level: 'slide' | 'layout' | 'master', out: Rendered): void {
  for (const node of tree.children) {
    const name = local(node.name)
    if (name === 'AlternateContent') {
      const pick = child(node, 'Fallback') ?? child(node, 'Choice')
      if (pick) renderTree(pick, frame, ctx, level, out)
      continue
    }
    try {
      if (name === 'sp' || name === 'cxnSp') renderShape(node, frame, ctx, level, out)
      else if (name === 'pic') renderPicture(node, frame, ctx, level, out)
      else if (name === 'grpSp') renderGroup(node, frame, ctx, level, out)
      else if (name === 'graphicFrame') renderGraphicFrame(node, frame, ctx, out)
    } catch (error) {
      out.html += `<div class="pp-missing" style="left:0;top:0;width:100%;height:100%">${escapeHtml(error instanceof Error ? error.message : String(error))}</div>`
    }
  }
}

function renderGroup(node: XmlNode, frame: Frame, ctx: Ctx, level: 'slide' | 'layout' | 'master', out: Rendered): void {
  const x = xfrmOf(path(node, 'grpSpPr', 'xfrm'))
  if (!x) {
    renderTree(node, frame, ctx, level, out)
    return
  }
  const style = boxStyle(x, frame)
  const flips = [x.flipH ? 'scaleX(-1)' : '', x.flipV ? 'scaleY(-1)' : ''].filter(Boolean)
  if (flips.length) style.push(`transform:${x.rot ? `rotate(${round(x.rot)}deg) ` : ''}${flips.join(' ')}`)
  const inner: Rendered = { html: '', text: [] }
  renderTree(node, { x: x.chX, y: x.chY, cx: x.chCx || x.cx, cy: x.chCy || x.cy }, ctx, level, inner)
  out.html += `<div class="pp-grp" style="${style.join(';')}">${inner.html}</div>`
  out.text.push(...inner.text)
}

function renderShape(node: XmlNode, frame: Frame, ctx: Ctx, level: 'slide' | 'layout' | 'master', out: Rendered): void {
  const ph = placeholderOf(node)
  // Placeholders on a layout or master are prompts for slides, not content.
  if (ph && level !== 'slide') return
  const spPr = child(node, 'spPr')
  const style = child(node, 'style')
  const inherit = inherited(ph, ctx, level)
  const x = xfrmOf(child(spPr, 'xfrm')) ?? inherit.nodes.map((n) => xfrmOf(path(n, 'spPr', 'xfrm'))).find(Boolean) ?? null
  if (!x) return
  const txBody = child(node, 'txBody')
  const hasText = Boolean(txBody && children(txBody, 'p').some((p) => p.children.some((r) => ['r', 'fld'].includes(local(r.name)) && (child(r, 't')?.text ?? '') !== '')))
  const spPrs = [spPr, ...inherit.nodes.map((n) => child(n, 'spPr'))].filter((n): n is XmlNode => n !== null)
  const geomNode = firstChild(spPrs, ['prstGeom', 'custGeom'])
  const fillNode = firstChild(spPrs, [...FILL_TAGS])
  let background = fillNode ? fillCss(fillNode, ctx) : styleRefFill(style, ctx)
  const ln = firstChild(spPrs, ['ln'])
  const line = lineOf(ln, style, ctx)
  if (!background && !line && !hasText) return
  const w = x.cx / EMU_PER_PT
  const h = x.cy / EMU_PER_PT
  const box = boxStyle(x, frame)
  let geometry = ''
  const flip = [x.flipH ? 'scaleX(-1)' : '', x.flipV ? 'scaleY(-1)' : ''].filter(Boolean).join(' ')
  const shadow = shadowCss(spPr, style, ctx)
  const prst = geomNode && local(geomNode.name) === 'prstGeom' ? attr(geomNode, 'prst') ?? 'rect' : null
  const custom = geomNode && local(geomNode.name) === 'custGeom' ? customPath(geomNode, w, h) : null
  const svgPath = custom?.d ?? (prst ? presetPath(prst, w, h, adjOf(geomNode)) : null)
  if (background || line) {
    if (svgPath) {
      const open = custom ? custom.open : OPEN_PATHS.has(prst ?? '')
      const fill = open ? 'none' : svgFill(fillNode, background, ctx)
      const strokeAttrs = line
        ? ` stroke="${line.color}" stroke-width="${round(line.width)}"${dashArray(line.dash, line.width) ? ` stroke-dasharray="${dashArray(line.dash, line.width)}"` : ''} stroke-linejoin="round"`
        : ''
      const markers = line && open ? arrowMarkers(line, out) : ''
      const svgStyle = [flip ? `transform:${flip}` : '', shadow ? `filter:${shadow}` : ''].filter(Boolean).join(';')
      geometry = `<svg class="pp-geom" viewBox="0 0 ${round(Math.max(w, 0.01))} ${round(Math.max(h, 0.01))}" preserveAspectRatio="none"${svgStyle ? ` style="${svgStyle}"` : ''}>${markers.split('|')[0]}<path d="${svgPath}" fill="${fill}"${strokeAttrs}${markers.split('|')[1] ?? ''}/></svg>`
    } else {
      const geom: string[] = []
      if (background) geom.push(`background:${background}`)
      if (line) {
        const half = pt(line.width / 2)
        geom.push(`inset:calc(-1*${half})`, `border:${pt(line.width)} ${line.dash === 'solid' ? 'solid' : line.dash.toLowerCase().includes('dot') && !line.dash.toLowerCase().includes('dash') ? 'dotted' : 'dashed'} ${line.color}`)
      }
      if (prst === 'ellipse' || prst === 'flowChartConnector') geom.push('border-radius:50%')
      else if (prst === 'roundRect' || prst === 'flowChartAlternateProcess') geom.push(`border-radius:${pt(Math.min(w, h) * ((adjOf(geomNode).adj ?? 16667) / 100000))}`)
      if (flip) geom.push(`transform:${flip}`)
      if (shadow) geom.push(`filter:${shadow}`)
      geometry = `<div class="pp-geom" style="${geom.join(';')}"></div>`
    }
  }
  let text = ''
  if (txBody && hasText) {
    const fontRef = colorListStyle(colorElement(child(style, 'fontRef')))
    const lists = [...inherit.lists, fontRef, masterTextStyle(ph, ctx), ctx.presentationStyle].filter((n): n is XmlNode => n !== null)
    const rendered = renderTextBody(txBody, { lists, bodyPrs: inherit.bodyPrs }, ctx, { bodyPr: inherit.bodyPrs })
    text = rendered.html
    out.text.push(rendered.text)
    // Text in a shape PowerPoint flips stays readable: only a vertical flip turns it over.
    const txXfrm = xfrmOf(child(node, 'txXfrm'))
    if (txXfrm) {
      const inner = boxStyle({ ...txXfrm, x: txXfrm.x - x.x, y: txXfrm.y - x.y }, { x: 0, y: 0, cx: x.cx, cy: x.cy })
      text = `<div class="pp-grp" style="${inner.join(';')}">${text}</div>`
    }
  }
  out.html += `<div class="pp-sp" style="${box.join(';')}">${geometry}${text}</div>`
}

let markerCount = 0

/** SVG arrowheads for a line's ends: `<defs>…|marker attributes`. */
function arrowMarkers(line: NonNullable<Line>, _out: Rendered): string {
  if (!line.head && !line.tail) return '|'
  const id = `pp-m${markerCount++}`
  const marker = (name: string, type: string, orient: string) => {
    const shape = type === 'oval' ? '<circle cx="5" cy="5" r="4"/>' : type === 'diamond' ? '<path d="M5,0 L10,5 L5,10 L0,5 Z"/>' : '<path d="M0,0 L10,5 L0,10 Z"/>'
    return `<marker id="${id}${name}" viewBox="0 0 10 10" refX="${type === 'triangle' || type === 'arrow' || type === 'stealth' ? 9 : 5}" refY="5" markerWidth="4" markerHeight="4" orient="${orient}" fill="${line.color}">${shape}</marker>`
  }
  let defs = '<defs>'
  let attrs = ''
  if (line.head) {
    defs += marker('h', line.head, 'auto-start-reverse')
    attrs += ` marker-start="url(#${id}h)"`
  }
  if (line.tail) {
    defs += marker('t', line.tail, 'auto')
    attrs += ` marker-end="url(#${id}t)"`
  }
  return `${defs}</defs>|${attrs}`
}

function renderPicture(node: XmlNode, frame: Frame, ctx: Ctx, level: 'slide' | 'layout' | 'master', out: Rendered): void {
  const ph = placeholderOf(node)
  if (ph && level !== 'slide') return
  const spPr = child(node, 'spPr')
  const inherit = inherited(ph, ctx, level)
  const x = xfrmOf(child(spPr, 'xfrm')) ?? inherit.nodes.map((n) => xfrmOf(path(n, 'spPr', 'xfrm'))).find(Boolean) ?? null
  if (!x) return
  const blipFill = child(node, 'blipFill')
  const url = blipFill ? blipUrl(blipFill, ctx) : null
  const box = boxStyle(x, frame)
  const name = attr(path(node, 'nvPicPr', 'cNvPr'), 'descr') ?? attr(path(node, 'nvPicPr', 'cNvPr'), 'name') ?? ''
  const nvPr = path(node, 'nvPicPr', 'nvPr')
  const isMovie = Boolean(child(nvPr, 'videoFile') || child(nvPr, 'audioFile') || findDeep(nvPr ?? node, 'media'))
  if (!url) {
    out.html += `<div class="pp-sp" style="${box.join(';')}"><div class="pp-missing">${isMovie ? '影音' : '圖片'}(網頁版無法顯示這種格式,請下載 PPTX 查看)</div></div>`
    return
  }
  const crop = child(blipFill, 'srcRect')
  const [l, t, r, b] = ['l', 't', 'r', 'b'].map((side) => (num(crop, side) ?? 0) / 100000)
  const visibleW = 1 - l - r || 1
  const visibleH = 1 - t - b || 1
  const imgStyle = [
    `left:${round((-l / visibleW) * 100, 4)}%`, `top:${round((-t / visibleH) * 100, 4)}%`,
    `width:${round((1 / visibleW) * 100, 4)}%`, `height:${round((1 / visibleH) * 100, 4)}%`
  ]
  const flip = [x.flipH ? 'scaleX(-1)' : '', x.flipV ? 'scaleY(-1)' : ''].filter(Boolean).join(' ')
  const clip: string[] = []
  const prst = attr(child(spPr, 'prstGeom'), 'prst')
  if (prst === 'ellipse') clip.push('border-radius:50%')
  else if (prst === 'roundRect') clip.push(`border-radius:${pt((Math.min(x.cx, x.cy) / EMU_PER_PT) * ((adjOf(child(spPr, 'prstGeom')).adj ?? 16667) / 100000))}`)
  if (flip) clip.push(`transform:${flip}`)
  const alpha = num(findDeep(blipFill!, 'alphaModFix'), 'amt')
  if (alpha !== undefined) clip.push(`opacity:${round(alpha / 100000)}`)
  const shadow = shadowCss(spPr, child(node, 'style'), ctx)
  if (shadow) clip.push(`filter:${shadow}`)
  const line = lineOf(child(spPr, 'ln'), child(node, 'style'), ctx)
  const border = line ? `<div class="pp-geom" style="inset:calc(-1*${pt(line.width / 2)});border:${pt(line.width)} solid ${line.color}${clip[0]?.startsWith('border-radius') ? `;${clip[0]}` : ''}"></div>` : ''
  out.html += `<div class="pp-sp" style="${box.join(';')}"><div class="pp-pic" style="${clip.join(';')}"><img src="${escapeHtml(url)}" alt="${escapeHtml(name)}" style="${imgStyle.join(';')}" decoding="async" draggable="false"></div>${border}</div>`
}

// ---------------------------------------------------------------------------------------------
// Graphic frames: tables, charts, SmartArt, OLE

function renderGraphicFrame(node: XmlNode, frame: Frame, ctx: Ctx, out: Rendered): void {
  const x = xfrmOf(child(node, 'xfrm'))
  if (!x) return
  const data = path(node, 'graphic', 'graphicData')
  const uri = attr(data, 'uri') ?? ''
  const box = boxStyle(x, frame)
  if (uri.endsWith('/table')) {
    const table = child(data, 'tbl')
    if (table) {
      const rendered = renderTable(table, ctx)
      out.html += `<div class="pp-sp" style="${box.join(';')}">${rendered.html}</div>`
      out.text.push(rendered.text)
    }
    return
  }
  if (uri.endsWith('/diagram')) {
    const ids = child(data, 'relIds')
    const dataPart = ctx.pkg.rel(ctx.part, relAttr(ids, 'dm'))
    const dataXml = dataPart ? ctx.pkg.xml(dataPart.target) : null
    const drawingId = attr(dataXml ? findDeep(dataXml, 'dataModelExt') : null, 'relId')
    const drawing = ctx.pkg.rel(ctx.part, drawingId)
    const drawingXml = drawing ? ctx.pkg.xml(drawing.target) : null
    const tree = path(drawingXml, 'drawing', 'spTree') ?? (drawingXml ? findDeep(drawingXml, 'spTree') : null)
    if (tree && drawing) {
      const inner: Rendered = { html: '', text: [] }
      renderTree(tree, { x: 0, y: 0, cx: x.cx, cy: x.cy }, { ...ctx, part: drawing.target }, 'slide', inner)
      out.html += `<div class="pp-grp" style="${box.join(';')}">${inner.html}</div>`
      out.text.push(...inner.text)
      return
    }
    out.html += missing(box, 'SmartArt 圖形')
    return
  }
  if (uri.endsWith('/chart')) {
    const rel = ctx.pkg.rel(ctx.part, relAttr(child(data, 'chart'), 'id'))
    const chart = rel ? ctx.pkg.xml(rel.target) : null
    const rendered = chart ? renderChart(chart, x, { ...ctx, part: rel!.target }) : null
    if (rendered) {
      out.html += `<div class="pp-sp" style="${box.join(';')}">${rendered.html}</div>`
      out.text.push(rendered.text)
    } else out.html += missing(box, '圖表')
    return
  }
  // OLE objects and the like carry a picture of themselves.
  const picture = findDeep(data ?? node, 'pic')
  if (picture) {
    const pictureX = xfrmOf(path(picture, 'spPr', 'xfrm'))
    const placed: XmlNode = pictureX
      ? picture
      : { ...picture, children: [...picture.children.filter((c) => local(c.name) !== 'spPr'), { name: 'p:spPr', attrs: {}, text: '', children: [child(node, 'xfrm')!] }] }
    renderPicture(placed, frame, ctx, 'slide', out)
    return
  }
  out.html += missing(box, '內嵌物件')
}

const missing = (box: string[], what: string) =>
  `<div class="pp-sp" style="${box.join(';')}"><div class="pp-missing">${escapeHtml(what)}(網頁版未支援,請下載 PPTX 查看)</div></div>`

type CellStyle = { fill?: string; color?: XmlNode; bold?: boolean; borders: Partial<Record<'left' | 'right' | 'top' | 'bottom' | 'insideH' | 'insideV', string>> }

function tablePart(styleNode: XmlNode | null, ctx: Ctx): CellStyle | null {
  if (!styleNode) return null
  const tcStyle = child(styleNode, 'tcStyle')
  const tcTx = child(styleNode, 'tcTxStyle')
  const out: CellStyle = { borders: {} }
  const fill = child(tcStyle, 'fill')
  if (fill) out.fill = fillCss(firstFill(fill), ctx) || undefined
  const fillRef = child(tcStyle, 'fillRef')
  if (!fill && fillRef) out.fill = fillCss(ctx.theme.fills[(num(fillRef, 'idx') ?? 1) - 1] ?? null, { ...ctx, phColor: colorIn(fillRef, ctx) ?? undefined }) || undefined
  for (const side of ['left', 'right', 'top', 'bottom', 'insideH', 'insideV'] as const) {
    const ln = child(child(child(tcStyle, 'tcBdr'), side), 'ln')
    const lnRef = child(child(child(tcStyle, 'tcBdr'), side), 'lnRef')
    const line = ln ? lineOf(ln, null, ctx) : lnRef ? lineOf(null, { name: 'p:style', attrs: {}, text: '', children: [lnRef] }, ctx) : undefined
    if (line !== undefined) out.borders[side] = line ? `${pt(line.width)} solid ${line.color}` : 'none'
  }
  const color = colorElement(tcTx) ?? colorElement(child(tcTx, 'fontRef'))
  if (color) out.color = color
  if (attr(tcTx, 'b') === 'on') out.bold = true
  return out
}

function renderTable(table: XmlNode, ctx: Ctx): { html: string; text: string } {
  const tblPr = child(table, 'tblPr')
  const grid = children(child(table, 'tblGrid'), 'gridCol').map((col) => (num(col, 'w') ?? 0) / EMU_PER_PT)
  const rows = children(table, 'tr')
  const styleId = child(tblPr, 'tableStyleId')?.text.trim() ?? ''
  const styleDef = children(ctx.tableStyles, 'tblStyle').find((style) => attr(style, 'styleId') === styleId) ?? null
  const flag = (name: string) => truthy(attr(tblPr, name))
  const part = (name: string) => tablePart(child(styleDef, name), ctx)
  const whole = part('wholeTbl')
  const parts = {
    band1H: flag('bandRow') ? part('band1H') : null, band2H: flag('bandRow') ? part('band2H') : null,
    band1V: flag('bandCol') ? part('band1V') : null, band2V: flag('bandCol') ? part('band2V') : null,
    firstRow: flag('firstRow') ? part('firstRow') : null, lastRow: flag('lastRow') ? part('lastRow') : null,
    firstCol: flag('firstCol') ? part('firstCol') : null, lastCol: flag('lastCol') ? part('lastCol') : null
  }
  const texts: string[] = []
  let html = `<table class="pp-tbl"><colgroup>${grid.map((w) => `<col style="width:${pt(w)}">`).join('')}</colgroup>`
  rows.forEach((row, r) => {
    html += `<tr style="height:${emuPt(num(row, 'h') ?? 0)}">`
    let c = 0
    for (const cell of children(row, 'tc')) {
      const col = c
      c += Math.max(1, Number(attr(cell, 'gridSpan') ?? 1))
      if (truthy(attr(cell, 'hMerge')) || truthy(attr(cell, 'vMerge'))) continue
      const isFirstRow = r === 0 && parts.firstRow !== null
      const isLastRow = r === rows.length - 1 && parts.lastRow !== null
      const bodyRow = r - (parts.firstRow ? 1 : 0)
      const layers = [
        whole,
        !isFirstRow && !isLastRow ? (bodyRow % 2 === 0 ? parts.band1H : parts.band2H) : null,
        col % 2 === 0 ? parts.band1V : parts.band2V,
        col === 0 ? parts.firstCol : null,
        col === grid.length - 1 ? parts.lastCol : null,
        isLastRow ? parts.lastRow : null,
        isFirstRow ? parts.firstRow : null
      ].filter((layer): layer is CellStyle => layer !== null)
      const tcPr = child(cell, 'tcPr')
      const style: string[] = []
      let fill = ''
      let color: XmlNode | undefined
      let bold = false
      const borders: Record<string, string> = {}
      for (const layer of layers) {
        if (layer.fill) fill = layer.fill
        if (layer.color) color = layer.color
        if (layer.bold) bold = true
        const b = layer.borders
        if (b.left || b.insideV) borders.left = (col === 0 ? b.left : b.insideV) ?? borders.left
        if (b.right || b.insideV) borders.right = (col + 1 >= grid.length ? b.right : b.insideV) ?? borders.right
        if (b.top || b.insideH) borders.top = (r === 0 ? b.top : b.insideH) ?? borders.top
        if (b.bottom || b.insideH) borders.bottom = (r === rows.length - 1 ? b.bottom : b.insideH) ?? borders.bottom
      }
      const ownFill = firstFill(tcPr)
      if (ownFill) fill = fillCss(ownFill, ctx)
      for (const [side, name] of [['left', 'lnL'], ['right', 'lnR'], ['top', 'lnT'], ['bottom', 'lnB']] as const) {
        const ln = child(tcPr, name)
        if (ln) {
          const line = lineOf(ln, null, ctx)
          borders[side] = line ? `${pt(line.width)} solid ${line.color}` : 'none'
        }
      }
      if (fill) style.push(`background:${fill}`)
      for (const [side, value] of Object.entries(borders)) if (value) style.push(`border-${side}:${value}`)
      const margins = ['marT', 'marR', 'marB', 'marL'].map((name, at) => Number(attr(tcPr, name) ?? (at % 2 === 0 ? 45720 : 91440)) / EMU_PER_PT)
      style.push(`padding:${margins.map(pt).join(' ')}`)
      const anchor = attr(tcPr, 'anchor')
      style.push(`vertical-align:${anchor === 'ctr' ? 'middle' : anchor === 'b' ? 'bottom' : 'top'}`)
      const txBody = child(cell, 'txBody')
      let content = ''
      if (txBody) {
        const lists = [colorListStyle(color ?? null, bold), child(path(ctx.master.xml, 'txStyles'), 'otherStyle'), ctx.presentationStyle].filter((n): n is XmlNode => n !== null)
        const rendered = renderTextBody(
          { ...txBody, children: [{ name: 'a:bodyPr', attrs: { lIns: '0', tIns: '0', rIns: '0', bIns: '0' }, text: '', children: [] }, ...txBody.children.filter((n) => local(n.name) !== 'bodyPr')] },
          { lists, bodyPrs: [] }, ctx, { bodyPr: [] }
        )
        content = rendered.html
        texts.push(rendered.text)
      }
      const span = Number(attr(cell, 'gridSpan') ?? 1)
      const rowSpan = Number(attr(cell, 'rowSpan') ?? 1)
      html += `<td${span > 1 ? ` colspan="${span}"` : ''}${rowSpan > 1 ? ` rowspan="${rowSpan}"` : ''} style="${style.join(';')}">${content}</td>`
    }
    html += '</tr>'
  })
  return { html: `${html}</table>`, text: texts.join('\n') }
}

// ---------------------------------------------------------------------------------------------
// Charts: bar, line, area, pie and doughnut from the values cached in the chart part

type Series = { name: string; values: (number | null)[]; color: string }

const SERIES_DEFAULTS = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6']

function cachedValues(node: XmlNode | null): (string | null)[] {
  const cache = node ? findDeep(node, 'numCache') ?? findDeep(node, 'strCache') ?? findDeep(node, 'multiLvlStrCache') : null
  if (!cache) return []
  const count = num(child(cache, 'ptCount'), 'val') ?? 0
  const out: (string | null)[] = Array.from({ length: Math.min(count, 2000) }, () => null)
  // Multi-level categories: the first level holds the labels next to the axis.
  const points = local(cache.name) === 'multiLvlStrCache' ? children(child(cache, 'lvl'), 'pt') : children(cache, 'pt')
  for (const point of points) {
    const at = num(point, 'idx') ?? -1
    if (at >= 0 && at < out.length) out[at] = child(point, 'v')?.text ?? null
  }
  return out
}

function chartTitle(chart: XmlNode): string {
  const title = path(chart, 'chart', 'title')
  if (!title) return ''
  const runs: string[] = []
  const walk = (node: XmlNode) => {
    for (const item of node.children) {
      if (local(item.name) === 't') runs.push(item.text)
      else walk(item)
    }
  }
  walk(title)
  return runs.join('')
}

function renderChart(chart: XmlNode, x: Xfrm, ctx: Ctx): { html: string; text: string } | null {
  const plot = path(chart, 'chart', 'plotArea')
  if (!plot) return null
  const kinds = ['barChart', 'bar3DChart', 'lineChart', 'line3DChart', 'areaChart', 'pieChart', 'pie3DChart', 'doughnutChart']
  const group = plot.children.find((node) => kinds.includes(local(node.name)))
  if (!group) return null
  const kind = local(group.name).replace('3D', '')
  let categories: string[] = []
  const series: Series[] = children(group, 'ser').map((ser, at) => {
    const name = cachedValues(child(ser, 'tx'))[0] ?? `數列 ${at + 1}`
    if (!categories.length) categories = cachedValues(child(ser, 'cat')).map((v) => v ?? '')
    const values = cachedValues(child(ser, 'val')).map((v) => (v === null || v === '' || !Number.isFinite(Number(v)) ? null : Number(v)))
    const fill = colorIn(path(ser, 'spPr', 'solidFill'), ctx) ?? colorIn(path(ser, 'spPr', 'ln', 'solidFill'), ctx)
    const fallback = hexToRgba(ctx.theme.colors[SERIES_DEFAULTS[at % 6]] ?? '4472C4')!
    return { name, values, color: cssColor(fill ?? fallback) }
  })
  if (!series.length) return null
  const length = Math.max(categories.length, ...series.map((s) => s.values.length))
  categories = Array.from({ length }, (_, at) => categories[at] ?? String(at + 1))
  const title = chartTitle(chart)
  const W = x.cx / EMU_PER_PT
  const H = x.cy / EMU_PER_PT
  const textColor = cssColor(hexToRgba(ctx.theme.colors[ctx.clrMap.tx1 ?? 'dk1'] ?? '404040') ?? { r: 64, g: 64, b: 64, a: 1 })
  const font = fontStack(undefined, undefined, ctx.theme)
  const fs = 10
  let svg = ''
  const top = title ? 24 : 8
  if (title) svg += `<text x="${round(W / 2)}" y="16" text-anchor="middle" font-size="13" font-weight="700" fill="${textColor}">${escapeHtml(title)}</text>`
  const legendH = series.length > 1 || kind === 'pieChart' || kind === 'doughnutChart' ? 18 : 0
  if (kind === 'pieChart' || kind === 'doughnutChart') {
    const values = series[0].values.map((v) => Math.max(0, v ?? 0))
    const total = values.reduce((a, b) => a + b, 0) || 1
    const r = Math.max(4, Math.min(W, H - top - legendH) / 2 - 6)
    const cx = W / 2
    const cy = top + (H - top - legendH) / 2
    let angle = -Math.PI / 2
    const points = series[0]
    const colors = values.map((_, at) => {
      const dPt = children(child(group, 'ser'), 'dPt').find((d) => num(child(d, 'idx'), 'val') === at)
      const c = colorIn(path(dPt, 'spPr', 'solidFill'), ctx) ?? hexToRgba(ctx.theme.colors[SERIES_DEFAULTS[at % 6]] ?? '4472C4')
      return c ? cssColor(c) : points.color
    })
    values.forEach((value, at) => {
      const sweep = (value / total) * Math.PI * 2
      const end = angle + sweep
      const large = sweep > Math.PI ? 1 : 0
      const [x1, y1, x2, y2] = [cx + r * Math.cos(angle), cy + r * Math.sin(angle), cx + r * Math.cos(end), cy + r * Math.sin(end)]
      svg += sweep >= Math.PI * 2 - 1e-6
        ? `<circle cx="${round(cx)}" cy="${round(cy)}" r="${round(r)}" fill="${colors[at]}"/>`
        : `<path d="M${round(cx)},${round(cy)} L${round(x1)},${round(y1)} A${round(r)},${round(r)} 0 ${large} 1 ${round(x2)},${round(y2)} Z" fill="${colors[at]}" stroke="#fff" stroke-width="1"/>`
      angle = end
    })
    if (kind === 'doughnutChart') svg += `<circle cx="${round(cx)}" cy="${round(cy)}" r="${round(r * 0.5)}" fill="#fff"/>`
    svg += legend(categories.map((name, at) => ({ name, color: colors[at] })), W, H - 6, fs, textColor)
  } else {
    const all = series.flatMap((s) => s.values).filter((v): v is number => v !== null)
    const stacked = ['stacked', 'percentStacked'].includes(attr(child(group, 'grouping'), 'val') ?? '')
    const sums = categories.map((_, at) => series.reduce((sum, s) => sum + Math.max(0, s.values[at] ?? 0), 0))
    let max = Math.max(0, ...(stacked ? sums : all))
    const min = Math.min(0, ...all)
    if (max === min) max = min + 1
    const step = niceStep((max - min) / 5)
    const top2 = Math.ceil(max / step) * step
    const bottom = Math.floor(min / step) * step
    // Bar charts may run sideways (barDir="bar"): the same drawing with the axes swapped, first category at the bottom.
    const horizontal = kind === 'barChart' && attr(child(group, 'barDir'), 'val') === 'bar'
    const left = horizontal ? Math.min(W * 0.3, 70) : 36
    const plotTop = top + 4
    const plotBottom = H - legendH - 20
    const plotH = Math.max(1, plotBottom - plotTop)
    const plotW = Math.max(1, W - left - 8)
    const catLength = horizontal ? plotH : plotW
    const valLength = horizontal ? plotW : plotH
    const dist = (value: number) => ((value - bottom) / (top2 - bottom || 1)) * valLength
    const point = (c: number, d: number): [number, number] => (horizontal ? [left + d, plotBottom - c] : [left + c, plotBottom - d])
    const n = Math.max(1, categories.length || series[0].values.length)
    const slot = catLength / n
    for (let v = bottom; v <= top2 + step / 2; v += step) {
      const [x1, y1] = point(0, dist(v))
      const [x2, y2] = point(catLength, dist(v))
      svg += `<line x1="${round(x1)}" x2="${round(x2)}" y1="${round(y1)}" y2="${round(y2)}" stroke="#d9d9d9" stroke-width="0.5"/>`
      svg += horizontal
        ? `<text x="${round(x1)}" y="${round(plotBottom + 13)}" text-anchor="middle" font-size="${fs}" fill="${textColor}">${escapeHtml(formatNumber(v))}</text>`
        : `<text x="${left - 4}" y="${round(y1 + 3)}" text-anchor="end" font-size="${fs}" fill="${textColor}">${escapeHtml(formatNumber(v))}</text>`
    }
    categories.forEach((name, at) => {
      const [x, y] = point(slot * (at + 0.5), 0)
      svg += horizontal
        ? `<text x="${round(left - 4)}" y="${round(y + 3)}" text-anchor="end" font-size="${fs}" fill="${textColor}">${escapeHtml(name)}</text>`
        : `<text x="${round(x)}" y="${round(plotBottom + 13)}" text-anchor="middle" font-size="${fs}" fill="${textColor}">${escapeHtml(name)}</text>`
    })
    if (kind === 'barChart') {
      const width = stacked ? slot * 0.6 : (slot * 0.7) / series.length
      const base: number[] = categories.map(() => 0)
      series.forEach((s, si) => {
        s.values.forEach((value, at) => {
          if (value === null) return
          const c0 = slot * at + slot * (stacked ? 0.2 : 0.15) + (stacked ? 0 : width * si)
          const from = stacked ? base[at] : 0
          const to = from + value
          if (stacked) base[at] = to
          const [ax, ay] = point(c0, dist(from))
          const [bx, by] = point(c0 + width, dist(to))
          svg += `<rect x="${round(Math.min(ax, bx))}" y="${round(Math.min(ay, by))}" width="${round(Math.abs(bx - ax))}" height="${round(Math.abs(by - ay))}" fill="${s.color}"/>`
        })
      })
    } else {
      series.forEach((s) => {
        const pts = s.values.map((value, at) => (value === null ? null : point(slot * (at + 0.5), dist(value)).map((v) => round(v)).join(','))).filter(Boolean)
        const [sx, sy] = point(slot * 0.5, dist(0))
        const [ex] = point(slot * (s.values.length - 0.5), dist(0))
        if (kind === 'areaChart') svg += `<path d="M${round(sx)},${round(sy)} L${pts.join(' L')} L${round(ex)},${round(sy)} Z" fill="${s.color}" opacity="0.7"/>`
        else svg += `<polyline points="${pts.join(' ')}" fill="none" stroke="${s.color}" stroke-width="2"/>`
      })
    }
    if (legendH) svg += legend(series.map((s) => ({ name: s.name, color: s.color })), W, H - 6, fs, textColor)
  }
  const text = [title, ...series.map((s) => s.name), ...categories].filter(Boolean).join(' ')
  return {
    html: `<svg class="pp-geom" viewBox="0 0 ${round(W)} ${round(H)}" font-family="${font.replace(/"/g, '')}" role="img" aria-label="${escapeHtml(title || '圖表')}">${svg}</svg>`,
    text
  }
}

function legend(items: { name: string; color: string }[], W: number, y: number, fs: number, color: string): string {
  const widths = items.map((item) => 14 + [...item.name].reduce((sum, ch) => sum + (/[⺀-￯]/.test(ch) ? fs : fs * 0.55), 0) + 10)
  let x = (W - widths.reduce((a, b) => a + b, 0)) / 2
  return items
    .map((item, at) => {
      const out = `<rect x="${round(x)}" y="${round(y - fs + 2)}" width="8" height="8" fill="${item.color}"/><text x="${round(x + 11)}" y="${round(y)}" font-size="${fs}" fill="${color}">${escapeHtml(item.name)}</text>`
      x += widths[at]
      return out
    })
    .join('')
}

function niceStep(raw: number): number {
  if (!(raw > 0)) return 1
  const power = 10 ** Math.floor(Math.log10(raw))
  const unit = raw / power
  return (unit <= 1 ? 1 : unit <= 2 ? 2 : unit <= 5 ? 5 : 10) * power
}

const formatNumber = (value: number) => (Math.abs(value) >= 1000 ? value.toLocaleString('en-US') : String(round(value, 2)))

// ---------------------------------------------------------------------------------------------
// Backgrounds and notes

function backgroundOf(xml: XmlNode | null, ctx: Ctx): string | null {
  const bg = path(xml, 'cSld', 'bg')
  if (!bg) return null
  const bgPr = child(bg, 'bgPr')
  if (bgPr) return fillCss(firstFill(bgPr), ctx) || 'transparent'
  const ref = child(bg, 'bgRef')
  if (ref) {
    const idx = num(ref, 'idx') ?? 0
    const refCtx = { ...ctx, phColor: colorIn(ref, ctx) ?? undefined }
    const fill = idx >= 1001 ? ctx.theme.bgFills[idx - 1001] : idx > 0 ? ctx.theme.fills[idx - 1] : undefined
    const css = fill ? fillCss(fill, refCtx) : ''
    return css || (refCtx.phColor ? cssColor(refCtx.phColor) : null)
  }
  return null
}

function notesHtml(pkg: Package, slidePart: string): { html: string; text: string } {
  const rel = pkg.relOfType(slidePart, 'notesSlide')
  const notes = rel ? pkg.xml(rel.target) : null
  const tree = path(notes, 'cSld', 'spTree')
  if (!tree) return { html: '', text: '' }
  const bodies = children(tree, 'sp').filter((sp) => placeholderOf(sp)?.type === 'body')
  const paragraphs: string[] = []
  const texts: string[] = []
  for (const sp of bodies) {
    for (const p of children(child(sp, 'txBody'), 'p')) {
      let html = ''
      let text = ''
      for (const run of p.children) {
        const name = local(run.name)
        if (name === 'br') {
          html += '<br>'
          text += '\n'
        } else if (name === 'r' || name === 'fld') {
          const value = child(run, 't')?.text ?? ''
          const rPr = child(run, 'rPr')
          let piece = escapeHtml(value)
          if (truthy(attr(rPr, 'b'))) piece = `<strong>${piece}</strong>`
          if (truthy(attr(rPr, 'i'))) piece = `<em>${piece}</em>`
          html += piece
          text += value
        }
      }
      if (text.trim()) {
        paragraphs.push(`<p>${html}</p>`)
        texts.push(text)
      }
    }
  }
  return { html: paragraphs.join(''), text: texts.join('\n') }
}

// ---------------------------------------------------------------------------------------------
// The deck

function themeOf(pkg: Package, masterPart: string): Theme {
  const rel = pkg.relOfType(masterPart, 'theme')
  const xml = rel ? pkg.xml(rel.target) : null
  const elements = path(xml, 'themeElements')
  const colors: Record<string, string> = {}
  for (const node of child(elements, 'clrScheme')?.children ?? []) {
    const value = node.children[0]
    if (!value) continue
    const name = local(value.name)
    colors[local(node.name)] = name === 'srgbClr' ? attr(value, 'val') ?? '' : name === 'sysClr' ? attr(value, 'lastClr') ?? '' : ''
  }
  const fonts = child(elements, 'fontScheme')
  const face = (node: XmlNode | null) => ({ latin: attr(child(node, 'latin'), 'typeface') || undefined, ea: attr(child(node, 'ea'), 'typeface') || undefined })
  const styles = child(elements, 'fmtScheme')
  return {
    colors,
    major: face(child(fonts, 'majorFont')),
    minor: face(child(fonts, 'minorFont')),
    fills: child(styles, 'fillStyleLst')?.children ?? [],
    bgFills: child(styles, 'bgFillStyleLst')?.children ?? [],
    lines: child(styles, 'lnStyleLst')?.children ?? []
  }
}

const DEFAULT_CLR_MAP: Record<string, string> = { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' }

const clrMapOf = (node: XmlNode | null): Record<string, string> | null => (node ? { ...DEFAULT_CLR_MAP, ...node.attrs } : null)

export type PptxSlide = { html: string; text: string; notes: string }
export type PptxDeck = { width: number; height: number; slides: PptxSlide[]; title: string }

/** Reads a deck: its slide size and every shown slide as HTML. `mediaBase` prefixes picture URLs. */
export function readPptx(data: Buffer, mediaBase: string): PptxDeck {
  const pkg = new Package(data)
  const presentationPart = pkg.relOfType('', 'officeDocument')?.target ?? 'ppt/presentation.xml'
  const presentation = pkg.xml(presentationPart)
  if (!presentation) throw new Error('這不是 PowerPoint 檔(找不到 presentation.xml)')
  const size = child(presentation, 'sldSz')
  const width = num(size, 'cx') ?? 12192000
  const height = num(size, 'cy') ?? 6858000
  const presentationStyle = child(presentation, 'defaultTextStyle')
  const tableStylesRel = pkg.relOfType(presentationPart, 'tableStyles')
  const tableStyles = tableStylesRel ? pkg.xml(tableStylesRel.target) : null
  const slideParts = children(child(presentation, 'sldIdLst'), 'sldId')
    .map((sld) => pkg.rel(presentationPart, relAttr(sld, 'id')))
    .filter((rel): rel is Rel => Boolean(rel && !rel.external))
    .map((rel) => rel.target)
  const shown = slideParts.filter((part) => attr(pkg.xml(part), 'show') !== '0')
  const numberOf = new Map(shown.map((part, at) => [part, at + 1]))
  const masters = new Map<string, Master>()
  const masterOf = (part: string): Master => {
    let master = masters.get(part)
    if (!master) {
      const xml = pkg.xml(part)
      if (!xml) throw new Error(`缺少母片 ${part}`)
      master = { part, xml, clrMap: clrMapOf(child(xml, 'clrMap')) ?? DEFAULT_CLR_MAP, theme: themeOf(pkg, part) }
      masters.set(part, master)
    }
    return master
  }
  const media = (part: string) => officeMediaUrl(mediaBase, part)
  const docTitle = pkg.xml('docProps/core.xml')
  const title = docTitle ? (findDeep(docTitle, 'title')?.text.trim() ?? '') : ''

  // One slide that cannot be read shows why on its own face; the rest of the deck still draws.
  const slides = shown.map((part, at): PptxSlide => {
    try {
      return drawSlide(part, at)
    } catch (error) {
      const number = at + 1
      const message = `第 ${number} 頁無法顯示:${error instanceof Error ? error.message : String(error)}`
      return {
        html: `<section class="slide layout-pptx" id="slide-${number}" data-number="${number}" aria-label="第 ${number} 頁"><div class="pp-canvas"><div class="pp-missing">${escapeHtml(message)}</div></div></section>\n`,
        text: '',
        notes: ''
      }
    }
  })
  return { width, height, slides, title }

  function drawSlide(part: string, at: number): PptxSlide {
    const xml = pkg.xml(part)
    if (!xml) throw new Error(`缺少 ${part}`)
    const layoutPart = pkg.relOfType(part, 'slideLayout')?.target
    const layoutXml = layoutPart ? pkg.xml(layoutPart) : null
    const masterPart = layoutPart ? pkg.relOfType(layoutPart, 'slideMaster')?.target : undefined
    if (!layoutPart || !layoutXml || !masterPart) throw new Error(`第 ${at + 1} 頁缺少版面配置`)
    const master = masterOf(masterPart)
    const layout: Layout = { part: layoutPart, xml: layoutXml, master }
    const override = (node: XmlNode | null) => clrMapOf(child(child(node, 'clrMapOvr'), 'overrideClrMapping'))
    const base: Omit<Ctx, 'part' | 'clrMap'> = {
      pkg, theme: master.theme, slideNumber: at + 1, slideOf: (target) => numberOf.get(target), media, master, layout, presentationStyle, tableStyles
    }
    const masterCtx: Ctx = { ...base, part: master.part, clrMap: master.clrMap }
    const layoutCtx: Ctx = { ...base, part: layoutPart, clrMap: override(layoutXml) ?? master.clrMap }
    const slideCtx: Ctx = { ...base, part, clrMap: override(xml) ?? layoutCtx.clrMap }
    const frame: Frame = { x: 0, y: 0, cx: width, cy: height }
    const out: Rendered = { html: '', text: [] }
    const background = backgroundOf(xml, slideCtx) ?? backgroundOf(layoutXml, layoutCtx) ?? backgroundOf(master.xml, masterCtx) ?? '#fff'
    const showMaster = attr(xml, 'showMasterSp') !== '0'
    if (showMaster && attr(layoutXml, 'showMasterSp') !== '0') {
      const tree = path(master.xml, 'cSld', 'spTree')
      const masterOut: Rendered = { html: '', text: [] }
      if (tree) renderTree(tree, frame, masterCtx, 'master', masterOut)
      out.html += masterOut.html
    }
    if (showMaster) {
      const tree = path(layoutXml, 'cSld', 'spTree')
      const layoutOut: Rendered = { html: '', text: [] }
      if (tree) renderTree(tree, frame, layoutCtx, 'layout', layoutOut)
      out.html += layoutOut.html
    }
    const tree = path(xml, 'cSld', 'spTree')
    if (tree) renderTree(tree, frame, slideCtx, 'slide', out)
    const notes = notesHtml(pkg, part)
    const number = at + 1
    const html = `<section class="slide layout-pptx" id="slide-${number}" data-number="${number}" aria-label="第 ${number} 頁" style="background:${background}"><div class="pp-canvas">${out.html}</div>${notes.html ? `<aside class="slide-notes" hidden>${notes.html}</aside>` : ''}</section>\n`
    return { html, text: out.text.filter(Boolean).join('\n'), notes: notes.text }
  }
}

/** The searchable text of a deck: every shown slide and its notes. */
export function pptxText(data: Buffer): string {
  const deck = readPptx(data, '')
  return [deck.title, ...deck.slides.flatMap((slide) => [slide.text, slide.notes])].filter(Boolean).join('\n')
}

/**
 * The deck page for a PowerPoint file. Pictures are addressed relative to the page through the
 * Office media route, which resolves inside the version's own path. A file that cannot be read
 * still gets a page: it says why, and the throw lands in the viewer's error log.
 */
export function renderPptxDeck(data: Buffer, fileName: string, title: string, assets: DeckAssets): string {
  const mediaBase = officeMediaBase(fileName)
  let deck: PptxDeck
  try {
    deck = readPptx(data, mediaBase)
  } catch (error) {
    const message = `PowerPoint: ${error instanceof Error ? error.message : String(error)}`
    const slide = `<section class="slide layout-pptx" id="slide-1" data-number="1" aria-label="第 1 頁"><div class="pp-canvas"><div class="pp-missing">${escapeHtml(message)}</div></div></section>`
    return deckPage(slide, title, assets, failureScript(message))
  }
  const widthPt = round(deck.width / EMU_PER_PT, 3)
  const heightPt = round(deck.height / EMU_PER_PT, 3)
  const head = `<style>:root{--aspect:${round(deck.width / deck.height, 6)};--slide-pt:${widthPt}}@page{size:${widthPt}pt ${heightPt}pt}@media print{.slide,body.overview .slide{width:${widthPt}pt;height:${heightPt}pt}}</style>`
  const html = deck.slides.length
    ? deck.slides.map((slide) => slide.html).join('')
    : '<section class="slide layout-pptx" id="slide-1" data-number="1" aria-label="第 1 頁"><div class="pp-canvas"><div class="pp-missing">這份簡報沒有可顯示的投影片</div></div></section>'
  return deckPage(html, title, assets, head)
}

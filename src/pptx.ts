import { posix } from 'node:path'
import { escapeHtml } from './content-render.ts'
import { chartTitle, renderChart } from './office-chart.ts'
import { adjustValues, customGeometry, fillShade, presetGeometry, type Geometry } from './office-geometry.ts'
import { chineseCounting, circled } from './office-numbers.ts'
import { PATTERN_PT, PATTERNS, patternSvg, patternTile } from './office-patterns.ts'
import { failureScript, isDrawablePart, officeMediaBase, officeMediaUrl } from './office.ts'
import { deckPage, type DeckAssets } from './slides.ts'
import { attr, child, children, local, num, parseXml, path, relAttr, XmlError, type XmlNode } from './xml.ts'
import { ZipError, ZipReader } from './zip.ts'

/**
 * A PowerPoint file published as is (`*.pptx`), drawn on the server as slides the deck runtime pages
 * through: every shape is absolutely placed in slide-relative units, text stays text. Pictures load
 * from the file itself through the Office media route (office.ts), inside the artifact's CSP scope. What it does not draw (charts it cannot read, OLE, video) shows as a
 * labelled box; the original file is always one download away.
 */
const EMU_PER_PT = 12700
/** Office's default theme accents, for a file without a theme of its own. */
const OFFICE_ACCENTS = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47']

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

  /** A part the deck can draw without (theme, table styles, notes, properties): null when it cannot be read. */
  optional(name: string): XmlNode | null {
    try {
      return this.xml(name)
    } catch {
      return null
    }
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

/** A theme's heading or body fonts; `scripts` holds the font for each script (`Hant`, `Jpan`…) when `ea` is left empty. */
type ThemeFonts = { latin?: string; ea?: string; scripts: Record<string, string> }

type Theme = {
  colors: Record<string, string>
  major: ThemeFonts
  minor: ThemeFonts
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

/** An sRGB channel (0–255) in linear light (0–1), and back. */
const toLinear = (channel: number) => {
  const c = clamp01(channel / 255)
  return c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4
}
const fromLinear = (linear: number) => {
  const c = clamp01(linear)
  return 255 * (c <= 0.0031308 ? c * 12.92 : 1.055 * c ** (1 / 2.4) - 0.055)
}

function applyModifiers(color: Rgba, element: XmlNode): Rgba {
  let current = { ...color }
  for (const mod of element.children) {
    const value = (num(mod, 'val') ?? 0) / 100000
    const name = local(mod.name)
    if (name === 'alpha') current.a = clamp01(value)
    else if (name === 'alphaMod') current.a = clamp01(current.a * value)
    else if (name === 'alphaOff') current.a = clamp01(current.a + value)
    else if (name === 'tint' || name === 'shade') {
      // PowerPoint mixes towards white (tint) or black (shade) in linear light, not in sRGB: 20 % of 4472C4 is E9EBF5.
      const mix = (channel: number) => {
        const linear = toLinear(channel)
        return fromLinear(name === 'tint' ? 1 - (1 - linear) * clamp01(value) : linear * clamp01(value))
      }
      current = { ...current, r: mix(current.r), g: mix(current.g), b: mix(current.b) }
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
  /** How much HTML the slide may still add; shared by every copy of the context. */
  budget: Budget
  /** Inside a SmartArt drawing: another SmartArt there is a box, never a second drawing (they can lead back to themselves). */
  inDiagram?: boolean
  /** The flips of the groups around a shape, which its text undoes so it stays readable. */
  flipH?: boolean
  flipV?: boolean
  /** How many groups the shape sits in. */
  depth?: number
}

type Budget = { left: number; noticed: boolean }

/** HTML characters one slide, and the whole deck, may produce: a small crafted file must not grow into a page that stalls the server. */
const SLIDE_BUDGET = 4_000_000
const DECK_BUDGET = 24_000_000

/** The labelled box a slide shows once where it stopped drawing. */
function budgetNotice(ctx: Ctx): string {
  if (ctx.budget.noticed) return ''
  ctx.budget.noticed = true
  return '<div class="pp-missing" style="top:auto;height:auto">這頁內容太多,其餘部分請下載 PPTX 查看</div>'
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

/**
 * CSS `background` for a fill element; '' for none. Tile sizes are in --pt, which only the slide canvas
 * defines; a background outside it (the slide's own) passes its `box` in points to get percentages.
 */
function fillCss(fill: XmlNode | null, ctx: Ctx, box?: { w: number; h: number }): string {
  if (!fill) return ''
  const name = local(fill.name)
  if (name === 'solidFill') {
    const color = colorIn(fill, ctx)
    return color ? cssColor(color) : ''
  }
  if (name === 'gradFill') {
    const stops = gradientStops(fill, ctx)
    if (stops.length === 0) return ''
    if (stops.length === 1) return cssColor(stops[0].color)
    const list = stops.map((stop) => `${cssColor(stop.color)} ${round(stop.pos)}%`).join(',')
    if (child(fill, 'path')) return `radial-gradient(ellipse at center,${list})`
    const angle = (num(child(fill, 'lin'), 'ang') ?? 0) / 60000
    return `linear-gradient(${round(angle + 90)}deg,${list})`
  }
  if (name === 'blipFill') {
    const picture = pictureFill(fill, ctx)
    if (!picture) return ''
    if (picture.tile) {
      const { x, y, w, h } = picture.tile
      return `url('${picture.url}') ${tileGeometry(x, y, w, h, box)} repeat`
    }
    // The picture's crop and its place in the shape, as CSS background size and position percentages.
    const { x, y, w, h } = picture.stretch
    const position = (offset: number, size: number) => `${round(Math.abs(1 - size) < 1e-6 ? 0 : (offset / (1 - size)) * 100, 4)}%`
    return `url('${picture.url}') ${position(x, w)} ${position(y, h)}/${round(w * 100, 4)}% ${round(h * 100, 4)}% no-repeat`
  }
  if (name === 'pattFill') {
    const pattern = patternOf(fill, ctx)
    return pattern ? `url('data:image/svg+xml,${encodeURIComponent(pattern.svg)}') ${tileGeometry(0, 0, PATTERN_PT, PATTERN_PT, box)} repeat` : ''
  }
  return ''
}

/** CSS background position and size of a tile at x, y, w × h points: in --pt, or as percentages of `box`. */
function tileGeometry(x: number, y: number, w: number, h: number, box?: { w: number; h: number }): string {
  if (!box) return `${pt(x)} ${pt(y)}/${pt(w)} ${pt(h)}`
  // A percentage position places that fraction of the tile on that fraction of the box.
  const position = (offset: number, size: number, of: number) => (offset && of !== size ? `${round((offset / (of - size)) * 100, 4)}%` : '0%')
  return `${position(x, w, box.w)} ${position(y, h, box.h)}/${pct(w, box.w)} ${pct(h, box.h)}`
}

type Stop = { pos: number; color: Rgba }

const gradientStops = (fill: XmlNode, ctx: Ctx): Stop[] =>
  children(child(fill, 'gsLst'), 'gs')
    .map((gs) => ({ pos: clamp01((num(gs, 'pos') ?? 0) / 100000) * 100, color: colorIn(gs, ctx) }))
    .filter((stop): stop is Stop => stop.color !== null)
    .sort((a, b) => a.pos - b.pos)

// ---------------------------------------------------------------------------------------------
// Pattern fills: Office's 8 × 8 hatches (src/office-patterns.ts).

/** A pattern fill as an 8 × 8 SVG (background, then the foreground pixels) and its two colours. */
function patternOf(fill: XmlNode, ctx: Ctx): { svg: string; inner: string } | null {
  const fg = colorIn(child(fill, 'fgClr'), ctx) ?? { r: 0, g: 0, b: 0, a: 1 }
  const bg = colorIn(child(fill, 'bgClr'), ctx) ?? { r: 255, g: 255, b: 255, a: 1 }
  const bits = PATTERNS[attr(fill, 'prst') ?? ''] ?? PATTERNS.pct5
  return { svg: patternSvg(cssColor(fg), cssColor(bg), bits), inner: patternTile(cssColor(fg), cssColor(bg), bits) }
}

// ---------------------------------------------------------------------------------------------
// Picture fills

/** A picture fill: either stretched over a box (x, y, w, h as fractions of the shape) or tiled (in points). */
type PictureFill = { url: string; stretch: { x: number; y: number; w: number; h: number }; tile?: { x: number; y: number; w: number; h: number }; opacity?: number }

function pictureFill(fill: XmlNode, ctx: Ctx): PictureFill | null {
  const url = blipUrl(fill, ctx)
  if (!url) return null
  const inset = (node: XmlNode | null) => ['l', 't', 'r', 'b'].map((side) => (num(node, side) ?? 0) / 100000)
  // srcRect crops the picture; stretch/fillRect places what is left inside the shape.
  const [cl, ct, cr, cb] = inset(child(fill, 'srcRect'))
  const [fl, ft, fr, fb] = inset(path(fill, 'stretch', 'fillRect'))
  const placedW = 1 - fl - fr || 1
  const placedH = 1 - ft - fb || 1
  const w = placedW / (1 - cl - cr || 1)
  const h = placedH / (1 - ct - cb || 1)
  const out: PictureFill = { url, stretch: { x: fl - cl * w, y: ft - ct * h, w, h } }
  const tile = child(fill, 'tile')
  if (tile) {
    const size = pictureSize(fill, ctx) ?? { w: 96, h: 96 }
    out.tile = {
      x: (num(tile, 'tx') ?? 0) / EMU_PER_PT, y: (num(tile, 'ty') ?? 0) / EMU_PER_PT,
      w: Math.max(1, size.w * ((num(tile, 'sx') ?? 100000) / 100000)), h: Math.max(1, size.h * ((num(tile, 'sy') ?? 100000) / 100000))
    }
  }
  const alpha = num(findDeep(fill, 'alphaModFix') ?? fill, 'amt')
  if (alpha !== undefined) out.opacity = clamp01(alpha / 100000)
  return out
}

/** A picture's size in points, from the header of a PNG, JPEG or GIF (at its own resolution, else 96 dpi). */
function pictureSize(fill: XmlNode, ctx: Ctx): { w: number; h: number } | null {
  const rel = ctx.pkg.rel(ctx.part, relAttr(child(fill, 'blip'), 'embed'))
  let data: Buffer | null = null
  try {
    data = rel && !rel.external ? ctx.pkg.zip.read(rel.target) : null
  } catch {
    return null
  }
  if (!data || data.length < 24) return null
  const toPt = (px: number, dpi: number) => (px * 72) / (dpi > 10 && dpi < 3000 ? dpi : 96)
  if (data.readUInt32BE(0) === 0x89504e47) {
    let dpiX = 96
    let dpiY = 96
    const phys = data.indexOf('pHYs')
    if (phys > 0 && phys + 13 <= data.length && data[phys + 12] === 1) {
      dpiX = data.readUInt32BE(phys + 4) * 0.0254
      dpiY = data.readUInt32BE(phys + 8) * 0.0254
    }
    return { w: toPt(data.readUInt32BE(16), dpiX), h: toPt(data.readUInt32BE(20), dpiY) }
  }
  if (data.toString('latin1', 0, 3) === 'GIF') return { w: toPt(data.readUInt16LE(6), 96), h: toPt(data.readUInt16LE(8), 96) }
  if (data[0] === 0xff && data[1] === 0xd8) {
    let dpiX = 96
    let dpiY = 96
    for (let at = 2; at + 9 < data.length; ) {
      if (data[at] !== 0xff) return null
      const marker = data[at + 1]
      const length = data.readUInt16BE(at + 2)
      if (marker === 0xe0 && data.toString('latin1', at + 4, at + 8) === 'JFIF' && data[at + 11] === 1) {
        dpiX = data.readUInt16BE(at + 12)
        dpiY = data.readUInt16BE(at + 14)
      }
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
        return { w: toPt(data.readUInt16BE(at + 7), dpiX), h: toPt(data.readUInt16BE(at + 5), dpiY) }
      }
      at += 2 + length
    }
  }
  return null
}

/** Why a picture cannot be shown: its part is missing, it lives at an outside address, or browsers cannot draw its format. */
type NoPicture = 'missing' | 'external' | 'format'

/** The URL of the picture a `blip` names, or why there is none. */
function blipSource(blipFill: XmlNode, ctx: Ctx): { url: string } | { none: NoPicture; target?: string } {
  const blip = child(blipFill, 'blip')
  // An SVG picture carries its PNG fallback in r:embed and the vector in an extension.
  const svg = blip ? findDeep(blip, 'svgBlip') : null
  let none: NoPicture = 'missing'
  let target: string | undefined
  for (const id of [relAttr(svg, 'embed'), relAttr(blip, 'embed')]) {
    const rel = ctx.pkg.rel(ctx.part, id)
    if (!rel) continue
    if (rel.external) none = 'external'
    else if (!ctx.pkg.zip.has(rel.target)) continue
    else if (isDrawablePart(rel.target)) return { url: ctx.media(rel.target) }
    else (none = 'format'), (target = rel.target)
  }
  // A linked picture (r:link) is fetched from its address by PowerPoint; the page never reaches outside.
  if (none === 'missing' && (relAttr(blip, 'link') || relAttr(svg, 'link'))) none = 'external'
  return { none, target }
}

/** The URL of a picture part a `blip` names, or null when browsers cannot show it. */
function blipUrl(blipFill: XmlNode, ctx: Ctx): string | null {
  const source = blipSource(blipFill, ctx)
  return 'url' in source ? source.url : null
}

/** The label for a picture that cannot be shown, worded as Word and Excel pages word it. */
function noPictureLabel(none: NoPicture, target?: string): string {
  if (none === 'missing') return '圖片遺失(檔案裡找不到這張圖片)'
  if (none === 'external') return '外部連結的圖片(圖片在文件之外,網頁版不載入外部連結)'
  const ext = posix.extname(target ?? '').slice(1).toUpperCase().replace(/[^A-Z0-9]/g, '').slice(0, 8)
  return `${ext ? `${ext} 圖片` : '圖片'}(網頁版未支援,請下載原始檔查看)`
}

/** What a reader is told when something in the file cannot be drawn: the reader's message as is when it is already Chinese, else its kind. */
function readerMessage(error: unknown): string {
  const message = error instanceof Error ? error.message : String(error)
  if (/[一-鿿]/.test(message)) return message
  if (error instanceof ZipError) return '檔案的壓縮資料無法讀取,可能已損毀'
  if (error instanceof XmlError) return '內容的 XML 無法解析,檔案可能已損毀'
  if (error instanceof RangeError) return '內容的層級太深,無法顯示'
  return '讀取時發生錯誤'
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

/** The theme fill a shape's style refers to (fillRef), with the context that colours its phClr. */
function styleRefFill(style: XmlNode | null, ctx: Ctx): { fill: XmlNode; ctx: Ctx } | null {
  const ref = child(style, 'fillRef')
  const idx = num(ref, 'idx') ?? 0
  if (!ref || idx === 0) return null
  const list = idx >= 1000 ? ctx.theme.bgFills : ctx.theme.fills
  const fill = list[(idx >= 1000 ? idx - 1000 : idx) - 1]
  return fill ? { fill, ctx: { ...ctx, phColor: colorIn(ref, ctx) ?? undefined } } : null
}

type Line = { width: number; color: string; dash: string; compound: string; head?: string; tail?: string } | null

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
  const compound = sources.map(([node]) => attr(node, 'cmpd')).find(Boolean) ?? 'sng'
  return { width, color: cssColor(color), dash, compound, head: end('headEnd'), tail: end('tailEnd') }
}

/** A CSS border for a line: a double or triple line as `double` (CSS has nothing finer), dashes as dashed or dotted. */
function borderCss(line: NonNullable<Line>): string {
  const dash = line.dash.toLowerCase()
  const style = line.compound !== 'sng' && line.width >= 1.5 ? 'double' : dash === 'solid' ? 'solid' : dash.includes('dot') && !dash.includes('dash') ? 'dotted' : 'dashed'
  return `${pt(line.width)} ${style} ${line.color}`
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

/**
 * The ECMA-376 geometry of a shape's `prstGeom` or `custGeom` in a box of cx × cy EMU (paths and text rectangle in
 * points). Null for a preset name the standard does not define; a custom shape with no path draws nothing.
 */
function geometryOf(node: XmlNode, cx: number, cy: number): Geometry | null {
  if (local(node.name) === 'custGeom') return customGeometry(node, cx, cy) ?? { paths: [], textRect: { l: 0, t: 0, r: cx / EMU_PER_PT, b: cy / EMU_PER_PT } }
  return presetGeometry(attr(node, 'prst') ?? 'rect', cx, cy, adjustValues(node))
}

let paintCount = 0

/**
 * How an SVG path is painted with a fill: a colour, or a gradient, picture or pattern defined beside
 * it. `w` × `h` is the path's box in points, which is also its user space.
 */
function svgPaint(fillNode: XmlNode | null, ctx: Ctx, w: number, h: number): { fill: string; defs: string } {
  const none = { fill: 'none', defs: '' }
  if (!fillNode) return none
  const name = local(fillNode.name)
  const id = `pp-f${paintCount++}`
  if (name === 'solidFill') {
    const color = colorIn(fillNode, ctx)
    return color ? { fill: cssColor(color), defs: '' } : none
  }
  if (name === 'gradFill') {
    const stops = gradientStops(fillNode, ctx)
    if (stops.length === 0) return none
    if (stops.length === 1) return { fill: cssColor(stops[0].color), defs: '' }
    const list = stops.map((stop) => `<stop offset="${round(stop.pos)}%" stop-color="${cssColor(stop.color)}"/>`).join('')
    const radial = child(fillNode, 'path')
    if (radial) {
      const to = child(radial, 'fillToRect')
      const [l, t, r, b] = ['l', 't', 'r', 'b'].map((side) => (num(to, side) ?? 50000) / 100000)
      const cx = round((l + 1 - r) / 2, 4)
      const cy = round((t + 1 - b) / 2, 4)
      // To the farthest corner, as PowerPoint's path gradients reach.
      const radius = round(Math.hypot(Math.max(cx, 1 - cx), Math.max(cy, 1 - cy)), 4)
      return { fill: `url(#${id})`, defs: `<radialGradient id="${id}" cx="${cx}" cy="${cy}" r="${radius}">${list}</radialGradient>` }
    }
    const angle = (((num(child(fillNode, 'lin'), 'ang') ?? 0) / 60000) * Math.PI) / 180
    const [dx, dy] = [Math.cos(angle) / 2, Math.sin(angle) / 2]
    const at = (v: number) => round(v, 4)
    return { fill: `url(#${id})`, defs: `<linearGradient id="${id}" x1="${at(0.5 - dx)}" y1="${at(0.5 - dy)}" x2="${at(0.5 + dx)}" y2="${at(0.5 + dy)}">${list}</linearGradient>` }
  }
  if (name === 'blipFill') {
    const picture = pictureFill(fillNode, ctx)
    if (!picture) return none
    const opacity = picture.opacity !== undefined ? ` opacity="${round(picture.opacity)}"` : ''
    const href = escapeHtml(picture.url)
    if (picture.tile) {
      const { x, y, w: tw, h: th } = picture.tile
      return {
        fill: `url(#${id})`,
        defs: `<pattern id="${id}" patternUnits="userSpaceOnUse" x="${round(x)}" y="${round(y)}" width="${round(tw)}" height="${round(th)}"><image href="${href}" width="${round(tw)}" height="${round(th)}" preserveAspectRatio="none"${opacity}/></pattern>`
      }
    }
    const { x, y, w: sw, h: sh } = picture.stretch
    return {
      fill: `url(#${id})`,
      defs: `<pattern id="${id}" patternUnits="userSpaceOnUse" width="${round(w)}" height="${round(h)}"><image href="${href}" x="${round(x * w)}" y="${round(y * h)}" width="${round(sw * w)}" height="${round(sh * h)}" preserveAspectRatio="none"${opacity}/></pattern>`
    }
  }
  if (name === 'pattFill') {
    const pattern = patternOf(fillNode, ctx)
    if (!pattern) return none
    const scale = PATTERN_PT / 8
    return { fill: `url(#${id})`, defs: `<pattern id="${id}" patternUnits="userSpaceOnUse" width="${PATTERN_PT}" height="${PATTERN_PT}"><g transform="scale(${scale})" shape-rendering="crispEdges">${pattern.inner}</g></pattern>` }
  }
  return none
}

// ---------------------------------------------------------------------------------------------
// Text

const WINGDINGS: Record<string, string> = {
  'l': '●', 'n': '■', 'q': '❑', 'u': '◆', 'v': '❖', 'Ø': '➢', 'ü': '✓', 'þ': '☑', '§': '■', 'ª': '◆', 'Ÿ': '•', 'o': '○', 'p': '□', 'à': '→', 'è': '➔'
}

const safeFont = (name: string | undefined) => (name ?? '').replace(/[^\p{L}\p{N} ._-]/gu, '').trim()

type FontClass = 'sans' | 'serif' | 'kai' | 'mono'

const latinClass = (name: string): FontClass =>
  /Consolas|Courier|Mono|Menlo/i.test(name) ? 'mono' : /Times|Georgia|Cambria|Garamond|Palatino|Antiqua|Baskerville|Bodoni|Didot|Century|Serif|Minion/i.test(name) ? 'serif' : 'sans'
const eastAsianClass = (name: string): FontClass =>
  /楷|Kai/i.test(name) ? 'kai' : /細明|明體|明体|明朝|Ming|Song|宋|Sun\b|SimSun|Mincho|Batang|Gungsuh|바탕|궁서|Myungjo|Serif/i.test(name) ? 'serif' : 'sans'

/** Fonts with the same metrics as Office's, where a reader's device has them (Linux, LibreOffice installs). */
const METRIC_TWINS: Record<string, string> = {
  calibri: "'Carlito'", cambria: "'Caladea'", arial: "'Liberation Sans','Arimo'", 'times new roman': "'Liberation Serif','Tinos'", 'courier new': "'Liberation Mono','Cousine'"
}
const LATIN_FALLBACKS: Record<FontClass, string> = {
  sans: "'Helvetica Neue',Helvetica,Arial", serif: "'Times New Roman',Times", kai: "'Times New Roman',Times", mono: "Menlo,Consolas,'Courier New'"
}
/** Fonts on Apple, Android and Windows devices for each East Asian script, by the kind of letter. */
const EAST_ASIAN_FALLBACKS: Record<string, Record<FontClass, string>> = {
  Hant: {
    sans: "'PingFang TC','Heiti TC','Noto Sans TC','Noto Sans CJK TC','Microsoft JhengHei'",
    serif: "'PMingLiU','Songti TC','Noto Serif TC','Noto Serif CJK TC'",
    kai: "'DFKai-SB','BiauKai','Kaiti TC'",
    mono: "'PingFang TC','Heiti TC','Noto Sans TC','Noto Sans CJK TC','Microsoft JhengHei'"
  },
  Hans: {
    sans: "'PingFang SC','Heiti SC','Noto Sans SC','Noto Sans CJK SC','Microsoft YaHei'",
    serif: "'SimSun','Songti SC','Noto Serif SC','Noto Serif CJK SC'",
    kai: "'KaiTi','Kaiti SC'",
    mono: "'PingFang SC','Heiti SC','Noto Sans SC','Noto Sans CJK SC','Microsoft YaHei'"
  },
  Jpan: {
    sans: "'Hiragino Sans','Hiragino Kaku Gothic ProN','Noto Sans JP','Noto Sans CJK JP','Yu Gothic','Meiryo'",
    serif: "'Hiragino Mincho ProN','Noto Serif JP','Noto Serif CJK JP','Yu Mincho','MS Mincho'",
    kai: "'Hiragino Mincho ProN','Noto Serif JP','Noto Serif CJK JP','Yu Mincho','MS Mincho'",
    mono: "'Hiragino Sans','Hiragino Kaku Gothic ProN','Noto Sans JP','Noto Sans CJK JP','Yu Gothic','Meiryo'"
  },
  Hang: {
    sans: "'Apple SD Gothic Neo','Noto Sans KR','Noto Sans CJK KR','Malgun Gothic'",
    serif: "'AppleMyungjo','Noto Serif KR','Noto Serif CJK KR','Batang'",
    kai: "'AppleMyungjo','Noto Serif KR','Noto Serif CJK KR','Batang'",
    mono: "'Apple SD Gothic Neo','Noto Sans KR','Noto Sans CJK KR','Malgun Gothic'"
  }
}

/** The East Asian script of a language tag (zh-TW → Hant), or undefined for other languages. */
function eastAsianScript(lang: string | undefined): string | undefined {
  const tag = (lang ?? '').toLowerCase()
  if (/^zh-(tw|hk|mo|hant)/.test(tag)) return 'Hant'
  if (/^zh/.test(tag)) return 'Hans'
  if (/^ja/.test(tag)) return 'Jpan'
  if (/^ko/.test(tag)) return 'Hang'
  return undefined
}

/**
 * The CSS font stack for a run: its Latin font and the Latin fonts like it, then its East Asian font
 * and the ones like it for its script, so each kind of letter falls back within its own family. An
 * East Asian font left empty in the theme comes from the theme's font for the run's language
 * (`lang`, else `altLang`), as PowerPoint picks it; a run in neither is read as Traditional Chinese.
 */
function fontStack(latin: string | undefined, ea: string | undefined, theme: Theme, lang?: string, altLang?: string): string {
  const script = eastAsianScript(lang) ?? eastAsianScript(altLang) ?? 'Hant'
  const eastAsian = (set: ThemeFonts) => set.ea || set.scripts[script]
  const resolve = (name: string | undefined) => {
    if (!name) return undefined
    if (name === '+mj-lt') return theme.major.latin
    if (name === '+mn-lt') return theme.minor.latin
    if (name === '+mj-ea') return eastAsian(theme.major)
    if (name === '+mn-ea') return eastAsian(theme.minor)
    return name.startsWith('+') ? undefined : name
  }
  const latinName = safeFont(resolve(latin) ?? theme.minor.latin)
  const eaName = safeFont(resolve(ea) ?? eastAsian(theme.minor))
  const eaKind = eaName ? eastAsianClass(eaName) : latinName ? (latinClass(latinName) === 'serif' ? 'serif' : 'sans') : 'sans'
  const stack: string[] = []
  if (latinName) stack.push(`'${latinName}'`, METRIC_TWINS[latinName.toLowerCase()] ?? '', LATIN_FALLBACKS[latinClass(latinName)])
  if (eaName) stack.push(`'${eaName}'`)
  stack.push(EAST_ASIAN_FALLBACKS[script][eaKind], eaKind === 'sans' || eaKind === 'mono' ? 'sans-serif' : 'serif')
  const seen = new Set<string>()
  return stack
    .flatMap((item) => item.split(','))
    .filter((item) => item && !seen.has(item.toLowerCase()) && seen.add(item.toLowerCase()))
    .join(',')
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
  let used = 0
  let stopped = false
  for (const [index, p] of paragraphs.entries()) {
    if (stopped || used > ctx.budget.left) {
      stopped = true
      break
    }
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
      // Outlined text (a:ln on the run): the outline is drawn whether or not the letters are filled.
      const outline = firstChild(nodes, ['ln'])
      const stroke = outline ? lineOf(outline, null, ctx) : null
      if (stroke) style.push(`-webkit-text-stroke:${pt(stroke.width)} ${stroke.color}`)
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
      style.push(`font-family:${fontStack(latin, ea, ctx.theme, firstAttr(nodes, 'lang'), firstAttr(nodes, 'altLang'))}`)
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
      // One paragraph may hold any number of runs: the budget is checked per run, not only per paragraph.
      if (used + html.length > ctx.budget.left) {
        stopped = true
        break
      }
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
          // PowerPoint's own range for a list start; a crafted one would make the label (and the page) huge.
          counter = { level, type, value: Math.min(32767, Math.max(1, Math.trunc(num(kind, 'startAt') ?? 1))) - 1 }
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
    const paragraph = `<p class="pp-p" style="${pStyle.join(';')}">${bullet}${html || '<br>'}</p>`
    used += paragraph.length
    out.push({ html: paragraph, text })
  }
  if (stopped) {
    const notice = budgetNotice(ctx)
    if (notice) out.push({ html: notice, text: '' })
  }

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
    if (value < 1 || value > 3999) return String(value)
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
  const base = type.startsWith('alphaLc') ? alpha(n) : type.startsWith('alphaUc') ? alpha(n).toUpperCase() : type.startsWith('romanLc') ? roman(n)
    : type.startsWith('romanUc') ? roman(n).toUpperCase() : type.startsWith('circleNum') ? circled(n) : /ea1Cht|ea1Chs|ea1Jpn|taiwan|chinese/i.test(type) ? chineseCounting(n) : String(n)
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
    if (ctx.budget.left <= 0) {
      out.html += budgetNotice(ctx)
      return
    }
    const name = local(node.name)
    const before = out.html.length
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
      out.html += `<div class="pp-missing" style="left:0;top:0;width:100%;height:100%">${escapeHtml(readerMessage(error))}</div>`
    }
    // A group's children were counted as they were drawn.
    if (name !== 'grpSp') ctx.budget.left -= out.html.length - before
  }
}

/** The transform that undoes the flips of the groups around a shape, for what must stay readable. */
const counterFlip = (ctx: Ctx): string => [ctx.flipH ? 'scaleX(-1)' : '', ctx.flipV ? 'scaleY(-1)' : ''].filter(Boolean).join(' ')

const unflipped = (html: string, ctx: Ctx): string => {
  const flip = counterFlip(ctx)
  return flip ? `<div class="pp-grp" style="left:0;top:0;width:100%;height:100%;transform:${flip}">${html}</div>` : html
}

/** Groups nested deeper than this are a labelled box: browsers stop building a page about 500 elements deep. */
const GROUP_DEPTH = 64

function renderGroup(node: XmlNode, frame: Frame, ctx: Ctx, level: 'slide' | 'layout' | 'master', out: Rendered): void {
  const depth = (ctx.depth ?? 0) + 1
  if (depth > GROUP_DEPTH) {
    out.html += '<div class="pp-missing" style="left:0;top:0;width:100%;height:100%">群組的層數太多,裡面的內容請下載 PPTX 查看</div>'
    return
  }
  const x = xfrmOf(path(node, 'grpSpPr', 'xfrm'))
  if (!x) {
    renderTree(node, frame, { ...ctx, depth }, level, out)
    return
  }
  const style = boxStyle(x, frame)
  const flips = [x.flipH ? 'scaleX(-1)' : '', x.flipV ? 'scaleY(-1)' : ''].filter(Boolean)
  if (flips.length) style.push(`transform:${x.rot ? `rotate(${round(x.rot)}deg) ` : ''}${flips.join(' ')}`)
  const inner: Rendered = { html: '', text: [] }
  const flipped = { ...ctx, depth, flipH: Boolean(ctx.flipH) !== x.flipH, flipV: Boolean(ctx.flipV) !== x.flipV }
  renderTree(node, { x: x.chX, y: x.chY, cx: x.chCx || x.cx, cy: x.chCy || x.cy }, flipped, level, inner)
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
  const themed = fillNode ? null : styleRefFill(style, ctx)
  const background = fillNode ? fillCss(fillNode, ctx) : themed ? fillCss(themed.fill, themed.ctx) : ''
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
  // A plain rectangle stays a CSS box (its fills, patterns and borders are CSS's own); every other shape is its
  // ECMA-376 geometry drawn as SVG. A preset name the standard does not define is a labelled box.
  const plain = !geomNode || prst === 'rect'
  const shape = plain ? null : geometryOf(geomNode!, x.cx, x.cy)
  const unknown = !plain && shape === null
  if (background || line) {
    if (shape) {
      const paint = svgPaint(fillNode ?? themed?.fill ?? null, themed?.ctx ?? ctx, w, h)
      const strokeAttrs = line
        ? ` stroke="${line.color}" stroke-width="${round(line.width)}"${dashArray(line.dash, line.width) ? ` stroke-dasharray="${dashArray(line.dash, line.width)}"` : ''} stroke-linejoin="round"`
        : ''
      const [markerDefs, markerAttrs] = (line && shape.paths.some((p) => p.stroke && !p.closed) ? arrowMarkers(line, out) : '|').split('|')
      // Every fill first, in order, then every stroke (the standard lists chartPlus's lines before its box). A preset
      // fills each path as the standard says (smileyFace's open eye arcs are filled, its lines are not); an open path
      // of a custom shape is a line and takes no fill. Arrowheads go on open paths.
      const custom = local(geomNode!.name) === 'custGeom'
      let paths = ''
      if (paint.fill !== 'none') {
        for (const p of shape.paths) {
          if (p.fill === 'none' || (!p.closed && custom)) continue
          paths += `<path d="${p.d}" fill="${paint.fill}"/>`
          const shade = fillShade(p.fill)
          if (shade) paths += `<path d="${p.d}" fill="${shade.color}" fill-opacity="${shade.opacity}"/>`
        }
      }
      if (line) for (const p of shape.paths) if (p.stroke) paths += `<path d="${p.d}" fill="none"${strokeAttrs}${p.closed ? '' : markerAttrs ?? ''}/>`
      // A line with no height (or no width) gets a box as thick as its stroke, centred on it, so both axes keep one scale.
      const thick = line ? Math.max(line.width, 0.75) : 1
      const flatH = h < 0.01
      const flatW = w < 0.01
      const viewBox = [flatW ? -thick / 2 : 0, flatH ? -thick / 2 : 0, flatW ? thick : w, flatH ? thick : h].map((v) => round(v)).join(' ')
      const svgStyle = [
        flatH ? `top:${pt(-thick / 2)};height:${pt(thick)}` : '', flatW ? `left:${pt(-thick / 2)};width:${pt(thick)}` : '',
        flip ? `transform:${flip}` : '', shadow ? `filter:${shadow}` : ''
      ].filter(Boolean).join(';')
      if (paths) geometry = `<svg class="pp-geom" viewBox="${viewBox}" preserveAspectRatio="none"${svgStyle ? ` style="${svgStyle}"` : ''}>${markerDefs}${paint.defs ? `<defs>${paint.defs}</defs>` : ''}${paths}</svg>`
    } else if (plain || unknown) {
      const geom: string[] = []
      if (background) geom.push(`background:${background}`)
      if (line) {
        const half = pt(line.width / 2)
        geom.push(`inset:calc(-1*${half})`, `border:${pt(line.width)} ${line.dash === 'solid' ? 'solid' : line.dash.toLowerCase().includes('dot') && !line.dash.toLowerCase().includes('dash') ? 'dotted' : 'dashed'} ${line.color}`)
      }
      if (flip) geom.push(`transform:${flip}`)
      if (shadow) geom.push(`filter:${shadow}`)
      geometry = `<div class="pp-geom" style="${geom.join(';')}"></div>`
    }
  }
  if (unknown) geometry += `<div class="pp-missing" style="top:auto;height:auto">${escapeHtml(`未支援的圖形:${prst}`)}</div>`
  let text = ''
  if (txBody && hasText) {
    const fontRef = colorListStyle(colorElement(child(style, 'fontRef')))
    const lists = [...inherit.lists, fontRef, masterTextStyle(ph, ctx), ctx.presentationStyle].filter((n): n is XmlNode => n !== null)
    const rendered = renderTextBody(txBody, { lists, bodyPrs: inherit.bodyPrs }, ctx, { bodyPr: inherit.bodyPrs })
    text = unflipped(rendered.html, ctx)
    out.text.push(rendered.text)
    // Text in a shape PowerPoint flips stays readable: only a vertical flip turns it over.
    const txXfrm = xfrmOf(child(node, 'txXfrm'))
    if (txXfrm) {
      const inner = boxStyle({ ...txXfrm, x: txXfrm.x - x.x, y: txXfrm.y - x.y }, { x: 0, y: 0, cx: x.cx, cy: x.cy })
      text = `<div class="pp-grp" style="${inner.join(';')}">${text}</div>`
    } else if (shape) {
      // Text sits in the shape's text rectangle, as PowerPoint lays it out (a flip moves the rectangle, not the letters).
      const { l, t, r, b } = shape.textRect
      if (r > l && b > t && (l > 0.01 || t > 0.01 || w - r > 0.01 || h - b > 0.01)) {
        const left = x.flipH ? w - r : l
        const top = x.flipV ? h - b : t
        text = `<div class="pp-grp" style="left:${pct(left, w)};top:${pct(top, h)};width:${pct(r - l, w)};height:${pct(b - t, h)}">${text}</div>`
      }
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
  const source = blipFill ? blipSource(blipFill, ctx) : { none: 'missing' as const }
  const url = 'url' in source ? source.url : null
  const box = boxStyle(x, frame)
  const name = attr(path(node, 'nvPicPr', 'cNvPr'), 'descr') ?? attr(path(node, 'nvPicPr', 'cNvPr'), 'name') ?? ''
  const nvPr = path(node, 'nvPicPr', 'nvPr')
  const isMovie = Boolean(child(nvPr, 'videoFile') || child(nvPr, 'audioFile') || findDeep(nvPr ?? node, 'media'))
  if (!url) {
    const label = isMovie ? '影音(網頁版無法播放,請下載 PPTX 查看)' : escapeHtml('none' in source ? noPictureLabel(source.none, source.target) : noPictureLabel('missing'))
    out.html += `<div class="pp-sp" style="${box.join(';')}"><div class="pp-missing">${label}</div></div>`
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
  // A picture is cut to its shape. The path is in points while the slide scales, so it clips through an SVG
  // clipPath in the picture's own bounding box (0–1), scaled from the shape's box.
  const geomNode = child(spPr, 'prstGeom') ?? child(spPr, 'custGeom')
  const shape = geomNode && attr(geomNode, 'prst') !== 'rect' ? geometryOf(geomNode, x.cx, x.cy) : null
  const outline = shape?.paths.filter((p) => p.closed && p.fill !== 'none') ?? []
  const w = x.cx / EMU_PER_PT
  const h = x.cy / EMU_PER_PT
  let clipDefs = ''
  if (outline.length && w > 0 && h > 0) {
    const id = `pp-c${paintCount++}`
    clipDefs = `<svg width="0" height="0" style="position:absolute" aria-hidden="true"><clipPath id="${id}" clipPathUnits="objectBoundingBox"><path transform="scale(${round(1 / w, 8)} ${round(1 / h, 8)})" d="${outline.map((p) => p.d).join(' ')}"/></clipPath></svg>`
    clip.push(`clip-path:url(#${id})`)
  }
  if (flip) clip.push(`transform:${flip}`)
  const alpha = num(findDeep(blipFill!, 'alphaModFix'), 'amt')
  if (alpha !== undefined) clip.push(`opacity:${round(alpha / 100000)}`)
  const shadow = shadowCss(spPr, child(node, 'style'), ctx)
  if (shadow) clip.push(`filter:${shadow}`)
  const line = lineOf(child(spPr, 'ln'), child(node, 'style'), ctx)
  const stroke = line && shape
    ? shape.paths.filter((p) => p.stroke).map((p) => `<path d="${p.d}" fill="none" stroke="${line.color}" stroke-width="${round(line.width)}"/>`).join('')
    : ''
  const border = !line ? '' : stroke
    ? `<svg class="pp-geom" viewBox="0 0 ${round(w)} ${round(h)}" preserveAspectRatio="none"${flip ? ` style="transform:${flip}"` : ''}>${stroke}</svg>`
    : `<div class="pp-geom" style="inset:calc(-1*${pt(line.width / 2)});border:${pt(line.width)} solid ${line.color}"></div>`
  // A video or sound shows its poster frame, labelled so it does not pass for a still picture.
  const label = isMovie ? '<div class="pp-missing" style="top:auto;height:auto">影音(網頁版無法播放,請下載 PPTX 查看)</div>' : ''
  out.html += `<div class="pp-sp" style="${box.join(';')}">${clipDefs}<div class="pp-pic" style="${clip.join(';')}"><img src="${escapeHtml(url)}" alt="${escapeHtml(name)}" style="${imgStyle.join(';')}" decoding="async" draggable="false"></div>${border}${label}</div>`
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
      out.html += `<div class="pp-sp" style="${box.join(';')}">${unflipped(rendered.html, ctx)}</div>`
      out.text.push(rendered.text)
    }
    return
  }
  if (uri.endsWith('/diagram')) {
    if (ctx.inDiagram) {
      out.html += missing(box, 'SmartArt 圖形')
      return
    }
    const ids = child(data, 'relIds')
    const dataPart = ctx.pkg.rel(ctx.part, relAttr(ids, 'dm'))
    const dataXml = dataPart ? ctx.pkg.xml(dataPart.target) : null
    const drawingId = attr(dataXml ? findDeep(dataXml, 'dataModelExt') : null, 'relId')
    const drawing = ctx.pkg.rel(ctx.part, drawingId)
    const drawingXml = drawing ? ctx.pkg.xml(drawing.target) : null
    const tree = path(drawingXml, 'drawing', 'spTree') ?? (drawingXml ? findDeep(drawingXml, 'spTree') : null)
    if (tree && drawing) {
      const inner: Rendered = { html: '', text: [] }
      renderTree(tree, { x: 0, y: 0, cx: x.cx, cy: x.cy }, { ...ctx, part: drawing.target, inDiagram: true }, 'slide', inner)
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
    const chartCtx = { ...ctx, part: rel?.target ?? ctx.part }
    const rendered = chart
      ? renderChart(chart, {
          widthPt: x.cx / EMU_PER_PT,
          heightPt: x.cy / EMU_PER_PT,
          resolveColor: (holder) => {
            const color = colorIn(holder, chartCtx)
            return color ? cssColor(color) : null
          },
          themeColor: (i) => `#${ctx.theme.colors[`accent${i + 1}`] ?? OFFICE_ACCENTS[i % 6]}`,
          fontFamily: fontStack(undefined, undefined, ctx.theme),
          textColor: cssColor(hexToRgba(ctx.theme.colors[ctx.clrMap.tx1 ?? 'dk1'] ?? '404040') ?? { r: 64, g: 64, b: 64, a: 1 })
        })
      : null
    if (rendered) {
      out.html += `<div class="pp-sp" style="${box.join(';')}"><div class="pp-geom pp-chart">${unflipped(rendered.svg, ctx)}</div></div>`
      out.text.push(rendered.text)
    } else out.html += missing(box, chart && chartTitle(chart) ? `圖表「${chartTitle(chart)}」` : '圖表')
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

type Side = 'left' | 'right' | 'top' | 'bottom' | 'insideH' | 'insideV'
type CellStyle = { fill?: string; color?: XmlNode; bold?: boolean; borders: Partial<Record<Side, string>> }

const SIDES: Side[] = ['left', 'right', 'top', 'bottom', 'insideH', 'insideV']

/** A theme fill (`fillRef`), coloured by the colour inside the reference. */
const fillRefCss = (ref: XmlNode, ctx: Ctx): string =>
  fillCss(ctx.theme.fills[(num(ref, 'idx') ?? 1) - 1] ?? null, { ...ctx, phColor: colorIn(ref, ctx) ?? undefined })

function tablePart(styleNode: XmlNode | null, ctx: Ctx): CellStyle | null {
  if (!styleNode) return null
  const tcStyle = child(styleNode, 'tcStyle')
  const tcTx = child(styleNode, 'tcTxStyle')
  const out: CellStyle = { borders: {} }
  const fill = child(tcStyle, 'fill')
  if (fill) out.fill = fillCss(firstFill(fill), ctx) || (child(fill, 'noFill') ? 'transparent' : undefined)
  const fillRef = child(tcStyle, 'fillRef')
  if (!fill && fillRef) out.fill = fillRefCss(fillRef, ctx) || undefined
  for (const side of SIDES) {
    const ln = child(child(child(tcStyle, 'tcBdr'), side), 'ln')
    const lnRef = child(child(child(tcStyle, 'tcBdr'), side), 'lnRef')
    const line = ln ? lineOf(ln, null, ctx) : lnRef ? lineOf(null, node('p:style', {}, [lnRef]), ctx) : undefined
    if (line !== undefined) out.borders[side] = line ? borderCss(line) : 'none'
  }
  const color = colorElement(tcTx) ?? colorElement(child(tcTx, 'fontRef'))
  if (color) out.color = color
  const bold = attr(tcTx, 'b')
  if (bold === 'on' || bold === 'off') out.bold = bold === 'on'
  return out
}

// ---------------------------------------------------------------------------------------------
// The table styles PowerPoint has built in. A deck names them by id and seldom carries their
// definition (python-pptx and pptxgenjs never do), so they are rebuilt here from the pattern every
// family follows: one layout of parts, coloured by an accent or, for the plain variant, by text colours.

/** A scheme colour with modifiers: `accent1:tint20000`. */
const schemeColor = (spec: string): string => {
  const [val, ...mods] = spec.split(':')
  return `<a:schemeClr val="${val}">${mods.map((mod) => mod.replace(/^([a-z]+)(\d+)$/i, '<a:$1 val="$2"/>')).join('')}</a:schemeClr>`
}
const lnOf = (w: number, color: string, cmpd = 'sng') => `<a:ln w="${w}" cmpd="${cmpd}"><a:solidFill>${schemeColor(color)}</a:solidFill></a:ln>`
const lnRefOf = (idx: number, color: string) => `<a:lnRef idx="${idx}">${schemeColor(color)}</a:lnRef>`
const NO_LINE = '<a:ln><a:noFill/></a:ln>'
const allSides = (line: string): Partial<Record<Side, string>> => Object.fromEntries(SIDES.map((side) => [side, line]))
const edges = (line: string): Partial<Record<Side, string>> => ({ left: line, right: line, top: line, bottom: line })

type PartSpec = { b?: boolean; tx?: string; fill?: string; fillRef?: [number, string]; bdr?: Partial<Record<Side, string>> }

const part = (name: string, spec: PartSpec): string => {
  const tx = spec.b || spec.tx ? `<a:tcTxStyle${spec.b ? ' b="on"' : ''}>${spec.tx ? schemeColor(spec.tx) : ''}</a:tcTxStyle>` : ''
  const sides = SIDES.filter((side) => spec.bdr?.[side]).map((side) => `<a:${side}>${spec.bdr![side]}</a:${side}>`).join('')
  const fill = spec.fill === 'none' ? '<a:fill><a:noFill/></a:fill>' : spec.fill ? `<a:fill><a:solidFill>${schemeColor(spec.fill)}</a:solidFill></a:fill>` : ''
  const fillRef = spec.fillRef ? `<a:fillRef idx="${spec.fillRef[0]}">${schemeColor(spec.fillRef[1])}</a:fillRef>` : ''
  return `<a:${name}>${tx}<a:tcStyle><a:tcBdr>${sides}</a:tcBdr>${fill}${fillRef}</a:tcStyle></a:${name}>`
}
const tableBackground = (idx: number, color: string) => `<a:tblBg><a:fillRef idx="${idx}">${schemeColor(color)}</a:fillRef></a:tblBg>`

/** One family of built-in styles: `accent` is accent1–6, or null for the variant without an accent. */
type Family = (accent: string | null, second: string | null) => string

const noStyle = (grid: boolean) => part('wholeTbl', { tx: 'tx1', bdr: allSides(grid ? lnOf(12700, 'tx1') : NO_LINE), fill: 'none' })

const FAMILIES: [ids: string, build: Family][] = [
  // Themed Style 1; its plain variant is No Style, No Grid.
  ['2D5ABB26-0587-4C30-8999-92F81FD0307C 3C2FFA5D-87B4-456A-9821-1D502468CF0F 284E427A-3D55-4303-BF80-6455036E1DE7 69C7853C-536D-4A76-A0AE-DD22124D55A5 775DCB02-9BB8-47FD-8907-85C794F793BA 35758FB7-9AC5-4552-8A53-C91805E547FA 08FB837D-C827-4EFA-A057-4D05807E0F7C', (a) => {
    if (!a) return noStyle(false)
    const ref1 = lnRefOf(1, a)
    const ref2 = lnRefOf(2, a)
    return tableBackground(2, a) + part('wholeTbl', { tx: 'dk1', bdr: allSides(ref1), fill: 'none' }) +
      part('band1H', { fill: `${a}:alpha40000` }) + part('band1V', { bdr: { top: ref1, bottom: ref1 }, fill: `${a}:alpha40000` }) +
      part('lastCol', { b: true, bdr: { left: ref2, right: ref1, top: ref1, bottom: ref1, insideH: ref1, insideV: NO_LINE } }) +
      part('firstCol', { b: true, bdr: { left: ref1, right: ref2, top: ref1, bottom: ref1, insideH: ref1, insideV: NO_LINE } }) +
      part('lastRow', { b: true, bdr: { left: ref1, right: ref1, top: ref2, bottom: ref2, insideH: NO_LINE, insideV: NO_LINE }, fill: 'none' }) +
      part('firstRow', { b: true, tx: 'lt1', bdr: { left: ref1, right: ref1, top: ref1, bottom: lnRefOf(2, 'lt1'), insideH: NO_LINE, insideV: NO_LINE }, fill: a })
  }],
  // Themed Style 2; its plain variant is No Style, Table Grid.
  ['5940675A-B579-460E-94D1-54222C63F5DA D113A9D2-9D6B-4929-AA2D-F23B5EE8CBE7 18603FDC-E32A-4AB5-989C-0864C3EAD2B8 306799F8-075E-4A3A-A7F6-7FBC6576F1A4 E269D01E-BC32-4049-B463-5C60D7B0CCD2 327F97BB-C833-4FB7-BDE5-3F7075034690 638B1855-1B75-4FBE-930C-398BA8C253C6', (a) => {
    if (!a) return noStyle(true)
    return tableBackground(3, a) + part('wholeTbl', { tx: 'lt1', bdr: { ...edges(lnRefOf(1, `${a}:tint50000`)), insideH: NO_LINE, insideV: NO_LINE }, fill: 'none' }) +
      part('band1H', { fill: 'lt1:alpha20000' }) + part('band1V', { fill: 'lt1:alpha20000' }) +
      part('lastCol', { b: true, bdr: { left: lnRefOf(2, 'lt1') } }) + part('firstCol', { b: true, bdr: { right: lnRefOf(2, 'lt1') } }) +
      part('lastRow', { b: true, bdr: { top: lnRefOf(2, 'lt1') }, fill: 'none' }) +
      part('seCell', { bdr: { left: NO_LINE, top: NO_LINE } }) + part('swCell', { bdr: { right: NO_LINE, top: NO_LINE } }) +
      part('firstRow', { b: true, bdr: { bottom: lnRefOf(3, 'lt1') }, fill: 'none' }) + part('neCell', { bdr: { bottom: NO_LINE } })
  }],
  // Light Style 1
  ['9D7B26C5-4107-4FEC-AEDC-1716B250A1EF 3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5 0E3FDE45-AF77-4B5C-9715-49D594BDF05E C083E6E3-FA7D-4D7B-A595-EF9225AFEA82 D27102A9-8310-4765-A935-A1911B00CA55 5FD0F851-EC5A-4D38-B0AD-8093EC10F338 68D230F3-CF80-4859-8CE7-A43EE81993B5', (accent) => {
    const a = accent ?? 'tx1'
    const rule = lnOf(12700, a)
    return part('wholeTbl', { tx: 'tx1', bdr: { ...allSides(NO_LINE), top: rule, bottom: rule }, fill: 'none' }) +
      part('band1H', { fill: `${a}:alpha20000` }) + part('band1V', { fill: `${a}:alpha20000` }) +
      part('lastCol', { b: true }) + part('firstCol', { b: true }) +
      part('lastRow', { b: true, bdr: { top: rule }, fill: 'none' }) + part('firstRow', { b: true, bdr: { bottom: rule }, fill: 'none' })
  }],
  // Light Style 2
  ['7E9639D4-E3E2-4D34-9284-5A2195B3D0D7 69012ECD-51FC-41F1-AA8D-1B2483CD663E 72833802-FEF1-4C79-8D5D-14CF1EAF98D9 F2DE63D5-997A-4646-A377-4702673A728D 17292A2E-F333-43FB-9621-5CBBE7FDCDCB 5A111915-BE36-4E01-A7E5-04B1672EAD32 912C8C85-51F0-491E-9774-3900AFEF0FD7', (accent) => {
    const a = accent ?? 'tx1'
    const ref = lnRefOf(1, a)
    return part('wholeTbl', { tx: 'tx1', bdr: { ...edges(ref), insideH: NO_LINE, insideV: NO_LINE }, fill: 'none' }) +
      part('band1H', { bdr: { top: ref, bottom: ref } }) + part('band1V', { bdr: { left: ref, right: ref } }) + part('band2V', { bdr: { left: ref, right: ref } }) +
      part('lastCol', { b: true }) + part('firstCol', { b: true }) +
      part('lastRow', { b: true, bdr: { top: lnOf(50800, a, 'dbl') } }) + part('firstRow', { b: true, tx: 'bg1', fillRef: [1, a] })
  }],
  // Light Style 3
  ['616DA210-FB5B-4158-B5E0-FEB733F419BA BC89EF96-8CEA-46FF-86C4-4CE0E7609802 5DA37D80-6434-44D0-A028-1B22A696006F 8799B23B-EC83-4686-B30A-512413B5E67A ED083AE6-46FA-4A59-8FB0-9F97EB10719F BDBED569-4797-4DF1-A0F4-6AAB3CD982D8 E8B1032C-EA38-4F05-BA0D-38AFFFC7BED3', (accent) => {
    const a = accent ?? 'tx1'
    return part('wholeTbl', { tx: 'tx1', bdr: allSides(lnOf(12700, a)), fill: 'none' }) +
      part('band1H', { fill: `${a}:alpha20000` }) + part('band1V', { fill: `${a}:alpha20000` }) +
      part('lastCol', { b: true }) + part('firstCol', { b: true }) +
      part('lastRow', { b: true, bdr: { top: lnOf(50800, a, 'dbl') }, fill: 'none' }) + part('firstRow', { b: true, bdr: { bottom: lnOf(25400, a) }, fill: 'none' })
  }],
  // Medium Style 1
  ['793D81CF-94F2-401A-BA57-92F5A7B2D0C5 B301B821-A1FF-4177-AEE7-76D212191A09 9DCAF9ED-07DC-4A11-8D7F-57B35C25682E 1FECB4D8-DB02-4DC6-A0A2-4F2EBAE1DC90 1E171933-4619-4E11-9A3F-F7608DF75F80 FABFCF23-3B69-468F-B69F-88F6DE6A72F2 10A1B5D5-9B99-4C35-A422-299274C87663', (accent) => {
    const a = accent ?? 'dk1'
    const rule = lnOf(12700, a)
    return part('wholeTbl', { tx: 'dk1', bdr: { ...edges(rule), insideH: rule, insideV: NO_LINE }, fill: 'lt1' }) +
      part('band1H', { fill: `${a}:tint20000` }) + part('band1V', { fill: `${a}:tint20000` }) +
      part('lastCol', { b: true }) + part('firstCol', { b: true }) +
      part('lastRow', { b: true, bdr: { top: lnOf(50800, a, 'dbl') }, fill: 'lt1' }) + part('firstRow', { b: true, tx: 'lt1', fill: a })
  }],
  // Medium Style 2: PowerPoint's default.
  ['073A0DAA-6AF3-43AB-8588-CEC1D06C72B9 5C22544A-7EE6-4342-B048-85BDC9FD1C3A 21E4AEA4-8DFA-4A89-87EB-49C32662AFE0 F5AB1C69-6EDB-4FF4-983F-18BD219EF322 00A15C55-8517-42AA-B614-E9B94910E393 7DF18680-E054-41AD-8BC1-D1AEF772440D 93296810-A885-4BE3-A3E7-6D5BEEA58F35', (accent) => {
    const a = accent ?? 'dk1'
    return part('wholeTbl', { tx: 'dk1', bdr: allSides(lnOf(12700, 'lt1')), fill: `${a}:tint20000` }) +
      part('band1H', { fill: `${a}:tint40000` }) + part('band1V', { fill: `${a}:tint40000` }) +
      part('lastCol', { b: true, tx: 'lt1', fill: a }) + part('firstCol', { b: true, tx: 'lt1', fill: a }) +
      part('lastRow', { b: true, tx: 'lt1', bdr: { top: lnOf(38100, 'lt1') }, fill: a }) + part('firstRow', { b: true, tx: 'lt1', bdr: { bottom: lnOf(38100, 'lt1') }, fill: a })
  }],
  // Medium Style 3
  ['8EC20E35-A176-4012-BC5E-935CFFF8708E 6E25E649-3F16-4E02-A733-19D2CDBF48F0 85BE263C-DBD7-4A20-BB59-AAB30ACAA65A EB344D84-9AFB-497E-A393-DC336BA19D2E EB9631B5-78F2-41C9-869B-9F39066F8104 74C1A8A3-306A-4EB7-A6B1-4F7E0EB9C5D6 2A488322-F2BA-4B5B-9748-0D474271808F', (accent) => {
    const a = accent ?? 'dk1'
    const rule = lnOf(25400, 'dk1')
    return part('wholeTbl', { tx: 'dk1', bdr: { ...allSides(NO_LINE), top: rule, bottom: rule }, fill: 'lt1' }) +
      part('band1H', { fill: 'dk1:tint20000' }) + part('band1V', { fill: 'dk1:tint20000' }) +
      part('lastCol', { b: true, tx: 'lt1', fill: a }) + part('firstCol', { b: true, tx: 'lt1', fill: a }) +
      part('lastRow', { b: true, bdr: { top: lnOf(50800, 'dk1', 'dbl') }, fill: 'lt1' }) +
      part('seCell', { b: true, tx: 'dk1' }) + part('swCell', { b: true, tx: 'dk1' }) +
      part('firstRow', { b: true, tx: 'lt1', bdr: { bottom: rule }, fill: a })
  }],
  // Medium Style 4
  ['D7AC3CCA-C797-4891-BE02-D94E43425B78 69CF1AB2-1976-4502-BF36-3FF5EA218861 8A107856-5554-42FB-B03E-39F5DBC370BA 0505E3EF-67EA-436B-97B2-0124C06EBD24 C4B1156A-380E-4F78-BDF5-A606A8083BF9 22838BEF-8BB2-4498-84A7-C5851F593DF1 16D9F66E-5EB9-4882-86FB-DCBF35E3C3E4', (accent) => {
    const a = accent ?? 'dk1'
    return part('wholeTbl', { tx: 'dk1', bdr: allSides(lnOf(12700, a)), fill: `${a}:tint20000` }) +
      part('band1H', { fill: `${a}:tint40000` }) + part('band1V', { fill: `${a}:tint40000` }) +
      part('lastCol', { b: true }) + part('firstCol', { b: true }) +
      part('lastRow', { b: true, bdr: { top: lnOf(25400, a) }, fill: `${a}:tint20000` }) + part('firstRow', { b: true, fill: `${a}:tint20000` })
  }],
  // Dark Style 1: the plain variant lightens dark 1 where the accented ones darken their accent.
  ['E8034E78-7F5D-4C2E-B375-FC64B27BC917 125E5076-3810-47DD-B79F-674D7AD40C01 37CE84F3-28C3-443E-9E96-99CF82512B78 D03447BB-5D67-496B-8E87-E561075AD55C E929F9F4-4A8F-4326-A1B4-22849713DDAB 8FD4443E-F989-4FC4-A0C8-D5A2AF1F390B AF606853-7671-496A-8E4F-DF71F8EC918B', (a) => {
    const [body, band, cols, total] = a ? [a, `${a}:shade60000`, `${a}:shade60000`, `${a}:shade40000`] : ['dk1:tint20000', 'dk1:tint40000', 'dk1:tint60000', 'dk1:tint60000']
    const rule = lnOf(25400, 'lt1')
    return part('wholeTbl', { tx: 'lt1', bdr: allSides(NO_LINE), fill: body }) +
      part('band1H', { fill: band }) + part('band1V', { fill: band }) +
      part('lastCol', { b: true, bdr: { left: rule }, fill: cols }) + part('firstCol', { b: true, bdr: { right: rule }, fill: cols }) +
      part('lastRow', { b: true, bdr: { top: rule }, fill: total }) +
      part('seCell', { bdr: { left: NO_LINE } }) + part('swCell', { bdr: { right: NO_LINE } }) +
      part('firstRow', { b: true, bdr: { bottom: rule }, fill: 'dk1' }) +
      part('neCell', { bdr: { left: NO_LINE } }) + part('nwCell', { bdr: { right: NO_LINE } })
  }],
  // Dark Style 2: accents in pairs, the second one for the header.
  ['5202B0CA-FC54-4496-8BCA-5EF66A818D29 0660B408-B3CF-4A94-85FC-2B1E0A45F4A2 91EBBBCC-DAD2-459C-BE2E-F6DE35CF9A28 46F890A9-2807-4EBB-B81D-B2AA78EC7F39', (accent, second) => {
    const a = accent ?? 'dk1'
    return part('wholeTbl', { tx: 'dk1', bdr: allSides(NO_LINE), fill: `${a}:tint20000` }) +
      part('band1H', { fill: `${a}:tint40000` }) + part('band1V', { fill: `${a}:tint40000` }) +
      part('lastCol', { b: true }) + part('firstCol', { b: true }) +
      part('lastRow', { b: true, bdr: { top: lnOf(50800, 'dk1', 'dbl') }, fill: `${a}:tint20000` }) + part('firstRow', { b: true, tx: 'lt1', fill: second ?? 'dk1' })
  }]
]

const BUILTIN_STYLES = new Map<string, () => string>()
for (const [ids, build] of FAMILIES) {
  const list = ids.split(' ')
  const paired = list.length === 4
  list.forEach((id, at) => {
    const accent = at === 0 ? null : `accent${paired ? at * 2 - 1 : at}`
    const second = at === 0 || !paired ? null : `accent${at * 2}`
    BUILTIN_STYLES.set(`{${id}}`, () => build(accent, second))
  })
}
const builtinCache = new Map<string, XmlNode>()

/** A built-in table style by its id (any letter case), or null. */
function builtinTableStyle(id: string): XmlNode | null {
  const key = id.toUpperCase()
  const build = BUILTIN_STYLES.get(key)
  if (!build) return null
  let style = builtinCache.get(key)
  if (!style) {
    style = parseXml(`<a:tblStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" styleId="${key}">${build()}</a:tblStyle>`)
    builtinCache.set(key, style)
  }
  return style
}

/** The style a table uses: its own id from the deck's table styles or PowerPoint's built-in ones, else the deck's default. */
function tableStyleOf(tblPr: XmlNode | null, ctx: Ctx): XmlNode | null {
  const find = (id: string) => (id ? (children(ctx.tableStyles, 'tblStyle').find((style) => attr(style, 'styleId')?.toUpperCase() === id.toUpperCase()) ?? builtinTableStyle(id)) : null)
  return find(child(tblPr, 'tableStyleId')?.text.trim() ?? '') ?? find(attr(ctx.tableStyles, 'def') ?? '')
}

/** PowerPoint's own table limit; a crafted file with more cells draws these and says the rest is in the file. */
const TABLE_MAX = 75

/** Which edges of a part's area a cell is on: the whole table, one row, one column, or the cell itself. */
type Area = 'table' | 'row' | 'col' | 'cell'

function renderTable(table: XmlNode, ctx: Ctx): { html: string; text: string } {
  const tblPr = child(table, 'tblPr')
  const allGrid = children(child(table, 'tblGrid'), 'gridCol')
  const allRows = children(table, 'tr')
  const grid = allGrid.slice(0, TABLE_MAX).map((col) => (num(col, 'w') ?? 0) / EMU_PER_PT)
  const rows = allRows.slice(0, TABLE_MAX)
  let cut = allGrid.length > TABLE_MAX || allRows.length > TABLE_MAX
  let used = 0
  const styleDef = tableStyleOf(tblPr, ctx)
  const flag = (name: string) => truthy(attr(tblPr, name))
  const part = (name: string) => tablePart(child(styleDef, name), ctx)
  const whole = part('wholeTbl')
  const parts = {
    band1H: flag('bandRow') ? part('band1H') : null, band2H: flag('bandRow') ? part('band2H') : null,
    band1V: flag('bandCol') ? part('band1V') : null, band2V: flag('bandCol') ? part('band2V') : null,
    firstRow: flag('firstRow') ? part('firstRow') : null, lastRow: flag('lastRow') ? part('lastRow') : null,
    firstCol: flag('firstCol') ? part('firstCol') : null, lastCol: flag('lastCol') ? part('lastCol') : null,
    nwCell: part('nwCell'), neCell: part('neCell'), swCell: part('swCell'), seCell: part('seCell')
  }
  // Behind every cell: the table's own fill, else its style's table background.
  const tblBg = child(styleDef, 'tblBg')
  const ownBackground = firstFill(tblPr)
  const bgRef = child(tblBg, 'fillRef')
  const background = ownBackground ? fillCss(ownBackground, ctx) : child(tblBg, 'fill') ? fillCss(firstFill(child(tblBg, 'fill')), ctx) : bgRef ? fillRefCss(bgRef, ctx) : ''
  const texts: string[] = []
  let html = `<table class="pp-tbl"${background ? ` style="background:${background}"` : ''}><colgroup>${grid.map((w) => `<col style="width:${pt(w)}">`).join('')}</colgroup>`
  const lastRowAt = rows.length - 1
  const lastColAt = grid.length - 1
  for (const [r, row] of rows.entries()) {
    if (used > ctx.budget.left) {
      html += `<tr><td>${budgetNotice(ctx)}</td></tr>`
      break
    }
    html += `<tr style="height:${emuPt(num(row, 'h') ?? 0)}">`
    let c = 0
    for (const cell of children(row, 'tc')) {
      const col = c
      if (col >= TABLE_MAX) {
        cut = true
        break
      }
      const span = Math.max(1, Math.trunc(Number(attr(cell, 'gridSpan') ?? 1)) || 1)
      const rowSpan = Math.max(1, Math.trunc(Number(attr(cell, 'rowSpan') ?? 1)) || 1)
      c += span
      if (truthy(attr(cell, 'hMerge')) || truthy(attr(cell, 'vMerge'))) continue
      const colEnd = Math.min(col + span - 1, lastColAt)
      const rowEnd = Math.min(r + rowSpan - 1, lastRowAt)
      const isFirstRow = r === 0 && parts.firstRow !== null
      const isLastRow = rowEnd === lastRowAt && parts.lastRow !== null
      const isFirstCol = col === 0 && parts.firstCol !== null
      const isLastCol = colEnd === lastColAt && parts.lastCol !== null
      // Bands count from the first row and column after the header ones.
      const bodyRow = r - (parts.firstRow ? 1 : 0)
      const bodyCol = col - (parts.firstCol ? 1 : 0)
      // In PowerPoint's order: later parts win.
      const layers: [CellStyle | null, Area][] = [
        [whole, 'table'],
        [!isFirstRow && !isLastRow ? (bodyRow % 2 === 0 ? parts.band1H : parts.band2H) : null, 'row'],
        [!isFirstCol && !isLastCol ? (bodyCol % 2 === 0 ? parts.band1V : parts.band2V) : null, 'col'],
        [isLastCol ? parts.lastCol : null, 'col'],
        [isFirstCol ? parts.firstCol : null, 'col'],
        [isLastRow ? parts.lastRow : null, 'row'],
        [isFirstRow ? parts.firstRow : null, 'row'],
        [isLastRow && isLastCol ? parts.seCell : null, 'cell'],
        [isLastRow && isFirstCol ? parts.swCell : null, 'cell'],
        [isFirstRow && isLastCol ? parts.neCell : null, 'cell'],
        [isFirstRow && isFirstCol ? parts.nwCell : null, 'cell']
      ]
      const tcPr = child(cell, 'tcPr')
      const style: string[] = []
      let fill = ''
      let color: XmlNode | undefined
      let bold = false
      const borders: Record<string, string> = {}
      for (const [layer, area] of layers) {
        if (!layer) continue
        if (layer.fill) fill = layer.fill
        if (layer.color) color = layer.color
        if (layer.bold !== undefined) bold = layer.bold
        const b = layer.borders
        // A part's outer edges are the edges of its own area; inside it, its inside lines.
        const atLeft = area === 'col' || area === 'cell' || col === 0
        const atRight = area === 'col' || area === 'cell' || colEnd === lastColAt
        const atTop = area === 'row' || area === 'cell' || r === 0
        const atBottom = area === 'row' || area === 'cell' || rowEnd === lastRowAt
        const set = (side: string, value: string | undefined) => {
          if (value !== undefined) borders[side] = value
        }
        set('left', atLeft ? b.left : b.insideV)
        set('right', atRight ? b.right : b.insideV)
        set('top', atTop ? b.top : b.insideH)
        set('bottom', atBottom ? b.bottom : b.insideH)
      }
      const ownFill = firstFill(tcPr)
      if (ownFill) fill = fillCss(ownFill, ctx) || (local(ownFill.name) === 'noFill' ? 'transparent' : fill)
      for (const [side, name] of [['left', 'lnL'], ['right', 'lnR'], ['top', 'lnT'], ['bottom', 'lnB']] as const) {
        const ln = child(tcPr, name)
        if (ln) {
          const line = lineOf(ln, null, ctx)
          borders[side] = line ? borderCss(line) : 'none'
        }
      }
      if (fill && fill !== 'transparent') style.push(`background:${fill}`)
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
      const td = `<td${span > 1 ? ` colspan="${Math.min(span, TABLE_MAX)}"` : ''}${rowSpan > 1 ? ` rowspan="${Math.min(rowSpan, TABLE_MAX)}"` : ''} style="${style.join(';')}">${content}</td>`
      used += td.length
      html += td
    }
    html += '</tr>'
  }
  if (cut) html += `<tr><td colspan="${grid.length || 1}" style="padding:${pt(4)};font-size:${pt(12)};color:#59636e">表格太大,只顯示前 ${TABLE_MAX} 列、${TABLE_MAX} 欄,完整內容請下載 PPTX 查看</td></tr>`
  return { html: `${html}</table>`, text: texts.join('\n') }
}

// ---------------------------------------------------------------------------------------------
// Backgrounds and notes

function backgroundOf(xml: XmlNode | null, ctx: Ctx, box: { w: number; h: number }): string | null {
  const bg = path(xml, 'cSld', 'bg')
  if (!bg) return null
  const bgPr = child(bg, 'bgPr')
  if (bgPr) return fillCss(firstFill(bgPr), ctx, box) || 'transparent'
  const ref = child(bg, 'bgRef')
  if (ref) {
    const idx = num(ref, 'idx') ?? 0
    const refCtx = { ...ctx, phColor: colorIn(ref, ctx) ?? undefined }
    const fill = idx >= 1001 ? ctx.theme.bgFills[idx - 1001] : idx > 0 ? ctx.theme.fills[idx - 1] : undefined
    const css = fill ? fillCss(fill, refCtx, box) : ''
    return css || (refCtx.phColor ? cssColor(refCtx.phColor) : null)
  }
  return null
}

function notesHtml(pkg: Package, slidePart: string): { html: string; text: string } {
  const rel = pkg.relOfType(slidePart, 'notesSlide')
  const notes = rel ? pkg.optional(rel.target) : null
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
  const xml = rel ? pkg.optional(rel.target) : null
  const elements = path(xml, 'themeElements')
  const colors: Record<string, string> = {}
  for (const node of child(elements, 'clrScheme')?.children ?? []) {
    const value = node.children[0]
    if (!value) continue
    const name = local(value.name)
    colors[local(node.name)] = name === 'srgbClr' ? attr(value, 'val') ?? '' : name === 'sysClr' ? attr(value, 'lastClr') ?? '' : ''
  }
  const fonts = child(elements, 'fontScheme')
  const face = (node: XmlNode | null): ThemeFonts => ({
    latin: attr(child(node, 'latin'), 'typeface') || undefined,
    ea: attr(child(node, 'ea'), 'typeface') || undefined,
    scripts: Object.fromEntries(children(node, 'font').map((font) => [attr(font, 'script') ?? '', attr(font, 'typeface') ?? '']).filter(([script, face]) => script && face))
  })
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
/** `errors`: why slides could not be drawn, as thrown, for the viewer's error log. */
export type PptxDeck = { width: number; height: number; slides: PptxSlide[]; title: string; errors?: string[] }

/** Reads a deck: its slide size and every shown slide as HTML. `mediaBase` prefixes picture URLs. */
export function readPptx(data: Buffer, mediaBase: string): PptxDeck {
  // Password-protected decks and the old binary .ppt are compound files, not zips.
  if (data.length >= 8 && data.readUInt32BE(0) === 0xd0cf11e0) throw new Error('這是舊版 PowerPoint (.ppt) 或受密碼保護的檔案,網頁版無法開啟。')
  const pkg = new Package(data)
  const presentationPart = pkg.relOfType('', 'officeDocument')?.target ?? 'ppt/presentation.xml'
  const presentation = pkg.xml(presentationPart)
  if (!presentation) throw new Error('這不是 PowerPoint 檔(找不到 presentation.xml)')
  const size = child(presentation, 'sldSz')
  const width = num(size, 'cx') ?? 12192000
  const height = num(size, 'cy') ?? 6858000
  const presentationStyle = child(presentation, 'defaultTextStyle')
  const tableStylesRel = pkg.relOfType(presentationPart, 'tableStyles')
  const tableStyles = tableStylesRel ? pkg.optional(tableStylesRel.target) : null
  const slideParts = children(child(presentation, 'sldIdLst'), 'sldId')
    .map((sld) => pkg.rel(presentationPart, relAttr(sld, 'id')))
    .filter((rel): rel is Rel => Boolean(rel && !rel.external))
    .map((rel) => rel.target)
  // A slide that cannot be read counts as shown, so it keeps its number and says why on its own face.
  const shown = slideParts.filter((part) => attr(pkg.optional(part), 'show') !== '0')
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
  const docTitle = pkg.optional('docProps/core.xml')
  const deckBudget = { left: DECK_BUDGET }
  const title = docTitle ? (findDeep(docTitle, 'title')?.text.trim() ?? '') : ''
  const errors: string[] = []

  // One slide that cannot be read shows why on its own face; the rest of the deck still draws.
  const slides = shown.map((part, at): PptxSlide => {
    const budget: Budget = { left: Math.min(SLIDE_BUDGET, deckBudget.left), noticed: false }
    try {
      const slide = drawSlide(part, at, budget)
      deckBudget.left -= slide.html.length
      return slide
    } catch (error) {
      const number = at + 1
      errors.push(`第 ${number} 頁:${error instanceof Error ? error.message : String(error)}`)
      const message = `第 ${number} 頁無法顯示:${readerMessage(error)}`
      return {
        html: `<section class="slide layout-pptx" id="slide-${number}" data-number="${number}" aria-label="第 ${number} 頁"><div class="pp-canvas"><div class="pp-missing">${escapeHtml(message)}</div></div></section>\n`,
        text: '',
        notes: ''
      }
    }
  })
  return { width, height, slides, title, errors }

  function drawSlide(part: string, at: number, budget: Budget): PptxSlide {
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
      pkg, theme: master.theme, slideNumber: at + 1, slideOf: (target) => numberOf.get(target), media, master, layout, presentationStyle, tableStyles, budget
    }
    const masterCtx: Ctx = { ...base, part: master.part, clrMap: master.clrMap }
    const layoutCtx: Ctx = { ...base, part: layoutPart, clrMap: override(layoutXml) ?? master.clrMap }
    const slideCtx: Ctx = { ...base, part, clrMap: override(xml) ?? layoutCtx.clrMap }
    const frame: Frame = { x: 0, y: 0, cx: width, cy: height }
    const out: Rendered = { html: '', text: [] }
    const box = { w: width / EMU_PER_PT, h: height / EMU_PER_PT }
    const background = backgroundOf(xml, slideCtx, box) ?? backgroundOf(layoutXml, layoutCtx, box) ?? backgroundOf(master.xml, masterCtx, box) ?? '#fff'
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

/** Why a deck cannot be read, in words for the reader. */
function failureReason(error: unknown): string {
  if (error instanceof ZipError) return '檔案不是有效的 .pptx(zip 格式),可能已損毀,或副檔名不符。'
  if (error instanceof XmlError) return '簡報內容的 XML 無法解析,檔案可能已損毀。'
  if (error instanceof Error && /[一-鿿]/.test(error.message)) return error.message
  return '讀取這份簡報時發生錯誤。'
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
    const reason = failureReason(error)
    const detail = readerMessage(error)
    const box = `<div class="pp-missing" role="alert" style="flex-direction:column;gap:calc(6*var(--pt))"><strong style="font-size:calc(20*var(--pt))">這份 PowerPoint 檔無法顯示</strong><span>${escapeHtml(reason)}</span>${detail === reason ? '' : `<code>${escapeHtml(detail)}</code>`}<span>可以下載原始檔,用 PowerPoint 開啟查看。</span></div>`
    const slide = `<section class="slide layout-pptx" id="slide-1" data-number="1" aria-label="第 1 頁"><div class="pp-canvas">${box}</div></section>`
    return deckPage(slide, title, assets, failureScript(message))
  }
  const widthPt = round(deck.width / EMU_PER_PT, 3)
  const heightPt = round(deck.height / EMU_PER_PT, 3)
  // A slide that could not be drawn says so on its face; why lands in the viewer's error log.
  const errors = deck.errors?.length ? failureScript(`PowerPoint: ${deck.errors.join('; ')}`.slice(0, 500)) : ''
  const head = `${errors}<style>:root{--aspect:${round(deck.width / deck.height, 6)};--slide-pt:${widthPt}}@page{size:${widthPt}pt ${heightPt}pt}@media print{.slide,body.overview .slide{width:${widthPt}pt;height:${heightPt}pt}}</style>`
  const html = deck.slides.length
    ? deck.slides.map((slide) => slide.html).join('')
    : '<section class="slide layout-pptx" id="slide-1" data-number="1" aria-label="第 1 頁"><div class="pp-canvas"><div class="pp-missing">這份簡報沒有可顯示的投影片</div></div></section>'
  return deckPage(html, title, assets, head)
}

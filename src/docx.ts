import { posix } from 'node:path'
import { escapeHtml, injectBridge } from './content-render.ts'
import { chartTitle, renderChart } from './office-chart.ts'
import { fillShade, shapeGeometry } from './office-geometry.ts'
import { AIUEO, AIUEO_HALF, cardinalText, chineseCounting, circled, chineseLegal, CHOSUNG, CN_DIGITS, GANADA, IROHA, IROHA_HALF, letters, ordinal, ordinalText, repeating, roman, RUSSIAN } from './office-numbers.ts'
import { PATTERN_PT, PATTERNS, patternCover, patternSvg, patternTile as sharedTile } from './office-patterns.ts'
import { failureScript, isDrawablePart, officeMediaBase, officeMediaUrl } from './office.ts'
import { attr, child, children, local, num, parseXml, path, relAttr, XmlError, type XmlNode } from './xml.ts'
import { ZipReader, ZipError } from './zip.ts'

/**
 * A Word file published as is (`*.docx`), drawn on the server as one continuous HTML page: a paper
 * column per section at the section's own page width, every look read from the file (styles.xml with
 * its basedOn chains, the theme's fonts and colours, numbering.xml, table styles). Text stays text;
 * pictures load from the file through the Office media route (office.ts). What it does not draw
 * (charts, SmartArt, OLE objects) shows as a labelled box; the original file is one download away.
 */
const EMU_PER_PT = 12700
/** Office's default theme accents, for a file without a theme of its own. */
const OFFICE_ACCENTS = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47']
const LIMITS = { blocks: 20_000, cells: 50_000, cellBlocks: 100_000, notes: 2_000, depth: 12, chain: 20 }
/** The largest list start Word accepts; a file's bigger number is clamped to it. */
const MAX_LIST_START = 32_767
/** Table-style regions that are one row tall, one column wide, or one cell. */
const ROW_REGIONS = new Set(['firstRow', 'lastRow', 'band1Horz', 'band2Horz'])
const COL_REGIONS = new Set(['firstCol', 'lastCol', 'band1Vert', 'band2Vert'])
const CORNER_REGIONS = new Set(['nwCell', 'neCell', 'swCell', 'seCell'])
const TAB = '\u0001'

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
      this.parts.set(name, text === null ? null : parseXml(text.replace(/^﻿/, '')))
    }
    return this.parts.get(name) ?? null
  }

  rels(part: string): Map<string, Rel> {
    let map = this.relsCache.get(part)
    if (map) return map
    map = new Map()
    this.relsCache.set(part, map)
    const dir = posix.dirname(part)
    let doc: XmlNode | null = null
    try {
      doc = this.xml(posix.join(dir, '_rels', `${posix.basename(part)}.rels`))
    } catch {
      doc = null
    }
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
    return map
  }

  rel(part: string, id: string | undefined): Rel | undefined {
    return id ? this.rels(part).get(id) : undefined
  }

  relOfType(part: string, suffix: string): Rel | undefined {
    for (const rel of this.rels(part).values()) if (rel.type.endsWith(`/${suffix}`) && !rel.external) return rel
    return undefined
  }
}

// ---------------------------------------------------------------------------------------------
// Colours and numbers

type Rgba = { r: number; g: number; b: number; a: number }

const round = (value: number, digits = 2) => {
  const factor = 10 ** digits
  return Math.round(value * factor) / factor
}
const clamp = (value: number, min: number, max: number) => Math.min(max, Math.max(min, value))
const clamp01 = (value: number) => clamp(value, 0, 1)
const pt = (value: number) => `${round(clamp(value, -5000, 5000))}pt`

const hexToRgba = (hex: string | undefined): Rgba | null => {
  if (!hex || !/^[0-9a-fA-F]{6}$/.test(hex)) return null
  return { r: parseInt(hex.slice(0, 2), 16), g: parseInt(hex.slice(2, 4), 16), b: parseInt(hex.slice(4, 6), 16), a: 1 }
}

const cssColor = ({ r, g, b, a }: Rgba): string => {
  const [R, G, B] = [r, g, b].map((v) => Math.round(clamp(v, 0, 255)))
  return a >= 1 ? `rgb(${R},${G},${B})` : `rgba(${R},${G},${B},${round(clamp01(a), 3)})`
}

const isDark = (color: Rgba | null | undefined) => Boolean(color && color.a > 0.5 && (0.299 * color.r + 0.587 * color.g + 0.114 * color.b) / 255 < 0.45)

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

/** DrawingML colour modifiers (text boxes and shapes inside a Word file use DrawingML). */
function applyModifiers(color: Rgba, element: XmlNode): Rgba {
  let current = { ...color }
  for (const mod of element.children) {
    const value = (num(mod, 'val') ?? 0) / 100000
    const name = local(mod.name)
    if (name === 'alpha') current.a = clamp01(value)
    else if (name === 'tint') current = { ...current, r: current.r + (255 - current.r) * (1 - value), g: current.g + (255 - current.g) * (1 - value), b: current.b + (255 - current.b) * (1 - value) }
    else if (name === 'shade') current = { ...current, r: current.r * value, g: current.g * value, b: current.b * value }
    else if (name === 'lumMod' || name === 'lumOff' || name === 'satMod') {
      let [h, s, l] = rgbToHsl(current)
      if (name === 'lumMod') l = clamp01(l * value)
      else if (name === 'lumOff') l = clamp01(l + value)
      else s = clamp01(s * value)
      current = hslToRgb(h, s, l, current.a)
    }
  }
  return current
}

const mix = (fg: Rgba, bg: Rgba, amount: number): Rgba => ({
  r: fg.r * amount + bg.r * (1 - amount),
  g: fg.g * amount + bg.g * (1 - amount),
  b: fg.b * amount + bg.b * (1 - amount),
  a: 1
})

const BLACK: Rgba = { r: 0, g: 0, b: 0, a: 1 }
const WHITE: Rgba = { r: 255, g: 255, b: 255, a: 1 }

const HIGHLIGHT: Record<string, string> = {
  yellow: 'FFFF00', green: '00FF00', cyan: '00FFFF', magenta: 'FF00FF', blue: '0000FF', red: 'FF0000', darkBlue: '000080', darkCyan: '008080',
  darkGreen: '008000', darkMagenta: '800080', darkRed: '800000', darkYellow: '808000', darkGray: '808080', lightGray: 'C0C0C0', black: '000000', white: 'FFFFFF'
}

const PRESET_COLORS: Record<string, string> = {
  black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', gray: '808080', grey: '808080', silver: 'C0C0C0',
  navy: '000080', teal: '008080', maroon: '800000', olive: '808000', lime: '00FF00', aqua: '00FFFF', fuchsia: 'FF00FF', purple: '800080', orange: 'FFA500'
}

/** w:themeColor names → theme slots, before settings' clrSchemeMapping is applied. */
const WORD_THEME_SLOTS: Record<string, string> = {
  dark1: 'dk1', light1: 'lt1', dark2: 'dk2', light2: 'lt2', text1: 'dk1', background1: 'lt1', text2: 'dk2', background2: 'lt2',
  accent1: 'accent1', accent2: 'accent2', accent3: 'accent3', accent4: 'accent4', accent5: 'accent5', accent6: 'accent6', hyperlink: 'hlink', followedHyperlink: 'folHlink'
}
const SCHEME_SLOTS: Record<string, string> = { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' }

const truthy = (node: XmlNode | null | undefined): boolean => {
  const value = attr(node, 'val')
  return value === undefined || !['0', 'false', 'off', 'none'].includes(value)
}

// ---------------------------------------------------------------------------------------------
// Fonts

const safeFont = (name: string | undefined) => (name ?? '').replace(/[^\p{L}\p{N} ._-]/gu, '').trim().slice(0, 80)

type FontKind = 'sans' | 'serif' | 'mono' | 'kai'
const LATIN_FALLBACK: Record<FontKind, string> = {
  sans: "Arial,'Helvetica Neue'",
  serif: "'Times New Roman',Times",
  mono: "Menlo,Consolas,'Courier New'",
  kai: "'Times New Roman',Times"
}
const CJK_FALLBACK: Record<FontKind, string> = {
  sans: "'PingFang TC','Heiti TC','Noto Sans TC','Noto Sans CJK TC','Microsoft JhengHei'",
  serif: "'Songti TC','Noto Serif TC','Noto Serif CJK TC','PMingLiU'",
  mono: "'PingFang TC','Noto Sans TC','Noto Sans CJK TC','Microsoft JhengHei'",
  kai: "'BiauKai','Kaiti TC','DFKai-SB','STKaiti'"
}
const CJK_SERIF_FONT = /細明|明體|明朝|宋|ming|song|sun\b|simsun|mincho|batang|gungsuh|바탕|궁서/i
const CJK_SANS_FONT = /黑|ゴシック|gothic|heiti|jhenghei|yahei|simhei|meiryo|メイリオ|gulim|dotum|malgun|굴림|돋움/i
const GENERIC: Record<FontKind, string> = { sans: 'sans-serif', serif: 'serif', mono: 'monospace', kai: 'serif' }

const SYMBOL_FONT = /^(symbol|wingdings( [23])?|webdings)$/i

const WINGDINGS: Record<number, string> = {
  0x6c: '●', 0x6e: '■', 0x71: '❑', 0x75: '◆', 0x76: '❖', 0xd8: '➢', 0xfc: '✓', 0xfb: '✗', 0xfd: '☒', 0xfe: '☑', 0xa7: '▪', 0x9f: '•',
  0x6f: '□', 0x70: '◻', 0xa8: '◻', 0xe0: '→', 0xe8: '➔', 0xf0: '⇨', 0x4a: '☺', 0x4c: '☹', 0x28: '☎', 0x2a: '✉', 0x46: '☞', 0xa1: '○', 0xa4: '◉', 0x77: '⬥', 0x73: '◆', 0x74: '◆', 0x22: '✂'
}
const SYMBOL_GREEK_LOWER = 'αβχδεφγηιϕκλμνοπθρστυϖωξψζ'
const SYMBOL_GREEK_UPPER = 'ΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟΠΘΡΣΤΥςΩΞΨΖ'
const SYMBOL: Record<number, string> = {
  0xb7: '•', 0xae: '→', 0xac: '←', 0xad: '↑', 0xaf: '↓', 0xb4: '×', 0xb8: '÷', 0xa3: '≤', 0xb3: '≥', 0xb9: '≠', 0xbb: '≈', 0xa5: '∞', 0xd6: '√',
  0xe5: '∑', 0xf2: '∫', 0xb1: '±', 0xb0: '°', 0xd7: '⋅', 0xa7: '♣', 0xa8: '♦', 0xa9: '♥', 0xaa: '♠', 0xde: '⇒', 0xdb: '⇔', 0xce: '∈', 0xc7: '∩', 0xc8: '∪', 0xb6: '∂', 0x22: '∀', 0x24: '∃'
}

/** A character typed in a symbol font, as the Unicode character it shows; null when not known. */
function symbolChar(font: string, code: number): string | null {
  const cp = code >= 0xf000 && code <= 0xf0ff ? code - 0xf000 : code
  if (/^symbol$/i.test(font)) {
    if (cp >= 0x61 && cp <= 0x7a) return SYMBOL_GREEK_LOWER[cp - 0x61]
    if (cp >= 0x41 && cp <= 0x5a) return SYMBOL_GREEK_UPPER[cp - 0x41]
    return SYMBOL[cp] ?? null
  }
  if (/wingdings|webdings/i.test(font)) return WINGDINGS[cp] ?? null
  return null
}

// ---------------------------------------------------------------------------------------------
// Number formats

const listStart = (n: number | undefined) => (n === undefined || !Number.isFinite(n) ? undefined : clamp(Math.trunc(n), 0, MAX_LIST_START))

function formatNumber(n: number, format: string): string {
  switch (format) {
    case 'none':
      return ''
    case 'decimalZero':
      return n < 10 && n >= 0 ? `0${n}` : String(n)
    case 'upperRoman':
      return roman(n)
    case 'lowerRoman':
      return roman(n).toLowerCase()
    case 'upperLetter':
      return letters(n)
    case 'lowerLetter':
      return letters(n).toLowerCase()
    case 'ordinal':
      return ordinal(n)
    case 'cardinalText':
      return cardinalText(n)
    case 'ordinalText':
      return ordinalText(n)
    case 'hex':
      return n >= 0 ? n.toString(16).toUpperCase() : String(n)
    case 'chicago':
      return repeating(n, ['*', '†', '‡', '§'])
    case 'aiueo':
      return repeating(n, AIUEO_HALF)
    case 'aiueoFullWidth':
      return repeating(n, AIUEO)
    case 'iroha':
      return repeating(n, IROHA_HALF)
    case 'irohaFullWidth':
      return repeating(n, IROHA)
    case 'ganada':
      return repeating(n, GANADA)
    case 'chosung':
      return repeating(n, CHOSUNG)
    case 'russianLower':
      return repeating(n, RUSSIAN)
    case 'russianUpper':
      return repeating(n, RUSSIAN).toUpperCase()
    case 'chineseCounting':
    case 'chineseCountingThousand':
    case 'taiwaneseCounting':
    case 'taiwaneseCountingThousand':
    case 'japaneseCounting':
    case 'koreanCounting':
      return chineseCounting(n)
    case 'chineseLegalSimplified':
    case 'ideographLegalTraditional':
    case 'japaneseLegal':
      return chineseLegal(n)
    case 'ideographDigital':
    case 'taiwaneseDigital':
    case 'japaneseDigitalTenThousand':
      return String(n).replace(/\d/g, (d) => CN_DIGITS[Number(d)])
    case 'ideographTraditional':
      return n > 0 ? '甲乙丙丁戊己庚辛壬癸'[(n - 1) % 10] : String(n)
    case 'ideographZodiac':
      return n > 0 ? '子丑寅卯辰巳午未申酉戌亥'[(n - 1) % 12] : String(n)
    case 'ideographEnclosedCircle':
      return n >= 1 && n <= 10 ? String.fromCodePoint(0x3280 + n - 1) : String(n)
    case 'decimalEnclosedCircle':
    case 'decimalEnclosedCircleChinese':
      return circled(n)
    case 'decimalEnclosedParen':
      return n >= 1 && n <= 20 ? String.fromCodePoint(0x2474 + n - 1) : `(${n})`
    case 'decimalEnclosedFullstop':
      return n >= 1 && n <= 20 ? String.fromCodePoint(0x2488 + n - 1) : `${n}.`
    case 'decimalFullWidth':
    case 'decimalFullWidth2':
      return String(n).replace(/\d/g, (d) => String.fromCodePoint(0xff10 + Number(d)))
    default:
      return String(n)
  }
}

// ---------------------------------------------------------------------------------------------
// Properties: what a run or paragraph looks like after its styles are applied

type RunProps = {
  b?: boolean
  i?: boolean
  u?: string
  uColor?: Rgba
  strike?: boolean
  dstrike?: boolean
  color?: Rgba | 'auto'
  highlight?: Rgba | null
  shd?: Rgba | null
  size?: number
  ascii?: string
  hAnsi?: string
  ea?: string
  vert?: string
  caps?: boolean
  smallCaps?: boolean
  spacing?: number
  position?: number
  vanish?: boolean
  rtl?: boolean
  em?: string
  shadow?: boolean
  outline?: boolean
}

type Len = { pt?: number; em?: number }
type TabStop = { val: string; pos: number; leader: string }
type ParaProps = {
  jc?: string
  left?: Len
  right?: Len
  firstLine?: Len
  hanging?: Len
  before?: number
  after?: number
  line?: number
  lineRule?: string
  contextual?: boolean
  borders: Record<string, XmlNode>
  shd?: Rgba | null
  bidi?: boolean
  pageBreakBefore?: boolean
  outline?: number
  tabs: TabStop[]
  snap?: boolean
}

const lenCss = (len: Len | undefined, negate = false): string | null => {
  if (!len) return null
  const sign = negate ? -1 : 1
  if (len.em !== undefined) return `${round(clamp(len.em * sign, -100, 100))}em`
  if (len.pt !== undefined) return pt(len.pt * sign)
  return null
}

type StyleCond = { pPr: XmlNode | null; rPr: XmlNode | null; tblPr: XmlNode | null; tcPr: XmlNode | null }
type StyleDef = StyleCond & { id: string; type: string; basedOn?: string; conds: Map<string, StyleCond> }

type Level = { start: number; fmt: string; text: string; pPr: XmlNode | null; rPr: XmlNode | null; suff: string; isLgl: boolean; restart?: number }
type NumDef = { abstractId: string; overrides: Map<number, { start?: number; level?: Level }> }

type Field = { instr: string; separated: boolean }
type NoteKind = 'footnote' | 'endnote'
type Story = { part: string; kind: 'body' | 'header' | 'footer' | 'note'; fields: Field[]; note?: { kind: NoteKind; n: number }; overflow?: number }
type Layer = { pPr: (XmlNode | null)[]; rPr: (XmlNode | null)[] }
type Geometry = { w: number; h: number; top: number; right: number; bottom: number; left: number; header: number; footer: number; cols: number; colSpace: number; grid: number; vertical: boolean; type: string }
/** `groupFill`: the shape properties of the group a drawing sits in, for shapes that take the group's fill (`grpFill`). */
/** `flipH`/`flipV`: the enclosing groups mirror what they hold (the flips of every group around a drawing, combined). */
type Scope = { story: Story; layer: Layer; depth: number; width: number; dark: boolean; grid: number; sheet: Geometry | null; groupFill?: XmlNode | null; flipH?: boolean; flipV?: boolean }
/** A shape's fill: a colour, one of Office's 8 × 8 hatches, or a picture from the file (stretched over the shape, or tiled). */
type Paint = { color: Rgba } | { pattern: { fg: Rgba; bg: Rgba; bits: string } } | { picture: string; tile: boolean }
type Link = { href?: string; title?: string } | null
/**
 * A paragraph's inline content; `pins` are drawings placed relative to the paragraph itself, drawn at its start;
 * `floats` are drawings the text wraps around, which sit at the paragraph's top wherever their anchor run is.
 */
type Inline = { parts: string[]; text: string[]; visual: boolean; pins: string[]; floats: string[] }
/** Where a floating drawing sits, in points: `relH`/`relV` are Word's relativeFrom values. */
type Placement = { relH: string; relV: string; alignH?: string; alignV?: string; offH: number; offV: number; w: number; h: number; behind: boolean; rotate?: number }

/** Whether the groups around a drawing mirror it along exactly one axis, which turns its rotation the other way. */
const mirrored = (scope: Scope) => Boolean(scope.flipH) !== Boolean(scope.flipV)

/** An offset along one side: an alignment inside [start, start + area], or start + offset. */
const alignIn = (align: string | undefined, start: number, area: number, size: number, offset: number) => {
  if (align === 'center') return start + (area - size) / 2
  if (align === 'right' || align === 'bottom' || align === 'outside') return start + area - size
  if (align === 'left' || align === 'top' || align === 'inside') return start
  return start + offset
}
/** A VML length (CSS-like, unitless means pixels) in points. */
const vmlLength = (value: string | undefined): number => {
  const match = /^(-?[\d.]+)(pt|in|cm|mm|px|pc)?$/.exec((value ?? '').trim())
  if (!match) return 0
  const factor = { pt: 1, in: 72, cm: 72 / 2.54, mm: 72 / 25.4, px: 0.75, pc: 12 }[match[2] ?? 'px'] ?? 0.75
  return clamp(Number(match[1]) * factor, -4000, 4000) || 0
}

/** A hatch tile's SVG content, in the shape's colours. */
const patternTile = ({ fg, bg, bits }: { fg: Rgba; bg: Rgba; bits: string }): string => sharedTile(cssColor(fg), cssColor(bg), bits)

/** The colour a fill reads as from a distance, for choosing the text colour on it; null for a picture. */
function paintTone(paint: Paint): Rgba | null {
  if ('color' in paint) return paint.color
  if ('picture' in paint) return null
  return mix(paint.pattern.fg, paint.pattern.bg, patternCover(paint.pattern.bits))
}

/** A fill as a CSS background, for shapes drawn as plain boxes. */
function paintCss(paint: Paint): string {
  if ('color' in paint) return `background:${cssColor(paint.color)}`
  if ('picture' in paint) return `background:url('${esc(paint.picture)}') 0 0${paint.tile ? ' repeat' : '/100% 100% no-repeat'}`
  const svg = patternSvg(cssColor(paint.pattern.fg), cssColor(paint.pattern.bg), paint.pattern.bits)
  return `background:url('data:image/svg+xml,${encodeURIComponent(svg)}') 0 0/${PATTERN_PT}pt ${PATTERN_PT}pt repeat`
}

const PAGE_BREAK = '<span class="dx-pgbr" role="separator" aria-label="分頁"></span>'
type ParaInfo = { node: XmlNode; styleId: string; props: ParaProps; base: RunProps; mark: RunProps; num?: { numId: string; ilvl: number }; sig: string }

const EMPTY_LAYER: Layer = { pPr: [], rPr: [] }

/** Element children of a mc:AlternateContent to draw: the Choice when this renderer can draw it, else the Fallback. */
const DRAWABLE_PREFIXES = new Set(['wps', 'wpg', 'wpc', 'wp14', 'w14', 'w15', 'w16', 'w16se', 'w16cid', 'w16cex', 'w16du', 'w16sdtdh', 'a14', 'v', 'o', 'w10', 'pic', 'a', 'r', 'w', 'wp', 'mc'])

function findGraphicUri(node: XmlNode, budget = { left: 5000 }): string[] {
  const found: string[] = []
  const walk = (current: XmlNode) => {
    if (--budget.left < 0) return
    if (local(current.name) === 'graphicData') found.push(attr(current, 'uri') ?? '')
    for (const next of current.children) walk(next)
  }
  walk(node)
  return found
}

function chooseAlternate(node: XmlNode): XmlNode[] {
  const fallback = child(node, 'Fallback')
  for (const choice of children(node, 'Choice')) {
    const requires = (attr(choice, 'Requires') ?? '').split(/\s+/).filter(Boolean)
    if (!requires.every((prefix) => DRAWABLE_PREFIXES.has(prefix))) continue
    if (fallback && findGraphicUri(choice).some((uri) => /chart|diagram/i.test(uri))) continue
    return choice.children
  }
  return fallback?.children ?? []
}

/** Ids for bookmarks and the links to them: deterministic, and only [A-Za-z0-9_-]. */
const anchorId = (name: string) => `bm-${name.slice(0, 80).replace(/[^A-Za-z0-9_-]/g, (char) => `-${char.codePointAt(0)!.toString(16)}-`)}`

const cleanText = (text: string) => text.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
/** Every string from the file reaches the page through here: control characters dropped (one is the tab marker), then escaped. */
const esc = (text: string) => escapeHtml(cleanText(text))

const UNSUPPORTED_NOTE = '網頁版未支援,請下載原始檔查看'
/** A labelled box for something the page does not draw; `note` says why (by default: the web view does not draw it). */
const unsupported = (label: string, w?: number, h?: number, note = UNSUPPORTED_NOTE) => {
  const size = w && w > 0 ? `width:${pt(clamp(w, 40, 1000))};` : ''
  const height = h && h > 0 ? `min-height:${pt(clamp(h, 0, 160))};` : ''
  return `<span class="dx-unsup" style="${size}${height}">${esc(label)}(${esc(note)})</span>`
}
const MISSING_PICTURE = ['圖片遺失', '檔案裡找不到這張圖片'] as const
const EXTERNAL_PICTURE = ['外部連結的圖片', '圖片在文件之外,網頁版不載入外部連結'] as const

const hasHan = (text: string) => /[\u3400-\u9fff]/.test(text)
/** Why something could not be read, in words for the reader: the parser's own message only when it is Chinese. */
function reason(error: unknown): string {
  const message = error instanceof Error ? error.message : String(error)
  if (hasHan(message)) return message
  if (error instanceof XmlError) return 'XML 格式有誤'
  if (error instanceof ZipError) return '壓縮檔內容損毀'
  if (error instanceof RangeError) return '內容層次太深或太大'
  return '內容無法解析'
}

const textOf = (node: XmlNode | null | undefined): string => {
  if (!node) return ''
  let out = node.text
  for (const next of node.children) out += textOf(next)
  return out
}

function parseHyperlinkInstr(instr: string): Link {
  const tokens = instr.match(/\\[a-zA-Z]|"[^"]*"|\S+/g) ?? []
  if (!/^HYPERLINK$/i.test(tokens[0] ?? '')) return null
  let url: string | undefined
  let anchor: string | undefined
  for (let at = 1; at < tokens.length; at++) {
    const token = tokens[at]
    const value = (next: string | undefined) => (next ?? '').replace(/^"|"$/g, '')
    if (token === '\\l') anchor = value(tokens[++at])
    else if (token === '\\o' || token === '\\t' || token === '\\s') at++
    else if (!token.startsWith('\\') && url === undefined) url = value(token)
  }
  if (anchor) return { href: `#${anchorId(anchor)}` }
  return url ? { title: url } : null
}

// ---------------------------------------------------------------------------------------------
// Embedded HTML (w:altChunk): html-docx-js and other generators put the whole body in one HTML or
// MHT part. It is untrusted markup, so it is rebuilt through a whitelist: text and structure only.

const CHUNK_LIMITS = { bytes: 16 * 1024 * 1024, tags: 50_000, image: 8 * 1024 * 1024 }

const ENTITIES: Record<string, string> = {
  amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: '\u00a0', copy: '©', reg: '®', trade: '™', hellip: '…', mdash: '—', ndash: '–',
  lsquo: '‘', rsquo: '’', ldquo: '“', rdquo: '”', sbquo: '‚', bdquo: '„', bull: '•', middot: '·', times: '×', divide: '÷', deg: '°',
  plusmn: '±', euro: '€', pound: '£', yen: '¥', cent: '¢', sect: '§', para: '¶', laquo: '«', raquo: '»', shy: '\u00ad', ensp: '\u2002',
  emsp: '\u2003', thinsp: '\u2009', zwnj: '\u200c', zwj: '\u200d', iexcl: '¡', iquest: '¿', frac12: '½', frac14: '¼', frac34: '¾',
  sup1: '¹', sup2: '²', sup3: '³', micro: 'µ', larr: '←', rarr: '→', uarr: '↑', darr: '↓', harr: '↔', le: '≤', ge: '≥', ne: '≠',
  infin: '∞', minus: '−', prime: '′', Prime: '″', dagger: '†', Dagger: '‡', permil: '‰', ordf: 'ª', ordm: 'º', not: '¬', macr: '¯',
  acute: '´', cedil: '¸', uml: '¨', szlig: 'ß', agrave: 'à', aacute: 'á', acirc: 'â', atilde: 'ã', auml: 'ä', aring: 'å', aelig: 'æ',
  ccedil: 'ç', egrave: 'è', eacute: 'é', ecirc: 'ê', euml: 'ë', igrave: 'ì', iacute: 'í', icirc: 'î', iuml: 'ï', ntilde: 'ñ', ograve: 'ò',
  oacute: 'ó', ocirc: 'ô', otilde: 'õ', ouml: 'ö', oslash: 'ø', ugrave: 'ù', uacute: 'ú', ucirc: 'û', uuml: 'ü', yacute: 'ý', yuml: 'ÿ',
  Agrave: 'À', Aacute: 'Á', Acirc: 'Â', Atilde: 'Ã', Auml: 'Ä', Aring: 'Å', AElig: 'Æ', Ccedil: 'Ç', Egrave: 'È', Eacute: 'É', Ecirc: 'Ê',
  Euml: 'Ë', Ntilde: 'Ñ', Ograve: 'Ò', Oacute: 'Ó', Ocirc: 'Ô', Otilde: 'Õ', Ouml: 'Ö', Oslash: 'Ø', Ugrave: 'Ù', Uacute: 'Ú', Ucirc: 'Û', Uuml: 'Ü'
}

const decodeEntities = (text: string) =>
  text.replace(/&(#[xX][0-9a-fA-F]{1,6}|#\d{1,7}|[A-Za-z][A-Za-z0-9]{1,31});?/g, (whole, body: string) => {
    if (body[0] === '#') {
      const code = body[1] === 'x' || body[1] === 'X' ? parseInt(body.slice(2), 16) : parseInt(body.slice(1), 10)
      return code > 0 && code < 0x110000 && !(code >= 0xd800 && code <= 0xdfff) ? String.fromCodePoint(code) : '\ufffd'
    }
    return ENTITIES[body] ?? whole
  })

/** Tags kept as they are (attributes dropped, except colspan/rowspan); a few spellings map onto them. */
const CHUNK_TAGS = new Set(['p', 'br', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'b', 'strong', 'i', 'em', 'u', 's', 'sub', 'sup', 'ul', 'ol', 'li', 'table', 'thead', 'tbody', 'tr', 'th', 'td', 'blockquote', 'pre', 'code'])
const CHUNK_ALIAS: Record<string, string> = { strike: 's', del: 's' }
/** Raw-text elements: their content is not markup and is never shown. */
const CHUNK_RAW = new Set(['script', 'style', 'textarea', 'title', 'xmp', 'iframe', 'noembed', 'noframes', 'noscript', 'plaintext'])
/** Elements dropped with everything inside them. */
const CHUNK_DROP = new Set(['head', 'template', 'svg', 'math', 'select', 'object', 'video', 'audio', 'canvas', 'map', 'datalist'])
/** Dropped block containers: their text stays, on lines of its own. */
const CHUNK_BLOCKS = new Set(['div', 'section', 'article', 'header', 'footer', 'main', 'nav', 'aside', 'figure', 'figcaption', 'address', 'center', 'dl', 'dt', 'dd', 'hr', 'caption', 'form', 'fieldset', 'legend', 'details', 'summary', 'body', 'html', 'tfoot', 'colgroup', 'col', 'option', 'label'])
const CHUNK_PARA_CLOSERS = new Set(['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol', 'table', 'blockquote', 'pre'])
const CHUNK_IMAGE = /^data:image\/(png|jpe?g|gif|webp|bmp|svg\+xml);base64,[A-Za-z0-9+/=\s]*$/i
const CHUNK_LINK = /^(https?|mailto|ftp):/i

/**
 * HTML rebuilt from scratch through a whitelist: only CHUNK_TAGS, links as text with their address as
 * a title, pictures only from data: URIs of image types (`resolve` turns an MHT part's address into
 * one). Every other tag, attribute, script, style and handler is dropped; text is escaped; the result
 * is balanced, so nothing in it can reach outside the box it is put in.
 */
function sanitizeChunk(source: string, resolve?: (src: string) => string | undefined): { html: string; text: string; cut: boolean } {
  let out = ''
  const text: string[] = []
  const stack: { name: string; close: string }[] = []
  let at = 0
  let tags = 0
  let drop: string | null = null
  let dropDepth = 0
  let lineHasContent = false
  let pendingBreak = false
  const has = (name: string, stopAt: string[]) => {
    for (let i = stack.length - 1; i >= 0; i--) {
      if (stack[i].name === name) return i
      if (stopAt.includes(stack[i].name)) return -1
    }
    return -1
  }
  const closeTo = (index: number) => {
    while (stack.length > index) out += stack.pop()!.close
  }
  const boundary = () => {
    if (lineHasContent) pendingBreak = true
  }
  const block = () => {
    pendingBreak = false
    lineHasContent = false
  }
  const content = () => {
    if (pendingBreak) out += '<br>'
    pendingBreak = false
    lineHasContent = true
  }
  const open = (name: string, html: string, close: string) => {
    out += html
    stack.push({ name, close })
  }
  while (at < source.length) {
    const lt = source.indexOf('<', at)
    const chunk = lt < 0 ? source.slice(at) : source.slice(at, lt)
    if (chunk && !drop) {
      const decoded = decodeEntities(chunk)
      if (decoded.trim()) content()
      out += esc(decoded)
      text.push(decoded)
    }
    if (lt < 0) break
    at = lt
    if (source.startsWith('<!--', at)) {
      const end = source.indexOf('-->', at + 4)
      at = end < 0 ? source.length : end + 3
      continue
    }
    const next = source[at + 1] ?? ''
    if (next === '!' || next === '?') {
      const end = source.indexOf('>', at)
      at = end < 0 ? source.length : end + 1
      continue
    }
    const closing = next === '/'
    const tagMatch = /^[A-Za-z][^\s/>]*/.exec(source.slice(at + (closing ? 2 : 1), at + (closing ? 2 : 1) + 64))
    if (!tagMatch) {
      if (!drop) {
        content()
        out += '&lt;'
        text.push('<')
      }
      at++
      continue
    }
    // Attributes: name, then an optional value in double, single or no quotes.
    let cursor = at + (closing ? 2 : 1) + tagMatch[0].length
    const attrs: Record<string, string> = {}
    const attrPattern = /\s*([^\s/>=][^\s/>=]*)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]*)))?|\s*\/(?!>)|\s+/y
    while (cursor < source.length && source[cursor] !== '>' && !source.startsWith('/>', cursor)) {
      attrPattern.lastIndex = cursor
      const m = attrPattern.exec(source)
      if (!m || m[0] === '') {
        cursor++
        continue
      }
      if (m[1] && Object.keys(attrs).length < 100) attrs[m[1].toLowerCase()] ??= decodeEntities(m[2] ?? m[3] ?? m[4] ?? '')
      cursor = attrPattern.lastIndex
    }
    if (cursor >= source.length) break
    at = cursor + (source[cursor] === '>' ? 1 : 2)
    if (++tags > CHUNK_LIMITS.tags) {
      closeTo(0)
      return { html: out, text: text.join(''), cut: true }
    }
    const raw = tagMatch[0].toLowerCase()
    if (!closing && CHUNK_RAW.has(raw)) {
      const end = new RegExp(`</${raw}`, 'ig')
      end.lastIndex = at
      const found = end.exec(source)
      if (!found) break
      const after = source.indexOf('>', found.index)
      at = after < 0 ? source.length : after + 1
      continue
    }
    if (drop) {
      if (raw === drop) dropDepth += closing ? -1 : 1
      if (dropDepth <= 0) drop = null
      continue
    }
    if (!closing && CHUNK_DROP.has(raw)) {
      drop = raw
      dropDepth = 1
      continue
    }
    if (CHUNK_BLOCKS.has(raw)) {
      boundary()
      continue
    }
    const name = CHUNK_ALIAS[raw] ?? raw
    if (name === 'a') {
      if (closing) {
        const index = has('a', ['td', 'th', 'table'])
        if (index >= 0) closeTo(index)
      } else {
        const href = (attrs.href ?? '').trim()
        open('a', CHUNK_LINK.test(href) ? `<span class="dx-lnk" title="${esc(href.slice(0, 2000))}">` : '<span class="dx-lnk">', '</span>')
      }
      continue
    }
    if (name === 'img') {
      if (closing) continue
      const src = (attrs.src ?? '').trim()
      const url = (CHUNK_IMAGE.test(src) ? src : resolve?.(src)) ?? ''
      if (!CHUNK_IMAGE.test(url) || url.length > CHUNK_LIMITS.image * 1.4) continue
      const size = (key: string) => (/^\d{1,5}$/.test(attrs[key] ?? '') ? ` ${key}="${clamp(Number(attrs[key]), 1, 4000)}"` : '')
      content()
      out += `<img class="dx-img" src="${esc(url.replace(/\s+/g, ''))}" alt="${esc((attrs.alt ?? '').slice(0, 500))}"${size('width')}${size('height')}>`
      continue
    }
    if (!CHUNK_TAGS.has(name)) continue
    if (name === 'br') {
      out += '<br>'
      pendingBreak = false
      lineHasContent = false
      continue
    }
    if (closing) {
      // As in HTML, an end tag does not reach past a cell or table it is not inside of.
      const index = has(name, name === 'table' ? [] : name === 'td' || name === 'th' || name === 'tr' || name === 'thead' || name === 'tbody' ? ['table'] : ['td', 'th', 'table'])
      if (index >= 0) closeTo(index)
      if (CHUNK_PARA_CLOSERS.has(name) || ['li', 'tr', 'td', 'th'].includes(name)) block()
      continue
    }
    // Opening a block closes what HTML's own rules would close.
    if (CHUNK_PARA_CLOSERS.has(name) || name === 'li') {
      const p = has('p', ['li', 'td', 'th', 'blockquote', 'table'])
      if (p >= 0) closeTo(p)
    }
    if (/^h[1-6]$/.test(name)) {
      for (let i = stack.length - 1; i >= 0 && !['li', 'td', 'th', 'blockquote', 'table'].includes(stack[i].name); i--) {
        if (/^h[1-6]$/.test(stack[i].name)) {
          closeTo(i)
          break
        }
      }
    }
    if (name === 'li') {
      const li = has('li', ['ul', 'ol'])
      if (li >= 0) closeTo(li)
    } else if (name === 'tr') {
      const tr = has('tr', ['table', 'thead', 'tbody'])
      if (tr >= 0) closeTo(tr)
    } else if (name === 'td' || name === 'th') {
      const cell = Math.max(has('td', ['tr', 'table']), has('th', ['tr', 'table']))
      if (cell >= 0) closeTo(cell)
    } else if (name === 'thead' || name === 'tbody') {
      const section = Math.max(has('thead', ['table']), has('tbody', ['table']))
      if (section >= 0) closeTo(section)
    }
    let attributes = ''
    if (name === 'td' || name === 'th') {
      for (const key of ['colspan', 'rowspan']) {
        const value = attrs[key]
        if (value && /^\s*\d{1,4}\s*$/.test(value) && Number(value) > 1) attributes += ` ${key}="${clamp(Number(value), 1, 1000)}"`
      }
    }
    if (CHUNK_PARA_CLOSERS.has(name) || ['li', 'tr', 'td', 'th', 'thead', 'tbody'].includes(name)) block()
    open(name, `<${name}${attributes}>`, `</${name}>`)
  }
  closeTo(0)
  return { html: out, text: text.join(''), cut: false }
}

/** Text from bytes: a BOM, else the declared charset, else a `<meta charset>`, else UTF-8. */
function decodeChunk(bytes: Buffer, charset?: string): string {
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) return bytes.subarray(3).toString('utf8')
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes.subarray(2))
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes.subarray(2))
  const label = charset || /<meta[^>]+charset\s*=\s*["']?\s*([\w.:-]+)/i.exec(bytes.subarray(0, 4096).toString('latin1'))?.[1]
  try {
    return new TextDecoder(label || 'utf-8').decode(bytes)
  } catch {
    return bytes.toString('utf8')
  }
}

type MimePart = { headers: Map<string, string>; body: string }

function mimePart(block: string): MimePart {
  const split = /\r?\n\r?\n/.exec(block)
  const head = split ? block.slice(0, split.index) : block
  const headers = new Map<string, string>()
  for (const line of head.replace(/\r?\n[ \t]+/g, ' ').split(/\r?\n/)) {
    const colon = line.indexOf(':')
    if (colon > 0) headers.set(line.slice(0, colon).trim().toLowerCase(), line.slice(colon + 1).trim())
  }
  return { headers, body: split ? block.slice(split.index + split[0].length) : '' }
}

const mimeParam = (value: string | undefined, name: string) => new RegExp(`(?:^|;)\\s*${name}\\s*=\\s*(?:"([^"]*)"|([^\\s;]+))`, 'i').exec(value ?? '')?.slice(1).find((v) => v !== undefined)

function mimeBody(part: MimePart): Buffer {
  const encoding = (part.headers.get('content-transfer-encoding') ?? '').toLowerCase()
  if (encoding === 'base64') return Buffer.from(part.body.replace(/[^A-Za-z0-9+/=]/g, ''), 'base64')
  if (encoding === 'quoted-printable') {
    const bytes = part.body.replace(/=\r?\n/g, '').replace(/=([0-9A-Fa-f]{2})/g, (_, hex: string) => String.fromCharCode(parseInt(hex, 16)))
    return Buffer.from(bytes, 'latin1')
  }
  return Buffer.from(part.body, 'latin1')
}

/** An MHT (MIME web archive) part: its HTML, and its pictures as data: URIs looked up by address. */
function readMht(bytes: Buffer): { html: string; resolve: (src: string) => string | undefined } | null {
  const top = mimePart(bytes.toString('latin1'))
  const boundary = mimeParam(top.headers.get('content-type'), 'boundary')
  const parts: MimePart[] = []
  if (!boundary) parts.push(top)
  else {
    for (const block of top.body.split(`--${boundary}`).slice(1)) {
      if (block.startsWith('--')) break
      parts.push(mimePart(block.replace(/^[ \t]*\r?\n/, '')))
    }
  }
  const typeOf = (part: MimePart) => (part.headers.get('content-type') ?? 'text/plain').split(';')[0].trim().toLowerCase()
  const page = parts.find((part) => typeOf(part) === 'text/html')
  if (!page) return null
  const images = new Map<string, string>()
  for (const part of parts) {
    const type = typeOf(part)
    if (!/^image\/(png|jpe?g|gif|webp|bmp|svg\+xml)$/.test(type) || part.body.length > CHUNK_LIMITS.image * 1.4) continue
    const url = `data:${type};base64,${mimeBody(part).toString('base64')}`
    const location = part.headers.get('content-location')
    const id = part.headers.get('content-id')?.replace(/^<|>$/g, '')
    if (location) images.set(location, url)
    if (id) images.set(`cid:${id}`, url)
  }
  const resolve = (src: string) => {
    if (images.has(src)) return images.get(src)
    const base = src.split(/[\\/]/).pop()
    const matches = [...images.keys()].filter((key) => key.split(/[\\/]/).pop() === base)
    return matches.length === 1 ? images.get(matches[0]) : undefined
  }
  return { html: decodeChunk(mimeBody(page), mimeParam(page.headers.get('content-type'), 'charset')), resolve }
}

// ---------------------------------------------------------------------------------------------
// The renderer

type FontInfo = { alt?: string; kind?: FontKind }

class DocxRenderer {
  private readonly pkg: Package
  private readonly media: (part: string) => string
  readonly notices: string[] = []
  readonly text: string[] = []
  private readonly classes = new Map<string, string>()
  private readonly fontStacks = new Map<string, string>()
  private mainPart = 'word/document.xml'
  private theme = { colors: {} as Record<string, string>, major: {} as { latin?: string; ea?: string; cs?: string }, minor: {} as { latin?: string; ea?: string; cs?: string } }
  private themeSlots: Record<string, string> = { ...WORD_THEME_SLOTS }
  private styles = new Map<string, StyleDef>()
  private defaultStyle: Record<string, string> = {}
  private rPrDefault: XmlNode | null = null
  private pPrDefault: XmlNode | null = null
  private readonly chains = new Map<string, StyleDef[]>()
  private abstracts = new Map<string, Level[]>()
  private abstractLinks = new Map<string, string>()
  private nums = new Map<string, NumDef>()
  private readonly counters = new Map<string, (number | undefined)[]>()
  private readonly usedNums = new Set<string>()
  private fonts = new Map<string, FontInfo>()
  defaultTab = 36
  private noteFormats: Record<NoteKind, string> = { footnote: 'decimal', endnote: 'lowerRoman' }
  private readonly noteParts: Record<NoteKind, { part: string; notes: Map<string, XmlNode> } | null> = { footnote: null, endnote: null }
  private readonly noteNumbers: Record<NoteKind, Map<string, number>> = { footnote: new Map(), endnote: new Map() }
  private blocks = 0
  private cells = 0
  private cutBlocks = false
  private cutCells = false
  /** Paragraphs and tables inside table cells: they have their own budget, so a long table cannot
   *  use up the document's and hide what follows it, and a table is only ever cut between rows. */
  private cellBlocks = 0
  private cutCellBlocks = false
  private inCell = 0

  constructor(data: Buffer, mediaBase: string) {
    this.pkg = new Package(data)
    this.media = (part) => officeMediaUrl(mediaBase, part)
  }

  /** A part that cannot be read leaves a note on the page; the document draws without it. */
  private optional<T>(label: string, read: () => T): T | null {
    try {
      return read()
    } catch (error) {
      this.notices.push(`${label}無法讀取,已略過:${reason(error)}`)
      return null
    }
  }

  cls(css: string): string {
    if (!css) return ''
    let name = this.classes.get(css)
    if (!name) {
      name = `s${this.classes.size}`
      this.classes.set(css, name)
    }
    return name
  }

  css(): string {
    let out = ''
    for (const [css, name] of this.classes) out += `.${name}{${css}}`
    return out
  }

  // --- parts ---------------------------------------------------------------------------------

  read(): { html: string; title: string } {
    this.mainPart = this.pkg.relOfType('', 'officeDocument')?.target ?? 'word/document.xml'
    const document = this.pkg.xml(this.mainPart)
    if (!document || local(document.name) !== 'document') throw new Error('這不是 Word 文件(找不到 word/document.xml)')
    const body = child(document, 'body')
    if (!body) throw new Error('文件沒有內文 (w:body)')
    this.optional('設定 (settings.xml)', () => this.readSettings())
    this.optional('佈景主題 (theme)', () => this.readTheme())
    this.optional('樣式 (styles.xml)', () => this.readStyles())
    this.optional('編號 (numbering.xml)', () => this.readNumbering())
    this.optional('字型表 (fontTable.xml)', () => this.readFontTable())
    for (const kind of ['footnote', 'endnote'] as NoteKind[]) {
      this.optional(kind === 'footnote' ? '註腳 (footnotes.xml)' : '章節附註 (endnotes.xml)', () => this.readNotes(kind))
    }
    const core = this.optional('文件摘要', () => this.pkg.xml('docProps/core.xml'))
    const title = core ? textOf(children(core, 'title')[0]).trim() : ''

    // Paper colour and fill image from w:background.
    const background = child(document, 'background')
    let paper: Rgba = WHITE
    let paperImage = ''
    if (background) {
      const color = this.wordColor(background)
      if (color && color !== 'auto') paper = color
      const fill = path(background, 'background', 'fill')
      const rel = this.pkg.rel(this.mainPart, relAttr(fill, 'id'))
      if (rel && !rel.external && isDrawablePart(rel.target)) {
        paperImage = ` url('${this.media(rel.target)}') ${attr(fill, 'type') === 'tile' ? 'repeat' : 'center/cover no-repeat'}`
      }
    }

    // Sections: a paragraph carrying w:sectPr ends one; the body's own w:sectPr is the last.
    const sections: { sectPr: XmlNode | null; blocks: XmlNode[] }[] = []
    let current: XmlNode[] = []
    for (const node of body.children) {
      if (local(node.name) === 'sectPr') continue
      current.push(node)
      const sectPr = local(node.name) === 'p' ? path(node, 'pPr', 'sectPr') : null
      if (sectPr) {
        sections.push({ sectPr, blocks: current })
        current = []
      }
    }
    sections.push({ sectPr: child(body, 'sectPr'), blocks: current })
    const sheets: { geometry: Geometry; parts: { geometry: Geometry; blocks: XmlNode[] }[] }[] = []
    for (const section of sections) {
      const geometry = this.geometry(section.sectPr)
      const last = sheets[sheets.length - 1]
      if (!section.blocks.length && last) continue
      if (last && geometry.type === 'continuous' && Math.abs(last.geometry.w - geometry.w) < 1) last.parts.push({ geometry, blocks: section.blocks })
      else sheets.push({ geometry, parts: [{ geometry, blocks: section.blocks }] })
    }
    // Header and footer: the first section that names one (a cover section often has none).
    const firstWith = (kind: string) => sections.find((section) => children(section.sectPr, `${kind}Reference`).length)?.sectPr ?? null
    const firstGeometry = sheets[0].geometry
    const dark = isDark(paper)

    const sheetHtml: string[] = []
    sheets.forEach((sheet, at) => {
      this.page = { index: 0, anchored: false, bottom: 0 }
      this.placedOnSheet = false
      let inner = ''
      if (at === 0) inner += this.headerFooter(firstWith('header'), 'header', firstGeometry, dark)
      for (const part of sheet.parts) {
        const g = part.geometry
        const scope: Scope = {
          story: { part: this.mainPart, kind: 'body', fields: [] },
          layer: EMPTY_LAYER,
          depth: 0,
          width: Math.max(72, g.w - g.left - g.right),
          dark,
          grid: g.grid,
          sheet: g
        }
        let blocks = this.renderBlocks(part.blocks, scope)
        if (g.cols > 1) blocks = `<div class="dx-cols" style="column-count:${g.cols};column-gap:${pt(g.colSpace)}">${blocks}</div>`
        if (g.vertical) blocks = `<div class="dx-vert" style="height:${pt(Math.max(72, g.h - g.top - g.bottom))}">${blocks}</div>`
        inner += `<div class="dx-body">${blocks}</div>`
      }
      if (at === sheets.length - 1) {
        inner += this.renderNotes(sheet.geometry, dark)
        inner += this.headerFooter(firstWith('footer'), 'footer', firstGeometry, dark)
      }
      const g = sheet.geometry
      const vars = `--pw:${pt(g.w)};--mt:${pt(g.top)};--mr:${pt(g.right)};--mb:${pt(g.bottom)};--ml:${pt(g.left)};--hd:${pt(Math.min(g.header, g.top))};--fd:${pt(Math.min(g.footer, g.bottom))};--paper:${cssColor(paper)}`
      // A sheet with placed drawings keeps its margins in proportion on a phone, so text stays where the art expects it.
      const share = (value: number) => `${round((value / g.w) * 100, 3)}cqw`
      const placed = this.placedOnSheet ? `;--plt:${share(g.top)};--plr:${share(g.right)};--plb:${share(g.bottom)};--pll:${share(g.left)}` : ''
      const style = `${vars}${placed};aspect-ratio:${round(g.w)}/${round(g.h)};background:${cssColor(paper)}${paperImage}`
      sheetHtml.push(`<div class="dx-sheet${dark ? ' dx-dark' : ''}${placed ? ' dx-placed' : ''}" style="${esc(style)}"><div class="dx-in">${inner}</div></div>`)
    })
    if (this.cutBlocks) this.notices.push(`文件太長:只顯示前 ${LIMITS.blocks} 個段落與表格,其餘請下載原始檔查看。`)
    if (this.cutCells) this.notices.push(`表格太大:只顯示前 ${LIMITS.cells} 個儲存格,其餘請下載原始檔查看。`)
    if (this.cutCellBlocks) this.notices.push(`表格內容太多:只顯示表格裡前 ${LIMITS.cellBlocks} 個段落,其餘請下載原始檔查看。`)
    const notices = this.notices.map((notice) => `<p class="dx-notice">${esc(notice)}</p>`).join('')
    return { html: notices + sheetHtml.join(''), title }
  }

  private readSettings(): void {
    const rel = this.pkg.relOfType(this.mainPart, 'settings')
    const settings = rel ? this.pkg.xml(rel.target) : null
    if (!settings) return
    const tab = num(child(settings, 'defaultTabStop'), 'val')
    if (tab && tab > 0) this.defaultTab = clamp(tab / 20, 6, 200)
    const mapping = child(settings, 'clrSchemeMapping')
    if (mapping) {
      const slot = (value: string | undefined) => (value ? WORD_THEME_SLOTS[value] : undefined)
      const pairs: [string, string][] = [['bg1', 'background1'], ['t1', 'text1'], ['bg2', 'background2'], ['t2', 'text2']]
      for (const [key, name] of pairs) {
        const target = slot(attr(mapping, key))
        if (target) this.themeSlots[name] = target
      }
    }
    for (const kind of ['footnote', 'endnote'] as NoteKind[]) {
      const format = attr(path(settings, `${kind}Pr`, 'numFmt'), 'val')
      if (format) this.noteFormats[kind] = format
    }
    this.themeLang = attr(child(settings, 'themeFontLang'), 'eastAsia') ?? ''
  }

  private themeLang = ''

  private readTheme(): void {
    const rel = this.pkg.relOfType(this.mainPart, 'theme')
    const xml = rel ? this.pkg.xml(rel.target) : null
    const elements = path(xml, 'themeElements')
    if (!elements) return
    for (const slot of child(elements, 'clrScheme')?.children ?? []) {
      const value = slot.children[0]
      if (!value) continue
      const name = local(value.name)
      const hex = name === 'srgbClr' ? attr(value, 'val') : name === 'sysClr' ? (attr(value, 'lastClr') ?? (attr(value, 'val') === 'window' ? 'FFFFFF' : '000000')) : undefined
      if (hex && /^[0-9a-fA-F]{6}$/.test(hex)) this.theme.colors[local(slot.name)] = hex
    }
    const lang = this.themeLang.toLowerCase()
    const scripts = lang.startsWith('zh-cn') || lang.startsWith('zh-sg') ? ['Hans', 'Hant'] : lang.startsWith('ja') ? ['Jpan', 'Hant'] : lang.startsWith('ko') ? ['Hang', 'Hant'] : ['Hant', 'Hans', 'Jpan']
    const face = (node: XmlNode | null) => {
      const scriptFont = (script: string) => children(node, 'font').find((font) => attr(font, 'script') === script)
      const ea = attr(child(node, 'ea'), 'typeface') || scripts.map((script) => attr(scriptFont(script), 'typeface')).find(Boolean)
      return { latin: attr(child(node, 'latin'), 'typeface') || undefined, ea: ea || undefined, cs: attr(child(node, 'cs'), 'typeface') || undefined }
    }
    const fonts = child(elements, 'fontScheme')
    this.theme.major = face(child(fonts, 'majorFont'))
    this.theme.minor = face(child(fonts, 'minorFont'))
  }

  private readStyles(): void {
    const rel = this.pkg.relOfType(this.mainPart, 'styles')
    const xml = rel ? this.pkg.xml(rel.target) : null
    if (!xml) return
    const defaults = child(xml, 'docDefaults')
    this.rPrDefault = path(defaults, 'rPrDefault', 'rPr')
    this.pPrDefault = path(defaults, 'pPrDefault', 'pPr')
    for (const style of children(xml, 'style')) {
      const id = attr(style, 'styleId')
      if (!id) continue
      const type = attr(style, 'type') ?? 'paragraph'
      const conds = new Map<string, StyleCond>()
      for (const cond of children(style, 'tblStylePr')) {
        const condType = attr(cond, 'type')
        if (condType) conds.set(condType, { pPr: child(cond, 'pPr'), rPr: child(cond, 'rPr'), tblPr: child(cond, 'tblPr'), tcPr: child(cond, 'tcPr') })
      }
      this.styles.set(id, {
        id, type, basedOn: attr(child(style, 'basedOn'), 'val'), conds,
        pPr: child(style, 'pPr'), rPr: child(style, 'rPr'), tblPr: child(style, 'tblPr'), tcPr: child(style, 'tcPr')
      })
      const isDefault = attr(style, 'default')
      if ((isDefault === '1' || isDefault === 'true' || isDefault === 'on') && !this.defaultStyle[type]) this.defaultStyle[type] = id
    }
  }

  /** A style and everything it is based on, base first. Cycles and absurd depths are cut. */
  private chain(id: string | undefined, type: string): StyleDef[] {
    const start = id ?? this.defaultStyle[type]
    if (!start) return []
    const key = `${type}:${start}`
    const cached = this.chains.get(key)
    if (cached) return cached
    const out: StyleDef[] = []
    const seen = new Set<string>()
    let current = this.styles.get(start) ?? (id ? this.styles.get(this.defaultStyle[type] ?? '') : undefined)
    while (current && !seen.has(current.id) && out.length < LIMITS.chain) {
      seen.add(current.id)
      out.unshift(current)
      current = current.basedOn ? this.styles.get(current.basedOn) : undefined
    }
    this.chains.set(key, out)
    return out
  }

  private readNumbering(): void {
    const rel = this.pkg.relOfType(this.mainPart, 'numbering')
    const xml = rel ? this.pkg.xml(rel.target) : null
    if (!xml) return
    const levelOf = (lvl: XmlNode): Level => ({
      start: listStart(num(child(lvl, 'start'), 'val')) ?? 0,
      fmt: attr(child(lvl, 'numFmt'), 'val') ?? 'decimal',
      text: attr(child(lvl, 'lvlText'), 'val') ?? '',
      pPr: child(lvl, 'pPr'),
      rPr: child(lvl, 'rPr'),
      suff: attr(child(lvl, 'suff'), 'val') ?? 'tab',
      isLgl: Boolean(child(lvl, 'isLgl')) && truthy(child(lvl, 'isLgl')),
      restart: num(child(lvl, 'lvlRestart'), 'val')
    })
    for (const abstract of children(xml, 'abstractNum')) {
      const id = attr(abstract, 'abstractNumId')
      if (id === undefined) continue
      const levels: Level[] = []
      for (const lvl of children(abstract, 'lvl')) {
        const ilvl = num(lvl, 'ilvl') ?? 0
        if (ilvl >= 0 && ilvl < 9) levels[ilvl] = levelOf(lvl)
      }
      this.abstracts.set(id, levels)
      const link = attr(child(abstract, 'numStyleLink'), 'val')
      if (link) this.abstractLinks.set(id, link)
    }
    for (const numNode of children(xml, 'num')) {
      const id = attr(numNode, 'numId')
      const abstractId = attr(child(numNode, 'abstractNumId'), 'val')
      if (id === undefined || abstractId === undefined) continue
      const overrides = new Map<number, { start?: number; level?: Level }>()
      for (const override of children(numNode, 'lvlOverride')) {
        const ilvl = num(override, 'ilvl') ?? 0
        if (!(ilvl >= 0 && ilvl < 9)) continue
        const lvl = child(override, 'lvl')
        overrides.set(ilvl, { start: listStart(num(child(override, 'startOverride'), 'val')), level: lvl ? levelOf(lvl) : undefined })
      }
      this.nums.set(id, { abstractId, overrides })
    }
  }

  /** The abstract numbering a num draws from, following numStyleLink through numbering styles. */
  private abstractOf(numId: string): string | undefined {
    let abstractId = this.nums.get(numId)?.abstractId
    for (let hops = 0; abstractId !== undefined && hops < 5; hops++) {
      const link = this.abstractLinks.get(abstractId)
      if (!link) break
      const linked = attr(path(this.styles.get(link)?.pPr, 'numPr', 'numId'), 'val')
      const next = linked ? this.nums.get(linked)?.abstractId : undefined
      if (!next || next === abstractId) break
      abstractId = next
    }
    return abstractId
  }

  private level(numId: string, ilvl: number): Level | null {
    const def = this.nums.get(numId)
    if (!def) return null
    const override = def.overrides.get(ilvl)
    const abstractId = this.abstractOf(numId)
    const base = override?.level ?? (abstractId !== undefined ? this.abstracts.get(abstractId)?.[ilvl] : undefined)
    if (!base) return null
    return override?.start !== undefined ? { ...base, start: override.start } : base
  }

  /** Advances the counters for one numbered paragraph and returns its label. */
  private nextLabel(numId: string, ilvl: number): string {
    const def = this.nums.get(numId)
    const abstractId = this.abstractOf(numId)
    const level = this.level(numId, ilvl)
    if (!def || abstractId === undefined || !level) return ''
    let counters = this.counters.get(abstractId)
    if (!counters) {
      counters = []
      this.counters.set(abstractId, counters)
    }
    if (!this.usedNums.has(numId)) {
      this.usedNums.add(numId)
      // A num that overrides a start restarts its list there, though it shares the abstract list.
      for (const [at, override] of def.overrides) {
        if (override.start !== undefined || override.level) for (let deeper = at; deeper < 9; deeper++) counters[deeper] = undefined
      }
    }
    counters[ilvl] = counters[ilvl] === undefined ? level.start : counters[ilvl]! + 1
    for (let deeper = ilvl + 1; deeper < 9; deeper++) {
      const restart = this.level(numId, deeper)?.restart
      if (restart === 0 || (restart !== undefined && restart - 1 < ilvl)) continue
      counters[deeper] = undefined
    }
    if (level.fmt === 'bullet') return level.text
    return level.text.replace(/%([1-9])/g, (_, k: string) => {
      const at = Number(k) - 1
      const at_level = this.level(numId, at)
      const value = counters![at] ?? at_level?.start ?? 1
      return formatNumber(value, level.isLgl && at < ilvl ? 'decimal' : (at_level?.fmt ?? 'decimal'))
    })
  }

  private readFontTable(): void {
    const rel = this.pkg.relOfType(this.mainPart, 'fontTable')
    const xml = rel ? this.pkg.xml(rel.target) : null
    for (const font of children(xml, 'font')) {
      const name = attr(font, 'name')
      if (!name) continue
      const family = attr(child(font, 'family'), 'val')
      const kind: FontKind | undefined = family === 'roman' ? 'serif' : family === 'modern' ? 'mono' : family === 'swiss' ? 'sans' : undefined
      this.fonts.set(name, { alt: attr(child(font, 'altName'), 'val'), kind })
    }
  }

  private readNotes(kind: NoteKind): void {
    const rel = this.pkg.relOfType(this.mainPart, `${kind}s`)
    const xml = rel ? this.pkg.xml(rel.target) : null
    if (!xml || !rel) return
    const notes = new Map<string, XmlNode>()
    for (const note of children(xml, kind)) {
      const type = attr(note, 'type')
      const id = attr(note, 'id')
      if (id !== undefined && (!type || type === 'normal')) notes.set(id, note)
    }
    this.noteParts[kind] = { part: rel.target, notes }
  }

  private geometry(sectPr: XmlNode | null): Geometry {
    const size = child(sectPr, 'pgSz')
    const margin = child(sectPr, 'pgMar')
    const twips = (node: XmlNode | null, name: string, fallback: number) => {
      const value = num(node, name)
      return value === undefined ? fallback / 20 : Math.abs(value) / 20
    }
    let w = twips(size, 'w', 12240)
    let h = twips(size, 'h', 15840)
    if (attr(size, 'orient') === 'landscape' && w < h) [w, h] = [h, w]
    w = clamp(w, 144, 4000)
    h = clamp(h, 144, 8000)
    const cols = child(sectPr, 'cols')
    const grid = child(sectPr, 'docGrid')
    const gridType = attr(grid, 'type')
    return {
      w,
      h,
      top: clamp(twips(margin, 'top', 1440), 0, h / 3),
      bottom: clamp(twips(margin, 'bottom', 1440), 0, h / 3),
      left: clamp(twips(margin, 'left', 1440) + twips(margin, 'gutter', 0), 0, w / 3),
      right: clamp(twips(margin, 'right', 1440), 0, w / 3),
      header: twips(margin, 'header', 720),
      footer: twips(margin, 'footer', 720),
      cols: clamp(Math.floor(num(cols, 'num') ?? 1), 1, 6),
      colSpace: clamp(twips(cols, 'space', 720), 0, 200),
      grid: gridType === 'lines' || gridType === 'linesAndChars' ? clamp((num(grid, 'linePitch') ?? 0) / 20, 0, 100) : 0,
      vertical: ['tbRl', 'tbRlV', 'tb'].includes(attr(child(sectPr, 'textDirection'), 'val') ?? ''),
      type: attr(child(sectPr, 'type'), 'val') ?? 'nextPage'
    }
  }

  private headerFooter(sectPr: XmlNode | null, kind: 'header' | 'footer', geometry: Geometry, dark: boolean): string {
    const refs = children(sectPr, `${kind}Reference`)
    // The header is drawn once, where the first page has it: a section with a different first page (titlePg) shows
    // its first-page header there. The footer closes the document, so it stays the default one.
    const titlePage = kind === 'header' && child(sectPr, 'titlePg') !== null && truthy(child(sectPr, 'titlePg'))
    const ref = (titlePage ? refs.find((candidate) => attr(candidate, 'type') === 'first') : undefined) ?? refs.find((candidate) => attr(candidate, 'type') === 'default') ?? refs[0]
    const rel = ref ? this.pkg.rel(this.mainPart, relAttr(ref, 'id')) : undefined
    if (!rel || rel.external) return ''
    const label = kind === 'header' ? '頁首' : '頁尾'
    const html = this.optional(label, () => {
      const xml = this.pkg.xml(rel.target)
      if (!xml) return ''
      const scope: Scope = {
        story: { part: rel.target, kind, fields: [] }, layer: EMPTY_LAYER, depth: 0,
        width: Math.max(72, geometry.w - geometry.left - geometry.right), dark, grid: 0, sheet: geometry
      }
      return this.renderBlocks(xml.children, scope)
    })
    return html ? `<div class="dx-${kind}" aria-label="${label}">${html}</div>` : ''
  }

  private renderNotes(geometry: Geometry, dark: boolean): string {
    let out = ''
    for (const kind of ['footnote', 'endnote'] as NoteKind[]) {
      const source = this.noteParts[kind]
      const numbers = this.noteNumbers[kind]
      if (!source || !numbers.size) continue
      let html = ''
      const prefix = kind === 'footnote' ? 'fn' : 'en'
      for (const [id, n] of numbers) {
        const note = source.notes.get(id)
        if (!note) continue
        const scope: Scope = {
          story: { part: source.part, kind: 'note', fields: [], note: { kind, n } }, layer: EMPTY_LAYER, depth: 0,
          width: Math.max(72, geometry.w - geometry.left - geometry.right), dark, grid: 0, sheet: null
        }
        let body: string
        try {
          body = this.renderBlocks(note.children, scope)
        } catch {
          body = unsupported(`註解 ${n}`)
        }
        html += `<div class="dx-note" id="${prefix}-${n}">${body}</div>`
      }
      if (html) out += `<div class="dx-notes" aria-label="${kind === 'footnote' ? '註腳' : '章節附註'}">${html}</div>`
    }
    return out
  }

  // --- properties ----------------------------------------------------------------------------

  private themeColor(name: string | undefined, tint?: string, shade?: string): Rgba | null {
    if (!name) return null
    const slot = this.themeSlots[name] ?? SCHEME_SLOTS[name] ?? name
    const color = hexToRgba(this.theme.colors[slot])
    if (!color) return null
    let [h, s, l] = rgbToHsl(color)
    const tintValue = tint ? parseInt(tint, 16) : NaN
    const shadeValue = shade ? parseInt(shade, 16) : NaN
    if (Number.isFinite(tintValue)) l = l * (tintValue / 255) + (1 - tintValue / 255)
    if (Number.isFinite(shadeValue)) l = l * (shadeValue / 255)
    return Number.isFinite(tintValue) || Number.isFinite(shadeValue) ? hslToRgb(h, s, clamp01(l), 1) : color
  }

  /** w:color and friends: a theme colour when the theme has it, else the hex value; `auto` follows the background. */
  private wordColor(node: XmlNode | null, prefix = ''): Rgba | 'auto' | null {
    if (!node) return null
    const themed = this.themeColor(attr(node, `${prefix}themeColor`), attr(node, `${prefix}themeTint`), attr(node, `${prefix}themeShade`))
    if (themed) return themed
    const value = attr(node, prefix ? `${prefix}color` : 'color') ?? attr(node, 'val')
    if (!value || value === 'auto') return 'auto'
    return hexToRgba(value)
  }

  /** w:shd as one colour; null means explicitly none, undefined means not given. */
  private shading(node: XmlNode | null): Rgba | null | undefined {
    if (!node) return undefined
    const val = attr(node, 'val') ?? 'clear'
    if (val === 'nil') return null
    let fill = this.themeColor(attr(node, 'themeFill'), attr(node, 'themeFillTint'), attr(node, 'themeFillShade'))
    if (!fill) {
      const raw = attr(node, 'fill')
      fill = raw && raw !== 'auto' ? hexToRgba(raw) : null
    }
    const rawColor = attr(node, 'color')
    const color = this.themeColor(attr(node, 'themeColor'), attr(node, 'themeTint'), attr(node, 'themeShade')) ?? (rawColor && rawColor !== 'auto' ? hexToRgba(rawColor) : null)
    if (val === 'clear') return fill
    if (val === 'solid') return color ?? BLACK
    const pct = /^pct(\d+)$/.exec(val)
    if (pct) return mix(color ?? BLACK, fill ?? WHITE, clamp01(Number(pct[1]) / 100))
    return fill
  }

  private themeFont(name: string | undefined): string | undefined {
    switch (name) {
      case 'minorHAnsi':
      case 'minorAscii':
        return this.theme.minor.latin
      case 'majorHAnsi':
      case 'majorAscii':
        return this.theme.major.latin
      case 'minorEastAsia':
        return this.theme.minor.ea
      case 'majorEastAsia':
        return this.theme.major.ea
      case 'minorBidi':
        return this.theme.minor.cs
      case 'majorBidi':
        return this.theme.major.cs
      default:
        return undefined
    }
  }

  private mergeRun(base: RunProps, rPr: XmlNode | null | undefined): RunProps {
    if (!rPr) return base
    const p: RunProps = { ...base }
    for (const c of rPr.children) {
      switch (local(c.name)) {
        case 'b':
          p.b = truthy(c)
          break
        case 'i':
          p.i = truthy(c)
          break
        case 'u': {
          p.u = attr(c, 'val') ?? 'single'
          const color = attr(c, 'color') || attr(c, 'themeColor') ? this.wordColor(c) : null
          p.uColor = color && color !== 'auto' ? color : undefined
          break
        }
        case 'strike':
          p.strike = truthy(c)
          break
        case 'dstrike':
          p.dstrike = truthy(c)
          break
        case 'color': {
          const color = this.wordColor(c)
          if (color) p.color = color
          break
        }
        case 'highlight': {
          const value = attr(c, 'val') ?? 'none'
          p.highlight = value === 'none' ? null : (hexToRgba(HIGHLIGHT[value]) ?? p.highlight)
          break
        }
        case 'shd': {
          const shd = this.shading(c)
          if (shd !== undefined) p.shd = shd
          break
        }
        case 'sz': {
          const size = num(c, 'val')
          if (size && size > 0) p.size = clamp(size / 2, 1, 400)
          break
        }
        case 'rFonts': {
          const pick = (key: 'ascii' | 'hAnsi' | 'eastAsia', themeKey: string) => this.themeFont(attr(c, themeKey)) ?? attr(c, key)
          const ascii = pick('ascii', 'asciiTheme')
          const hAnsi = pick('hAnsi', 'hAnsiTheme')
          const ea = pick('eastAsia', 'eastAsiaTheme')
          if (ascii) p.ascii = ascii
          if (hAnsi) p.hAnsi = hAnsi
          if (ea) p.ea = ea
          break
        }
        case 'vertAlign':
          p.vert = attr(c, 'val')
          break
        case 'caps':
          p.caps = truthy(c)
          break
        case 'smallCaps':
          p.smallCaps = truthy(c)
          break
        case 'spacing': {
          const value = num(c, 'val')
          if (value !== undefined) p.spacing = clamp(value / 20, -20, 100)
          break
        }
        case 'position': {
          const value = num(c, 'val')
          if (value !== undefined) p.position = clamp(value / 2, -100, 100)
          break
        }
        case 'vanish':
          p.vanish = truthy(c)
          break
        case 'rtl':
          p.rtl = truthy(c)
          break
        case 'em':
          p.em = attr(c, 'val')
          break
        case 'shadow':
          p.shadow = truthy(c)
          break
        case 'outline':
          p.outline = truthy(c)
          break
      }
    }
    return p
  }

  private mergePara(base: ParaProps, pPr: XmlNode | null | undefined): ParaProps {
    if (!pPr) return base
    const p: ParaProps = { ...base, borders: { ...base.borders }, tabs: base.tabs }
    for (const c of pPr.children) {
      switch (local(c.name)) {
        case 'jc':
          p.jc = attr(c, 'val')
          break
        case 'ind': {
          const len = (twipsName: string, charsName: string): Len | undefined => {
            const chars = num(c, charsName)
            if (chars) return { em: chars / 100 }
            const twips = num(c, twipsName)
            return twips === undefined ? undefined : { pt: twips / 20 }
          }
          const left = len('left', 'leftChars') ?? len('start', 'startChars')
          const right = len('right', 'rightChars') ?? len('end', 'endChars')
          const firstLine = len('firstLine', 'firstLineChars') ?? len('first-line', 'firstLineChars')
          const hanging = len('hanging', 'hangingChars')
          if (left) p.left = left
          if (right) p.right = right
          if (hanging && (hanging.pt || hanging.em)) {
            p.hanging = hanging
            p.firstLine = undefined
          } else if (firstLine) {
            p.firstLine = firstLine
            p.hanging = undefined
          } else if (hanging) p.hanging = undefined
          break
        }
        case 'spacing': {
          const side = (name: 'before' | 'after') => {
            const auto = attr(c, `${name}Autospacing`)
            if (auto === '1' || auto === 'true' || auto === 'on') return 14
            const lines = num(c, `${name}Lines`)
            if (lines) return (lines / 100) * 12
            const twips = num(c, name)
            return twips === undefined ? undefined : twips / 20
          }
          const before = side('before')
          const after = side('after')
          if (before !== undefined) p.before = clamp(before, 0, 1000)
          if (after !== undefined) p.after = clamp(after, 0, 1000)
          const line = num(c, 'line')
          if (line !== undefined && line > 0) {
            p.line = line
            p.lineRule = attr(c, 'lineRule') ?? 'auto'
          }
          break
        }
        case 'contextualSpacing':
          p.contextual = truthy(c)
          break
        case 'pBdr':
          for (const side of c.children) p.borders[local(side.name)] = side
          break
        case 'shd': {
          const shd = this.shading(c)
          if (shd !== undefined) p.shd = shd
          break
        }
        case 'bidi':
          p.bidi = truthy(c)
          break
        case 'pageBreakBefore':
          p.pageBreakBefore = truthy(c)
          break
        case 'outlineLvl': {
          const level = num(c, 'val')
          if (level !== undefined) p.outline = level
          break
        }
        case 'snapToGrid':
          p.snap = truthy(c)
          break
        case 'tabs': {
          let tabs = [...p.tabs]
          for (const tab of children(c, 'tab')) {
            const pos = num(tab, 'pos') ?? 0
            tabs = tabs.filter((existing) => existing.pos !== pos)
            const val = attr(tab, 'val') ?? 'left'
            if (val !== 'clear') tabs.push({ val, pos, leader: attr(tab, 'leader') ?? 'none' })
          }
          p.tabs = tabs.sort((a, b) => a.pos - b.pos).slice(0, 40)
          break
        }
      }
    }
    return p
  }

  private fontStack(p: RunProps, skipSymbol = false): string {
    let names = [p.ascii, p.ea, p.hAnsi].map(safeFont).filter(Boolean)
    if (skipSymbol) names = names.filter((name) => !SYMBOL_FONT.test(name))
    if (!names.length) names = [this.theme.minor.latin, this.theme.minor.ea].map(safeFont).filter(Boolean)
    names = [...new Set(names)]
    const key = names.join('|') + (safeFont(p.ea) ? `#${safeFont(p.ea)}` : '')
    const cached = this.fontStacks.get(key)
    if (cached !== undefined) return cached
    const kindOf = (name: string | undefined): FontKind => {
      if (!name) return 'sans'
      if (/楷|kai/i.test(name)) return 'kai'
      // East Asian names say what the face looks like better than the font table's family, which
      // calls fixed-pitch Ming and Mincho faces (細明體, MS Mincho) "modern".
      if (CJK_SERIF_FONT.test(name)) return 'serif'
      if (CJK_SANS_FONT.test(name)) return 'sans'
      const known = this.fonts.get(name)?.kind
      if (known) return known
      if (/serif|times|georgia|cambria|garamond|book antiqua|palatino/i.test(name) && !/sans/i.test(name)) return 'serif'
      if (/consolas|courier|mono|menlo/i.test(name)) return 'mono'
      return 'sans'
    }
    const parts: string[] = []
    for (const name of names) {
      parts.push(`'${name}'`)
      const alt = safeFont(this.fonts.get(name)?.alt)
      if (alt && alt !== name) parts.push(`'${alt}'`)
    }
    const latin = kindOf(names[0])
    const cjk = kindOf(safeFont(p.ea) || names[0])
    const stack = [...new Set(parts)].join(',') + `,${LATIN_FALLBACK[latin]},${CJK_FALLBACK[cjk]},${GENERIC[latin]}`
    this.fontStacks.set(key, stack)
    return stack
  }

  private runCss(p: RunProps, dark: boolean, skipSymbol = false): string {
    const out: string[] = [`font-family:${this.fontStack(p, skipSymbol)}`]
    let size = p.size
    if (p.vert === 'superscript' || p.vert === 'subscript') {
      out.push(`vertical-align:${p.vert === 'superscript' ? 'super' : 'sub'}`)
      if (size) size = size * 0.65
      else out.push('font-size:65%')
    } else if (p.position) out.push(`vertical-align:${pt(p.position)}`)
    if (size) out.push(`font-size:${pt(size)}`)
    if (p.b) out.push('font-weight:700')
    if (p.i) out.push('font-style:italic')
    const lines: string[] = []
    if (p.u && p.u !== 'none') lines.push('underline')
    if (p.strike || p.dstrike) lines.push('line-through')
    if (lines.length) {
      out.push(`text-decoration-line:${lines.join(' ')}`)
      const u = p.u ?? ''
      const style = p.dstrike || /double/i.test(u) ? 'double' : /dotted|dot(?!Dash)/i.test(u) ? 'dotted' : /dash/i.test(u) ? 'dashed' : /wav/i.test(u) ? 'wavy' : ''
      if (style) out.push(`text-decoration-style:${style}`)
      if (p.uColor && !p.strike && !p.dstrike) out.push(`text-decoration-color:${cssColor(p.uColor)}`)
      if (/thick|heavy/i.test(u)) out.push('text-decoration-thickness:2px')
    }
    const background = p.highlight ?? p.shd ?? null
    if (background) out.push(`background-color:${cssColor(background)}`)
    if (p.color && p.color !== 'auto') out.push(`color:${cssColor(p.color)}`)
    else if (background ? isDark(background) : dark) out.push('color:rgb(255,255,255)')
    else if (background) out.push('color:rgb(0,0,0)')
    if (p.caps) out.push('text-transform:uppercase')
    if (p.smallCaps) out.push('font-variant:small-caps')
    if (p.spacing) out.push(`letter-spacing:${pt(p.spacing)}`)
    if (p.rtl) out.push('direction:rtl;unicode-bidi:embed')
    if (p.em && p.em !== 'none') {
      const mark = p.em === 'comma' ? 'filled sesame' : p.em === 'circle' ? 'open circle' : 'filled dot'
      out.push(`text-emphasis:${mark}`)
      if (p.em === 'underDot') out.push('text-emphasis-position:under right')
    }
    if (p.shadow) out.push('text-shadow:1px 1px 1px rgba(0,0,0,.35)')
    if (p.outline) out.push('-webkit-text-stroke:.6px currentColor;color:transparent')
    return out.join(';')
  }

  private borderCss(node: XmlNode | undefined): string | null {
    if (!node) return null
    const val = attr(node, 'val') ?? 'single'
    if (val === 'nil' || val === 'none') return 'none'
    let width = clamp((num(node, 'sz') ?? 4) / 8, 0.25, 12)
    const style = /double|triple|thinThick|thickThin/i.test(val) ? 'double' : /dotted/i.test(val) ? 'dotted' : /dash/i.test(val) ? 'dashed' : val === 'inset' || val === 'outset' ? val : 'solid'
    if (style === 'double') width = Math.max(width * 3, 2.25)
    const color = this.wordColor(node)
    return `${pt(width)} ${style} ${color && color !== 'auto' ? cssColor(color) : 'rgb(0,0,0)'}`
  }

  // --- blocks --------------------------------------------------------------------------------

  private expand(nodes: XmlNode[], depth = 0): XmlNode[] {
    const out: XmlNode[] = []
    if (depth > 20) return out
    for (const node of nodes) {
      switch (local(node.name)) {
        case 'p':
        case 'tbl':
        case 'bookmarkStart':
        case 'altChunk':
          out.push(node)
          break
        case 'sdt':
          out.push(...this.expand(child(node, 'sdtContent')?.children ?? [], depth + 1))
          break
        case 'customXml':
        case 'ins':
        case 'moveTo':
        case 'smartTag':
          out.push(...this.expand(node.children, depth + 1))
          break
        case 'AlternateContent':
          out.push(...this.expand(chooseAlternate(node), depth + 1))
          break
      }
    }
    return out
  }

  private paraInfo(p: XmlNode, scope: Scope): ParaInfo {
    const pPr = child(p, 'pPr')
    const styleId = attr(child(pPr, 'pStyle'), 'val') ?? this.defaultStyle.paragraph ?? ''
    const chain = this.chain(styleId || undefined, 'paragraph')
    let numId: string | undefined
    let ilvl: number | undefined
    for (const numPr of [...chain.map((style) => child(style.pPr, 'numPr')), child(pPr, 'numPr')]) {
      if (!numPr) continue
      numId = attr(child(numPr, 'numId'), 'val') ?? numId
      ilvl = num(child(numPr, 'ilvl'), 'val') ?? ilvl
    }
    let props: ParaProps = { borders: {}, tabs: [] }
    props = this.mergePara(props, this.pPrDefault)
    // In a table, the table style outranks the default paragraph style but not any other style (as Word does).
    const plain = styleId === (this.defaultStyle.paragraph ?? '')
    const ordered = <T,>(merge: (into: T, node: XmlNode | null) => T, into: T, layers: (XmlNode | null)[], styles: (XmlNode | null)[]) => {
      for (const node of plain ? [...styles, ...layers] : [...layers, ...styles]) into = merge(into, node)
      return into
    }
    props = ordered((into, node) => this.mergePara(into, node), props, scope.layer.pPr, chain.map((style) => style.pPr))
    let numbering: ParaInfo['num']
    const level = numId && numId !== '0' ? clamp(Math.floor(ilvl ?? 0), 0, 8) : 0
    if (numId && numId !== '0') {
      const lvl = this.level(numId, level)
      if (lvl) {
        props = this.mergePara(props, lvl.pPr)
        numbering = { numId, ilvl: level }
      }
    }
    props = this.mergePara(props, pPr)
    let base: RunProps = { size: 10 }
    base = this.mergeRun(base, this.rPrDefault)
    base = ordered((into, node) => this.mergeRun(into, node), base, scope.layer.rPr, chain.map((style) => style.rPr))
    const mark = this.mergeRun(base, child(pPr, 'rPr'))
    const borders = Object.keys(props.borders).length ? JSON.stringify([Object.entries(props.borders).map(([side, node]) => [side, node.attrs]), props.left, props.right]) : ''
    return { node: p, styleId, props, base, mark, num: numbering, sig: borders }
  }

  renderBlocks(nodes: XmlNode[], scope: Scope): string {
    const items = this.expand(nodes)
    const infos = items.map((node) => {
      if (local(node.name) !== 'p') return null
      try {
        return this.paraInfo(node, scope)
      } catch {
        return null
      }
    })
    let html = ''
    // The body's top level is cut into page segments at its page breaks, so a drawing placed on a page lands on it.
    const sheet = scope.story.kind === 'body' && scope.depth === 0 && !this.inCell ? scope.sheet : null
    let segment = ''
    const emit = (piece: string) => {
      if (sheet) segment += piece
      else html += piece
    }
    const nextPage = () => {
      html += this.closePage(segment, sheet!)
      segment = ''
      this.page.index++
    }
    let lastBlock = items.length - 1
    while (lastBlock >= 0 && local(items[lastBlock].name) === 'bookmarkStart') lastBlock--
    for (let at = 0; at < items.length; at++) {
      const node = items[at]
      const name = local(node.name)
      if (name === 'bookmarkStart') {
        const bookmark = attr(node, 'name')
        if (bookmark && bookmark !== '_GoBack') emit(`<span id="${anchorId(bookmark)}"></span>`)
        continue
      }
      if (this.inCell) {
        // Rows are cut whole before this budget runs out; this stops only one row that alone holds too much.
        if (this.cellBlocks >= LIMITS.cellBlocks + LIMITS.blocks) {
          this.cutCellBlocks = true
          break
        }
        this.cellBlocks++
      } else {
        if (this.blocks >= LIMITS.blocks) {
          this.cutBlocks = true
          break
        }
        this.blocks++
      }
      try {
        if (name === 'p') {
          const info = infos[at]
          if (!info) throw new Error('段落格式無法讀取')
          const prev = at > 0 ? infos[at - 1] : null
          const prevIsTable = at > 0 && local(items[at - 1].name) === 'tbl'
          const next = at + 1 < items.length ? infos[at + 1] : null
          const hasNext = at < lastBlock
          const same = Boolean(prev && prev.styleId === info.styleId)
          const before = info.props.contextual && same ? 0 : (info.props.before ?? 0)
          const prevAfter = prev ? (prev.props.contextual && same ? 0 : (prev.props.after ?? 0)) : 0
          const top = prevIsTable || !prev ? before : prevAfter + before
          const bottom = hasNext ? 0 : (info.props.after ?? 0)
          const groupTop = Boolean(info.sig && prev && prev.sig === info.sig)
          const groupBottom = Boolean(info.sig && next && next.sig === info.sig)
          let pageBreak = sheet ? this.pageBreak(info, at === 0) : null
          if (pageBreak === 'start' && (segment || html)) nextPage()
          // A page break inside the paragraph: what follows it starts the next page, so a drawing placed
          // after it lands there. Each piece after a break is drawn as the paragraph's continuation.
          let rest: ParaInfo | null = null
          let current = info
          let pieces = 0
          while (pageBreak === 'end' && pieces < 50) {
            const split = splitAtPageBreak(current.node)
            if (!split) break
            emit(this.renderParagraph({ ...current, node: split[0] }, scope, pieces ? 0 : top, 0, groupTop, false, at === 0 && !pieces))
            nextPage()
            pieces++
            // The rest continues the same list item: the item's indent, but no label of its own (wherever
            // the numbering comes from, the paragraph or its style).
            const continued = this.paraInfo(split[1], scope)
            rest = continued.num ? { ...continued, num: undefined, props: { ...continued.props, firstLine: undefined, hanging: undefined } } : continued
            current = rest
            pageBreak = this.pageBreak(current, true) === 'end' ? 'end' : null
          }
          let block = rest
            ? this.renderParagraph(current, scope, 0, bottom, false, groupBottom, false)
            : this.renderParagraph(info, scope, top, bottom, groupTop, groupBottom, at === 0)
          if (pageBreak === 'start') {
            // The divider belongs between the pages, not inside the new one.
            const lead = /^(<div class="dx-p[^"]*"[^>]*>)?<span class="dx-pgbr" role="separator" aria-label="分頁"><\/span>/.exec(block)
            if (lead) {
              html += segment + PAGE_BREAK
              segment = ''
              block = (lead[1] ?? '') + block.slice(lead[0].length)
            }
          }
          emit(block)
          if (pageBreak === 'end') nextPage()
        } else if (name === 'tbl') {
          const prev = at > 0 ? infos[at - 1] : null
          emit(this.renderTable(node, scope, prev ? (prev.props.after ?? 0) : 0))
        } else if (name === 'altChunk') {
          emit(this.altChunk(node, scope))
        }
      } catch (error) {
        emit(`<div class="dx-p">${unsupported(`此${name === 'tbl' ? '表格' : name === 'altChunk' ? '嵌入內容' : '段落'}無法顯示:${reason(error)}`)}</div>`)
      }
    }
    if (sheet) html += this.closePage(segment, sheet)
    return html
  }

  /** A part's content type from [Content_Types].xml (an Override, else the Default for its extension). */
  private contentType(part: string): string {
    let types: XmlNode | null = null
    try {
      types = this.pkg.xml('[Content_Types].xml')
    } catch {}
    const name = `/${part}`.toLowerCase()
    const override = children(types, 'Override').find((node) => (attr(node, 'PartName') ?? '').toLowerCase() === name)
    if (override) return (attr(override, 'ContentType') ?? '').toLowerCase()
    const ext = posix.extname(part).slice(1).toLowerCase()
    return (attr(children(types, 'Default').find((node) => (attr(node, 'Extension') ?? '').toLowerCase() === ext), 'ContentType') ?? '').toLowerCase()
  }

  /** w:altChunk: an HTML, MHT or plain-text part drawn in place through the whitelist; anything else is a labelled box. */
  private altChunk(node: XmlNode, scope: Scope): string {
    const rel = this.pkg.rel(scope.story.part, relAttr(node, 'id'))
    const label = '嵌入的外部內容'
    if (!rel || rel.external || !this.pkg.zip.has(rel.target)) return `<div class="dx-p">${unsupported(label, undefined, undefined, '檔案裡找不到這段內容')}</div>`
    const bytes = this.pkg.zip.read(rel.target)
    if (!bytes || bytes.length > CHUNK_LIMITS.bytes) return `<div class="dx-p">${unsupported(label, undefined, undefined, '內容太大,請下載原始檔查看')}</div>`
    const type = this.contentType(rel.target)
    const ext = posix.extname(rel.target).toLowerCase()
    const head = bytes.subarray(0, 1024).toString('latin1').replace(/^\u00ef\u00bb\u00bf/, '')
    let source: string | null = null
    let resolve: ((src: string) => string | undefined) | undefined
    if (type === 'message/rfc822' || type === 'multipart/related' || ext === '.mht' || ext === '.mhtml' || /^\s*(MIME-Version|Content-Type)\s*:/i.test(head)) {
      const mht = readMht(bytes)
      if (mht) ({ html: source, resolve } = mht)
    } else if (type === 'text/html' || type === 'application/xhtml+xml' || /^\.x?html?$/.test(ext) || /^\s*(<!doctype html|<html|<body|<\?xml[^>]*>\s*<html)/i.test(head)) {
      source = decodeChunk(bytes)
    } else if (type === 'text/plain' || ext === '.txt') {
      const plain = decodeChunk(bytes).replace(/\r\n?/g, '\n').trim()
      const paragraphs = plain.split(/\n{2,}/).map((block) => `<p>${block.split('\n').map(esc).join('<br>')}</p>`)
      this.text.push(plain)
      return `<div class="dx-html ${this.chunkClass(scope)}">${paragraphs.join('')}</div>`
    }
    if (source === null) return `<div class="dx-p">${unsupported(label)}</div>`
    const chunk = sanitizeChunk(source, resolve)
    const text = chunk.text.replace(/\s+/g, ' ').trim()
    if (text) this.text.push(text)
    if (chunk.cut) this.notices.push('嵌入的 HTML 內容太長:只顯示前段,其餘請下載原始檔查看。')
    return `<div class="dx-html ${this.chunkClass(scope)}">${chunk.html}</div>`
  }

  /** The document's default paragraph look, for embedded HTML (whose own styles are not kept). */
  private chunkClass(scope: Scope): string {
    const info = this.paraInfo({ name: 'w:p', attrs: {}, children: [], text: '' }, scope)
    return this.cls(`${this.runCss(info.mark, scope.dark)};font-size:${pt(info.mark.size ?? 10)}`)
  }

  private paraCss(info: ParaInfo, scope: Scope, top: number, bottom: number, groupTop: boolean, groupBottom: boolean): string {
    const p = info.props
    const out: string[] = [`font-family:${this.fontStack(info.mark)}`, `font-size:${pt(info.mark.size ?? 10)}`]
    const jc = p.jc
    if (jc === 'center') out.push('text-align:center')
    else if (jc === 'right' || jc === 'end') out.push('text-align:end')
    else if (jc === 'both' || jc === 'lowKashida' || jc === 'mediumKashida' || jc === 'highKashida' || jc === 'thaiDistribute') out.push('text-align:justify')
    else if (jc === 'distribute') out.push('text-align:justify;text-align-last:justify')
    else if (jc === 'left' || jc === 'start') out.push('text-align:start')
    if (p.bidi) out.push('direction:rtl')
    const left = lenCss(p.left)
    const right = lenCss(p.right)
    if (left && left !== '0pt') out.push(`margin-inline-start:${left}`)
    if (right && right !== '0pt') out.push(`margin-inline-end:${right}`)
    const indent = p.hanging ? lenCss(p.hanging, true) : lenCss(p.firstLine)
    if (indent && indent !== '0pt' && indent !== '-0pt') out.push(`text-indent:${indent}`)
    if (groupTop) {
      if (top) out.push(`padding-top:${pt(top)}`)
    } else if (top) out.push(`margin-top:${pt(top)}`)
    if (bottom) out.push(`margin-bottom:${pt(bottom)}`)
    const grid = p.snap === false ? 0 : scope.grid
    if (p.line !== undefined) {
      const rule = p.lineRule ?? 'auto'
      if (rule === 'exact') out.push(`line-height:${pt(clamp(p.line / 20, 1, 1000))}`)
      else if (rule === 'atLeast') out.push(`line-height:max(${pt(clamp(p.line / 20, 1, 1000))},1.2em)`)
      else {
        const multiple = clamp(p.line / 240, 0.3, 10)
        out.push(grid ? `line-height:max(${pt(grid * multiple)},${round(1.2 * multiple, 3)}em)` : `line-height:${round(1.2 * multiple, 3)}`)
      }
    } else if (grid) out.push(`line-height:max(${pt(grid)},1.2em)`)
    const sides: [string, string, string][] = [['top', 'top', 'between'], ['bottom', 'bottom', ''], ['left', 'left', 'start'], ['right', 'right', 'end']]
    for (const [side, key, alt] of sides) {
      let node: XmlNode | undefined = p.borders[key] ?? (side === 'left' || side === 'right' ? p.borders[alt] : undefined)
      if (side === 'top' && groupTop) node = p.borders.between
      if (side === 'bottom' && groupBottom) node = undefined
      const css = this.borderCss(node)
      if (css && css !== 'none') {
        out.push(`border-${side}:${css}`)
        const space = clamp(num(node, 'space') ?? 0, 0, 31)
        if (space) out.push(`padding-${side}:${pt(space)}`)
      }
    }
    if (p.shd) out.push(`background-color:${cssColor(p.shd)}`)
    return out.join(';')
  }

  private renderParagraph(info: ParaInfo, scope: Scope, top: number, bottom: number, groupTop: boolean, groupBottom: boolean, first: boolean): string {
    const dark = info.props.shd ? isDark(info.props.shd) : scope.dark
    const local_scope = dark === scope.dark ? scope : { ...scope, dark }
    const out: Inline = { parts: [], text: [], visual: false, pins: [], floats: [] }
    this.inline(info.node.children.filter((node) => local(node.name) !== 'pPr'), local_scope, info.base, null, out)
    const text = out.text.join('')
    if (info.mark.vanish && !text && !out.visual) return ''
    let numHtml = ''
    if (info.num) {
      const label = this.nextLabel(info.num.numId, info.num.ilvl)
      const level = this.level(info.num.numId, info.num.ilvl)
      if (label && level) {
        let props = this.mergeRun(info.mark, level.rPr)
        props = { ...props, u: undefined, highlight: undefined }
        const font = safeFont(props.ascii) || safeFont(props.hAnsi)
        let shown = label
        let skipSymbol = false
        if (SYMBOL_FONT.test(font)) {
          const mapped = [...label].map((char) => symbolChar(font, char.codePointAt(0)!) ?? (char.codePointAt(0)! >= 0xf000 ? '•' : char)).join('')
          shown = mapped
          skipSymbol = true
        } else shown = [...label].map((char) => (char.codePointAt(0)! >= 0xf000 && char.codePointAt(0)! <= 0xf0ff ? '•' : char)).join('')
        const hanging = info.props.hanging ? lenCss(info.props.hanging) : null
        let layout = ''
        if (level.suff === 'tab') layout = hanging ? `;display:inline-block;box-sizing:border-box;min-width:${hanging};padding-inline-end:.4em;text-indent:0` : ';display:inline-block;min-width:1.5em;text-indent:0;padding-inline-end:.5em'
        const space = level.suff === 'space' ? ' ' : ''
        numHtml = `<span class="dx-num ${this.cls(this.runCss(props, dark, skipSymbol) + layout)}">${esc(cleanText(shown))}${space}</span>`
      }
    }
    if (text) this.text.push(text)
    // Floats go first, so the text wraps them from the paragraph's first line, as Word lays them out.
    let content = out.floats.join('') + out.parts.join('')
    const css = this.paraCss(info, scope, top, bottom, groupTop, groupBottom)
    const segments = content.split(TAB)
    let flex = false
    if (segments.length > 1) {
      const stops = info.props.tabs
      const aligned = stops.some((stop) => ['right', 'end', 'center', 'decimal'].includes(stop.val))
      const gap = '<span class="dx-tab"></span>'
      if (aligned) {
        flex = true
        const count = segments.length - 1
        let html = `<span class="dx-seg">${numHtml}${segments[0]}</span>`
        for (let j = 0; j < count; j++) {
          const stop = stops[stops.length - (count - j)]
          if (stop && ['right', 'end', 'center', 'decimal'].includes(stop.val)) {
            const leader = ['dot', 'hyphen', 'underscore', 'heavy', 'middleDot'].includes(stop.leader) ? stop.leader : 'none'
            html += `<span class="dx-fill dx-ld-${leader}"></span>`
          } else html += gap
          html += `<span class="dx-seg">${segments[j + 1]}</span>`
        }
        content = html
      } else content = numHtml + segments.join(gap)
    } else content = numHtml + content
    if (!text && !out.visual && !numHtml) content += '<br>'
    if (out.pins.length) {
      // The pin sits at the paragraph's top left; the paragraph's indent is undone, as Word measures from the column.
      const indent = lenCss(info.props.left)
      const undo = indent && indent !== '0pt' ? ` style="margin-inline-start:calc(-1 * ${indent})"` : ''
      content = `<span class="dx-pin"${undo}>${out.pins.join('')}</span>${content}`
    }
    const outline = info.props.outline
    const role = outline !== undefined && outline >= 0 && outline < 9 ? ` role="heading" aria-level="${Math.min(outline + 1, 6)}"` : ''
    const breakBefore = info.props.pageBreakBefore && !first && scope.story.kind === 'body' ? PAGE_BREAK : ''
    return `${breakBefore}<div class="dx-p${flex ? ' dx-flex' : ''} ${this.cls(css)}"${role}>${content}</div>`
  }

  // --- inline content ------------------------------------------------------------------------

  private visible(scope: Scope): boolean {
    return scope.story.fields.every((field) => field.separated)
  }

  private fieldLink(scope: Scope): Link {
    for (let at = scope.story.fields.length - 1; at >= 0; at--) {
      const field = scope.story.fields[at]
      const link = parseHyperlinkInstr(field.instr.trim())
      if (link) return link
    }
    return null
  }

  private wrapLink(html: string, link: Link): string {
    if (!link) return html
    if (link.href) return `<a class="dx-lnk" href="${esc(link.href)}">${html}</a>`
    return `<span class="dx-lnk" title="${esc(link.title ?? '')}">${html}</span>`
  }

  private inline(nodes: XmlNode[], scope: Scope, base: RunProps, link: Link, out: Inline, depth = 0): void {
    if (depth > 30) return
    for (const node of nodes) {
      switch (local(node.name)) {
        case 'r':
          this.run(node, scope, base, link, out)
          break
        case 'hyperlink': {
          const rel = this.pkg.rel(scope.story.part, relAttr(node, 'id'))
          const anchor = attr(node, 'anchor')
          const target: Link = rel?.external ? { title: rel.target } : anchor ? { href: `#${anchorId(anchor)}` } : link
          this.inline(node.children, scope, base, target, out, depth + 1)
          break
        }
        case 'fldSimple': {
          const fieldLink = parseHyperlinkInstr((attr(node, 'instr') ?? '').trim())
          this.inline(node.children, scope, base, fieldLink ?? link, out, depth + 1)
          break
        }
        case 'sdt':
          this.inline(child(node, 'sdtContent')?.children ?? [], scope, base, link, out, depth + 1)
          break
        case 'ins':
        case 'moveTo':
        case 'smartTag':
        case 'customXml':
        case 'dir':
        case 'bdo':
          this.inline(node.children, scope, base, link, out, depth + 1)
          break
        case 'AlternateContent':
          this.inline(chooseAlternate(node), scope, base, link, out, depth + 1)
          break
        case 'bookmarkStart': {
          const name = attr(node, 'name')
          if (name && name !== '_GoBack') out.parts.push(`<span id="${anchorId(name)}"></span>`)
          break
        }
        case 'oMath':
        case 'oMathPara': {
          if (!this.visible(scope)) break
          const text: string[] = []
          const html = this.math(node, text, 0)
          const block = local(node.name) === 'oMathPara'
          out.parts.push(block ? `<span class="dx-math dx-mathpara">${html}</span>` : `<span class="dx-math">${html}</span>`)
          out.text.push(text.join(''))
          out.visual = true
          break
        }
        case 'subDoc':
          out.parts.push(unsupported('子文件'))
          out.visual = true
          break
      }
    }
  }

  private runItems(nodes: XmlNode[], depth = 0): XmlNode[] {
    const out: XmlNode[] = []
    for (const node of nodes) {
      if (local(node.name) === 'AlternateContent' && depth < 5) out.push(...this.runItems(chooseAlternate(node), depth + 1))
      else out.push(node)
    }
    return out
  }

  private run(r: XmlNode, scope: Scope, base: RunProps, link: Link, out: Inline): void {
    const rPr = child(r, 'rPr')
    let props = base
    const rStyle = attr(child(rPr, 'rStyle'), 'val')
    if (rStyle) for (const style of this.chain(rStyle, 'character')) props = this.mergeRun(props, style.rPr)
    props = this.mergeRun(props, rPr)
    const font = safeFont(props.ascii) || safeFont(props.hAnsi)
    const symbolRun = SYMBOL_FONT.test(font)
    let buffer = ''
    let bufferMapped = true
    const flush = () => {
      if (!buffer) return
      const runLink = link ?? this.fieldLink(scope)
      const css = this.runCss(props, scope.dark, symbolRun && bufferMapped)
      out.parts.push(this.wrapLink(`<span class="${this.cls(css)}">${esc(buffer)}</span>`, runLink))
      out.text.push(buffer)
      buffer = ''
      bufferMapped = true
    }
    const push = (html: string) => {
      flush()
      out.parts.push(this.wrapLink(html, link ?? this.fieldLink(scope)))
      out.visual = true
    }
    const fields = scope.story.fields
    for (const c of this.runItems(r.children)) {
      const name = local(c.name)
      if (name === 'fldChar') {
        flush()
        const type = attr(c, 'fldCharType')
        // Past 20 nested fields the frames are only counted, so begins and ends stay balanced.
        if (type === 'begin') {
          if (fields.length < 20) fields.push({ instr: '', separated: false })
          else scope.story.overflow = (scope.story.overflow ?? 0) + 1
        } else if (type === 'separate') {
          const top = fields[fields.length - 1]
          if (top && !scope.story.overflow) top.separated = true
        } else if (type === 'end') {
          if (scope.story.overflow) scope.story.overflow--
          else fields.pop()
        }
        continue
      }
      if (name === 'instrText') {
        const top = fields[fields.length - 1]
        if (top && !top.separated && top.instr.length < 2000) top.instr += c.text
        continue
      }
      if (props.vanish || !this.visible(scope)) continue
      switch (name) {
        case 't': {
          let text = cleanText(c.text)
          if (symbolRun) {
            text = [...text]
              .map((char) => {
                const mapped = symbolChar(font, char.codePointAt(0)!)
                if (mapped === null) bufferMapped = false
                return mapped ?? char
              })
              .join('')
          }
          buffer += text
          break
        }
        case 'tab':
        case 'ptab':
          flush()
          out.parts.push(TAB)
          out.visual = true
          break
        case 'br': {
          const type = attr(c, 'type')
          if (type === 'page') push(PAGE_BREAK)
          else if (type === 'column') push('<span class="dx-colbr"></span>')
          else push('<br>')
          break
        }
        case 'cr':
          push('<br>')
          break
        case 'sym': {
          const symFont = safeFont(attr(c, 'font'))
          const code = parseInt(attr(c, 'char') ?? '', 16)
          if (!Number.isFinite(code)) break
          const mapped = symbolChar(symFont, code)
          if (mapped) buffer += mapped
          else if (code > 0 && code < 0x110000) {
            flush()
            const symProps = { ...props, ascii: symFont, hAnsi: symFont, ea: symFont }
            const shown = String.fromCodePoint(code >= 0xf000 && code <= 0xf0ff ? code - 0xf000 : code)
            push(`<span class="${this.cls(this.runCss(symProps, scope.dark))}">${esc(cleanText(shown))}</span>`)
          }
          break
        }
        case 'noBreakHyphen':
          buffer += '‑'
          break
        case 'softHyphen':
          buffer += '­'
          break
        case 'drawing': {
          flush()
          let html: string
          try {
            html = this.drawing(c, scope, out)
          } catch {
            html = unsupported('圖片')
          }
          if (html) push(html)
          break
        }
        case 'pict':
        case 'object': {
          flush()
          let html: string
          try {
            html = this.vml(c, scope, out)
          } catch {
            html = ''
          }
          if (!html && name === 'object') html = unsupported('內嵌物件')
          if (html) push(html)
          break
        }
        case 'footnoteReference':
        case 'endnoteReference': {
          flush()
          const kind: NoteKind = name === 'footnoteReference' ? 'footnote' : 'endnote'
          const id = attr(c, 'id') ?? ''
          const numbers = this.noteNumbers[kind]
          let n = numbers.get(id)
          if (n === undefined && numbers.size < LIMITS.notes) {
            n = numbers.size + 1
            numbers.set(id, n)
          }
          const custom = attr(c, 'customMarkFollows')
          if (n === undefined || custom === '1' || custom === 'true' || custom === 'on') break
          const prefix = kind === 'footnote' ? 'fn' : 'en'
          const label = esc(formatNumber(n, this.noteFormats[kind]))
          const inner = `<a class="dx-fnref" href="#${prefix}-${n}" id="${prefix}ref-${n}">${label}</a>`
          const span = `<span class="${this.cls(this.runCss(props, scope.dark))}">${props.vert === 'superscript' ? inner : `<sup>${inner}</sup>`}</span>`
          out.parts.push(span)
          out.visual = true
          break
        }
        case 'footnoteRef':
        case 'endnoteRef': {
          flush()
          const note = scope.story.note
          if (!note) break
          const prefix = note.kind === 'footnote' ? 'fn' : 'en'
          const label = esc(formatNumber(note.n, this.noteFormats[note.kind]))
          const inner = `<a class="dx-fnref" href="#${prefix}ref-${note.n}">${label}</a>`
          out.parts.push(`<span class="${this.cls(this.runCss(props, scope.dark))}">${props.vert === 'superscript' ? inner : `<sup>${inner}</sup>`}</span>`)
          out.visual = true
          break
        }
        case 'ruby': {
          flush()
          const baseOut: Inline = { parts: [], text: [], visual: false, pins: [], floats: [] }
          const rtOut: Inline = { parts: [], text: [], visual: false, pins: [], floats: [] }
          this.inline(child(c, 'rubyBase')?.children ?? [], scope, props, null, baseOut, 1)
          this.inline(child(c, 'rt')?.children ?? [], scope, props, null, rtOut, 1)
          out.parts.push(`<ruby>${baseOut.parts.join('').split(TAB).join('')}<rt>${rtOut.parts.join('').split(TAB).join('')}</rt></ruby>`)
          out.text.push(baseOut.text.join(''))
          break
        }
      }
    }
    flush()
  }

  // --- equations -----------------------------------------------------------------------------

  private math(node: XmlNode, text: string[], depth: number): string {
    if (depth > 40) return ''
    const kids = (holder: XmlNode | null | undefined) => (holder ? holder.children.filter((c) => !local(c.name).endsWith('Pr')).map((c) => this.math(c, text, depth + 1)).join('') : '')
    const part = (name: string) => kids(child(node, name))
    const prop = (pr: string, name: string) => attr(path(node, pr, name), 'val')
    switch (local(node.name)) {
      case 'r': {
        const value = children(node, 't').map((t) => t.text).join('')
        text.push(value)
        return esc(cleanText(value))
      }
      case 't':
        text.push(node.text)
        return esc(cleanText(node.text))
      case 'f':
        return `<span class="m-frac"><span>${part('num')}</span><span>${part('den')}</span></span>`
      case 'sSup':
        return `${part('e')}<sup>${part('sup')}</sup>`
      case 'sSub':
        return `${part('e')}<sub>${part('sub')}</sub>`
      case 'sSubSup':
        return `${part('e')}<sub>${part('sub')}</sub><sup>${part('sup')}</sup>`
      case 'sPre':
        return `<sub>${part('sub')}</sub><sup>${part('sup')}</sup>${part('e')}`
      case 'rad': {
        const degree = part('deg')
        return `${degree ? `<sup>${degree}</sup>` : ''}√<span class="m-rad">${part('e')}</span>`
      }
      case 'd': {
        const begin = prop('dPr', 'begChr') ?? '('
        const end = prop('dPr', 'endChr') ?? ')'
        const separator = prop('dPr', 'sepChr') ?? '|'
        return `${esc(begin)}${children(node, 'e').map((e) => kids(e)).join(esc(separator))}${esc(end)}`
      }
      case 'nary': {
        const symbol = prop('naryPr', 'chr') ?? '∫'
        const sub = part('sub')
        const sup = part('sup')
        return `<span class="m-nary">${esc(symbol)}</span>${sub ? `<sub>${sub}</sub>` : ''}${sup ? `<sup>${sup}</sup>` : ''}${part('e')}`
      }
      case 'func':
        return `${part('fName')}⁡${part('e')}`
      case 'acc':
        return `${part('e')}${esc(prop('accPr', 'chr') ?? '̂')}`
      case 'bar':
        return `<span class="m-bar">${part('e')}</span>`
      case 'limLow':
        return `${part('e')}<sub>${part('lim')}</sub>`
      case 'limUpp':
        return `${part('e')}<sup>${part('lim')}</sup>`
      case 'eqArr':
        return children(node, 'e').map((e) => kids(e)).join('<br>')
      case 'm':
        return `<span class="m-mat">${children(node, 'mr').map((row) => `<span class="m-row">${children(row, 'e').map((e) => `<span>${kids(e)}</span>`).join('')}</span>`).join('')}</span>`
      default:
        return kids(node)
    }
  }

  // --- drawings ------------------------------------------------------------------------------

  private drawingColor(node: XmlNode | null | undefined): Rgba | null {
    if (!node) return null
    const name = local(node.name)
    let base: Rgba | null = null
    if (name === 'srgbClr') base = hexToRgba(attr(node, 'val'))
    else if (name === 'schemeClr') {
      const val = attr(node, 'val') ?? ''
      base = hexToRgba(this.theme.colors[SCHEME_SLOTS[val] ?? val])
    } else if (name === 'sysClr') base = hexToRgba(attr(node, 'lastClr') ?? (attr(node, 'val') === 'window' ? 'FFFFFF' : '000000'))
    else if (name === 'prstClr') base = hexToRgba(PRESET_COLORS[attr(node, 'val') ?? ''] ?? '000000')
    return base ? applyModifiers(base, node) : null
  }

  private fillColor(holder: XmlNode | null): Rgba | null | undefined {
    if (!holder) return undefined
    for (const c of holder.children) {
      const name = local(c.name)
      if (name === 'noFill') return null
      if (name === 'solidFill') return this.drawingColor(c.children[0])
      if (name === 'gradFill') return this.drawingColor(path(c, 'gsLst', 'gs')?.children[0])
      // An outline drawn with a hatch reads as its foreground.
      if (name === 'pattFill') return this.drawingColor(child(c, 'fgClr')?.children[0]) ?? BLACK
    }
    return undefined
  }

  /** A shape's fill (see `Paint`): null for no fill or a picture that cannot be shown, undefined when the shape names none. */
  private fillPaint(holder: XmlNode | null, scope: Scope): Paint | null | undefined {
    if (!holder) return undefined
    for (const c of holder.children) {
      const name = local(c.name)
      if (name === 'pattFill') {
        const fg = this.drawingColor(child(c, 'fgClr')?.children[0]) ?? BLACK
        const bg = this.drawingColor(child(c, 'bgClr')?.children[0]) ?? WHITE
        return { pattern: { fg, bg, bits: PATTERNS[attr(c, 'prst') ?? ''] ?? PATTERNS.pct5 } }
      }
      if (name === 'blipFill') {
        const rel = this.pkg.rel(scope.story.part, relAttr(child(c, 'blip'), 'embed'))
        if (!rel || rel.external || !this.pkg.zip.has(rel.target) || !isDrawablePart(rel.target)) return null
        return { picture: this.media(rel.target), tile: child(c, 'tile') !== null }
      }
      if (name === 'grpFill') return scope.groupFill ? (this.fillPaint(scope.groupFill, { ...scope, groupFill: null }) ?? null) : null
    }
    const color = this.fillColor(holder)
    return color === undefined ? undefined : color ? { color } : null
  }

  private paints = 0

  /** A fill as an SVG paint over a w × h (pt) viewBox, with the pattern it needs. */
  private svgPaint(paint: Paint, w: number, h: number): { fill: string; defs: string } {
    if ('color' in paint) return { fill: cssColor(paint.color), defs: '' }
    const id = `dx-f${++this.paints}`
    if ('pattern' in paint) {
      const tile = `<g transform="scale(${PATTERN_PT / 8})" shape-rendering="crispEdges">${patternTile(paint.pattern)}</g>`
      return { fill: `url(#${id})`, defs: `<pattern id="${id}" patternUnits="userSpaceOnUse" width="${PATTERN_PT}" height="${PATTERN_PT}">${tile}</pattern>` }
    }
    const [pw, ph] = [round(Math.max(w, 0.01)), round(Math.max(h, 0.01))]
    return { fill: `url(#${id})`, defs: `<pattern id="${id}" patternUnits="userSpaceOnUse" width="${pw}" height="${ph}"><image href="${esc(paint.picture)}" width="${pw}" height="${ph}" preserveAspectRatio="none"/></pattern>` }
  }

  private drawing(d: XmlNode, scope: Scope, out: Inline): string {
    const holder = child(d, 'inline') ?? child(d, 'anchor')
    if (!holder) return ''
    const extent = child(holder, 'extent')
    const w = clamp((num(extent, 'cx') ?? 0) / EMU_PER_PT, 0, 4000)
    const h = clamp((num(extent, 'cy') ?? 0) / EMU_PER_PT, 0, 8000)
    const docPr = child(holder, 'docPr')
    if (attr(docPr, 'hidden') === '1' || attr(docPr, 'hidden') === 'true') return ''
    const alt = attr(docPr, 'descr') || attr(docPr, 'title') || ''
    const content = path(holder, 'graphic', 'graphicData')?.children[0] ?? null
    if (local(holder.name) !== 'anchor') return this.graphic(content, w, h, scope, alt, false)
    const placed = this.anchorBox(holder, w, h, scope)
    if (placed && placed.mode !== 'flow') {
      // Placed drawings take no room in the text, so they do not count as the paragraph's content.
      ;(placed.mode === 'pin' ? out.pins : out.parts).push(`${placed.open}${this.graphic(content, w, h, scope, alt, true)}</span>`)
      return ''
    }
    const inner = this.graphic(content, w, h, scope, alt, false)
    if (!placed || !inner) return inner
    out.floats.push(`${placed.open}${inner}</span>`)
    out.visual = true
    return ''
  }

  /** Per sheet: which page segment (between page breaks) the body is on, and whether a drawing was placed on it. */
  private page = { index: 0, anchored: false, bottom: 0 }
  private placedOnSheet = false

  /**
   * An absolutely placed box for a drawing that floats free of the text (no wrapping, or behind it).
   * Page, margin and column positions are measured on the sheet (cqw: share of the sheet's width,
   * which scales with the page); on a later page of the same sheet, from that page's segment. Positions
   * relative to the paragraph go into a pin at the paragraph's start. Null when the story has no page.
   */
  private place(o: Placement, scope: Scope): { open: string; pin: boolean } | null {
    const g = scope.sheet
    const story = scope.story.kind
    if (!g || o.w <= 0 || story === 'note') return null
    const cq = (value: number) => `${round((clamp(value, -8000, 8000) / g.w) * 100, 3)}cqw`
    const fromPage = ['page', 'leftMargin', 'rightMargin', 'insideMargin', 'outsideMargin'].includes(o.relH)
    const [hs, ha] =
      o.relH === 'page' ? [0, g.w] : o.relH === 'leftMargin' || o.relH === 'insideMargin' ? [0, g.left] : o.relH === 'rightMargin' || o.relH === 'outsideMargin' ? [g.w - g.right, g.right] : [g.left, g.w - g.left - g.right]
    const x = alignIn(o.alignH, hs, ha, o.w, o.offH)
    const pin = !['page', 'margin', 'topMargin', 'bottomMargin', 'insideMargin', 'outsideMargin'].includes(o.relV)
    // The first page of a sheet places on the sheet itself; later pages on their segment, whose left is the text column's.
    const onSheet = !pin && (story !== 'body' || this.page.index === 0)
    const left = onSheet ? (fromPage ? cq(x) : `calc(var(--il) + ${cq(x - g.left)})`) : fromPage ? `calc(${cq(x)} - var(--il))` : cq(x - g.left)
    let vertical: string
    if (pin) vertical = `top:${cq(o.alignV ? 0 : o.offV)}`
    else {
      const [vs, va] =
        o.relV === 'page' ? [0, g.h] : o.relV === 'margin' ? [g.top, g.h - g.top - g.bottom] : o.relV === 'topMargin' || o.relV === 'insideMargin' ? [0, g.top] : [g.h - g.bottom, g.bottom]
      const y = alignIn(o.alignV, vs, va, o.h, o.offV)
      vertical = story === 'footer' ? `bottom:${cq(g.h - y - o.h)}` : `top:${cq(y)}`
      if (story === 'body' && this.page.index > 0) {
        this.page.anchored = true
        this.page.bottom = Math.max(this.page.bottom, clamp(y + o.h, 0, g.h))
      }
    }
    // Art behind the text or as wide as half the column (covers, frames, watermarks) needs the real margins; a corner logo does not.
    if (o.behind || o.w >= (g.w - g.left - g.right) / 2) this.placedOnSheet = true
    const rotate = o.rotate && Number.isFinite(o.rotate) ? `;transform:rotate(${round(o.rotate % 360)}deg)` : ''
    const style = `left:${left};${vertical};width:${cq(o.w)};${o.h > 0 ? `height:${cq(o.h)};` : ''}z-index:${o.behind ? -1 : 1}${rotate}`
    return { open: `<span class="dx-abs" style="${style}">`, pin }
  }

  /** Where a paragraph moves to a new page: at its start (page break before, or a break ahead of any content) or after it. */
  private pageBreak(info: ParaInfo, first: boolean): 'start' | 'end' | null {
    if (info.props.pageBreakBefore && !first) return 'start'
    let seen = false
    let found: 'start' | 'end' | null = null
    let budget = 5000
    const walk = (nodes: XmlNode[]) => {
      for (const node of nodes) {
        if (found || --budget < 0) return
        const name = local(node.name)
        if (name === 'br' && attr(node, 'type') === 'page') {
          found = seen ? 'end' : 'start'
          return
        }
        if (name === 'drawing' || name === 'pict' || name === 'object' || name === 'sym' || name === 'tab') seen = true
        else if (name === 't') seen ||= node.text.trim() !== ''
        else if (!NOT_ON_PAGE.has(name)) walk(node.children)
      }
    }
    walk(info.node.children)
    return found
  }

  /** A finished page segment: a later page that holds placed drawings gets its own box, a page tall and with its top margin. */
  private closePage(html: string, g: Geometry): string {
    const { index, anchored, bottom } = this.page
    this.page.anchored = false
    this.page.bottom = 0
    if (!index || !anchored || !html) return html
    const cq = (value: number) => `${round((value / g.w) * 100, 3)}cqw`
    return `<div class="dx-pg" style="padding-top:${cq(g.top)};min-height:${cq(Math.max(bottom, g.top))}">${html}</div>`
  }

  /** Where a floating drawing goes: placed (see `place`) when the text does not wrap it or it is behind the text, floated for square wrapping, else inline. */
  private anchorBox(anchor: XmlNode, w: number, h: number, scope: Scope): { open: string; mode: 'abs' | 'pin' | 'flow' } | null {
    const wrap = anchor.children.find((c) => local(c.name).startsWith('wrap'))
    const wrapName = wrap ? local(wrap.name) : 'wrapNone'
    const behind = attr(anchor, 'behindDoc') === '1' || attr(anchor, 'behindDoc') === 'true'
    const posH = child(anchor, 'positionH')
    const posV = child(anchor, 'positionV')
    const relH = attr(posH, 'relativeFrom') ?? 'column'
    const relV = attr(posV, 'relativeFrom') ?? 'paragraph'
    const alignH = child(posH, 'align')?.text.trim()
    const alignV = child(posV, 'align')?.text.trim()
    const offH = Number(child(posH, 'posOffset')?.text.trim() ?? '0') / EMU_PER_PT || 0
    const offV = Number(child(posV, 'posOffset')?.text.trim() ?? '0') / EMU_PER_PT || 0
    const g = scope.sheet
    const horizontal = (): number | null => {
      if (!g) return null
      const [start, area] =
        relH === 'page' ? [0, g.w] : relH === 'leftMargin' ? [0, g.left] : relH === 'rightMargin' ? [g.w - g.right, g.right] : relH === 'margin' || relH === 'column' ? [g.left, g.w - g.left - g.right] : [null, 0]
      return start === null ? null : alignIn(alignH, start, area, w, offH)
    }
    if (wrapName === 'wrapNone' || behind) {
      const simple = attr(anchor, 'simplePos') === '1' || attr(anchor, 'simplePos') === 'true'
      const at = child(anchor, 'simplePos')
      const placement: Placement = simple
        ? { relH: 'page', relV: 'page', offH: (num(at, 'x') ?? 0) / EMU_PER_PT, offV: (num(at, 'y') ?? 0) / EMU_PER_PT, w, h, behind }
        : { relH, relV, alignH, alignV, offH, offV, w, h, behind }
      const placed = this.place(placement, scope)
      if (placed) return { open: placed.open, mode: placed.pin ? 'pin' : 'abs' }
    }
    // A paragraph offset moves the drawing down from the paragraph's top; other references are approximated by that top.
    const down = relV === 'paragraph' && !alignV ? clamp(offV, 0, 1000) : 0
    const dist = (name: string) => pt(clamp((num(anchor, name) ?? 0) / EMU_PER_PT, 0, 72) + (name === 'distT' ? down : 0))
    const lower = down > 0 ? `;margin-top:${pt(down)}` : ''
    if (wrapName === 'wrapSquare' || wrapName === 'wrapTight' || wrapName === 'wrapThrough') {
      let side = alignH === 'right' || alignH === 'outside' ? 'right' : alignH === 'center' ? '' : 'left'
      if (!alignH && g) {
        const x = horizontal()
        if (x !== null && x + w / 2 > g.w / 2) side = 'right'
      }
      if (side) return { open: `<span class="dx-float" style="float:${side};margin:${dist('distT')} ${dist('distR')} ${dist('distB')} ${dist('distL')}">`, mode: 'flow' }
      return { open: `<span class="dx-figure" style="text-align:center${lower}">`, mode: 'flow' }
    }
    if (wrapName === 'wrapTopAndBottom') {
      const align = alignH === 'center' ? 'center' : alignH === 'right' || alignH === 'outside' ? 'right' : 'left'
      return { open: `<span class="dx-figure" style="text-align:${align}${lower}">`, mode: 'flow' }
    }
    return null
  }

  private graphic(node: XmlNode | null, w: number, h: number, scope: Scope, alt: string, fill: boolean): string {
    if (!node) return unsupported('圖形', w, h)
    const size = fill ? 'width:100%;height:100%' : `width:${pt(w)};max-width:100%`
    switch (local(node.name)) {
      case 'pic':
        return this.picture(node, scope, size, w, h, alt, fill)
      case 'wsp':
        return this.shape(node, scope, size, w, h, fill)
      case 'wgp':
      case 'grpSp':
      case 'wpc':
        return this.group(node, scope, size, w, h)
      case 'chart':
        return this.chart(node, scope, size, w, h, fill)
      case 'chartex':
        return unsupported('圖表', w, h)
      case 'relIds':
        return unsupported('SmartArt 圖形', w, h)
      default:
        return unsupported('圖形', w, h)
    }
  }

  /** A chart part (`c:chart` in a drawing) drawn by the shared chart renderer at the drawing's extent; a labelled box when it cannot be drawn. */
  private chart(node: XmlNode, scope: Scope, size: string, w: number, h: number, fill: boolean): string {
    const rel = this.pkg.rel(scope.story.part, relAttr(node, 'id'))
    const xml = rel && !rel.external && this.pkg.zip.has(rel.target) ? this.pkg.xml(rel.target) : null
    if (!xml) return unsupported('圖表', w, h, '檔案裡找不到這張圖表')
    const drawn = w > 0 && h > 0
      ? renderChart(xml, {
          widthPt: w,
          heightPt: h,
          resolveColor: (holder) => {
            const color = holder?.children.map((c) => this.drawingColor(c)).find(Boolean)
            return color ? cssColor(color) : null
          },
          themeColor: (i) => `#${this.theme.colors[`accent${i + 1}`] ?? OFFICE_ACCENTS[i % 6]}`,
          fontFamily: this.fontStack({}),
          // Office's chart text is the text colour at 65% (#595959 on black text).
          textColor: `#${!this.theme.colors.dk1 || this.theme.colors.dk1.toUpperCase() === '000000' ? '595959' : this.theme.colors.dk1}`
        })
      : null
    const title = chartTitle(xml)
    if (!drawn) return unsupported(title ? `圖表「${title}」` : '圖表', w, h)
    const ratio = !fill ? `;aspect-ratio:${round(w)}/${round(h)}` : ''
    return `<span class="dx-chart" style="${size}${ratio}">${drawn.svg}</span>`
  }

  /** The box for a picture that cannot be shown (missing, linked from outside the file, or in a format browsers do not draw); null when it can. */
  private missingPicture(rel: Rel | undefined, w: number, h: number): string | null {
    if (rel?.external) return unsupported(EXTERNAL_PICTURE[0], w, h, EXTERNAL_PICTURE[1])
    if (!rel || !this.pkg.zip.has(rel.target)) return unsupported(MISSING_PICTURE[0], w, h, MISSING_PICTURE[1])
    if (isDrawablePart(rel.target)) return null
    const ext = posix.extname(rel.target).slice(1).toUpperCase().replace(/[^A-Z0-9]/g, '').slice(0, 8)
    return unsupported(ext ? `${ext} 圖片` : '圖片', w, h)
  }

  private picture(pic: XmlNode, scope: Scope, size: string, w: number, h: number, alt: string, fill: boolean): string {
    const blip = path(pic, 'blipFill', 'blip')
    const rel = this.pkg.rel(scope.story.part, relAttr(blip, 'embed'))
    const missing = this.missingPicture(rel, w, h)
    if (missing !== null) return missing
    const xfrm = path(pic, 'spPr', 'xfrm')
    const transforms: string[] = []
    const rot = num(xfrm, 'rot')
    if (rot) transforms.push(`rotate(${round(mirrored(scope) ? -rot / 60000 : rot / 60000)}deg)`)
    if ((attr(xfrm, 'flipH') === '1') !== Boolean(scope.flipH)) transforms.push('scaleX(-1)')
    if ((attr(xfrm, 'flipV') === '1') !== Boolean(scope.flipV)) transforms.push('scaleY(-1)')
    const ratio = !fill && w > 0 && h > 0 ? `;aspect-ratio:${round(w)}/${round(h)}` : ''
    const transform = transforms.length ? `;transform:${transforms.join(' ')}` : ''
    return `<img class="dx-img" src="${esc(this.media(rel!.target))}" alt="${esc(alt)}" style="${size}${ratio}${transform}">`
  }

  private shape(wsp: XmlNode, scope: Scope, size: string, w: number, h: number, fill: boolean): string {
    const spPr = child(wsp, 'spPr')
    const style = child(wsp, 'style')
    let paint = this.fillPaint(spPr, scope)
    if (paint === undefined) {
      const ref = child(style, 'fillRef')
      const color = ref && (num(ref, 'idx') ?? 0) > 0 ? this.drawingColor(ref.children[0]) : null
      paint = color ? { color } : null
    }
    const background = paint ? paintTone(paint) : null
    const ln = child(spPr, 'ln')
    let border: string | null = null
    let lineColor: Rgba | null = null
    const lineFill = this.fillColor(ln)
    const lineWidth = clamp((num(ln, 'w') ?? 9525) / EMU_PER_PT, 0.25, 20)
    if (lineFill) lineColor = lineFill
    else if (lineFill === undefined) {
      const ref = child(style, 'lnRef')
      lineColor = ref && (num(ref, 'idx') ?? 0) > 0 ? this.drawingColor(ref.children[0]) : null
    }
    if (lineColor) border = `${pt(lineWidth)} solid ${cssColor(lineColor)}`
    const geomNode = child(spPr, 'prstGeom') ?? child(spPr, 'custGeom')
    const geometry = geomNode ? (local(geomNode.name) === 'custGeom' ? 'custom' : attr(geomNode, 'prst') ?? 'rect') : 'rect'
    const content = path(wsp, 'txbx', 'txbxContent')
    if ((geometry === 'line' || geometry === 'straightConnector1') && !content) {
      return border ? `<span class="dx-line" style="${fill ? 'width:100%' : `width:${pt(w)};max-width:100%`};border-top:${border}"></span>` : ''
    }
    if (!content && !paint && !border) return ''
    // A plain rectangle stays a CSS box; every other shape is its ECMA-376 geometry drawn behind the text, and the
    // text sits in the geometry's text rectangle. A preset the standard does not define keeps a box and says so.
    const shape = geometry === 'rect' ? null : shapeGeometry(spPr, w * EMU_PER_PT, h * EMU_PER_PT)
    const bodyPr = child(wsp, 'bodyPr')
    const rect = shape && w > 0 && h > 0 ? shape.textRect : { l: 0, t: 0, r: w, b: h }
    const inset = (name: string, fallback: number, extra: number) => pt(clamp((num(bodyPr, name) ?? fallback) / EMU_PER_PT, 0, 100) + clamp(extra, 0, 1000))
    const xfrm = path(spPr, 'xfrm')
    const vert = attr(bodyPr, 'vert')
    // vert270 turns the whole box over (its text reads bottom to top); the geometry turns back, and a flip mirrors
    // only the geometry, so the text stays readable and sits in the text rectangle the flip moved.
    const turned = vert === 'vert270'
    const flipH = (attr(xfrm, 'flipH') === '1') !== Boolean(scope.flipH)
    const flipV = (attr(xfrm, 'flipV') === '1') !== Boolean(scope.flipV)
    const swapX = flipH !== turned
    const swapY = flipV !== turned
    const css = [size, 'box-sizing:border-box']
    if (!fill && h) css.push(`min-height:${pt(h)}`)
    const [padL, padR] = swapX ? [w - rect.r, rect.l] : [rect.l, w - rect.r]
    const [padT, padB] = swapY ? [h - rect.b, rect.t] : [rect.t, h - rect.b]
    css.push(`padding:${inset('tIns', 45720, padT)} ${inset('rIns', 91440, padR)} ${inset('bIns', 45720, padB)} ${inset('lIns', 91440, padL)}`)
    let svg = ''
    if (shape) {
      css.push('position:relative;z-index:0')
      let paths = ''
      let defs = ''
      if (paint) {
        const svgFill = this.svgPaint(paint, w, h)
        defs = svgFill.defs
        for (const p of shape.paths) {
          if (p.fill === 'none' || (!p.closed && geometry === 'custom')) continue
          paths += `<path d="${p.d}" fill="${svgFill.fill}"/>`
          const shade = fillShade(p.fill)
          if (shade) paths += `<path d="${p.d}" fill="${shade.color}" fill-opacity="${shade.opacity}"/>`
        }
      }
      if (lineColor) for (const p of shape.paths) if (p.stroke) paths += `<path d="${p.d}" fill="none" stroke="${cssColor(lineColor)}" stroke-width="${round(lineWidth)}" stroke-linejoin="round"/>`
      const geomTransform = [turned ? 'rotate(180deg)' : '', flipH ? 'scaleX(-1)' : '', flipV ? 'scaleY(-1)' : ''].filter(Boolean).join(' ')
      const geomStyle = geomTransform ? ` style="transform:${geomTransform}"` : ''
      if (paths) svg = `<svg class="dx-geom" viewBox="0 0 ${round(Math.max(w, 0.01))} ${round(Math.max(h, 0.01))}" preserveAspectRatio="none" aria-hidden="true"${geomStyle}>${defs ? `<defs>${defs}</defs>` : ''}${paths}</svg>`
    } else {
      if (paint) css.push(paintCss(paint))
      if (border) css.push(`border:${border}`)
    }
    const unknown = geometry !== 'rect' && geometry !== 'custom' && !shape ? unsupported(`未支援的圖形:${geometry}`) : ''
    const anchor = attr(bodyPr, 'anchor')
    if (anchor === 'ctr') css.push('justify-content:center')
    else if (anchor === 'b') css.push('justify-content:flex-end')
    if (vert === 'vert' || vert === 'eaVert' || vert === 'wordArtVert' || turned) css.push('writing-mode:vertical-rl')
    const rot = num(xfrm, 'rot')
    const transforms = [rot ? `rotate(${round(mirrored(scope) ? -rot / 60000 : rot / 60000)}deg)` : '', turned ? 'rotate(180deg)' : ''].filter(Boolean)
    if (transforms.length) css.push(`transform:${transforms.join(' ')}`)
    let inner = ''
    if (content) {
      if (scope.depth >= LIMITS.depth) inner = unsupported('多層巢狀文字方塊')
      else {
        const innerScope: Scope = { ...scope, layer: EMPTY_LAYER, depth: scope.depth + 1, width: Math.max(36, rect.r - rect.l - 14), dark: background ? isDark(background) : scope.dark, sheet: null, groupFill: null, flipH: false, flipV: false }
        inner = this.renderBlocks(content.children, innerScope)
      }
    }
    return `<span class="dx-txbx" style="${css.join(';')}">${svg}${inner}${unknown}</span>`
  }

  private group(node: XmlNode, scope: Scope, size: string, w: number, h: number): string {
    if (scope.depth >= LIMITS.depth) return unsupported('多層群組圖形', w, h)
    const canvas = local(node.name) === 'wpc'
    const xfrm = path(node, 'grpSpPr', 'xfrm')
    const chOff = child(xfrm, 'chOff')
    const chExt = child(xfrm, 'chExt')
    const ox = canvas ? 0 : (num(chOff, 'x') ?? num(child(xfrm, 'off'), 'x') ?? 0)
    const oy = canvas ? 0 : (num(chOff, 'y') ?? num(child(xfrm, 'off'), 'y') ?? 0)
    const ew = canvas ? w * EMU_PER_PT : (num(chExt, 'cx') ?? num(child(xfrm, 'ext'), 'cx') ?? w * EMU_PER_PT)
    const eh = canvas ? h * EMU_PER_PT : (num(chExt, 'cy') ?? num(child(xfrm, 'ext'), 'cy') ?? h * EMU_PER_PT)
    if (!ew || !eh) return ''
    let items = ''
    let count = 0
    // A member that takes the group's fill takes this group's, or the enclosing group's when this one does too.
    const own = child(node, 'grpSpPr')
    const groupFill = own && !child(own, 'grpFill') ? own : scope.groupFill
    // A flipped group mirrors where its members sit and how each is drawn; text inside stays readable.
    const flipH = Boolean(scope.flipH) !== (attr(xfrm, 'flipH') === '1')
    const flipV = Boolean(scope.flipV) !== (attr(xfrm, 'flipV') === '1')
    const inner: Scope = { ...scope, depth: scope.depth + 1, sheet: null, groupFill, flipH, flipV }
    for (const c of node.children) {
      const name = local(c.name)
      if (!['wsp', 'pic', 'grpSp', 'wgp', 'graphicFrame'].includes(name)) continue
      if (++count > 500) break
      const childXfrm = name === 'grpSp' || name === 'wgp' ? path(c, 'grpSpPr', 'xfrm') : (path(c, 'spPr', 'xfrm') ?? child(c, 'xfrm'))
      const off = child(childXfrm, 'off')
      const ext = child(childXfrm, 'ext')
      const cx = num(ext, 'cx') ?? 0
      const cy = num(ext, 'cy') ?? 0
      const x0 = (((num(off, 'x') ?? 0) - ox) / ew) * 100
      const y0 = (((num(off, 'y') ?? 0) - oy) / eh) * 100
      const left = flipH ? 100 - x0 - (cx / ew) * 100 : x0
      const top = flipV ? 100 - y0 - (cy / eh) * 100 : y0
      const box = `left:${round(left, 3)}%;top:${round(top, 3)}%;width:${round((cx / ew) * 100, 3)}%;height:${round((cy / eh) * 100, 3)}%`
      const target = name === 'graphicFrame' ? (path(c, 'graphic', 'graphicData')?.children[0] ?? null) : c
      let html: string
      try {
        html = this.graphic(target, (cx / ew) * w, (cy / eh) * h, inner, '', true)
      } catch {
        html = unsupported('圖形')
      }
      if (html) items += `<span class="dx-gi" style="${box}">${html}</span>`
    }
    const ratio = w > 0 && h > 0 ? `;aspect-ratio:${round(w)}/${round(h)}` : ''
    return `<span class="dx-grp" style="${size}${ratio}">${items}</span>`
  }

  private vml(node: XmlNode, scope: Scope, out: Inline): string {
    return node.children.map((c) => this.vmlShape(c, scope, 0, out)).join('')
  }

  /** A VML shape (older files, watermarks, fallbacks). One positioned absolutely on the page is placed like a DrawingML anchor. */
  private vmlShape(node: XmlNode, scope: Scope, depth: number, out: Inline): string {
    if (depth > 10) return ''
    const name = local(node.name)
    if (name === 'group') return node.children.map((c) => this.vmlShape(c, scope, depth + 1, out)).join('')
    if (!['shape', 'rect', 'roundrect', 'oval', 'line', 'image', 'polyline', 'arc', 'curve'].includes(name)) return ''
    const style: Record<string, string> = {}
    for (const declaration of (attr(node, 'style') ?? '').split(';')) {
      const at = declaration.indexOf(':')
      if (at > 0) style[declaration.slice(0, at).trim().toLowerCase()] = declaration.slice(at + 1).trim()
    }
    if (style.visibility === 'hidden') return ''
    const length = (value: string | undefined) => Math.max(0, vmlLength(value))
    const w = length(style.width)
    const h = length(style.height)
    let placed: { open: string; pin: boolean } | null = null
    const wrapType = attr(child(node, 'wrap'), 'type')
    if (style.position === 'absolute' && depth === 0 && w > 0 && !['square', 'tight', 'through', 'topAndBottom'].includes(wrapType ?? '')) {
      const H: Record<string, string> = { margin: 'margin', page: 'page', 'left-margin-area': 'leftMargin', 'right-margin-area': 'rightMargin', 'inner-margin-area': 'leftMargin', 'outer-margin-area': 'rightMargin' }
      const V: Record<string, string> = { margin: 'margin', page: 'page', 'top-margin-area': 'topMargin', 'bottom-margin-area': 'bottomMargin', 'inner-margin-area': 'topMargin', 'outer-margin-area': 'bottomMargin' }
      const align = (value: string | undefined) => (value && value !== 'absolute' ? value : undefined)
      const rotation = /^(-?[\d.]+)(fd)?$/.exec(style.rotation ?? '')
      placed = this.place(
        {
          relH: H[style['mso-position-horizontal-relative'] ?? ''] ?? 'column',
          relV: V[style['mso-position-vertical-relative'] ?? ''] ?? 'paragraph',
          alignH: align(style['mso-position-horizontal']),
          alignV: align(style['mso-position-vertical']),
          offH: vmlLength(style['margin-left']) + vmlLength(style.left),
          offV: vmlLength(style['margin-top']) + vmlLength(style.top),
          w,
          h,
          behind: Number.parseInt(style['z-index'] ?? '', 10) < 0,
          rotate: rotation ? Number(rotation[1]) / (rotation[2] ? 65536 : 1) : undefined
        },
        scope
      )
    }
    const html = this.vmlContent(node, name, scope, w, h, placed !== null)
    if (!placed && html && style.position === 'absolute' && depth === 0 && ['square', 'tight', 'through', 'topAndBottom'].includes(wrapType ?? '')) {
      // Wrapped like a DrawingML anchor: floated to its side at the paragraph's top, or on a line of its own.
      const alignH = style['mso-position-horizontal']
      const left = vmlLength(style['margin-left']) + vmlLength(style.left)
      const toRight = alignH === 'right' || alignH === 'outside' || (alignH !== 'left' && alignH !== 'inside' && alignH !== 'center' && left + w / 2 > scope.width / 2)
      const side = wrapType === 'topAndBottom' || alignH === 'center' ? '' : toRight ? 'right' : 'left'
      // Above and below the text: on a line of its own, at its side as DrawingML's top-and-bottom wrap is.
      const align = alignH === 'center' ? 'center' : toRight ? 'right' : 'left'
      const down = ['text', 'paragraph'].includes(style['mso-position-vertical-relative'] ?? 'text') && !['top', 'center', 'bottom'].includes(style['mso-position-vertical'] ?? '') ? clamp(vmlLength(style['margin-top']) + vmlLength(style.top), 0, 1000) : 0
      const open = side ? `<span class="dx-float" style="float:${side};margin:${pt(down)} 9pt 0pt 9pt">` : `<span class="dx-figure" style="text-align:${align}${down > 0 ? `;margin-top:${pt(down)}` : ''}">`
      out.floats.push(`${open}${html}</span>`)
      out.visual = true
      return ''
    }
    if (!placed || !html) return html
    ;(placed.pin ? out.pins : out.parts).push(`${placed.open}${html}</span>`)
    return ''
  }

  /** What a VML shape shows; `fill` when it is placed in a box of its own size. */
  private vmlContent(node: XmlNode, name: string, scope: Scope, w: number, h: number, fill: boolean): string {
    const vmlColor = (value: string | undefined) => {
      const token = (value ?? '').trim().split(/\s/)[0].replace(/^#/, '')
      return hexToRgba(token.length === 3 ? token.replace(/./g, '$&$&') : token) ?? hexToRgba(PRESET_COLORS[token.toLowerCase()])
    }
    const width = fill ? 'width:100%;height:100%' : w ? `width:${pt(w)};max-width:100%` : ''
    const imagedata = child(node, 'imagedata')
    if (imagedata) {
      const rel = this.pkg.rel(scope.story.part, relAttr(imagedata, 'id') ?? relAttr(imagedata, 'relid'))
      const missing = this.missingPicture(rel, w, h)
      if (missing !== null) return missing
      const ratio = !fill && w > 0 && h > 0 ? `;aspect-ratio:${round(w)}/${round(h)}` : ''
      return `<img class="dx-img" src="${esc(this.media(rel!.target))}" alt="${esc(attr(imagedata, 'title') ?? '')}" style="${width || 'max-width:100%'}${ratio}">`
    }
    const textpath = child(node, 'textpath')
    if (textpath) {
      const text = cleanText(attr(textpath, 'string') ?? '')
      if (!text) return ''
      const color = vmlColor(attr(node, 'fillcolor')) ?? { r: 192, g: 192, b: 192, a: 1 }
      const opacity = Number(attr(child(node, 'fill'), 'opacity')?.replace(/f$/, '') ?? '')
      const alpha = Number.isFinite(opacity) && opacity > 0 ? (opacity > 1 ? opacity / 65536 : opacity) : 0.5
      this.text.push(text)
      const ink = `color:${cssColor({ ...color, a: clamp(alpha, 0.1, 1) })}`
      if (fill && scope.sheet) {
        // Word stretches the text to its box: as large as fits the width (wide characters count whole), within the height.
        const units = [...text].reduce((sum, char) => sum + (char.codePointAt(0)! > 0x2e80 ? 1 : 0.6), 0)
        const size = clamp(Math.min(h * 0.85 || 54, w / Math.max(units, 1)), 4, 400)
        return `<span class="dx-wordart dx-fit" style="${ink};font-size:${round((size / scope.sheet.w) * 100, 3)}cqw">${esc(text)}</span>`
      }
      return `<span class="dx-wordart" style="${ink};font-size:${pt(clamp(h * 0.6 || 24, 8, 54))}">${esc(text)}</span>`
    }
    const content = path(node, 'textbox', 'txbxContent')
    const filled = attr(node, 'filled') !== 'f' && attr(node, 'filled') !== 'false'
    const background = filled ? vmlColor(attr(node, 'fillcolor')) : null
    const stroked = attr(node, 'stroked') !== 'f' && attr(node, 'stroked') !== 'false'
    const strokeColor = vmlColor(attr(node, 'strokecolor')) ?? BLACK
    const strokeWidth = clamp(vmlLength(attr(node, 'strokeweight') ?? '0.75pt'), 0, 4000) || 0.75
    if (attr(node, 'hr') === 't' || name === 'line') {
      return `<span class="dx-line" style="width:${fill ? '100%' : w ? pt(w) : '100%'};max-width:100%;border-top:${pt(Math.max(fill ? 0 : h, strokeWidth, 0.75))} solid ${cssColor(background ?? strokeColor)}"></span>`
    }
    if (!content) {
      if (!background || !w || !h) return ''
      return `<span class="dx-txbx" style="${fill ? width : `width:${pt(w)};max-width:100%;min-height:${pt(h)}`};background:${cssColor(background)}"></span>`
    }
    const css = [width || 'width:100%', 'box-sizing:border-box', 'padding:3.6pt 7.2pt']
    if (h && !fill) css.push(`min-height:${pt(h)}`)
    if (background) css.push(`background:${cssColor(background)}`)
    if (stroked && name !== 'shape') css.push(`border:${pt(strokeWidth)} solid ${cssColor(strokeColor)}`)
    else if (stroked && attr(node, 'strokecolor')) css.push(`border:${pt(strokeWidth)} solid ${cssColor(strokeColor)}`)
    if (name === 'oval') css.push('border-radius:50%')
    if (scope.depth >= LIMITS.depth) return unsupported('多層巢狀文字方塊')
    const inner = this.renderBlocks(content.children, { ...scope, layer: EMPTY_LAYER, depth: scope.depth + 1, width: Math.max(36, (w || scope.width) - 14), dark: background ? isDark(background) : scope.dark, sheet: null, flipH: false, flipV: false })
    return `<span class="dx-txbx" style="${css.join(';')}">${inner}</span>`
  }

  // --- tables --------------------------------------------------------------------------------

  private renderTable(tbl: XmlNode, scope: Scope, marginTop: number): string {
    if (scope.depth >= LIMITS.depth) return `<div class="dx-p">${unsupported('多層巢狀表格')}</div>`
    const tblPr = child(tbl, 'tblPr')
    const styleId = attr(child(tblPr, 'tblStyle'), 'val')
    const chain = this.chain(styleId, 'table')
    const tblPrs = [...chain.map((style) => style.tblPr), tblPr]
    const tableBorders: Record<string, XmlNode> = {}
    const cellMargin: Record<string, number> = { top: 0, bottom: 0, left: 5.4, right: 5.4 }
    const marginInto = (target: Record<string, number>, holder: XmlNode | null) => {
      for (const side of holder?.children ?? []) {
        const key = { start: 'left', end: 'right' }[local(side.name)] ?? local(side.name)
        const width = num(side, 'w')
        if (width !== undefined && (attr(side, 'type') ?? 'dxa') === 'dxa') target[key] = clamp(width / 20, 0, 100)
      }
    }
    let tableShading: Rgba | null | undefined
    for (const props of tblPrs) {
      for (const side of child(props, 'tblBorders')?.children ?? []) tableBorders[{ start: 'left', end: 'right' }[local(side.name)] ?? local(side.name)] = side
      marginInto(cellMargin, child(props, 'tblCellMar'))
      const shd = this.shading(child(props, 'shd'))
      if (shd !== undefined) tableShading = shd
    }
    const lookNode = child(tblPr, 'tblLook')
    const lookBits = parseInt(attr(lookNode, 'val') ?? '', 16)
    const flag = (name: string, bit: number, fallback: boolean) => {
      const value = attr(lookNode, name)
      if (value !== undefined) return value === '1' || value === 'true' || value === 'on'
      if (Number.isFinite(lookBits)) return Boolean(lookBits & bit)
      return fallback
    }
    const look = {
      firstRow: flag('firstRow', 0x20, true), lastRow: flag('lastRow', 0x40, false), firstColumn: flag('firstColumn', 0x80, true),
      lastColumn: flag('lastColumn', 0x100, false), noHBand: flag('noHBand', 0x200, false), noVBand: flag('noVBand', 0x400, true)
    }
    let rowBand = 1
    let colBand = 1
    for (const props of tblPrs) {
      rowBand = num(child(props, 'tblStyleRowBandSize'), 'val') ?? rowBand
      colBand = num(child(props, 'tblStyleColBandSize'), 'val') ?? colBand
    }
    rowBand = Math.max(1, rowBand)
    colBand = Math.max(1, colBand)
    const condOf = (type: string): StyleCond[] => chain.map((style) => style.conds.get(type)).filter((cond): cond is StyleCond => Boolean(cond))

    // Layout: which grid columns each cell covers, so vertical merges line up across rows.
    const grid = children(child(tbl, 'tblGrid'), 'gridCol').map((col) => clamp((num(col, 'w') ?? 0) / 20, 0, 2000))
    const rowNodes: XmlNode[] = []
    const collect = (nodes: XmlNode[], depth: number, name: string, into: XmlNode[]) => {
      for (const node of nodes) {
        const tag = local(node.name)
        if (tag === name) into.push(node)
        else if (depth < 5 && tag === 'sdt') collect(child(node, 'sdtContent')?.children ?? [], depth + 1, name, into)
        else if (depth < 5 && (tag === 'customXml' || tag === 'ins' || tag === 'moveTo')) collect(node.children, depth + 1, name, into)
      }
    }
    collect(tbl.children, 0, 'tr', rowNodes)
    type Cell = { tc: XmlNode; col: number; span: number; merge: 'restart' | 'continue' | null; rowspan: number }
    const rows: { tr: XmlNode; cells: Cell[]; before: number; after: number }[] = []
    for (const tr of rowNodes) {
      const trPr = child(tr, 'trPr')
      if (child(trPr, 'del')) continue
      const before = clamp(num(child(trPr, 'gridBefore'), 'val') ?? 0, 0, 63)
      const after = clamp(num(child(trPr, 'gridAfter'), 'val') ?? 0, 0, 63)
      const tcs: XmlNode[] = []
      collect(tr.children, 0, 'tc', tcs)
      let col = before
      const cells: Cell[] = []
      for (const tc of tcs) {
        const tcPr = child(tc, 'tcPr')
        const span = clamp(num(child(tcPr, 'gridSpan'), 'val') ?? 1, 1, 63)
        const hMerge = attr(child(tcPr, 'hMerge'), 'val')
        if (child(tcPr, 'hMerge') && hMerge !== 'restart' && cells.length) {
          cells[cells.length - 1].span += span
          col += span
          continue
        }
        const vMerge = child(tcPr, 'vMerge')
        const merge = vMerge ? (attr(vMerge, 'val') === 'restart' ? 'restart' : 'continue') : null
        cells.push({ tc, col, span, merge, rowspan: 1 })
        col += span
      }
      rows.push({ tr, cells, before, after })
    }
    const columnCount = Math.max(grid.length, ...rows.map((row) => (row.cells.length ? row.cells[row.cells.length - 1].col + row.cells[row.cells.length - 1].span : 0) + row.after), 1)
    // A continuation with nothing to continue (no restart above it) is drawn as an ordinary cell.
    const covered = new Set<Cell>()
    for (let r = 0; r < rows.length; r++) {
      for (const cell of rows[r].cells) {
        if (cell.merge === 'continue' && !covered.has(cell)) cell.merge = 'restart'
        if (cell.merge !== 'restart') continue
        for (let below = r + 1; below < rows.length; below++) {
          const match = rows[below].cells.find((candidate) => candidate.col === cell.col)
          if (!match || match.merge !== 'continue') break
          covered.add(match)
          cell.rowspan++
        }
      }
    }

    // Widths: the grid as fractions of the table, the table as a share of the space it sits in.
    const gridTotal = grid.reduce((sum, width) => sum + width, 0)
    const tblW = child(tblPr, 'tblW')
    const tblWType = attr(tblW, 'type') ?? 'auto'
    const tblWValue = attr(tblW, 'w') ?? ''
    let widthCss = ''
    if (tblWType === 'pct') {
      const value = tblWValue.endsWith('%') ? parseFloat(tblWValue) : Number(tblWValue) / 50
      if (Number.isFinite(value) && value > 0) widthCss = `width:${round(clamp(value, 1, 100))}%`
    } else if (tblWType === 'dxa' || scope.depth === 0) {
      // A nested table with an automatic width sizes to its content, as Word's autofit does.
      const points = tblWType === 'dxa' && Number(tblWValue) > 0 ? Number(tblWValue) / 20 : gridTotal
      if (points > 0) widthCss = `width:${round(clamp((points / scope.width) * 100, 5, 100))}%`
    }
    const columns = Math.max(grid.length, columnCount)
    const tableCss = [widthCss, 'border-collapse:collapse']
    if (columns > 3) tableCss.push(`min-width:${round(Math.min(columns * 3.2, 60))}em`)
    if (attr(child(tblPr, 'tblLayout'), 'type') === 'fixed') tableCss.push('table-layout:fixed')
    if (child(tblPr, 'bidiVisual') && truthy(child(tblPr, 'bidiVisual'))) tableCss.push('direction:rtl')
    const jc = attr(child(tblPr, 'jc'), 'val') ?? attr(child(chain.length ? chain[chain.length - 1].tblPr : null, 'jc'), 'val')
    const wrapCss: string[] = []
    if (marginTop) wrapCss.push(`margin-top:${pt(marginTop)}`)
    if (jc === 'center') tableCss.push('margin-inline:auto')
    else if (jc === 'right' || jc === 'end') tableCss.push('margin-inline-start:auto')
    else {
      let indent: number | undefined
      for (const props of tblPrs) {
        const node = child(props, 'tblInd')
        if (node && (attr(node, 'type') ?? 'dxa') === 'dxa') indent = (num(node, 'w') ?? 0) / 20
      }
      if (indent) tableCss.push(`margin-inline-start:${pt(clamp(indent, -72, 500))}`)
    }
    const colgroup = gridTotal > 0 ? `<colgroup>${grid.map((width) => `<col style="width:${round((width / gridTotal) * 100, 3)}%">`).join('')}</colgroup>` : ''

    let body = ''
    const lastRow = rows.length - 1
    let cutRows = 0
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r]
      if (this.cells >= LIMITS.cells || this.cellBlocks >= LIMITS.cellBlocks) {
        if (this.cells >= LIMITS.cells) this.cutCells = true
        else this.cutCellBlocks = true
        cutRows = rows.length - r
        break
      }
      const trPr = child(row.tr, 'trPr')
      const heightNode = child(trPr, 'trHeight')
      const height = num(heightNode, 'val')
      const rowCss = height ? `height:${pt(clamp(height / 20, 0, 1000))}` : ''
      let cellsHtml = row.before ? `<td colspan="${row.before}" class="dx-gap"></td>` : ''
      const rowConds: string[] = []
      const headerRow = look.firstRow && r === 0
      const footerRow = look.lastRow && r === lastRow && r > 0
      if (!look.noHBand && !headerRow && !footerRow) rowConds.push(Math.floor((r - (look.firstRow ? 1 : 0)) / rowBand) % 2 === 0 ? 'band1Horz' : 'band2Horz')
      for (const cell of row.cells) {
        if (cell.merge === 'continue') continue
        this.cells++
        const c = cell.col
        const lastCol = cell.col + cell.span >= columnCount
        const types = ['wholeTable']
        const firstCol = look.firstColumn && c === 0
        const finalCol = look.lastColumn && lastCol && c > 0
        if (!look.noVBand && !firstCol && !finalCol) types.push(Math.floor((c - (look.firstColumn ? 1 : 0)) / colBand) % 2 === 0 ? 'band1Vert' : 'band2Vert')
        types.push(...rowConds)
        if (firstCol) types.push('firstCol')
        if (finalCol) types.push('lastCol')
        if (headerRow) types.push('firstRow')
        if (footerRow) types.push('lastRow')
        if (headerRow && firstCol) types.push('nwCell')
        if (headerRow && finalCol) types.push('neCell')
        if (footerRow && firstCol) types.push('swCell')
        if (footerRow && finalCol) types.push('seCell')
        const typed = types.flatMap((type) => (type === 'wholeTable' ? [...chain.map((style) => ({ pPr: style.pPr, rPr: style.rPr, tblPr: null, tcPr: style.tcPr })), ...condOf('wholeTable')] : condOf(type)).map((cond) => ({ cond, type })))
        const conds = typed.map((entry) => entry.cond)
        const tcPr = child(cell.tc, 'tcPr')
        // Borders: the table's edges or inside lines, then each conditional format's, then the cell's own.
        const bottomRow = r + cell.rowspan - 1 >= lastRow
        const sides: Record<string, XmlNode | undefined> = {
          top: r === 0 ? tableBorders.top : tableBorders.insideH,
          bottom: bottomRow ? tableBorders.bottom : tableBorders.insideH,
          left: c === 0 ? tableBorders.left : tableBorders.insideV,
          right: lastCol ? tableBorders.right : tableBorders.insideV
        }
        // A conditional format's inside lines run only between the rows (insideH) or columns (insideV)
        // of its own region: a header or footer row has none, a band two rows tall has one between its
        // rows, and the band's own edges keep their rules. Corner cells have no inside edge at all.
        const rowGroup = (row: number) => (look.firstRow && row === 0 ? 'head' : look.lastRow && row === lastRow && row > 0 ? 'foot' : `b${Math.floor((row - (look.firstRow ? 1 : 0)) / rowBand)}`)
        const colGroup = (col: number) => (look.firstColumn && col === 0 ? 'first' : look.lastColumn && col === columnCount - 1 && col > 0 ? 'last' : `b${Math.floor((col - (look.firstColumn ? 1 : 0)) / colBand)}`)
        const applyBorders = (holder: XmlNode | null, region: string | null) => {
          const rowRegion = region !== null && ROW_REGIONS.has(region)
          const colRegion = region !== null && COL_REGIONS.has(region)
          const corner = region !== null && CORNER_REGIONS.has(region)
          for (const side of holder?.children ?? []) {
            const key = { start: 'left', end: 'right' }[local(side.name)] ?? local(side.name)
            if (key === 'insideH' && region !== null) {
              if (corner) continue
              const insideTop = r > 0 && (!rowRegion || rowGroup(r - 1) === rowGroup(r))
              const insideBottom = !bottomRow && (!rowRegion || rowGroup(r + cell.rowspan) === rowGroup(r))
              if (insideTop) sides.top = side
              if (insideBottom) sides.bottom = side
            } else if (key === 'insideV' && region !== null) {
              if (corner) continue
              const insideLeft = c > 0 && (!colRegion || colGroup(c - 1) === colGroup(c))
              const insideRight = !lastCol && (!colRegion || colGroup(c + cell.span) === colGroup(c))
              if (insideLeft) sides.left = side
              if (insideRight) sides.right = side
            } else if (key in sides) sides[key] = side
          }
        }
        for (const { cond, type } of typed) {
          applyBorders(child(cond.tblPr, 'tblBorders'), type)
          applyBorders(child(cond.tcPr, 'tcBorders'), type)
        }
        applyBorders(child(tcPr, 'tcBorders'), null)
        let shading: Rgba | null | undefined = tableShading
        for (const cond of conds) {
          const shd = this.shading(child(cond.tcPr, 'shd'))
          if (shd !== undefined) shading = shd
        }
        const ownShading = this.shading(child(tcPr, 'shd'))
        if (ownShading !== undefined) shading = ownShading
        const margin = { ...cellMargin }
        for (const cond of conds) marginInto(margin, child(cond.tcPr, 'tcMar'))
        marginInto(margin, child(tcPr, 'tcMar'))
        let vAlign: string | undefined
        for (const cond of conds) vAlign = attr(child(cond.tcPr, 'vAlign'), 'val') ?? vAlign
        vAlign = attr(child(tcPr, 'vAlign'), 'val') ?? vAlign
        const css: string[] = [`padding:${pt(margin.top)} ${pt(margin.right)} ${pt(margin.bottom)} ${pt(margin.left)}`]
        for (const side of ['top', 'right', 'bottom', 'left']) {
          const value = this.borderCss(sides[side])
          if (value) css.push(`border-${side}:${value}`)
        }
        if (shading) css.push(`background-color:${cssColor(shading)}`)
        if (vAlign === 'center') css.push('vertical-align:middle')
        else if (vAlign === 'bottom') css.push('vertical-align:bottom')
        const tcW = child(tcPr, 'tcW')
        if (attr(tcW, 'type') === 'dxa' && gridTotal === 0) css.push(`width:${pt(clamp((num(tcW, 'w') ?? 0) / 20, 0, 2000))}`)
        const direction = attr(child(tcPr, 'textDirection'), 'val')
        const vertical = direction === 'tbRl' || direction === 'tbRlV' || direction === 'tb' ? 'writing-mode:vertical-rl' : direction === 'btLr' || direction === 'lrTbV' ? 'writing-mode:vertical-rl;transform:rotate(180deg)' : ''
        const layer: Layer = { pPr: conds.map((cond) => cond.pPr), rPr: conds.map((cond) => cond.rPr) }
        const span = cell.span > 1 ? ` colspan="${cell.span}"` : ''
        const rowspan = cell.rowspan > 1 ? ` rowspan="${cell.rowspan}"` : ''
        const cellWidth = gridTotal > 0 ? (grid.slice(cell.col, cell.col + cell.span).reduce((sum, width) => sum + width, 0) / gridTotal) * scope.width : scope.width / columns
        const cellScope: Scope = { ...scope, layer, depth: scope.depth + 1, width: Math.max(18, cellWidth - margin.left - margin.right), dark: shading ? isDark(shading) : scope.dark, sheet: null }
        let content: string
        this.inCell++
        try {
          content = this.renderBlocks(cell.tc.children.filter((node) => local(node.name) !== 'tcPr'), cellScope)
        } catch (error) {
          content = unsupported(`儲存格無法顯示:${reason(error)}`)
        } finally {
          this.inCell--
        }
        cellsHtml += `<td${span}${rowspan} class="${this.cls(css.join(';'))}"><div class="dx-cell"${vertical ? ` style="${vertical}"` : ''}>${content}</div></td>`
      }
      if (row.after) cellsHtml += `<td colspan="${row.after}" class="dx-gap"></td>`
      body += `<tr${rowCss ? ` style="${rowCss}"` : ''}>${cellsHtml}</tr>`
    }
    if (cutRows) body += `<tr><td colspan="${columnCount}" class="dx-cut">表格太大,其餘 ${cutRows} 列未顯示,請下載原始檔查看</td></tr>`
    return `<div class="dx-tblw"${wrapCss.length ? ` style="${wrapCss.join(';')}"` : ''}><table class="dx-tbl" style="${tableCss.filter(Boolean).join(';')}">${colgroup}<tbody>${body}</tbody></table></div>`
  }
}

// ---------------------------------------------------------------------------------------------
// Page

const PAGE_CSS = `:root{color-scheme:light}
html,body{margin:0;background:#e8eaed}
body{padding:16px 10px 40px;font-family:Arial,'PingFang TC','Noto Sans TC','Microsoft JhengHei',sans-serif;-webkit-text-size-adjust:100%;text-size-adjust:100%}
.dx-doc{max-width:100%;margin:0 auto}
.dx-sheet{--il:var(--ml);position:relative;box-sizing:border-box;width:min(100%,var(--pw));margin:0 auto 20px;color:#000;box-shadow:0 1px 3px rgba(60,64,67,.3),0 4px 12px rgba(60,64,67,.12);container-type:inline-size;isolation:isolate;overflow-wrap:break-word;overflow:clip}
.dx-in{box-sizing:border-box;min-height:100%;display:flex;flex-direction:column;padding:var(--mt) var(--mr) var(--mb) var(--il)}
.dx-body{min-width:0}
.dx-header{margin-top:calc(var(--hd) - var(--mt));min-height:calc(var(--mt) - var(--hd));margin-bottom:6pt}
.dx-footer{margin-top:auto;padding-top:12pt;margin-bottom:calc(var(--fd) - var(--mb))}
.dx-p{margin:0;white-space:pre-wrap;line-height:1.2;min-width:0}
.dx-p:after{content:"";display:table;clear:both}
.dx-flex{display:flex;align-items:baseline}
.dx-flex>.dx-seg{min-width:0;flex:0 1 auto}
.dx-flex>.dx-seg~.dx-seg{text-indent:0;flex:0 0 auto;max-width:70%}
.dx-fill{flex:1 1 1em;min-width:1em;margin:0 .2em;align-self:baseline}
.dx-ld-dot,.dx-ld-middleDot{border-bottom:1.5px dotted currentColor}
.dx-ld-hyphen{border-bottom:1px dashed currentColor}
.dx-ld-underscore{border-bottom:1px solid currentColor}
.dx-ld-heavy{border-bottom:2px solid currentColor}
.dx-tab{display:inline-block;width:var(--tab,36pt);text-indent:0}
.dx-pgbr{display:block;position:relative;height:0;border-top:1px dashed #b9bec6;margin:14px 0;text-indent:0}
.dx-pgbr:after{content:"分頁";position:absolute;left:50%;top:-.75em;transform:translateX(-50%);font:11px/1.4 sans-serif;color:#7d838c;background:var(--paper,#fff);padding:0 6px;white-space:nowrap}
.dx-colbr{display:block;break-after:column}
.dx-tblw{overflow-x:auto;max-width:100%;-webkit-overflow-scrolling:touch}
.dx-tbl{border-spacing:0;white-space:normal;text-indent:0}
td{vertical-align:top}
.dx-gap{border:none;padding:0}
.dx-cut{font:12px/1.5 sans-serif;color:#5f6368;padding:6px;text-align:center}
.dx-img{max-width:100%;height:auto;vertical-align:bottom}
.dx-float{max-width:60%}
.dx-figure{display:block}
.dx-abs{position:absolute;display:block}
.dx-abs>.dx-img{width:100%;height:100%}
.dx-pg{position:relative;box-sizing:border-box}
.dx-pin{display:block;position:absolute;width:0;height:0;text-indent:0}
.dx-wordart.dx-fit{display:flex;width:100%;height:100%;align-items:center;justify-content:center;line-height:1;overflow:visible}
.dx-html{white-space:normal;line-height:1.35;overflow-wrap:break-word}
.dx-html p,.dx-html ul,.dx-html ol,.dx-html blockquote,.dx-html pre,.dx-html table{margin:0 0 .6em}
.dx-html h1,.dx-html h2,.dx-html h3,.dx-html h4,.dx-html h5,.dx-html h6{margin:.6em 0 .4em;line-height:1.25}
.dx-html table{border-collapse:collapse;max-width:100%}
.dx-html th,.dx-html td{border:1px solid #9aa0a6;padding:2pt 5pt;vertical-align:top}
.dx-html blockquote{margin-inline:1.5em}
.dx-html pre{white-space:pre-wrap;font-family:Menlo,Consolas,'Courier New',monospace}
.dx-html code{font-family:Menlo,Consolas,'Courier New',monospace}
.dx-txbx{display:inline-flex;flex-direction:column;max-width:100%;vertical-align:top;text-indent:0;white-space:normal;overflow-wrap:break-word}
.dx-txbx>.dx-p,.dx-txbx>.dx-tblw{flex:none}
.dx-line{display:inline-block;height:0;vertical-align:middle}
.dx-grp{display:inline-block;position:relative;max-width:100%;vertical-align:top;text-indent:0}
.dx-gi{position:absolute;display:block}
.dx-gi>.dx-txbx{display:flex;overflow:hidden}
.dx-wordart{display:inline-block;font-weight:700;white-space:nowrap;max-width:100%;overflow:hidden}
.dx-geom{position:absolute;left:0;top:0;width:100%;height:100%;overflow:visible;z-index:-1}
.dx-chart{display:inline-block;max-width:100%;vertical-align:top;background:transparent}
.dx-chart>svg{display:block}
.dx-unsup{display:inline-block;max-width:100%;box-sizing:border-box;border:1px dashed #9aa0a6;background:#f6f7f9;color:#5f6368;font:12px/1.45 sans-serif;font-style:normal;font-weight:400;padding:6px 8px;vertical-align:middle;white-space:normal;text-indent:0;text-align:center}
.dx-notice{box-sizing:border-box;max-width:min(100%,640px);margin:0 auto 12px;background:#fff8e1;border:1px solid #ecd28b;color:#5d4500;font:13px/1.5 sans-serif;padding:8px 12px;border-radius:6px}
.dx-lnk{color:inherit;text-decoration:none}
span.dx-lnk{cursor:help}
.dx-fnref{color:inherit;text-decoration:none}
.dx-notes{margin-top:18pt;padding-top:6pt;position:relative}
.dx-notes:before{content:"";position:absolute;top:0;left:0;width:33%;border-top:.75pt solid currentColor}
.dx-note{font-size:90%}
.dx-cols{column-fill:balance}
.dx-vert{writing-mode:vertical-rl;overflow-x:auto;width:100%}
.dx-math{font-family:'Cambria Math','STIX Two Math','Latin Modern Math','Times New Roman',serif;white-space:nowrap}
.dx-mathpara{display:block;text-align:center;margin:4pt 0}
.m-frac{display:inline-flex;flex-direction:column;vertical-align:middle;text-align:center;font-size:90%;margin:0 .1em}
.m-frac>span+span{border-top:1px solid currentColor}
.m-rad,.m-bar{border-top:1px solid currentColor}
.m-nary{font-size:130%}
.m-mat{display:inline-table;vertical-align:middle;margin:0 .2em}.m-row{display:table-row}.m-row>span{display:table-cell;padding:0 .3em;text-align:center}
.dx-fail{box-sizing:border-box;max-width:560px;margin:10vh auto;background:#fff;border:1px solid #dadce0;border-radius:10px;padding:20px 22px;font:15px/1.6 sans-serif;color:#202124}
.dx-fail h1{font-size:18px;margin:0 0 8px}.dx-fail code{font-size:12px;color:#5f6368;word-break:break-all}
@media (max-width:640px){
body{padding:10px 6px 28px}
.dx-sheet{--il:min(var(--ml),16px)}
.dx-in{padding:min(var(--mt),26px) min(var(--mr),16px) min(var(--mb),26px) var(--il)}
.dx-sheet.dx-placed{--il:var(--pll)}
.dx-placed>.dx-in{padding:var(--plt) var(--plr) var(--plb) var(--il)}
.dx-header{margin-top:0;min-height:0}
.dx-footer{margin-bottom:0}
.dx-cols{column-count:1!important}
.dx-float{max-width:50%}
}
@media print{body{background:none;padding:0}.dx-sheet{box-shadow:none;margin:0}.dx-pgbr{border:0;break-after:page}.dx-pgbr:after{display:none}}`

/** The page as HTML; `mediaBase` prefixes picture URLs. Exposed for tests. */
export function readDocx(data: Buffer, mediaBase: string): { html: string; css: string; text: string; title: string; tab: number } {
  const renderer = new DocxRenderer(data, mediaBase)
  const { html, title } = renderer.read()
  return { html, css: renderer.css(), text: [title, ...renderer.text].filter(Boolean).join('\n'), title, tab: renderer.defaultTab }
}

function failureReason(error: unknown, data: Buffer): string {
  if (data.length >= 8 && data.readUInt32BE(0) === 0xd0cf11e0) return '這是舊版 Word (.doc) 或受密碼保護的檔案,網頁版無法開啟。'
  if (error instanceof ZipError) return '檔案不是有效的 .docx(zip 格式),可能已損毀,或副檔名不符。'
  if (error instanceof XmlError) return '文件內容的 XML 無法解析,檔案可能已損毀。'
  if (error instanceof Error && /[一-鿿]/.test(error.message)) return error.message
  return '讀取這份文件時發生錯誤。'
}

/** The whole page for a Word file. A file that cannot be read still gets a page that says why. */
export function renderDocxPage(data: Buffer, fileName: string, title: string): string {
  const head = (extra: string) =>
    `<!doctype html><html lang="zh-Hant"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>${esc(title)}</title><style>${PAGE_CSS}${extra}</style></head>`
  let doc: ReturnType<typeof readDocx>
  try {
    doc = readDocx(data, officeMediaBase(fileName))
  } catch (error) {
    const message = `Word: ${error instanceof Error ? error.message : String(error)}`
    return injectBridge(
      `${head('')}<body class="docx"><main class="dx-doc"><div class="dx-fail" role="alert"><h1>這份 Word 檔無法顯示</h1><p>${esc(failureReason(error, data))}</p>${hasHan(message) ? `<p><code>${esc(message)}</code></p>` : ''}<p>可以下載原始檔,用 Word 開啟查看。</p></div></main>${failureScript(message)}</body></html>`
    )
  }
  return injectBridge(`${head(`:root{--tab:${round(doc.tab)}pt}${doc.css}`)}<body class="docx"><main class="dx-doc">${doc.html}</main></body></html>`)
}

/**
 * Cuts a paragraph after its first page break that has content before it: [up to and including the break,
 * the rest]. Run and wrapper properties are kept on both sides; the rest carries no list numbering, since
 * it continues the same item.
 */
function splitAtPageBreak(p: XmlNode): [XmlNode, XmlNode] | null {
  const PROPS = new Set(['pPr', 'rPr'])
  const split = (node: XmlNode, depth: number): [XmlNode, XmlNode] | null => {
    if (depth > 12) return null
    for (let i = 0; i < node.children.length; i++) {
      const item = node.children[i]
      const props = node.children.filter((c) => PROPS.has(local(c.name)))
      if (local(item.name) === 'br' && attr(item, 'type') === 'page') {
        return [{ ...node, children: node.children.slice(0, i + 1) }, { ...node, children: [...props, ...node.children.slice(i + 1)] }]
      }
      // The same parts pageBreak() looks past: a deleted or moved-away break is not on the page.
      if (!NOT_ON_PAGE.has(local(item.name)) && item.children.length) {
        const inner = split(item, depth + 1)
        if (inner) return [{ ...node, children: [...node.children.slice(0, i), inner[0]] }, { ...node, children: [...props, inner[1], ...node.children.slice(i + 1)] }]
      }
    }
    return null
  }
  const pieces = split(p, 0)
  if (!pieces) return null
  const [first, rest] = pieces
  // Nothing after the break: the paragraph simply ends the page.
  const content = (node: XmlNode): boolean =>
    node.children.some((c) => !NOT_ON_PAGE.has(local(c.name)) && (['drawing', 'pict', 'object', 'sym', 'tab', 'br'].includes(local(c.name)) || (local(c.name) === 't' && c.text !== '') || content(c)))
  if (!content(rest)) return null
  const pPr = child(rest, 'pPr')
  const restPPr = pPr ? { ...pPr, children: pPr.children.filter((c) => local(c.name) !== 'pageBreakBefore') } : null
  return [first, { ...rest, children: rest.children.map((c) => (c === pPr && restPPr ? restPPr : c)) }]
}

/** Parts of a paragraph that are not on the page: properties, tracked deletions and moves, field codes. */
const NOT_ON_PAGE = new Set(['pPr', 'rPr', 'del', 'moveFrom', 'instrText', 'delText'])

/** The searchable text of a Word file: body, header, footer, notes and text boxes. */
export function docxText(data: Buffer): string {
  return readDocx(data, '').text
}

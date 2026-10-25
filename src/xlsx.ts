/**
 * An Excel file published as is (`*.xlsx`), drawn on the server as one HTML page: a tab bar of the
 * visible sheets and, for each, a grid the way Excel draws it (column letters, row numbers, frozen
 * panes, merged cells, the file's own fonts, fills, borders and number formats). Pictures load from
 * the file itself through the Office media route (office.ts). What it does not draw shows as a small
 * labelled box; one sheet, picture or chart that cannot be read never takes the others down.
 */
import { posix } from 'node:path'
import { escapeHtml, injectBridge } from './content-render.ts'
import { chartTitle, renderChart, type ChartResult } from './office-chart.ts'
import { fillShade, presetGeometry, shapeGeometry } from './office-geometry.ts'
import { chineseCount, CN_DIGITS, CN_LEGAL } from './office-numbers.ts'
import { CELL_PATTERNS, patternCover, patternSvg } from './office-patterns.ts'
import { failureScript, isDrawablePart, officeMediaBase, officeMediaUrl } from './office.ts'
import { attr, child, children, local, num, parseXml, path, relAttr, tagAttrs, type XmlNode } from './xml.ts'
import { ZipReader } from './zip.ts'

/** How much of a workbook one page draws. Past these the page says what was cut. */
export const XLSX_LIMITS = {
  rows: 2000, cols: 200, sheetCells: 100_000, workbookCells: 600_000, drawings: 300, textChars: 300_000, chartPoints: 2000,
  /** Every sheet draws at least its first this many rows, however the workbook budget was shared. */
  floorRows: 200,
  /** Excel's own limit for one cell, and the text one page may carry in all. */
  cellChars: 32_767, pageChars: 8_000_000
}

const EMU_PER_PX = 9525
const EMU_PER_PT = 12700
const UNSUPPORTED = '網頁版未支援,請下載原始檔查看'

// ---------------------------------------------------------------------------------------------
// Package parts and relationships

type Rel = { type: string; target: string; external: boolean }

class Package {
  readonly zip: ZipReader
  private readonly parts = new Map<string, XmlNode | null>()
  private readonly relsCache = new Map<string, Map<string, Rel>>()

  private readonly names = new Map<string, string>()

  constructor(data: Buffer) {
    this.zip = new ZipReader(data)
    // Part names are case-insensitive, and some tools write them with backslashes; Excel accepts both.
    for (const name of this.zip.names()) this.names.set(name.replace(/\\/g, '/').toLowerCase(), name)
  }

  /** The name a part is stored under, or the name asked for when there is no such part. */
  stored(name: string): string {
    return this.zip.has(name) ? name : (this.names.get(name.toLowerCase()) ?? name)
  }

  has(name: string): boolean {
    return this.zip.has(this.stored(name))
  }

  text(name: string): string | null {
    return this.zip.text(this.stored(name))
  }

  xml(name: string): XmlNode | null {
    if (!this.parts.has(name)) {
      const text = this.text(name)
      this.parts.set(name, text === null ? null : parseXml(text))
    }
    return this.parts.get(name) ?? null
  }

  rels(part: string): Map<string, Rel> {
    let map = this.relsCache.get(part)
    if (map) return map
    map = new Map()
    const dir = posix.dirname(part)
    let doc: XmlNode | null = null
    try {
      doc = this.xml(posix.join(dir, '_rels', `${posix.basename(part)}.rels`))
    } catch {}
    for (const rel of children(doc, 'Relationship')) {
      const id = attr(rel, 'Id')
      const target = attr(rel, 'Target')
      if (!id || target === undefined) continue
      const external = attr(rel, 'TargetMode') === 'External'
      let resolved = target
      if (!external) {
        resolved = target.startsWith('/') ? target.slice(1) : posix.normalize(posix.join(dir === '.' ? '' : dir, target))
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

type Rgb = { r: number; g: number; b: number }

const isHex = (value: string | undefined): value is string => typeof value === 'string' && /^[0-9a-fA-F]{6}$/.test(value)
const hexRgb = (hex: string): Rgb => ({ r: parseInt(hex.slice(0, 2), 16), g: parseInt(hex.slice(2, 4), 16), b: parseInt(hex.slice(4, 6), 16) })
const rgbHex = ({ r, g, b }: Rgb): string =>
  [r, g, b].map((v) => Math.round(Math.min(255, Math.max(0, v))).toString(16).padStart(2, '0')).join('').toUpperCase()
const clamp01 = (value: number) => Math.min(1, Math.max(0, value))

function rgbToHsl({ r, g, b }: Rgb): [number, number, number] {
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

function hslToRgb(h: number, s: number, l: number): Rgb {
  const hue = (p: number, q: number, t: number) => {
    t = ((t % 1) + 1) % 1
    if (t < 1 / 6) return p + (q - p) * 6 * t
    if (t < 1 / 2) return q
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6
    return p
  }
  if (s === 0) return { r: l * 255, g: l * 255, b: l * 255 }
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s
  const p = 2 * l - q
  const H = h / 360
  return { r: hue(p, q, H + 1 / 3) * 255, g: hue(p, q, H) * 255, b: hue(p, q, H - 1 / 3) * 255 }
}

/** SpreadsheetML tint: the luminance moves towards black (negative) or white (positive). */
function applyTint(hex: string, tint: number): string {
  if (!tint) return hex
  const [h, s, l] = rgbToHsl(hexRgb(hex))
  const lum = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint
  return rgbHex(hslToRgb(h, s, clamp01(lum)))
}

function mix(fg: string, bg: string, ratio: number): string {
  const a = hexRgb(fg)
  const b = hexRgb(bg)
  return rgbHex({ r: a.r * ratio + b.r * (1 - ratio), g: a.g * ratio + b.g * (1 - ratio), b: a.b * ratio + b.b * (1 - ratio) })
}

/** Readable text on a background: black or white, by luminance. */
const contrastOn = (hex: string) => {
  const { r, g, b } = hexRgb(hex)
  return 0.299 * r + 0.587 * g + 0.114 * b > 150 ? '000000' : 'FFFFFF'
}

const DEFAULT_INDEXED = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'
]

type Theme = { scheme: Record<string, string>; major: string; minor: string; majorEa: string; minorEa: string }

const DEFAULT_SCHEME: Record<string, string> = {
  dk1: '000000', lt1: 'FFFFFF', dk2: '44546A', lt2: 'E7E6E6', accent1: '4472C4', accent2: 'ED7D31', accent3: 'A5A5A5', accent4: 'FFC000',
  accent5: '5B9BD5', accent6: '70AD47', hlink: '0563C1', folHlink: '954F72'
}

/** `theme="N"` in SpreadsheetML counts lt1 before dk1, unlike the order clrScheme lists them in. */
const THEME_INDEX = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']

function readTheme(pkg: Package, workbookPart: string, warnings: string[]): Theme {
  const theme: Theme = { scheme: { ...DEFAULT_SCHEME }, major: '', minor: '', majorEa: '', minorEa: '' }
  const rel = pkg.relOfType(workbookPart, 'theme')
  if (!rel || rel.external) return theme
  try {
    const xml = pkg.xml(rel.target)
    const elements = child(xml, 'themeElements')
    for (const entry of child(elements, 'clrScheme')?.children ?? []) {
      const color = entry.children[0]
      const value = local(color?.name ?? '') === 'sysClr' ? attr(color, 'lastClr') : attr(color, 'val')
      if (isHex(value)) theme.scheme[local(entry.name)] = value.toUpperCase()
    }
    const fonts = child(elements, 'fontScheme')
    theme.major = attr(path(fonts, 'majorFont', 'latin'), 'typeface') ?? ''
    theme.minor = attr(path(fonts, 'minorFont', 'latin'), 'typeface') ?? ''
    theme.majorEa = attr(path(fonts, 'majorFont', 'ea'), 'typeface') ?? ''
    theme.minorEa = attr(path(fonts, 'minorFont', 'ea'), 'typeface') ?? ''
  } catch (error) {
    warnings.push(`佈景主題無法讀取,改用預設色彩(${reason(error)})`)
  }
  return theme
}

type Palette = { theme: Theme; indexed: string[] }

/** A SpreadsheetML `<color>`: rgb (ARGB, alpha ignored), theme + tint, or indexed. Null for automatic. */
function xColor(node: XmlNode | null | undefined, palette: Palette): string | null {
  if (!node) return null
  let hex: string | undefined
  const rgb = attr(node, 'rgb')
  const theme = num(node, 'theme')
  const indexed = num(node, 'indexed')
  if (rgb && /^[0-9a-fA-F]{6}([0-9a-fA-F]{2})?$/.test(rgb)) hex = rgb.slice(-6)
  else if (theme !== undefined) hex = palette.theme.scheme[THEME_INDEX[theme] ?? '']
  else if (indexed !== undefined) hex = indexed === 64 ? '000000' : indexed === 65 ? 'FFFFFF' : palette.indexed[indexed]
  if (!isHex(hex)) return null
  return applyTint(hex.toUpperCase(), Math.max(-1, Math.min(1, num(node, 'tint') ?? 0)))
}

// DrawingML colours (charts, shapes and text boxes inside drawings).

const PRESET_COLORS: Record<string, string> = {
  black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', gray: '808080', grey: '808080',
  darkGray: 'A9A9A9', lightGray: 'D3D3D3', orange: 'FFA500', purple: '800080', navy: '000080', teal: '008080', silver: 'C0C0C0'
}
const COLOR_TAGS = new Set(['srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'scrgbClr', 'hslClr'])
const SCHEME_ALIAS: Record<string, string> = { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' }

function drawingColor(holder: XmlNode | null | undefined, theme: Theme): { hex: string; alpha: number } | null {
  const element = holder?.children.find((candidate) => COLOR_TAGS.has(local(candidate.name)))
  if (!element) return null
  const name = local(element.name)
  let base: string | undefined
  if (name === 'srgbClr') base = attr(element, 'val')
  else if (name === 'schemeClr') {
    const val = attr(element, 'val') ?? ''
    base = theme.scheme[SCHEME_ALIAS[val] ?? val]
  } else if (name === 'sysClr') base = attr(element, 'lastClr') ?? (attr(element, 'val') === 'window' ? 'FFFFFF' : '000000')
  else if (name === 'prstClr') base = PRESET_COLORS[attr(element, 'val') ?? ''] ?? '000000'
  else base = '808080'
  if (!isHex(base)) return null
  let color = hexRgb(base)
  let alpha = 1
  for (const mod of element.children) {
    const value = (num(mod, 'val') ?? 0) / 100000
    const kind = local(mod.name)
    if (kind === 'alpha') alpha = clamp01(value)
    else if (kind === 'tint') color = { r: color.r + (255 - color.r) * (1 - value), g: color.g + (255 - color.g) * (1 - value), b: color.b + (255 - color.b) * (1 - value) }
    else if (kind === 'shade') color = { r: color.r * value, g: color.g * value, b: color.b * value }
    else if (kind === 'lumMod' || kind === 'lumOff' || kind === 'satMod') {
      let [h, s, l] = rgbToHsl(color)
      if (kind === 'lumMod') l = clamp01(l * value)
      else if (kind === 'lumOff') l = clamp01(l + value)
      else s = clamp01(s * value)
      color = hslToRgb(h, s, l)
    }
  }
  return { hex: rgbHex(color), alpha }
}

const cssDrawing = (color: { hex: string; alpha: number }) => {
  if (color.alpha >= 1) return `#${color.hex}`
  const { r, g, b } = hexRgb(color.hex)
  return `rgba(${r},${g},${b},${Math.round(color.alpha * 1000) / 1000})`
}

// ---------------------------------------------------------------------------------------------
// Fonts

const CJK_SANS = "'PingFang TC','Heiti TC','Noto Sans TC','Noto Sans CJK TC','Microsoft JhengHei'"
// First match wins: monospace names before serif ones ('Noto Sans Mono'), and a name that says
// 'Sans Serif' is sans.
const FONT_FALLBACKS: [RegExp, string][] = [
  [/Consolas|Courier|Mono|Menlo|Lucida Console|Typewriter|Inconsolata|Fira Code|Source Code|Cascadia|Code Pro|\bCode\b/i, 'ui-monospace,Menlo,Consolas,monospace'],
  [/Sans Serif/i, `${CJK_SANS},sans-serif`],
  [/楷|Kai/i, "'DFKai-SB','BiauKai','Kaiti TC',serif"],
  [
    /細明|明體|明朝|宋|Ming|Song|Sun|Mincho|Myungjo|Batang|Gungsuh|Serif|Times|Georgia|Cambria|Garamond|Palatino|Antiqua|Bookman|Baskerville|Bodoni|Century(?! Gothic)|Constantia|Minion|Perpetua|Centaur|Rockwell/i,
    "'PMingLiU','Songti TC','Noto Serif TC','Noto Serif CJK TC',serif"
  ]
]

/** Font names go into CSS: letters, digits, space and ._- only. */
const safeFont = (name: string | undefined) => (name ?? '').replace(/[^\p{L}\p{N} ._-]/gu, '').trim().slice(0, 80)

function fontStack(...candidates: (string | undefined)[]): string {
  const names = [...new Set(candidates.map(safeFont).filter(Boolean))]
  // The cell's own font decides the kind; the theme's East Asian font only when it says nothing.
  let fallback = `${CJK_SANS},sans-serif`
  for (const name of names) {
    const found = FONT_FALLBACKS.find(([pattern]) => pattern.test(name))
    if (found) {
      fallback = found[1]
      break
    }
  }
  return [...names.map((name) => `'${name}'`), fallback].join(',')
}

type Font = { b: boolean; i: boolean; u: string; strike: boolean; color: string | null; sz: number; name: string; vert: string }

const DEFAULT_FONT: Font = { b: false, i: false, u: '', strike: false, color: null, sz: 11, name: 'Calibri', vert: '' }

const flag = (node: XmlNode | null, name: string): boolean => {
  const found = child(node, name)
  return Boolean(found) && !['0', 'false', 'off'].includes(attr(found, 'val') ?? '1')
}

function readFont(node: XmlNode | null, palette: Palette, base: Font): Font {
  if (!node) return { ...base }
  const u = child(node, 'u')
  const underline = u ? (attr(u, 'val') ?? 'single') : ''
  const scheme = attr(child(node, 'scheme'), 'val')
  const themed = scheme === 'major' ? palette.theme.major : scheme === 'minor' ? palette.theme.minor : ''
  const size = num(child(node, 'sz'), 'val')
  return {
    b: flag(node, 'b'),
    i: flag(node, 'i'),
    u: underline === 'none' ? '' : underline,
    strike: flag(node, 'strike'),
    color: child(node, 'color') ? xColor(child(node, 'color'), palette) : base.color,
    sz: size && size > 0 && size < 410 ? size : base.sz,
    name: attr(child(node, 'name'), 'val') ?? attr(child(node, 'rFont'), 'val') ?? (themed || base.name),
    vert: attr(child(node, 'vertAlign'), 'val') ?? ''
  }
}

/** CSS for a font, leaving out what it shares with `base` (the grid's default font). */
function fontCss(font: Font, base: Font, theme: Theme, run = false): string {
  const out: string[] = []
  if (font.b !== base.b) out.push(`font-weight:${font.b ? 700 : 400}`)
  if (font.i !== base.i) out.push(`font-style:${font.i ? 'italic' : 'normal'}`)
  if (font.u !== base.u || font.strike !== base.strike) {
    const lines = [font.u ? 'underline' : '', font.strike ? 'line-through' : ''].filter(Boolean)
    out.push(`text-decoration:${lines.length ? lines.join(' ') : 'none'}${font.u.startsWith('double') ? ' double' : ''}`)
  }
  if (font.color !== base.color) out.push(`color:#${font.color ?? '000000'}`)
  if (font.sz !== base.sz) out.push(`font-size:${round(font.sz, 2)}pt`)
  if (font.name !== base.name) out.push(`font-family:${fontStack(font.name, theme.minorEa)}`)
  if (run && (font.vert === 'superscript' || font.vert === 'subscript')) out.push(`vertical-align:${font.vert === 'superscript' ? 'super' : 'sub'};font-size:smaller`)
  return out.join(';')
}

// ---------------------------------------------------------------------------------------------
// Number formats

const BUILTIN_FORMATS: Record<number, string> = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
  5: '"$"#,##0_);\\("$"#,##0\\)', 6: '"$"#,##0_);[Red]\\("$"#,##0\\)', 7: '"$"#,##0.00_);\\("$"#,##0.00\\)', 8: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
  9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
  // Locale-dependent ids follow Excel set to Traditional Chinese (zh-TW).
  14: 'yyyy/m/d', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'yyyy/m/d h:mm',
  27: '[$-404]e/m/d', 28: '[$-404]e"年"m"月"d"日"', 29: '[$-404]e"年"m"月"d"日"', 30: 'm/d/yy', 31: 'yyyy"年"m"月"d"日"', 32: 'hh"時"mm"分"',
  33: 'hh"時"mm"分"ss"秒"', 34: '上午/下午hh"時"mm"分"', 35: '上午/下午hh"時"mm"分"ss"秒"', 36: '[$-404]e/m/d',
  37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)',
  41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)', 42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)', 44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
  45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mm:ss.0', 48: '##0.0E+0', 49: '@',
  50: '[$-404]e/m/d', 51: '[$-404]e"年"m"月"d"日"', 52: '上午/下午hh"時"mm"分"', 53: '上午/下午hh"時"mm"分"ss"秒"', 54: '[$-404]e"年"m"月"d"日"',
  55: '上午/下午hh"時"mm"分"', 56: '上午/下午hh"時"mm"分"ss"秒"', 57: '[$-404]e/m/d', 58: '[$-404]e"年"m"月"d"日"'
}

type Tok =
  | { k: 'lit'; v: string }
  | { k: 'ph'; v: string }
  | { k: 'dot' | 'comma' | 'pct' | 'slash' | 'at' | 'gen' }
  | { k: 'exp'; v: string; e: string }
  | { k: 'dt'; v: string }
  | { k: 'ampm'; v: string }
  | { k: 'el'; v: string }
  | { k: 'sub'; n: number }

/** `dbnum`: [DBNum1]–[DBNum3], the number written in Chinese (counted, legal characters, or full-width digits); 0 for none. */
type Section = { toks: Tok[]; color: string | null; cond: [string, number] | null; date: boolean; text: boolean; locale: string; dbnum: number }

const NAMED_COLORS: Record<string, string> = {
  black: '000000', blue: '0000FF', cyan: '00FFFF', green: '00FF00', magenta: 'FF00FF', red: 'FF0000', white: 'FFFFFF', yellow: 'FFFF00'
}

function splitSections(code: string): string[] {
  const out: string[] = []
  let current = ''
  for (let i = 0; i < code.length; i++) {
    const ch = code[i]
    if (ch === '"') {
      const end = code.indexOf('"', i + 1)
      const stop = end < 0 ? code.length - 1 : end
      current += code.slice(i, stop + 1)
      i = stop
    } else if (ch === '\\' || ch === '_' || ch === '*') {
      current += code.slice(i, i + 2)
      i++
    } else if (ch === '[') {
      const end = code.indexOf(']', i)
      const stop = end < 0 ? code.length - 1 : end
      current += code.slice(i, stop + 1)
      i = stop
    } else if (ch === ';') {
      out.push(current)
      current = ''
    } else current += ch
  }
  out.push(current)
  return out.slice(0, 4)
}

function compileSection(source: string, palette: string[]): Section {
  const toks: Tok[] = []
  let color: string | null = null
  let cond: [string, number] | null = null
  let locale = ''
  let dbnum = 0
  const lit = (v: string) => toks.push({ k: 'lit', v })
  for (let i = 0; i < source.length; ) {
    const ch = source[i]
    const rest = source.slice(i)
    if (ch === '"') {
      const end = source.indexOf('"', i + 1)
      const stop = end < 0 ? source.length : end
      lit(source.slice(i + 1, stop))
      i = stop + 1
    } else if (ch === '\\') {
      lit(source[i + 1] ?? '')
      i += 2
    } else if (ch === '_') {
      lit(' ')
      i += 2
    } else if (ch === '*') {
      i += 2
    } else if (ch === '[') {
      const end = source.indexOf(']', i)
      if (end < 0) {
        lit(rest)
        break
      }
      const body = source.slice(i + 1, end)
      i = end + 1
      const lower = body.toLowerCase()
      let match: RegExpExecArray | null
      if (/^(h+|m+|s+)$/i.test(body)) toks.push({ k: 'el', v: lower })
      else if (NAMED_COLORS[lower]) color = NAMED_COLORS[lower]
      else if ((match = /^colou?r\s*(\d+)$/i.exec(body))) color = palette[Number(match[1]) + 7] ?? null
      else if ((match = /^(<=|>=|<>|<|>|=)\s*(-?\d+(?:\.\d+)?(?:e[+-]?\d+)?)$/i.exec(body))) cond = [match[1], Number(match[2])]
      else if ((match = /^dbnum([1-4])$/i.exec(body))) dbnum = Number(match[1])
      else if (body.startsWith('$')) {
        const inner = body.slice(1)
        const dash = inner.indexOf('-')
        const symbol = dash < 0 ? inner : inner.slice(0, dash)
        if (symbol) lit(symbol)
        if (dash >= 0) locale = inner.slice(dash + 1).toLowerCase()
      }
    } else if (/^general/i.test(rest)) {
      toks.push({ k: 'gen' })
      i += 7
    } else if (/^(am\/pm|a\/p)/i.test(rest)) {
      const v = /^(am\/pm|a\/p)/i.exec(rest)![0]
      toks.push({ k: 'ampm', v })
      i += v.length
    } else if (rest.startsWith('上午/下午')) {
      toks.push({ k: 'ampm', v: '上午/下午' })
      i += 5
    } else if (ch === '0' || ch === '#' || ch === '?') {
      toks.push({ k: 'ph', v: ch })
      i++
    } else if (ch === '.') {
      toks.push({ k: 'dot' })
      i++
    } else if (ch === ',') {
      toks.push({ k: 'comma' })
      i++
    } else if (ch === '%') {
      toks.push({ k: 'pct' })
      i++
    } else if (ch === '/') {
      toks.push({ k: 'slash' })
      i++
    } else if (ch === '@') {
      toks.push({ k: 'at' })
      i++
    } else if ((ch === 'E' || ch === 'e') && (source[i + 1] === '+' || source[i + 1] === '-')) {
      toks.push({ k: 'exp', v: source[i + 1], e: ch })
      i += 2
    } else if (/^[ymdhsgeYMDHSGE]$/.test(ch)) {
      const lower = ch.toLowerCase()
      let j = i
      while (j < source.length && source[j].toLowerCase() === lower) j++
      toks.push({ k: 'dt', v: lower.repeat(Math.min(j - i, 5)) })
      i = j
    } else if (/^aaa/i.test(rest)) {
      let j = i
      while (j < source.length && source[j].toLowerCase() === 'a') j++
      toks.push({ k: 'dt', v: j - i >= 4 ? 'aaaa' : 'aaa' })
      i = j
    } else if ((ch === 'B' || ch === 'b') && (source[i + 1] === '1' || source[i + 1] === '2')) {
      i += 2
    } else {
      lit(ch)
      i++
    }
  }
  const date = toks.some((tok) => tok.k === 'ampm' || tok.k === 'el' || (tok.k === 'dt' && tok.v[0] !== 'e' && tok.v[0] !== 'g')) ||
    (toks.some((tok) => tok.k === 'dt') && !toks.some((tok) => tok.k === 'ph'))
  const out: Tok[] = []
  if (date) {
    for (let i = 0; i < toks.length; i++) {
      const tok = toks[i]
      if (tok.k === 'dot' && toks[i + 1]?.k === 'ph') {
        const previous = [...out].reverse().find((t) => t.k !== 'lit')
        if (previous && ((previous.k === 'dt' && previous.v[0] === 's') || (previous.k === 'el' && previous.v[0] === 's'))) {
          let n = 0
          while (toks[i + 1 + n]?.k === 'ph') n++
          out.push({ k: 'sub', n: Math.min(n, 3) })
          i += n
          continue
        }
      }
      if (tok.k === 'ph') out.push({ k: 'lit', v: tok.v })
      else if (tok.k === 'dot') out.push({ k: 'lit', v: '.' })
      else if (tok.k === 'comma') out.push({ k: 'lit', v: ',' })
      else if (tok.k === 'pct') out.push({ k: 'lit', v: '%' })
      else if (tok.k === 'slash') out.push({ k: 'lit', v: '/' })
      else if (tok.k === 'exp') out.push({ k: 'lit', v: `${tok.e}${tok.v}` })
      else out.push(tok)
    }
    // An m right after hours or right before seconds is minutes.
    out.forEach((tok, at) => {
      if (tok.k !== 'dt' || tok.v[0] !== 'm' || tok.v.length > 2) return
      const before = out.slice(0, at).reverse().find((t) => t.k === 'dt' || t.k === 'el')
      const after = out.slice(at + 1).find((t) => t.k === 'dt' || t.k === 'el')
      const isHour = (t: Tok | undefined) => Boolean(t && (t.k === 'dt' || t.k === 'el') && t.v[0] === 'h')
      const isSecond = (t: Tok | undefined) => Boolean(t && (t.k === 'dt' || t.k === 'el') && t.v[0] === 's')
      if (isHour(before) || isSecond(after) || (isSecond(before) && !(after && after.k === 'dt' && /^[dy]/.test(after.v)))) out[at] = { k: 'dt', v: tok.v.toUpperCase() }
    })
  } else {
    for (const tok of toks) out.push(tok.k === 'dt' ? { k: 'lit', v: tok.v } : tok.k === 'ampm' ? { k: 'lit', v: tok.v } : tok)
  }
  return { toks: out, color, cond, date, text: out.some((tok) => tok.k === 'at'), locale, dbnum }
}

type Format = { sections: Section[]; date: boolean }

const formatCache = new Map<string, Format>()

function compileFormat(code: string, palette: string[]): Format {
  const key = `${palette === DEFAULT_INDEXED ? '' : palette.join('')}\u0000${code}`
  let format = formatCache.get(key)
  if (!format) {
    const sections = splitSections(code).map((section) => compileSection(section, palette))
    format = { sections, date: sections[0]?.date ?? false }
    if (formatCache.size > 2000) formatCache.clear()
    formatCache.set(key, format)
  }
  return format
}

const round = (value: number, digits = 3) => {
  const factor = 10 ** digits
  return Math.round(value * factor) / factor
}

/** Excel's General: up to 10 significant digits, scientific for very large or small numbers. */
export function formatGeneral(value: number): string {
  if (!Number.isFinite(value)) return '#NUM!'
  if (value === 0) return '0'
  const abs = Math.abs(value)
  const scientific = () => {
    const [mantissa, exponent] = value.toExponential(5).split('e')
    const e = Number(exponent)
    return `${mantissa.replace(/\.?0+$/, '')}E${e < 0 ? '-' : '+'}${String(Math.abs(e)).padStart(2, '0')}`
  }
  if (abs >= 1e11) return scientific()
  // What fits in eleven characters, as a default-width column shows it.
  const magnitude = Math.floor(Math.log10(abs))
  const decimals = Math.max(0, Math.min(20, 9 - magnitude))
  let text = abs.toFixed(decimals)
  if (text.includes('.')) text = text.replace(/\.?0+$/, '')
  if (text.length > 11) {
    const room = Math.max(0, 10 - text.indexOf('.'))
    text = text.includes('.') ? abs.toFixed(Math.min(room, 20)).replace(/\.?0+$/, '') : text
  }
  // Small numbers switch to scientific when the fixed form would drop their significant digits.
  const needed = Math.min(6, abs.toPrecision(10).replace(/e.*$/, '').replace(/^[0.]+/, '').replace('.', '').replace(/0+$/, '').length)
  const kept = text.replace(/^[0.]+/, '').replace('.', '').length
  if (abs < 1 && kept < needed) return scientific()
  return (value < 0 ? '-' : '') + text
}

/** toFixed that rounds the decimal number Excel shows (1.005 → 1.01), not its binary neighbour. */
function toFixed(value: number, digits: number): string {
  const scaled = value * 10 ** digits
  if (!(scaled < 1e15)) return value.toFixed(digits)
  const rounded = String(Math.round(Number(scaled.toPrecision(15)))).padStart(digits + 1, '0')
  return digits ? `${rounded.slice(0, -digits)}.${rounded.slice(-digits)}` : rounded
}

const groupThousands = (digits: string) => digits.replace(/\B(?=(\d{3})+(?!\d))/g, ',')

/** Fills digit placeholders of the integer part right to left; extra digits go to the leftmost one. */
function fillInteger(toks: Tok[], digits: string, thousands: boolean): string {
  const phs = toks.filter((tok) => tok.k === 'ph')
  if (thousands) {
    const zeros = phs.filter((tok) => tok.k === 'ph' && tok.v === '0').length
    let body = digits.padStart(zeros, '0')
    const spaces = phs.filter((tok) => tok.k === 'ph' && tok.v === '?').length
    if (body.length < spaces) body = body.padStart(spaces, ' ')
    body = groupThousands(body.trimStart()).padStart(body.length, ' ')
    let out = ''
    let placed = false
    for (const tok of toks) {
      if (tok.k === 'ph') {
        if (!placed) out += body
        placed = true
      } else if (tok.k === 'lit') out += tok.v
    }
    return out
  }
  const queue = digits.split('')
  const out: string[] = []
  const firstPh = toks.findIndex((tok) => tok.k === 'ph')
  for (let at = toks.length - 1; at >= 0; at--) {
    const tok = toks[at]
    if (tok.k === 'ph') {
      if (at === firstPh) {
        const restDigits = queue.splice(0).join('')
        out.unshift(restDigits || (tok.v === '0' ? '0' : tok.v === '?' ? ' ' : ''))
      } else {
        const digit = queue.pop()
        out.unshift(digit ?? (tok.v === '0' ? '0' : tok.v === '?' ? ' ' : ''))
      }
    } else if (tok.k === 'lit') out.unshift(tok.v)
  }
  return out.join('')
}

function fillDecimals(toks: Tok[], digits: string): string {
  let out = ''
  let at = 0
  const phCount = toks.filter((tok) => tok.k === 'ph').length
  const lastSignificant = digits.slice(0, phCount).replace(/0+$/, '').length
  for (const tok of toks) {
    if (tok.k === 'ph') {
      const digit = digits[at] ?? '0'
      if (at >= lastSignificant && tok.v !== '0') out += tok.v === '?' ? ' ' : ''
      else out += digit
      at++
    } else if (tok.k === 'lit') out += tok.v
  }
  return out
}

/** A non-negative number through a number section (sign is the caller's business). */
function formatNumberSection(toks: Tok[], value: number): string {
  if (toks.some((tok) => tok.k === 'gen')) {
    return toks.map((tok) => (tok.k === 'gen' ? formatGeneral(value) : tok.k === 'lit' ? tok.v : tok.k === 'ph' ? tok.v : '')).join('')
  }
  if (!toks.some((tok) => tok.k === 'ph')) {
    return toks.map((tok) => (tok.k === 'lit' ? tok.v : tok.k === 'pct' ? '%' : tok.k === 'dot' ? '.' : tok.k === 'comma' ? ',' : tok.k === 'slash' ? '/' : '')).join('')
  }
  const percents = toks.filter((tok) => tok.k === 'pct').length
  let v = value * 100 ** percents
  const pctLit = (list: Tok[]) => list.map((tok): Tok => (tok.k === 'pct' ? { k: 'lit', v: '%' } : tok))
  const expAt = toks.findIndex((tok) => tok.k === 'exp')
  if (expAt >= 0) {
    const mantissaToks = pctLit(toks.slice(0, expAt))
    const exponentToks = pctLit(toks.slice(expAt + 1))
    const dot = mantissaToks.findIndex((tok) => tok.k === 'dot')
    const intToks = dot < 0 ? mantissaToks : mantissaToks.slice(0, dot)
    const decToks = dot < 0 ? [] : mantissaToks.slice(dot + 1)
    const intCount = Math.max(1, intToks.filter((tok) => tok.k === 'ph').length)
    const decCount = Math.min(15, decToks.filter((tok) => tok.k === 'ph').length)
    let exponent = v === 0 ? 0 : Math.floor(Math.log10(v))
    // Several integer digits make the exponent a multiple of their count (engineering notation).
    if (intCount > 1) exponent = Math.floor(exponent / intCount) * intCount
    let mantissa = v / 10 ** exponent
    if (Number(toFixed(mantissa, decCount)) >= 10 ** intCount) {
      exponent += intCount
      mantissa = v / 10 ** exponent
    }
    const [intDigits, decDigits = ''] = toFixed(mantissa, decCount).split('.')
    const { v: sign, e: letter } = toks[expAt] as { v: string; e: string }
    const expText = fillInteger(exponentToks, String(Math.abs(exponent)), false)
    return `${fillInteger(intToks, intDigits === '0' ? '' : intDigits, false)}${dot >= 0 ? '.' : ''}${fillDecimals(decToks, decDigits)}${letter}${exponent < 0 ? '-' : sign === '+' ? '+' : ''}${expText}`
  }
  const slashAt = toks.findIndex((tok) => tok.k === 'slash')
  if (slashAt >= 0) return formatFraction(pctLit(toks), slashAt, v)
  const dot = toks.findIndex((tok) => tok.k === 'dot')
  let intToks = dot < 0 ? toks : toks.slice(0, dot)
  let decToks = dot < 0 ? [] : toks.slice(dot + 1)
  // Commas right after the last placeholder (of either part) scale by a thousand; between placeholders they group.
  const lastPh = intToks.map((tok) => tok.k).lastIndexOf('ph')
  let scale = 0
  for (let at = lastPh + 1; at < intToks.length && intToks[at].k === 'comma'; at++) scale++
  const lastDecPh = decToks.map((tok) => tok.k).lastIndexOf('ph')
  if (lastDecPh >= 0) for (let at = lastDecPh + 1; at < decToks.length && decToks[at].k === 'comma'; at++) scale++
  const firstPh = intToks.findIndex((tok) => tok.k === 'ph')
  const thousands = intToks.some((tok, at) => tok.k === 'comma' && at > firstPh && at < lastPh)
  v /= 1000 ** scale
  intToks = pctLit(intToks.filter((tok) => tok.k !== 'comma'))
  decToks = pctLit(decToks.filter((tok) => tok.k !== 'comma'))
  const decCount = Math.min(15, decToks.filter((tok) => tok.k === 'ph').length)
  if (v >= 1e21) return formatGeneral(v)
  const [intDigits, decDigits = ''] = toFixed(v, decCount).split('.')
  return `${fillInteger(intToks, intDigits === '0' ? '' : intDigits, thousands)}${dot >= 0 ? '.' : ''}${fillDecimals(decToks, decDigits)}`
}

function formatFraction(toks: Tok[], slashAt: number, value: number): string {
  // Numerator: the placeholders just before the slash; a whole part is any placeholders before those.
  let numStart = slashAt
  while (numStart > 0 && toks[numStart - 1].k === 'ph') numStart--
  const before = toks.slice(0, numStart)
  const hasWhole = before.some((tok) => tok.k === 'ph')
  const numToks = toks.slice(numStart, slashAt)
  let denEnd = slashAt + 1
  let denText = ''
  while (denEnd < toks.length) {
    const tok = toks[denEnd]
    if (tok.k === 'ph' || (tok.k === 'lit' && /^\d$/.test(tok.v))) {
      denText += tok.v
      denEnd++
    } else break
  }
  const after = toks.slice(denEnd)
  const fixed = /^\d+$/.test(denText) && Number(denText) > 0 ? Number(denText) : 0
  const whole = hasWhole ? Math.floor(value) : 0
  const frac = value - whole
  let numerator: number
  let denominator: number
  if (fixed) {
    denominator = fixed
    numerator = Math.round(frac * fixed)
  } else {
    const maxDen = 10 ** Math.max(1, Math.min(4, denText.length)) - 1
    ;[numerator, denominator] = bestFraction(frac, maxDen)
  }
  let wholeValue = whole
  if (hasWhole && numerator === denominator) {
    wholeValue += 1
    numerator = 0
  }
  const lits = (list: Tok[]) => list.map((tok) => (tok.k === 'lit' ? tok.v : '')).join('')
  const wholeText = hasWhole ? fillInteger(before, wholeValue === 0 && numerator !== 0 ? '' : String(wholeValue), false) : lits(before)
  if (hasWhole && numerator === 0) return `${wholeText}${' '.repeat(numToks.length + denText.length + 1)}${lits(after)}`
  const numText = String(numerator).padStart(numToks.length, numToks.some((tok) => tok.k === 'ph' && tok.v === '0') ? '0' : ' ')
  const denShown = fixed ? String(fixed) : String(denominator).padEnd(denText.length, ' ')
  return `${wholeText}${numText}/${denShown}${lits(after)}`
}

function bestFraction(value: number, maxDen: number): [number, number] {
  let best: [number, number] = [Math.round(value), 1]
  let bestError = Math.abs(value - best[0])
  for (let d = 1; d <= maxDen && bestError > 1e-12; d++) {
    const n = Math.round(value * d)
    const error = Math.abs(value - n / d)
    if (error < bestError - 1e-12) {
      best = [n, d]
      bestError = error
    }
  }
  return best
}

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
const ZH_DAYS = ['日', '一', '二', '三', '四', '五', '六']
const ZH_MONTHS = ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十', '十一', '十二']

type DateParts = { y: number; mo: number; d: number; dow: number; H: number; M: number; S: number; frac: number; units: number; prec: number }

function serialParts(serial: number, date1904: boolean, prec: number): DateParts {
  const perDay = 86400 * prec
  const units = Math.round(serial * perDay)
  const days = Math.floor(units / perDay)
  const rest = units - days * perDay
  const seconds = Math.floor(rest / prec)
  let y: number
  let mo: number
  let d: number
  let dow: number
  if (!date1904 && days === 60) {
    // Lotus 1-2-3's 29 February 1900, which Excel keeps for compatibility.
    ;[y, mo, d, dow] = [1900, 2, 29, 3]
  } else if (!date1904 && days === 0) {
    ;[y, mo, d, dow] = [1900, 1, 0, 6]
  } else {
    const base = date1904 ? Date.UTC(1904, 0, 1) : days < 60 ? Date.UTC(1899, 11, 31) : Date.UTC(1899, 11, 30)
    const at = new Date(base + days * 86400000)
    ;[y, mo, d] = [at.getUTCFullYear(), at.getUTCMonth() + 1, at.getUTCDate()]
    dow = !date1904 && days < 60 ? (at.getUTCDay() + 6) % 7 : at.getUTCDay()
  }
  return { y, mo, d, dow, H: Math.floor(seconds / 3600), M: Math.floor((seconds % 3600) / 60), S: seconds % 60, frac: rest - seconds * prec, units, prec }
}

function formatDateSection(section: Section, serial: number, date1904: boolean): string {
  if (serial < 0 || serial > 2958465.99999) return '#'.repeat(8)
  const subDigits = Math.max(0, ...section.toks.map((tok) => (tok.k === 'sub' ? tok.n : 0)))
  const p = serialParts(serial, date1904, 10 ** subDigits)
  const hour12 = section.toks.some((tok) => tok.k === 'ampm')
  // Beside [h], [m] or [s], a plain h, m or s of the same unit counts the whole duration too.
  const elapsed = new Set(section.toks.flatMap((tok) => (tok.k === 'el' ? [tok.v[0]] : [])))
  const total = Math.floor(p.units / p.prec)
  const zh = /^(?:[0-9a-f]*?0?(?:404|804|c04|1004|1404))$|^zh/.test(section.locale)
  const roc = /^(?:[0-9a-f]*?0?404)$|^zh-tw/.test(section.locale)
  const pad = (n: number, width: number) => String(n).padStart(width, '0')
  let out = ''
  for (const tok of section.toks) {
    if (tok.k === 'lit') out += tok.v
    else if (tok.k === 'sub') out += `.${pad(p.frac, tok.n)}`
    else if (tok.k === 'ampm') {
      const am = p.H < 12
      if (tok.v === '上午/下午') out += am ? '上午' : '下午'
      else if (tok.v.length === 3) out += am ? tok.v[0] : tok.v[2]
      else out += am ? 'AM' : 'PM'
    } else if (tok.k === 'el') {
      const value = tok.v[0] === 'h' ? Math.floor(total / 3600) : tok.v[0] === 'm' ? Math.floor(total / 60) : total
      out += pad(value, tok.v.length)
    } else if (tok.k === 'dt') {
      const v = tok.v
      const hour = hour12 ? (p.H % 12 || 12) : p.H
      switch (v[0]) {
        case 'y':
          out += v.length <= 2 ? pad(p.y % 100, 2) : pad(p.y, 4)
          break
        case 'e':
          out += String(roc && p.y > 1911 ? p.y - 1911 : p.y)
          break
        case 'g':
          out += roc ? '民國' : ''
          break
        case 'm':
          if (v.length === 1) out += String(p.mo)
          else if (v.length === 2) out += pad(p.mo, 2)
          else if (v.length === 3) out += zh ? `${p.mo}月` : MONTHS[p.mo - 1].slice(0, 3)
          else if (v.length === 4) out += zh ? `${ZH_MONTHS[p.mo - 1]}月` : MONTHS[p.mo - 1]
          else out += MONTHS[p.mo - 1][0]
          break
        case 'M': {
          const minutes = elapsed.has('m') ? Math.floor(total / 60) : p.M
          out += v.length === 1 ? String(minutes) : pad(minutes, 2)
          break
        }
        case 'd':
          if (v.length === 1) out += String(p.d)
          else if (v.length === 2) out += pad(p.d, 2)
          else if (v.length === 3) out += zh ? `週${ZH_DAYS[p.dow]}` : DAYS[p.dow].slice(0, 3)
          else out += zh ? `星期${ZH_DAYS[p.dow]}` : DAYS[p.dow]
          break
        case 'a':
          out += v.length >= 4 ? `星期${ZH_DAYS[p.dow]}` : ZH_DAYS[p.dow]
          break
        case 'h': {
          const hours = elapsed.has('h') ? Math.floor(total / 3600) : hour
          out += v.length === 1 ? String(hours) : pad(hours, 2)
          break
        }
        case 's': {
          const seconds = elapsed.has('s') ? total : p.S
          out += v.length === 1 ? String(seconds) : pad(seconds, 2)
          break
        }
      }
    }
  }
  return out
}

type Formatted = { text: string; color: string | null }

const FULL_WIDTH_DIGITS = '０１２３４５６７８９'

/**
 * The digits of a formatted number as [DBNum1] (一千二百三十四), [DBNum2] (壹仟貳佰參拾肆) or [DBNum3] (１２３４) write
 * them. A whole number is counted with its units (10 is 一十, as Excel writes it) and loses its thousands separators;
 * decimals, and the year of a date, are read digit by digit.
 */
function chineseDigits(text: string, dbnum: number, date: boolean): string {
  if (!dbnum) return text
  if (dbnum >= 3) return text.replace(/\d/g, (d) => FULL_WIDTH_DIGITS[Number(d)])
  const legal = dbnum === 2
  const digits = legal ? CN_LEGAL : CN_DIGITS
  const units = legal ? ['', '拾', '佰', '仟'] : ['', '十', '百', '千']
  const each = (run: string) => run.replace(/\d/g, (d) => digits[Number(d)])
  return text.replace(/(\d{1,3}(?:,\d{3})+|\d+)((?:\.\d+)?)/g, (_, whole: string, fraction: string) => {
    const plain = whole.replace(/,/g, '')
    const counted = (date && plain.length === 4) || plain.length > 8 ? each(plain) : chineseCount(Number(plain), digits, units, '萬', date)
    return counted + (fraction ? `.${each(fraction.slice(1))}` : '')
  })
}

/** A number, shown the way a number format code says. */
export function formatNumber(code: string, value: number, date1904 = false, palette: string[] = DEFAULT_INDEXED): Formatted {
  const format = compileFormat(code, palette)
  const sections = format.sections
  const numeric = sections.filter((section, at) => !(at === 3 || (section.text && !section.toks.some((tok) => tok.k === 'ph' || tok.k === 'gen' || tok.k === 'dt'))))
  let section: Section
  let shown = value
  let sign = ''
  if (numeric.some((s) => s.cond)) {
    const test = (s: Section) => {
      if (!s.cond) return true
      const [op, limit] = s.cond
      return op === '<' ? value < limit : op === '<=' ? value <= limit : op === '>' ? value > limit : op === '>=' ? value >= limit : op === '=' ? value === limit : value !== limit
    }
    const conditional = numeric.filter((s) => s.cond)
    section = conditional.find(test) ?? numeric.find((s) => !s.cond) ?? numeric[numeric.length - 1]
    // Conditional sections keep the sign: Excel shows the number, the condition only picks the look.
    shown = Math.abs(value)
    sign = value < 0 ? '-' : ''
  } else if (value < 0 && numeric.length >= 2) {
    section = numeric[1]
    shown = -value
  } else if (value === 0 && numeric.length >= 3) section = numeric[2]
  else {
    section = numeric[0] ?? sections[0]
    if (value < 0) {
      shown = -value
      sign = '-'
    }
  }
  if (!section || !numeric.length) return { text: formatGeneral(value), color: null }
  if (section.date) {
    // A negative duration keeps its sign; a negative date cannot be shown.
    if (value < 0 && section.toks.some((tok) => tok.k === 'el')) return { text: `-${formatDateSection(section, -value, date1904)}`, color: section.color }
    if (value < 0) return { text: '#'.repeat(8), color: section.color }
    return { text: chineseDigits(formatDateSection(section, value, date1904), section.dbnum, true), color: section.color }
  }
  if (section.toks.length === 0) return { text: '', color: section.color }
  const body = formatNumberSection(section.toks, shown)
  // A negative number that rounds to zero in a single-section format shows no sign, as in Excel.
  const allZero = !/[1-9]/.test(body)
  return { text: (sign && !allZero ? sign : '') + chineseDigits(body, section.dbnum, false), color: section.color }
}

/** Text through a number format: the text section if there is one, else the text as is. */
export function formatText(code: string, text: string, palette: string[] = DEFAULT_INDEXED): Formatted {
  const sections = compileFormat(code, palette).sections
  const section = sections[3] ?? sections.find((candidate) => candidate.text)
  if (!section) return { text, color: null }
  return { text: section.toks.map((tok) => (tok.k === 'at' ? text : tok.k === 'lit' ? tok.v : '')).join(''), color: section.color }
}

// ---------------------------------------------------------------------------------------------
// Styles

type Side = { style: string; color: string; weight: number } | null
/** `diagonal`: the line corner to corner, rising (up: bottom-left to top-right), falling (down), or both. */
type Border = { left: Side; right: Side; top: Side; bottom: Side; diagonal?: { up: boolean; down: boolean; side: NonNullable<Side> } | null }
type Align = { h: string; v: string; wrap: boolean; indent: number; rot: number; shrink: boolean }
type Xf = { fmt: number; font: number; fill: number; border: number; align: Align | null }
type Dxf = { font: Partial<Font> & { set: boolean }; fill: string; border: Border | null }

type Styles = {
  fonts: Font[]
  fills: string[]
  fillHex: (string | null)[]
  borders: Border[]
  xfs: Xf[]
  fmts: Map<number, string>
  dxfs: Dxf[]
  tableStyles: Map<string, Map<string, number>>
  defaultTableStyle: string
  palette: Palette
}

const BORDER_WEIGHT: Record<string, number> = {
  hair: 1, dotted: 2, dashDotDot: 3, dashDot: 4, dashed: 5, thin: 6, mediumDashDotDot: 7, slantDashDot: 8, mediumDashDot: 9, mediumDashed: 10, medium: 11, double: 12, thick: 13
}
const BORDER_CSS: Record<string, string> = {
  hair: '1px dotted', dotted: '1px dotted', dashDotDot: '1px dashed', dashDot: '1px dashed', dashed: '1px dashed', thin: '1px solid',
  mediumDashDotDot: '2px dashed', slantDashDot: '2px dashed', mediumDashDot: '2px dashed', mediumDashed: '2px dashed', medium: '2px solid', double: '3px double', thick: '3px solid'
}

function readSide(node: XmlNode | null, palette: Palette): Side {
  const style = attr(node, 'style')
  if (!node || !style || style === 'none' || !BORDER_CSS[style]) return null
  return { style, color: xColor(child(node, 'color'), palette) ?? '000000', weight: BORDER_WEIGHT[style] }
}

function readBorder(node: XmlNode | null, palette: Palette): Border {
  const on = (name: string) => ['1', 'true'].includes(attr(node, name) ?? '')
  const side = readSide(child(node, 'diagonal'), palette)
  return {
    left: readSide(child(node, 'left') ?? child(node, 'start'), palette),
    right: readSide(child(node, 'right') ?? child(node, 'end'), palette),
    top: readSide(child(node, 'top'), palette),
    bottom: readSide(child(node, 'bottom'), palette),
    diagonal: side && (on('diagonalUp') || on('diagonalDown')) ? { up: on('diagonalUp'), down: on('diagonalDown'), side } : null
  }
}

/** A cell's diagonal lines as a CSS background layer: an SVG stretched over the cell, its stroke one width at any size. */
function diagonalCss(diagonal: Border['diagonal']): string {
  if (!diagonal) return ''
  const { style, color } = diagonal.side
  const [width, kind] = BORDER_CSS[style].split(' ')
  const px = parseFloat(width)
  const dash = kind === 'dotted' ? ` stroke-dasharray="${px} ${px * 2}"` : kind === 'dashed' ? ` stroke-dasharray="${px * 4} ${px * 2}"` : ''
  const line = (x1: number, y1: number, x2: number, y2: number) =>
    `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="#${color}" stroke-width="${px}"${dash} vector-effect="non-scaling-stroke"/>`
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100" preserveAspectRatio="none">${diagonal.up ? line(0, 100, 100, 0) : ''}${diagonal.down ? line(0, 0, 100, 100) : ''}</svg>`
  return `url('data:image/svg+xml,${encodeURIComponent(svg)}') 0 0/100% 100% no-repeat`
}

/** A fill as a CSS background and, for solid-looking fills, its colour (to pick readable text). */
function readFill(node: XmlNode | null, palette: Palette): { css: string; hex: string | null } {
  const pattern = child(node, 'patternFill')
  if (pattern) {
    const fgNode = child(pattern, 'fgColor')
    const type = attr(pattern, 'patternType') ?? (fgNode ? 'solid' : 'none')
    if (type === 'none') return { css: '', hex: null }
    const fg = xColor(fgNode, palette) ?? '000000'
    const bg = xColor(child(pattern, 'bgColor'), palette) ?? 'FFFFFF'
    if (type === 'solid') return { css: `#${fg}`, hex: fg }
    // Any other pattern is its hatch in its two colours; the hex is how dark it reads, to pick the text colour.
    const bits = CELL_PATTERNS[type] ?? CELL_PATTERNS.mediumGray
    return { css: `url('data:image/svg+xml,${encodeURIComponent(patternSvg(`#${fg}`, `#${bg}`, bits))}') 0 0/8px 8px repeat`, hex: mix(fg, bg, patternCover(bits)) }
  }
  const gradient = child(node, 'gradientFill')
  if (gradient) {
    const stops = children(gradient, 'stop')
      .map((stop) => ({ at: clamp01(num(stop, 'position') ?? 0), hex: xColor(child(stop, 'color'), palette) }))
      .filter((stop): stop is { at: number; hex: string } => stop.hex !== null)
    if (!stops.length) return { css: '', hex: null }
    if (stops.length === 1) return { css: `#${stops[0].hex}`, hex: stops[0].hex }
    const list = stops.map((stop) => `#${stop.hex} ${round(stop.at * 100, 2)}%`).join(',')
    const css = attr(gradient, 'type') === 'path' ? `radial-gradient(${list})` : `linear-gradient(${round((num(gradient, 'degree') ?? 0) + 90, 2)}deg,${list})`
    return { css, hex: stops[0].hex }
  }
  return { css: '', hex: null }
}

function readAlign(node: XmlNode | null): Align | null {
  if (!node) return null
  return {
    h: attr(node, 'horizontal') ?? 'general',
    v: attr(node, 'vertical') ?? 'bottom',
    wrap: ['1', 'true'].includes(attr(node, 'wrapText') ?? ''),
    indent: Math.min(50, Math.max(0, num(node, 'indent') ?? 0)),
    rot: num(node, 'textRotation') ?? 0,
    shrink: ['1', 'true'].includes(attr(node, 'shrinkToFit') ?? '')
  }
}

function readXf(node: XmlNode, parent: Xf | undefined): Xf {
  return {
    fmt: num(node, 'numFmtId') ?? parent?.fmt ?? 0,
    font: num(node, 'fontId') ?? parent?.font ?? 0,
    fill: num(node, 'fillId') ?? parent?.fill ?? 0,
    border: num(node, 'borderId') ?? parent?.border ?? 0,
    align: readAlign(child(node, 'alignment')) ?? parent?.align ?? null
  }
}

function readStyles(pkg: Package, workbookPart: string, theme: Theme, warnings: string[]): Styles {
  const palette: Palette = { theme, indexed: DEFAULT_INDEXED }
  const styles: Styles = {
    fonts: [DEFAULT_FONT], fills: [''], fillHex: [null], borders: [], xfs: [], fmts: new Map(), dxfs: [], tableStyles: new Map(), defaultTableStyle: 'TableStyleMedium2', palette
  }
  const rel = pkg.relOfType(workbookPart, 'styles')
  const part = rel && !rel.external ? rel.target : 'xl/styles.xml'
  let xml: XmlNode | null = null
  try {
    xml = pkg.xml(part)
  } catch (error) {
    warnings.push(`樣式無法讀取,以預設樣式顯示(${reason(error)})`)
    return styles
  }
  if (!xml) return styles
  const custom = children(path(xml, 'colors', 'indexedColors'), 'rgbColor').map((node) => (attr(node, 'rgb') ?? '').slice(-6))
  if (custom.length) palette.indexed = DEFAULT_INDEXED.map((hex, at) => (isHex(custom[at]) ? custom[at].toUpperCase() : hex))
  for (const fmt of children(child(xml, 'numFmts'), 'numFmt')) {
    const id = num(fmt, 'numFmtId')
    const code = attr(fmt, 'formatCode')
    if (id !== undefined && code !== undefined) styles.fmts.set(id, code.slice(0, 500))
  }
  const fontNodes = children(child(xml, 'fonts'), 'font')
  const base = readFont(fontNodes[0] ?? null, palette, DEFAULT_FONT)
  styles.fonts = fontNodes.length ? fontNodes.map((node, at) => (at === 0 ? base : readFont(node, palette, { ...DEFAULT_FONT, name: base.name, sz: base.sz }))) : [DEFAULT_FONT]
  const fills = children(child(xml, 'fills'), 'fill').map((node) => readFill(node, palette))
  if (fills.length) {
    styles.fills = fills.map((fill) => fill.css)
    styles.fillHex = fills.map((fill) => fill.hex)
  }
  styles.borders = children(child(xml, 'borders'), 'border').map((node) => readBorder(node, palette))
  const styleXfs = children(child(xml, 'cellStyleXfs'), 'xf').map((node) => readXf(node, undefined))
  styles.xfs = children(child(xml, 'cellXfs'), 'xf').map((node) => readXf(node, styleXfs[num(node, 'xfId') ?? 0]))
  styles.dxfs = children(child(xml, 'dxfs'), 'dxf').map((node) => {
    const font = child(node, 'font')
    const fill = child(node, 'fill')
    // In a dxf the pattern colour of a solid fill may sit in bgColor.
    const pattern = child(fill, 'patternFill')
    const solid = pattern ? xColor(child(pattern, 'bgColor'), palette) ?? xColor(child(pattern, 'fgColor'), palette) : null
    return {
      font: font ? { set: true, b: flag(font, 'b') || undefined, i: flag(font, 'i') || undefined, color: xColor(child(font, 'color'), palette) } : { set: false },
      fill: solid ? `#${solid}` : fill ? readFill(fill, palette).css : '',
      border: child(node, 'border') ? readBorder(child(node, 'border'), palette) : null
    }
  })
  const tableStyles = child(xml, 'tableStyles')
  styles.defaultTableStyle = attr(tableStyles, 'defaultTableStyle') ?? styles.defaultTableStyle
  for (const style of children(tableStyles, 'tableStyle')) {
    const elements = new Map<string, number>()
    for (const element of children(style, 'tableStyleElement')) {
      const type = attr(element, 'type')
      const dxf = num(element, 'dxfId')
      if (type && dxf !== undefined) elements.set(type, dxf)
    }
    styles.tableStyles.set(attr(style, 'name') ?? '', elements)
  }
  return styles
}

// ---------------------------------------------------------------------------------------------
// Shared strings and rich text

type Run = { text: string; font: XmlNode | null }
type RichText = { text: string; runs: Run[] | null }

/** `_x000D_` and friends: characters XML cannot carry. */
const unescapeOoxml = (text: string) =>
  text.includes('_x') ? text.replace(/_x([0-9A-Fa-f]{4})_/g, (_, hex: string) => String.fromCharCode(parseInt(hex, 16))) : text

function richText(node: XmlNode | null): RichText {
  if (!node) return { text: '', runs: null }
  const runs = children(node, 'r')
  if (!runs.length) return { text: unescapeOoxml(child(node, 't')?.text ?? ''), runs: null }
  const list: Run[] = runs.map((run) => ({ text: unescapeOoxml(children(run, 't').map((t) => t.text).join('')), font: child(run, 'rPr') }))
  const plain = child(node, 't')
  if (plain) list.unshift({ text: unescapeOoxml(plain.text), font: null })
  return { text: list.map((run) => run.text).join(''), runs: list }
}

/** Reads `<si>` items only when a cell asks for them, so a huge string table costs one index. */
class SharedStrings {
  private readonly starts: number[] = []
  private readonly cache = new Map<number, RichText>()
  private readonly source: string
  private readonly end: number

  constructor(source: string) {
    this.source = source
    const pattern = /<(?:[\w.-]+:)?si[\s>/]/g
    for (let match = pattern.exec(source); match; match = pattern.exec(source)) this.starts.push(match.index)
    this.end = source.lastIndexOf('</')
  }

  get(index: number): RichText | null {
    if (!(index >= 0 && index < this.starts.length)) return null
    let item = this.cache.get(index)
    if (!item) {
      const stop = this.starts[index + 1] ?? this.end
      try {
        item = richText(parseXml(this.source.slice(this.starts[index], stop)))
      } catch {
        item = { text: '', runs: null }
      }
      if (this.cache.size > 50_000) this.cache.clear()
      this.cache.set(index, item)
    }
    return item
  }
}

// ---------------------------------------------------------------------------------------------
// Sheets

/** `vm`: the cell's value metadata (1-based), which is how Excel 365 places a picture in a cell. */
type Cell = { c: number; s: number | null; t: string; v: string | null; f: string | null; is: XmlNode | null; vm: number | null }
type RowInfo = { ht?: number; hidden: boolean; s?: number; cells: Map<number, Cell> }

/** `cut`: the part was too large to read whole, so only its first rows were read and what follows sheetData is lost. */
type ScannedSheet = { head: XmlNode; rows: Map<number, RowInfo>; maxRow: number; maxCol: number; totalRows: number; totalCols: number; cut: boolean; exact: boolean }

const colIndex = (letters: string): number => {
  let n = 0
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n
}

export const colName = (index: number): string => {
  let name = ''
  for (let n = index; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name
  return name
}

const parseRef = (ref: string | undefined): { r: number; c: number } | null => {
  const match = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(ref ?? '')
  return match ? { c: colIndex(match[1]), r: Number(match[2]) } : null
}

type Range = { r1: number; c1: number; r2: number; c2: number }

const parseRange = (ref: string | undefined): Range | null => {
  const [a, b] = (ref ?? '').split(':')
  const start = parseRef(a)
  const end = b ? parseRef(b) : start
  if (!start || !end) return null
  return { r1: Math.min(start.r, end.r), c1: Math.min(start.c, end.c), r2: Math.max(start.r, end.r), c2: Math.max(start.c, end.c) }
}

/** The `r` of a raw `<row …>` start tag, quoted either way and with any space around `=`. */
const rowNumber = (tag: string): number | null => {
  const value = tagAttrs(tag).r
  return value !== undefined && /^\d{1,7}$/.test(value) ? Number(value) : null
}

/** The number of the last row in a stretch of sheetData, looking back from its end. */
function lastRowNumber(body: string, prefix: string): number | null {
  const token = `<${prefix}row`
  for (let at = body.lastIndexOf(token); at >= 0; at = at ? body.lastIndexOf(token, at - 1) : -1) {
    if (!/[\s>/]/.test(body[at + token.length] ?? '')) continue
    const close = body.indexOf('>', at)
    return close < 0 ? null : rowNumber(body.slice(at, close + 1))
  }
  return null
}

/**
 * A worksheet is read without building a tree for all of it: the parts around `<sheetData>` are
 * parsed whole, rows one at a time, and reading stops at the row limit.
 */
function scanSheet(source: string, rowLimit: number, colLimit: number, cellLimit = XLSX_LIMITS.sheetCells, cut = false): ScannedSheet {
  // `<sheetData/>`, `<sheetData />` and `<x:sheetData>` alike; a slash just before `>` closes it.
  const open = /<((?:[\w.-]+:)?)sheetData(?=[\s/>])[^>]*>/.exec(source)
  const rows = new Map<number, RowInfo>()
  let maxRow = 0
  let maxCol = 0
  let totalRows = 0
  let totalCols = 0
  let truncated = cut
  let head: XmlNode
  if (!open) head = parseXml(source, undefined, cut)
  else if (open[0].endsWith('/>')) head = parseXml(source.slice(0, open.index) + source.slice(open.index + open[0].length))
  else {
    const closeTag = `</${open[1]}sheetData>`
    const found = source.indexOf(closeTag, open.index)
    if (found < 0 && !cut) throw new Error('工作表的 sheetData 沒有結束')
    const close = found < 0 ? source.length : found
    head = parseXml(source.slice(0, open.index) + (found < 0 ? '' : source.slice(close + closeTag.length)), undefined, cut)
    const body = source.slice(open.index + open[0].length, close)
    const rowPattern = /<((?:[\w.-]+:)?)row[\s>/]/g
    let previous = 0
    for (let match = rowPattern.exec(body); match; match = rowPattern.exec(body)) {
      const tagEnd = body.indexOf('>', match.index)
      if (tagEnd < 0) break
      let end = tagEnd + 1
      if (body[tagEnd - 1] !== '/') {
        const rowClose = `</${match[1]}row>`
        const rowEnd = body.indexOf(rowClose, tagEnd)
        // In a part read only in part, the row the read stopped in is left out.
        if (rowEnd < 0 && cut) break
        if (rowEnd < 0) throw new Error('工作表的列沒有結束')
        end = rowEnd + rowClose.length
      }
      rowPattern.lastIndex = end
      const rAttr = rowNumber(body.slice(match.index, tagEnd + 1))
      const r = rAttr ?? previous + 1
      previous = r
      // Past the row limit, or once the grid drawn so far would use up the cell budget, stop reading.
      if (r > rowLimit || (r - 1) * Math.max(10, Math.min(colLimit, maxCol + 1)) >= cellLimit) {
        truncated = true
        totalRows = Math.max(totalRows, r)
        const lastR = lastRowNumber(body, match[1])
        if (lastR) totalRows = Math.max(totalRows, lastR)
        break
      }
      const row = parseXml(body.slice(match.index, end))
      const info: RowInfo = { hidden: ['1', 'true'].includes(attr(row, 'hidden') ?? ''), cells: new Map() }
      const ht = num(row, 'ht')
      if (ht !== undefined && ht >= 0 && ht < 2000) info.ht = ht
      if (['1', 'true'].includes(attr(row, 'customFormat') ?? '')) info.s = num(row, 's')
      let col = 0
      for (const cellNode of children(row, 'c')) {
        const ref = parseRef(attr(cellNode, 'r'))
        col = ref ? ref.c : col + 1
        totalCols = Math.max(totalCols, col)
        if (col > colLimit) continue
        const f = child(cellNode, 'f')
        const cell: Cell = {
          c: col,
          s: num(cellNode, 's') ?? null,
          t: attr(cellNode, 't') ?? 'n',
          v: child(cellNode, 'v')?.text ?? null,
          f: f ? f.text : null,
          is: child(cellNode, 'is'),
          vm: num(cellNode, 'vm') ?? null
        }
        info.cells.set(col, cell)
        maxCol = Math.max(maxCol, col)
      }
      if (info.cells.size || info.s !== undefined) maxRow = Math.max(maxRow, r)
      rows.set(r, info)
    }
  }
  totalRows = Math.max(totalRows, maxRow)
  // The declared extent is only trusted when reading stopped early; tools often overstate it.
  const dimension = truncated ? parseRange(attr(child(head, 'dimension'), 'ref')) : null
  if (dimension) {
    totalRows = Math.max(totalRows, Math.min(dimension.r2, 1_048_576))
    totalCols = Math.max(totalCols, Math.min(dimension.c2, 16_384))
  }
  // A cut part with no declared extent only says how many rows there are at least.
  return { head, rows, maxRow, maxCol, totalRows, totalCols, cut, exact: !cut || Boolean(dimension) }
}

type SheetEntry = { name: string; part: string | null; kind: string; state: string; tab: number }

const isShown = (sheet: SheetEntry) => sheet.state !== 'hidden' && sheet.state !== 'veryHidden'

type Workbook = {
  pkg: Package
  part: string
  sheets: SheetEntry[]
  date1904: boolean
  /** calcPr fullCalcOnLoad: the writer left formula results for Excel to work out on open. */
  recalc: boolean
  activeTab: number
  theme: Theme
  styles: Styles
  strings: SharedStrings | null
  warnings: string[]
  media: (part: string) => string
  scanned: Map<string, ScannedSheet>
  /** Pictures placed in cells, by value-metadata index; read when a sheet first asks. */
  cellPictures?: (vm: number) => CellPicture | null
}

/** A picture placed in a cell (Excel 365): a part of the file, or a picture from the web. */
type CellPicture = { target: string; alt: string } | { web: true; alt: string }

/**
 * The chain Excel writes for a picture in a cell: the cell's vm → metadata.xml valueMetadata →
 * futureMetadata XLRICHVALUE → a rich value (rdrichvalue.xml) whose structure names the image key →
 * richValueRel.xml → that part's relationship → the image part.
 */
function readCellPictures(wb: Workbook): (vm: number) => CellPicture | null {
  const partOf = (type: string, fallback: string) => {
    const rel = wb.pkg.relOfType(wb.part, type)
    return rel && !rel.external ? rel.target : fallback
  }
  const xml = (part: string) => {
    try {
      return wb.pkg.xml(part)
    } catch {
      return null
    }
  }
  const metadata = xml(partOf('sheetMetadata', 'xl/metadata.xml'))
  if (!metadata) return () => null
  const types = children(child(metadata, 'metadataTypes'), 'metadataType').map((node) => attr(node, 'name') ?? '')
  const future = children(metadata, 'futureMetadata').find((node) => attr(node, 'name') === 'XLRICHVALUE')
  const futureBlocks = children(future, 'bk')
  const valueBlocks = children(child(metadata, 'valueMetadata'), 'bk')
  const values = children(xml(partOf('rdRichValue', 'xl/richData/rdrichvalue.xml')), 'rv')
  const structures = children(xml(partOf('rdRichValueStructure', 'xl/richData/rdrichvaluestructure.xml')), 's')
  const relPart = partOf('richValueRel', 'xl/richData/richValueRel.xml')
  const relIds = children(xml(relPart), 'rel').map((node) => relAttr(node, 'id'))
  return (vm) => {
    const rc = child(valueBlocks[vm - 1], 'rc')
    const type = types[(num(rc, 't') ?? 0) - 1]
    if (type !== 'XLRICHVALUE') return null
    const block = futureBlocks[num(rc, 'v') ?? -1]
    const rvb = block ? findDeep(block, 'rvb') : null
    const value = values[num(rvb, 'i') ?? -1]
    if (!value) return null
    const keys = children(structures[num(value, 's') ?? -1], 'k').map((key) => attr(key, 'n') ?? '')
    const fields = children(value, 'v').map((field) => field.text)
    const field = (name: string) => {
      const at = keys.indexOf(name)
      return at >= 0 ? fields[at] : undefined
    }
    const alt = (field('Text') ?? '').slice(0, 500)
    if (field('_rvRel:WebImageIdentifier') !== undefined) return { web: true, alt }
    const local = field('_rvRel:LocalImageIdentifier') ?? (keys.length ? undefined : fields[0])
    if (local === undefined) return null
    const rel = wb.pkg.rel(relPart, relIds[Number(local)])
    if (!rel) return null
    if (rel.external) return { web: true, alt }
    return { target: rel.target, alt }
  }
}

function cellPicture(wb: Workbook, cell: Cell): CellPicture | null {
  if (cell.vm === null) return null
  wb.cellPictures ??= readCellPictures(wb)
  try {
    return wb.cellPictures(cell.vm)
  } catch {
    return null
  }
}

/** A worksheet's cells, read once per page however many charts point at it. */
function scannedSheet(wb: Workbook, part: string): ScannedSheet {
  let sheet = wb.scanned.get(part)
  if (!sheet) {
    // A sheet too large to inflate whole still shows its first rows: only the start of the part is read.
    const read = wb.pkg.zip.readPrefix(wb.pkg.stored(part))
    if (read === null) throw new Error(`找不到 ${part}`)
    sheet = scanSheet(read.data.toString('utf8'), XLSX_LIMITS.rows, XLSX_LIMITS.cols, XLSX_LIMITS.sheetCells, read.cut)
    wb.scanned.set(part, sheet)
  }
  return sheet
}

/**
 * XlsxWriter saves 0 for every formula it did not compute and asks Excel to recalculate on open. A formula
 * that really is 0 looks the same, so the 0 is shown, marked as possibly not computed.
 */
const STALE_NOTE = '存檔時可能沒有計算這個公式,Excel 開啟時會重新計算'
const uncomputed = (wb: Workbook, cell: Cell) => wb.recalc && Boolean(cell.f) && cell.v === '0' && (cell.t === 'n' || !cell.t)

/** The values of `'Sheet'!$A$1:$A$9` as a chart would plot them: numbers raw, text as shown. */
function lookupRange(wb: Workbook, formula: string): (string | null)[] {
  const match = /^(?:'((?:[^']|'')+)'|([^'!(),]+))!(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)$/.exec(formula.trim())
  if (!match) return []
  const name = match[1]?.replace(/''/g, "'") ?? match[2]
  const entry = wb.sheets.find((sheet) => sheet.name === name)
  const range = parseRange(match[3].replace(/\$/g, ''))
  if (!entry?.part || !range) return []
  try {
    const sheet = scannedSheet(wb, entry.part)
    const out: (string | null)[] = []
    for (let r = range.r1; r <= range.r2 && out.length < XLSX_LIMITS.chartPoints; r++) {
      for (let c = range.c1; c <= range.c2 && out.length < XLSX_LIMITS.chartPoints; c++) {
        const cell = sheet.rows.get(r)?.cells.get(c)
        if (!cell) out.push(null)
        else if ((cell.t === 'n' || cell.t === 'str' || cell.t === 'b') && cell.v !== null) out.push(cell.v)
        else out.push(showCell(cell, cell.s ?? 0, wb, false).text || null)
      }
    }
    return out
  } catch {
    return []
  }
}

function openWorkbook(data: Buffer, mediaBase: string): Workbook {
  if (data.length >= 8 && data.readUInt32BE(0) === 0xd0cf11e0) {
    throw new Error('這是舊版 .xls 或加密的活頁簿,網頁版無法顯示,請下載原始檔查看')
  }
  let pkg: Package
  try {
    pkg = new Package(data)
  } catch (error) {
    throw new Error(`這不是 .xlsx 檔,或檔案已損壞(${reason(error)})`)
  }
  const rel = pkg.relOfType('', 'officeDocument')
  const part = rel && !rel.external && pkg.has(rel.target) ? rel.target : 'xl/workbook.xml'
  const xml = pkg.xml(part)
  if (!xml || local(xml.name) !== 'workbook') throw new Error('這不是 Excel 活頁簿(找不到 workbook.xml)')
  const warnings: string[] = []
  const pr = child(xml, 'workbookPr')
  const date1904 = ['1', 'true'].includes(attr(pr, 'date1904') ?? '')
  const recalc = ['1', 'true'].includes(attr(child(xml, 'calcPr'), 'fullCalcOnLoad') ?? '')
  const sheets = children(child(xml, 'sheets'), 'sheet').map((node, tab): SheetEntry => {
    const sheetRel = pkg.rel(part, relAttr(node, 'id'))
    const kind = sheetRel?.type.split('/').pop() ?? ''
    return { name: attr(node, 'name') ?? `工作表${tab + 1}`, part: sheetRel && !sheetRel.external ? sheetRel.target : null, kind, state: attr(node, 'state') ?? 'visible', tab }
  })
  const activeTab = num(path(xml, 'bookViews', 'workbookView'), 'activeTab') ?? 0
  const theme = readTheme(pkg, part, warnings)
  const styles = readStyles(pkg, part, theme, warnings)
  let strings: SharedStrings | null = null
  const stringsRel = pkg.relOfType(part, 'sharedStrings')
  try {
    const source = pkg.text(stringsRel && !stringsRel.external ? stringsRel.target : 'xl/sharedStrings.xml')
    if (source) strings = new SharedStrings(source)
  } catch (error) {
    warnings.push(`共用字串無法讀取,文字儲存格會是空白(${reason(error)})`)
  }
  return { pkg, part, sheets, date1904, recalc, activeTab, theme, styles, strings, warnings, media: (target) => officeMediaUrl(mediaBase, target), scanned: new Map() }
}

/**
 * What went wrong, for the page and the error log. The reader's own errors are already in
 * Traditional Chinese; a runtime error (a file shaped in a way the code did not expect) is named so.
 */
const reason = (error: unknown) => {
  const message = error instanceof Error ? error.message : String(error)
  const runtime = error instanceof TypeError || error instanceof RangeError || error instanceof SyntaxError || error instanceof ReferenceError
  return (runtime ? `檔案內容不符預期(${message})` : message).slice(0, 300)
}

// ---------------------------------------------------------------------------------------------
// Cell values

type Shown = { text: string; html: string; color: string | null; kind: 'num' | 'text' | 'bool' | 'err' | 'formula' | 'empty' | 'picture'; stale?: boolean }

function formatCode(wb: Workbook, xf: Xf | undefined): string {
  const id = xf?.fmt ?? 0
  return wb.styles.fmts.get(id) ?? BUILTIN_FORMATS[id] ?? 'General'
}

function runsHtml(runs: Run[], wb: Workbook, cellFont: Font): string {
  return runs
    .map((run) => {
      const text = escapeHtml(run.text)
      if (!run.font) return text
      const font = readFont(run.font, wb.styles.palette, { ...cellFont, b: false, i: false, u: '', strike: false, vert: '' })
      const css = fontCss(font, cellFont, wb.theme, true)
      return css ? `<span style="${css}">${text}</span>` : text
    })
    .join('')
}

function isoToSerial(value: string, date1904: boolean): number | null {
  const match = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?)?/.exec(value)
  if (!match) {
    const time = /^(\d{2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?/.exec(value)
    return time ? (Number(time[1]) * 3600 + Number(time[2]) * 60 + Number(time[3] ?? 0)) / 86400 : null
  }
  const ms = Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3]), Number(match[4] ?? 0), Number(match[5] ?? 0)) + Number(match[6] ?? 0) * 1000
  if (date1904) return (ms - Date.UTC(1904, 0, 1)) / 86400000
  const serial = (ms - Date.UTC(1899, 11, 30)) / 86400000
  return serial < 61 ? serial - 1 : serial
}

function showCell(cell: Cell, xfIndex: number, wb: Workbook, wantHtml: boolean): Shown {
  const xf = wb.styles.xfs[xfIndex]
  const code = formatCode(wb, xf)
  const palette = wb.styles.palette.indexed
  const plain = (full: string, kind: Shown['kind'], color: string | null = null): Shown => {
    const text = full.length > XLSX_LIMITS.cellChars ? full.slice(0, XLSX_LIMITS.cellChars) : full
    return { text, html: wantHtml ? escapeHtml(text) : '', color, kind }
  }
  const asText = (rich: RichText): Shown => {
    const formatted = formatText(code, rich.text, palette)
    if (formatted.text === rich.text && rich.runs && wantHtml && rich.text.length <= XLSX_LIMITS.cellChars) {
      const font = wb.styles.fonts[xf?.font ?? 0] ?? wb.styles.fonts[0]
      return { text: rich.text, html: runsHtml(rich.runs, wb, font), color: formatted.color, kind: 'text' }
    }
    return plain(formatted.text, 'text', formatted.color)
  }
  const value = cell.v
  switch (cell.t) {
    case 's': {
      const rich = value === null ? null : wb.strings?.get(Number(value))
      return rich ? asText(rich) : plain('', 'empty')
    }
    case 'inlineStr':
      return asText(richText(cell.is))
    case 'str':
      if (value === null) break
      return asText({ text: unescapeOoxml(value), runs: null })
    case 'b':
      if (value === null) break
      return plain(value === '1' || value === 'true' ? 'TRUE' : 'FALSE', 'bool')
    case 'e':
      if (value === null) break
      return plain(value, 'err')
    case 'd': {
      if (value === null) break
      const serial = isoToSerial(value, wb.date1904)
      if (serial === null) return plain(value, 'text')
      const formatted = formatNumber(compileFormat(code, palette).date ? code : 'yyyy/m/d h:mm', serial, wb.date1904, palette)
      return plain(formatted.text, 'num', formatted.color)
    }
    default: {
      if (value === null || value === '') break
      const number = Number(value)
      if (!Number.isFinite(number)) return plain(value, 'text')
      const formatted = formatNumber(code, number, wb.date1904, palette)
      return { ...plain(formatted.text, 'num', formatted.color), stale: uncomputed(wb, cell) }
    }
  }
  if (cell.f !== null && cell.f !== '') return plain(`=${cell.f}`, 'formula')
  return plain('', 'empty')
}

// ---------------------------------------------------------------------------------------------
// Tables (list objects): the file's own table style, or Excel's built-in look drawn from the theme

type TableLook = { fill?: string; color?: string; bold?: boolean; top?: Side; bottom?: Side }

function tableLooks(wb: Workbook, sheetPart: string, head: XmlNode): Map<number, TableLook> {
  const looks = new Map<number, TableLook>()
  for (const tablePart of children(child(head, 'tableParts'), 'tablePart')) {
    try {
      const rel = wb.pkg.rel(sheetPart, relAttr(tablePart, 'id'))
      if (!rel || rel.external) continue
      const table = wb.pkg.xml(rel.target)
      const range = parseRange(attr(table, 'ref'))
      if (!table || !range) continue
      const info = child(table, 'tableStyleInfo')
      const name = attr(info, 'name') ?? ''
      if (!name) continue
      const stripes = attr(info, 'showRowStripes') !== '0'
      const headerRows = num(table, 'headerRowCount') ?? 1
      const totalRows = num(table, 'totalsRowCount') ?? (attr(table, 'totalsRowShown') === '1' && num(table, 'totalsRowCount') ? 1 : 0)
      const look = (kind: 'header' | 'total' | 'band' | 'body'): TableLook => tableLook(wb, name, kind)
      for (let r = range.r1; r <= Math.min(range.r2, XLSX_LIMITS.rows); r++) {
        const offset = r - range.r1 - headerRows
        const kind = r < range.r1 + headerRows ? 'header' : r > range.r2 - totalRows ? 'total' : stripes && offset % 2 === 0 ? 'band' : 'body'
        const style = look(kind)
        for (let c = range.c1; c <= Math.min(range.c2, XLSX_LIMITS.cols); c++) looks.set(r * 20000 + c, style)
      }
    } catch {}
  }
  return looks
}

function tableLook(wb: Workbook, name: string, kind: 'header' | 'total' | 'band' | 'body'): TableLook {
  const custom = wb.styles.tableStyles.get(name)
  if (custom) {
    const pick = (type: string): TableLook => {
      const dxf = wb.styles.dxfs[custom.get(type) ?? -1]
      if (!dxf) return {}
      return { fill: dxf.fill || undefined, color: dxf.font.color ?? undefined, bold: dxf.font.b, top: dxf.border?.top, bottom: dxf.border?.bottom }
    }
    const whole = pick('wholeTable')
    const own = kind === 'header' ? pick('headerRow') : kind === 'total' ? pick('totalRow') : kind === 'band' ? pick('firstRowStripe') : {}
    return { ...whole, ...Object.fromEntries(Object.entries(own).filter(([, v]) => v !== undefined)) }
  }
  const match = /^TableStyle(Light|Medium|Dark)(\d+)$/.exec(name)
  if (!match) return {}
  const n = Number(match[2])
  const slot = (n - 1) % 7
  const accent = slot === 0 ? wb.theme.scheme.dk1 : wb.theme.scheme[`accent${slot}`] ?? '4472C4'
  const line = (hex: string): Side => ({ style: 'thin', color: hex, weight: 6 })
  const family = match[1]
  if (family === 'Light') {
    const filledHeader = n >= 8 && n <= 14
    if (kind === 'header') return filledHeader ? { fill: `#${accent}`, color: contrastOn(accent), bold: true } : { bold: true, bottom: line(accent) }
    if (kind === 'total') return { bold: true, top: { style: 'double', color: accent, weight: 12 } }
    if (kind === 'band') return { fill: `#${applyTint(accent, 0.8)}` }
    return {}
  }
  if (family === 'Medium') {
    if (kind === 'header') return { fill: `#${accent}`, color: contrastOn(accent), bold: true }
    if (kind === 'total') return { bold: true, top: { style: 'double', color: accent, weight: 12 } }
    if (kind === 'band') return { fill: `#${applyTint(accent, n >= 22 ? 0.6 : 0.8)}` }
    return n >= 22 ? { fill: `#${applyTint(accent, 0.8)}` } : {}
  }
  if (kind === 'header') return { fill: `#${wb.theme.scheme.dk1}`, color: contrastOn(wb.theme.scheme.dk1), bold: true }
  const body = applyTint(accent, kind === 'band' ? -0.25 : 0)
  return { fill: `#${body}`, color: contrastOn(body), bold: kind === 'total' }
}

// ---------------------------------------------------------------------------------------------
// Conditional formatting: what each rule decides from the values on the sheet

type CfLook = { fill?: string; color?: string; bold?: boolean; italic?: boolean; bar?: string; icon?: { glyph: string; color: string }; hideValue?: boolean }

const CF_CELLS = 60_000

/** What a formula rule sees in a cell, and what it may produce. */
type FValue = number | string | boolean | null | FError

class FError {
  readonly code: string
  constructor(code: string) {
    this.code = code
  }
}

/** A formula this evaluator does not cover (a range, another sheet, a function it lacks): the rule is left out and said so. */
class Unsupported extends Error {}

type FNode =
  | { k: 'lit'; v: FValue }
  | { k: 'ref'; r: number; c: number; absR: boolean; absC: boolean }
  | { k: 'un'; op: string; a: FNode }
  | { k: 'bin'; op: string; a: FNode; b: FNode }
  | { k: 'fn'; name: string; args: FNode[] }

type FToken = { t: 'num' | 'str' | 'ref' | 'fn' | 'bool' | 'op'; v: string }

const FORMULA_TOKEN = /\s+|(\d+(?:\.\d*)?(?:[eE][+-]?\d+)?|\.\d+(?:[eE][+-]?\d+)?)|"((?:[^"]|"")*)"|(\$?[A-Za-z]{1,3}\$?\d{1,7})(?![\w(!:])|([A-Za-z_][\w.]*)(?=\s*\()|(TRUE|FALSE)(?![\w(!:])|(<>|<=|>=|[-+*/^&=<>(),%])/y

/** Parses the formula of a conditional format: the small part of Excel's language these rules use. */
function compileFormula(source: string): FNode {
  const tokens: FToken[] = []
  FORMULA_TOKEN.lastIndex = 0
  let at = source.startsWith('=') ? 1 : 0
  while (at < source.length) {
    FORMULA_TOKEN.lastIndex = at
    const m = FORMULA_TOKEN.exec(source)
    if (!m || m[0] === '') throw new Unsupported()
    at = FORMULA_TOKEN.lastIndex
    if (m[1] !== undefined) tokens.push({ t: 'num', v: m[1] })
    else if (m[2] !== undefined) tokens.push({ t: 'str', v: m[2].replace(/""/g, '"') })
    else if (m[3] !== undefined) tokens.push({ t: 'ref', v: m[3] })
    else if (m[4] !== undefined) tokens.push({ t: 'fn', v: m[4].toUpperCase() })
    else if (m[5] !== undefined) tokens.push({ t: 'bool', v: m[5].toUpperCase() })
    else if (m[6] !== undefined) tokens.push({ t: 'op', v: m[6] })
    if (tokens.length > 500) throw new Unsupported()
  }
  let pos = 0
  const isOp = (...ops: string[]) => tokens[pos]?.t === 'op' && ops.includes(tokens[pos].v)
  const expect = (op: string) => {
    if (!isOp(op)) throw new Unsupported()
    pos++
  }
  const binary = (next: () => FNode, ...ops: string[]) => (): FNode => {
    let a = next()
    while (isOp(...ops)) {
      const op = tokens[pos++].v
      a = { k: 'bin', op, a, b: next() }
    }
    return a
  }
  const primary = (): FNode => {
    const token = tokens[pos++]
    if (!token) throw new Unsupported()
    if (token.t === 'num') return { k: 'lit', v: Number(token.v) }
    if (token.t === 'str') return { k: 'lit', v: token.v }
    if (token.t === 'bool') return { k: 'lit', v: token.v === 'TRUE' }
    if (token.t === 'ref') {
      const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(token.v)!
      return { k: 'ref', c: colIndex(m[2]), r: Number(m[4]), absC: m[1] === '$', absR: m[3] === '$' }
    }
    if (token.t === 'fn') {
      expect('(')
      const args: FNode[] = []
      if (!isOp(')')) {
        args.push(compare())
        while (isOp(',')) {
          pos++
          args.push(compare())
        }
      }
      expect(')')
      return { k: 'fn', name: token.v, args }
    }
    if (token.v === '(') {
      const inner = compare()
      expect(')')
      return inner
    }
    throw new Unsupported()
  }
  const percent = (): FNode => {
    let a = primary()
    while (isOp('%')) {
      pos++
      a = { k: 'bin', op: '/', a, b: { k: 'lit', v: 100 } }
    }
    return a
  }
  const unary = (): FNode => {
    if (isOp('-', '+')) {
      const op = tokens[pos++].v
      return { k: 'un', op, a: unary() }
    }
    return percent()
  }
  const power = binary(unary, '^')
  const product = binary(power, '*', '/')
  const sum = binary(product, '+', '-')
  const concat = binary(sum, '&')
  const compare: () => FNode = binary(concat, '=', '<>', '<', '>', '<=', '>=')
  const tree = compare()
  if (pos !== tokens.length) throw new Unsupported()
  return tree
}

const fNumber = (value: FValue): number => {
  if (value instanceof FError) throw value
  if (value === null) return 0
  if (typeof value === 'boolean') return value ? 1 : 0
  if (typeof value === 'number') return value
  const trimmed = value.trim()
  const parsed = Number(trimmed)
  if (trimmed === '' || !Number.isFinite(parsed)) throw new FError('#VALUE!')
  return parsed
}

const fText = (value: FValue): string => {
  if (value instanceof FError) throw value
  if (value === null) return ''
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
  return typeof value === 'number' ? formatGeneral(value) : value
}

const fBool = (value: FValue): boolean => {
  if (value instanceof FError) throw value
  if (value === null) return false
  if (typeof value === 'boolean') return value
  if (typeof value === 'number') return value !== 0
  const upper = value.toUpperCase()
  if (upper === 'TRUE' || upper === 'FALSE') return upper === 'TRUE'
  throw new FError('#VALUE!')
}

/** Excel's order across kinds: numbers, then text (ignoring case), then logicals; an empty cell is 0, "" or FALSE. */
function fCompare(a: FValue, b: FValue): number {
  if (a instanceof FError) throw a
  if (b instanceof FError) throw b
  const blank = (other: FValue): FValue => (typeof other === 'string' ? '' : typeof other === 'boolean' ? false : 0)
  const x = a === null ? blank(b) : a
  const y = b === null ? blank(a) : b
  const rank = (v: FValue) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2)
  if (rank(x) !== rank(y)) return rank(x) - rank(y)
  if (typeof x === 'string') {
    const [p, q] = [x.toLowerCase(), (y as string).toLowerCase()]
    return p < q ? -1 : p > q ? 1 : 0
  }
  const [p, q] = [Number(x), Number(y)]
  return p < q ? -1 : p > q ? 1 : 0
}

type FContext = { r: number; c: number; base: { r: number; c: number }; at: (r: number, c: number) => FValue }

function evaluateFormula(node: FNode, ctx: FContext): FValue {
  const ev = (n: FNode) => evaluateFormula(n, ctx)
  const place = (n: Extract<FNode, { k: 'ref' }>) => ({ r: n.absR ? n.r : n.r + ctx.r - ctx.base.r, c: n.absC ? n.c : n.c + ctx.c - ctx.base.c })
  switch (node.k) {
    case 'lit':
      return node.v
    case 'ref': {
      const { r, c } = place(node)
      return r < 1 || c < 1 ? new FError('#REF!') : ctx.at(r, c)
    }
    case 'un': {
      const value = fNumber(ev(node.a))
      return node.op === '-' ? -value : value
    }
    case 'bin': {
      const a = ev(node.a)
      const b = ev(node.b)
      switch (node.op) {
        case '&':
          return fText(a) + fText(b)
        case '+':
          return fNumber(a) + fNumber(b)
        case '-':
          return fNumber(a) - fNumber(b)
        case '*':
          return fNumber(a) * fNumber(b)
        case '/': {
          const d = fNumber(b)
          if (d === 0) throw new FError('#DIV/0!')
          return fNumber(a) / d
        }
        case '^':
          return fNumber(a) ** fNumber(b)
        default: {
          const order = fCompare(a, b)
          return node.op === '=' ? order === 0 : node.op === '<>' ? order !== 0 : node.op === '<' ? order < 0 : node.op === '>' ? order > 0 : node.op === '<=' ? order <= 0 : order >= 0
        }
      }
    }
    case 'fn': {
      const args = node.args
      const arg = (at: number) => (args[at] ? ev(args[at]) : null)
      const count = (min: number, max = min) => {
        if (args.length < min || args.length > max) throw new Unsupported()
      }
      switch (node.name) {
        case 'AND':
        case 'OR': {
          if (!args.length) throw new Unsupported()
          const values = args.map((a) => fBool(ev(a)))
          return node.name === 'AND' ? values.every(Boolean) : values.some(Boolean)
        }
        case 'NOT':
          count(1)
          return !fBool(arg(0))
        case 'IF':
          count(2, 3)
          return fBool(arg(0)) ? arg(1) : args[2] ? arg(2) : false
        case 'IFERROR':
          count(2)
          try {
            const value = arg(0)
            if (value instanceof FError) return arg(1)
            return value
          } catch (error) {
            if (error instanceof FError) return arg(1)
            throw error
          }
        case 'ISERROR':
        case 'ISERR':
        case 'ISNA':
          count(1)
          try {
            const value = arg(0)
            return value instanceof FError && (node.name !== 'ISNA' || value.code === '#N/A') && (node.name !== 'ISERR' || value.code !== '#N/A')
          } catch (error) {
            if (!(error instanceof FError)) throw error
            return (node.name !== 'ISNA' || error.code === '#N/A') && (node.name !== 'ISERR' || error.code !== '#N/A')
          }
        case 'ISBLANK':
          count(1)
          return arg(0) === null
        case 'ISNUMBER':
          count(1)
          return typeof arg(0) === 'number'
        case 'ISTEXT':
          count(1)
          return typeof arg(0) === 'string'
        case 'ISLOGICAL':
          count(1)
          return typeof arg(0) === 'boolean'
        case 'ISEVEN':
        case 'ISODD': {
          count(1)
          const even = Math.trunc(fNumber(arg(0))) % 2 === 0
          return node.name === 'ISEVEN' ? even : !even
        }
        case 'ROW':
        case 'COLUMN': {
          count(0, 1)
          const target = args[0]
          if (target && target.k !== 'ref') throw new Unsupported()
          const where = target ? place(target) : { r: ctx.r, c: ctx.c }
          return node.name === 'ROW' ? where.r : where.c
        }
        case 'MOD': {
          count(2)
          const [a, b] = [fNumber(arg(0)), fNumber(arg(1))]
          if (b === 0) throw new FError('#DIV/0!')
          return a - b * Math.floor(a / b)
        }
        case 'ABS':
          count(1)
          return Math.abs(fNumber(arg(0)))
        case 'INT':
          count(1)
          return Math.floor(fNumber(arg(0)))
        case 'ROUND': {
          count(1, 2)
          const scale = 10 ** Math.trunc(fNumber(arg(1)))
          const value = fNumber(arg(0)) * scale
          return (Math.sign(value) * Math.round(Math.abs(value))) / scale
        }
        case 'MIN':
        case 'MAX':
        case 'SUM':
        case 'AVERAGE': {
          if (!args.length) throw new Unsupported()
          const values = args.map((a) => fNumber(ev(a)))
          if (node.name === 'MIN') return Math.min(...values)
          if (node.name === 'MAX') return Math.max(...values)
          const total = values.reduce((s, v) => s + v, 0)
          return node.name === 'SUM' ? total : total / values.length
        }
        case 'LEN':
          count(1)
          return [...fText(arg(0))].length
        case 'TRIM':
          count(1)
          return fText(arg(0)).replace(/ +/g, ' ').trim()
        case 'UPPER':
          count(1)
          return fText(arg(0)).toUpperCase()
        case 'LOWER':
          count(1)
          return fText(arg(0)).toLowerCase()
        case 'LEFT':
        case 'RIGHT': {
          count(1, 2)
          const text = [...fText(arg(0))]
          const n = args[1] ? Math.max(0, Math.trunc(fNumber(arg(1)))) : 1
          return (node.name === 'LEFT' ? text.slice(0, n) : text.slice(Math.max(0, text.length - n))).join('')
        }
        case 'MID': {
          count(3)
          const text = [...fText(arg(0))]
          const start = Math.trunc(fNumber(arg(1)))
          if (start < 1) throw new FError('#VALUE!')
          return text.slice(start - 1, start - 1 + Math.max(0, Math.trunc(fNumber(arg(2))))).join('')
        }
        case 'EXACT':
          count(2)
          return fText(arg(0)) === fText(arg(1))
        case 'SEARCH':
        case 'FIND': {
          count(2, 3)
          let needle = fText(arg(0))
          let hay = fText(arg(1))
          if (node.name === 'SEARCH') {
            if (/[*?~]/.test(needle)) throw new Unsupported()
            needle = needle.toLowerCase()
            hay = hay.toLowerCase()
          }
          const from = args[2] ? Math.trunc(fNumber(arg(2))) : 1
          const found = from >= 1 ? hay.indexOf(needle, from - 1) : -1
          if (found < 0) throw new FError('#VALUE!')
          return found + 1
        }
        case 'VALUE':
          count(1)
          return fNumber(arg(0))
        case 'WEEKDAY': {
          count(1, 2)
          const serial = Math.floor(fNumber(arg(0)))
          const kind = args[1] ? fNumber(arg(1)) : 1
          if (kind !== 1 && kind !== 2 && kind !== 3) throw new Unsupported()
          const sunday1 = (((serial - 1) % 7) + 7) % 7 + 1
          return kind === 1 ? sunday1 : kind === 2 ? ((sunday1 + 5) % 7) + 1 : (sunday1 + 5) % 7
        }
        default:
          // TODAY(), NOW(), lookups and the rest: not evaluated here.
          throw new Unsupported()
      }
    }
  }
}

/** The icons of Excel's icon sets, lowest value first: a glyph and its colour. */
const ICON_SETS: Record<string, [string, string][]> = (() => {
  const [G, Y, R, K, S, P, B] = ['1E9E4A', 'F2B705', 'D9363E', '262626', '8C8C8C', 'F4A6A6', '4472C4']
  const gray = (set: [string, string][]) => set.map(([glyph]): [string, string] => [glyph, S])
  const arrows3: [string, string][] = [['↓', R], ['→', Y], ['↑', G]]
  const arrows4: [string, string][] = [['↓', R], ['↘', Y], ['↗', Y], ['↑', G]]
  const arrows5: [string, string][] = [['↓', R], ['↘', Y], ['→', Y], ['↗', Y], ['↑', G]]
  const lights: [string, string][] = [['●', R], ['●', Y], ['●', G]]
  const symbols: [string, string][] = [['✖', R], ['!', Y], ['✔', G]]
  return {
    '3Arrows': arrows3, '3ArrowsGray': gray(arrows3), '3Flags': [['⚑', R], ['⚑', Y], ['⚑', G]], '3TrafficLights1': lights, '3TrafficLights2': lights,
    '3Signs': [['◆', R], ['▲', Y], ['●', G]], '3Symbols': symbols, '3Symbols2': symbols, '3Stars': [['☆', Y], ['✭', Y], ['★', Y]], '3Triangles': [['▼', R], ['▬', Y], ['▲', G]],
    '4Arrows': arrows4, '4ArrowsGray': gray(arrows4), '4RedToBlack': [['●', K], ['●', S], ['●', P], ['●', R]], '4Rating': [['▂', B], ['▄', B], ['▆', B], ['█', B]],
    '4TrafficLights': [['●', K], ['●', R], ['●', Y], ['●', G]], '5Arrows': arrows5, '5ArrowsGray': gray(arrows5),
    '5Rating': [['▁', B], ['▂', B], ['▄', B], ['▆', B], ['█', B]], '5Quarters': [['○', K], ['◔', K], ['◑', K], ['◕', K], ['●', K]], '5Boxes': [['▁', B], ['▃', B], ['▅', B], ['▇', B], ['█', B]]
  }
})()

type CfResult = { looks: Map<number, CfLook>; skipped: Set<string> }

function conditionalLooks(wb: Workbook, head: XmlNode, rows: Map<number, RowInfo>, lastRow: number, lastCol: number): CfResult {
  const looks = new Map<number, CfLook>()
  const skipped = new Set<string>()
  let budget = CF_CELLS
  const numberAt = (r: number, c: number): number | null => {
    const cell = rows.get(r)?.cells.get(c)
    if (!cell || cell.v === null || (cell.t !== 'n' && cell.t !== 'str')) return null
    const value = Number(cell.v)
    return cell.v !== '' && Number.isFinite(value) ? value : null
  }
  /** What a formula sees in a cell: a number, text, a logical, an error or nothing. */
  const valueAt = (r: number, c: number): FValue => {
    const cell = rows.get(r)?.cells.get(c)
    if (!cell) return null
    const v = cell.v
    switch (cell.t) {
      case 's':
        return v === null ? null : (wb.strings?.get(Number(v))?.text ?? null)
      case 'inlineStr':
        return richText(cell.is).text
      case 'str':
        return v === null ? null : unescapeOoxml(v)
      case 'b':
        return v === null ? null : v === '1' || v === 'true'
      case 'e':
        return v === null ? null : new FError(v)
      case 'd':
        return v === null ? null : isoToSerial(v, wb.date1904)
      default: {
        if (v === null || v === '') return null
        const n = Number(v)
        return Number.isFinite(n) ? n : v
      }
    }
  }
  /** The text a cell shows, for the text rules. */
  const shownAt = (r: number, c: number): string | null => {
    const cell = rows.get(r)?.cells.get(c)
    return cell ? showCell(cell, cell.s ?? rows.get(r)?.s ?? 0, wb, false).text : null
  }
  const cellsOf = (sqref: string) => {
    const cells: [number, number][] = []
    let base: { r: number; c: number } | null = null
    for (const ref of sqref.split(/\s+/)) {
      const range = parseRange(ref)
      if (!range) continue
      base ??= { r: range.r1, c: range.c1 }
      for (let r = range.r1; r <= Math.min(range.r2, lastRow) && budget > 0; r++) {
        for (let c = range.c1; c <= Math.min(range.c2, lastCol) && budget > 0; c++) {
          cells.push([r, c])
          budget--
        }
      }
    }
    return { cells, base: base ?? { r: 1, c: 1 } }
  }
  type Rule = { node: XmlNode; priority: number; cells: [number, number][]; base: { r: number; c: number } }
  const rules: Rule[] = []
  const extended = new Set<string>()
  for (const block of children(head, 'conditionalFormatting')) {
    const { cells, base } = cellsOf(attr(block, 'sqref') ?? '')
    for (const node of children(block, 'cfRule')) {
      rules.push({ node, priority: num(node, 'priority') ?? 1e9, cells, base })
      for (const ext of children(child(node, 'extLst'), 'ext')) {
        const id = child(ext, 'id')?.text
        if (id) extended.add(id)
      }
    }
  }
  // Excel 2010 rules live in the sheet's extension list: icon sets with custom icons and the rest.
  for (const ext of children(child(head, 'extLst'), 'ext')) {
    for (const block of children(child(ext, 'conditionalFormattings'), 'conditionalFormatting')) {
      const { cells, base } = cellsOf(child(block, 'sqref')?.text ?? '')
      for (const node of children(block, 'cfRule')) {
        // A data bar's extension adds to a rule already read above.
        if (extended.has(attr(node, 'id') ?? '')) continue
        if (attr(node, 'type') === 'iconSet') rules.push({ node, priority: num(node, 'priority') ?? 1e9, cells, base })
        else if (cells.length) skipped.add('Excel 2010 新增的規則')
      }
    }
  }
  rules.sort((a, b) => a.priority - b.priority)
  const apply = (r: number, c: number, look: CfLook) => {
    const key = r * 20000 + c
    const had = looks.get(key) ?? {}
    // The rule with the higher priority (lower number) keeps what it set.
    looks.set(key, { ...look, ...had })
  }
  const dxfLook = (node: XmlNode): CfLook | null => {
    const dxf = wb.styles.dxfs[num(node, 'dxfId') ?? -1]
    return dxf ? { fill: dxf.fill || undefined, color: dxf.font.color ?? undefined, bold: dxf.font.b, italic: dxf.font.i } : null
  }
  const yes = (node: XmlNode | null | undefined, name: string, otherwise = false) => {
    const value = attr(node, name)
    return value === undefined ? otherwise : value === '1' || value === 'true'
  }
  for (const { node, cells, base } of rules) {
    try {
      const type = attr(node, 'type')
      const values = cells.map(([r, c]) => numberAt(r, c))
      const numbers = values.filter((v): v is number => v !== null)
      const each = (test: (at: number, r: number, c: number) => boolean) => {
        const look = dxfLook(node)
        if (!look) return
        cells.forEach(([r, c], at) => {
          if (test(at, r, c)) apply(r, c, look)
        })
      }
      if (type === 'cellIs') {
        // The operands are formulas too ("完成", $D$1, A1*2), worked out for each cell from the range's first cell.
        let formulas: FNode[]
        try {
          formulas = children(node, 'formula').map((f) => compileFormula(f.text))
          if (!formulas.length) throw new Unsupported()
        } catch {
          skipped.add('比較規則')
          continue
        }
        const op = attr(node, 'operator') ?? 'equal'
        let unsupported = false
        each((_, r, c) => {
          if (unsupported) return false
          try {
            const value = valueAt(r, c)
            if (value instanceof FError) return false
            const [x, y] = formulas.map((formula) => evaluateFormula(formula, { r, c, base, at: valueAt }))
            const vs = (operand: FValue | undefined) => fCompare(value, operand ?? null)
            if (op === 'between' || op === 'notBetween') {
              const [lo, hi] = fCompare(x, y ?? null) <= 0 ? [x, y] : [y, x]
              const inside = vs(lo) >= 0 && vs(hi) <= 0
              return op === 'between' ? inside : !inside
            }
            const order = vs(x)
            return op === 'greaterThan' ? order > 0 : op === 'greaterThanOrEqual' ? order >= 0 : op === 'lessThan' ? order < 0 : op === 'lessThanOrEqual' ? order <= 0 : op === 'notEqual' ? order !== 0 : order === 0
          } catch (error) {
            if (error instanceof Unsupported) unsupported = true
            return false
          }
        })
        if (unsupported) skipped.add('比較規則')
      } else if (type === 'expression') {
        let formula: FNode
        try {
          formula = compileFormula(child(node, 'formula')?.text ?? '')
        } catch {
          skipped.add('公式規則')
          continue
        }
        let unsupported = false
        each((_, r, c) => {
          if (unsupported) return false
          try {
            const result = evaluateFormula(formula, { r, c, base, at: valueAt })
            return typeof result === 'number' ? result !== 0 : result === true
          } catch (error) {
            if (error instanceof Unsupported) unsupported = true
            return false
          }
        })
        if (unsupported) skipped.add('公式規則')
      } else if (type === 'top10') {
        if (!numbers.length) continue
        const rank = num(node, 'rank') ?? 10
        const bottom = yes(node, 'bottom')
        const sorted = [...numbers].sort((x, y) => (bottom ? x - y : y - x))
        const n = yes(node, 'percent') ? Math.max(1, Math.floor((sorted.length * rank) / 100)) : Math.trunc(rank)
        if (n < 1) continue
        const edge = sorted[Math.min(n, sorted.length) - 1]
        each((at) => values[at] !== null && (bottom ? values[at]! <= edge : values[at]! >= edge))
      } else if (type === 'aboveAverage') {
        if (!numbers.length) continue
        const mean = numbers.reduce((s, v) => s + v, 0) / numbers.length
        const spread = Math.sqrt(numbers.reduce((s, v) => s + (v - mean) ** 2, 0) / numbers.length) * (num(node, 'stdDev') ?? 0)
        const above = yes(node, 'aboveAverage', true)
        const equal = yes(node, 'equalAverage')
        each((at) => {
          const v = values[at]
          if (v === null) return false
          return above ? (equal ? v >= mean + spread : v > mean + spread) : equal ? v <= mean - spread : v < mean - spread
        })
      } else if (type === 'containsText' || type === 'notContainsText' || type === 'beginsWith' || type === 'endsWith') {
        const needle = (attr(node, 'text') ?? '').toLowerCase()
        each((_, r, c) => {
          const text = (shownAt(r, c) ?? '').toLowerCase()
          return type === 'containsText' ? text.includes(needle) : type === 'notContainsText' ? !text.includes(needle) : type === 'beginsWith' ? text.startsWith(needle) : text.endsWith(needle)
        })
      } else if (type === 'containsBlanks' || type === 'notContainsBlanks') {
        each((_, r, c) => {
          const value = valueAt(r, c)
          const blank = value === null || (typeof value === 'string' && value.trim() === '')
          return type === 'containsBlanks' ? blank : !blank
        })
      } else if (type === 'containsErrors' || type === 'notContainsErrors') {
        each((_, r, c) => (valueAt(r, c) instanceof FError) === (type === 'containsErrors'))
      } else if (type === 'duplicateValues' || type === 'uniqueValues') {
        const keyOf = (r: number, c: number) => {
          const value = valueAt(r, c)
          if (value === null || value instanceof FError) return null
          return typeof value === 'string' ? `s${value.toLowerCase()}` : `${typeof value}${value}`
        }
        const keys = cells.map(([r, c]) => keyOf(r, c))
        const counts = new Map<string, number>()
        for (const key of keys) if (key !== null) counts.set(key, (counts.get(key) ?? 0) + 1)
        each((at) => keys[at] !== null && (counts.get(keys[at]!)! > 1) === (type === 'duplicateValues'))
      } else if (type === 'colorScale' || type === 'dataBar' || type === 'iconSet') {
        if (!numbers.length) continue
        const scale = child(node, type)
        const sorted = [...numbers].sort((x, y) => x - y)
        const threshold = (cfvo: XmlNode | undefined, fallback: number) => {
          const kind = attr(cfvo, 'type')
          const val = Number(attr(cfvo, 'val') ?? child(cfvo, 'f')?.text)
          if (kind === 'min' || kind === 'autoMin') return sorted[0]
          if (kind === 'max' || kind === 'autoMax') return sorted[sorted.length - 1]
          if ((kind === 'num' || kind === 'formula') && Number.isFinite(val)) return val
          if (kind === 'percent' && Number.isFinite(val)) return sorted[0] + ((sorted[sorted.length - 1] - sorted[0]) * val) / 100
          if (kind === 'percentile' && Number.isFinite(val)) return sorted[Math.min(sorted.length - 1, Math.max(0, Math.round(((sorted.length - 1) * val) / 100)))]
          return fallback
        }
        const cfvos = children(scale, 'cfvo')
        if (type === 'iconSet') {
          const setName = attr(scale, 'iconSet') ?? '3TrafficLights1'
          const custom = children(scale, 'cfIcon')
          const icons = cfvos.map((_, at) => {
            const own = custom[at]
            const set = own ? attr(own, 'iconSet') ?? setName : setName
            if (set === 'NoIcons') return null
            return ICON_SETS[set]?.[own ? (num(own, 'iconId') ?? at) : at] ?? null
          })
          if (!ICON_SETS[setName] && !custom.length) {
            skipped.add('圖示集')
            continue
          }
          if (yes(scale, 'reverse')) icons.reverse()
          const limits = cfvos.map((cfvo, at) => ({ at: threshold(cfvo, at ? sorted[sorted.length - 1] : sorted[0]), gte: yes(cfvo, 'gte', true) }))
          const hideValue = attr(scale, 'showValue') === '0' || attr(scale, 'showValue') === 'false'
          cells.forEach(([r, c], at) => {
            const v = values[at]
            if (v === null) return
            let index = 0
            limits.forEach((limit, i) => {
              if (i && (limit.gte ? v >= limit.at : v > limit.at)) index = i
            })
            const icon = icons[index]
            if (icon || hideValue) apply(r, c, { ...(icon ? { icon: { glyph: icon[0], color: icon[1] } } : {}), ...(hideValue ? { hideValue } : {}) })
          })
          continue
        }
        const colors = children(scale, 'color').map((color) => xColor(color, wb.styles.palette))
        if (type === 'colorScale') {
          const stops = cfvos.map((cfvo, at) => ({ at: threshold(cfvo, at ? sorted[sorted.length - 1] : sorted[0]), color: colors[at] })).filter((stop): stop is { at: number; color: string } => stop.color !== null)
          if (stops.length < 2) continue
          cells.forEach(([r, c], at) => {
            const v = values[at]
            if (v === null) return
            let i = 0
            while (i < stops.length - 2 && v > stops[i + 1].at) i++
            const [lo, hi] = [stops[i], stops[i + 1]]
            const t = hi.at === lo.at ? 1 : clamp01((v - lo.at) / (hi.at - lo.at))
            apply(r, c, { fill: `#${mix(hi.color, lo.color, t)}` })
          })
        } else {
          const color = colors[0] ?? '638EC6'
          // An automatic minimum starts the bars at zero, so the smallest value still shows a bar.
          const lo = attr(cfvos[0], 'type') === 'min' ? Math.min(0, sorted[0]) : threshold(cfvos[0], sorted[0])
          const hi = threshold(cfvos[1], sorted[sorted.length - 1])
          cells.forEach(([r, c], at) => {
            const v = values[at]
            if (v === null) return
            const pct = Math.round(clamp01(hi === lo ? 1 : (v - lo) / (hi - lo)) * 1000) / 10
            apply(r, c, { bar: `linear-gradient(90deg,#${color} 0,#${applyTint(color, 0.6)} ${pct}%,transparent ${pct}%)` })
          })
        }
      } else if (type === 'timePeriod') skipped.add('日期期間規則')
      else if (cells.length) skipped.add('其他規則')
    } catch {
      skipped.add('無法讀取的規則')
    }
  }
  return { looks, skipped }
}

// ---------------------------------------------------------------------------------------------
// Drawings: pictures, charts, shapes and text boxes, placed over the grid

/** Where grid lines fall, in px from the canvas corner; `frozen` is the far edge of the frozen panes. */
type Geometry = { colLeft: (col0: number) => number; rowTop: (row0: number) => number; frozen?: { x: number; y: number } }
type Box = { x: number; y: number; w: number; h: number }

const px = (value: number) => `${round(value, 1)}px`

function anchorBox(anchor: XmlNode, geo: Geometry): Box | null {
  const kind = local(anchor.name)
  const at = (node: XmlNode | null) => {
    if (!node) return null
    const col = Number(child(node, 'col')?.text ?? NaN)
    const row = Number(child(node, 'row')?.text ?? NaN)
    if (!Number.isFinite(col) || !Number.isFinite(row)) return null
    return {
      x: geo.colLeft(Math.max(0, col)) + Number(child(node, 'colOff')?.text ?? 0) / EMU_PER_PX,
      y: geo.rowTop(Math.max(0, row)) + Number(child(node, 'rowOff')?.text ?? 0) / EMU_PER_PX
    }
  }
  if (kind === 'twoCellAnchor') {
    const from = at(child(anchor, 'from'))
    const to = at(child(anchor, 'to'))
    if (!from || !to) return null
    return { x: from.x, y: from.y, w: Math.max(1, to.x - from.x), h: Math.max(1, to.y - from.y) }
  }
  const ext = child(anchor, 'ext')
  const w = (num(ext, 'cx') ?? 0) / EMU_PER_PX
  const h = (num(ext, 'cy') ?? 0) / EMU_PER_PX
  if (kind === 'oneCellAnchor') {
    const from = at(child(anchor, 'from'))
    return from ? { x: from.x, y: from.y, w: Math.max(1, w), h: Math.max(1, h) } : null
  }
  if (kind === 'absoluteAnchor') {
    const pos = child(anchor, 'pos')
    return { x: (num(pos, 'x') ?? 0) / EMU_PER_PX, y: (num(pos, 'y') ?? 0) / EMU_PER_PX, w: Math.max(1, w), h: Math.max(1, h) }
  }
  return null
}

/** The end of an anchor in grid units (0-based col/row), to grow the drawn range under it. */
function anchorEnd(anchor: XmlNode, colPx: (c: number) => number, rowPx: (r: number) => number): { col: number; row: number } | null {
  const kind = local(anchor.name)
  const cell = (node: XmlNode | null) => {
    const col = Number(child(node, 'col')?.text ?? NaN)
    const row = Number(child(node, 'row')?.text ?? NaN)
    return Number.isFinite(col) && Number.isFinite(row) ? { col: Math.max(0, col), row: Math.max(0, row) } : null
  }
  if (kind === 'twoCellAnchor') return cell(child(anchor, 'to'))
  const ext = child(anchor, 'ext')
  const start = kind === 'oneCellAnchor' ? cell(child(anchor, 'from')) : kind === 'absoluteAnchor' ? { col: 0, row: 0 } : null
  if (!start) return null
  // Walk the grid until the object's size is used up (1-based widths; bounded by the limits).
  const walk = (from: number, size: number, step: (i: number) => number, limit: number) => {
    let at = from
    for (let left = size; left > 0 && at < limit; at++) left -= Math.max(1, step(at + 1))
    return at
  }
  const offsetX = kind === 'absoluteAnchor' ? (num(child(anchor, 'pos'), 'x') ?? 0) / EMU_PER_PX : 0
  const offsetY = kind === 'absoluteAnchor' ? (num(child(anchor, 'pos'), 'y') ?? 0) / EMU_PER_PX : 0
  return {
    col: walk(start.col, (num(ext, 'cx') ?? 0) / EMU_PER_PX + offsetX, colPx, XLSX_LIMITS.cols),
    row: walk(start.row, (num(ext, 'cy') ?? 0) / EMU_PER_PX + offsetY, rowPx, XLSX_LIMITS.rows)
  }
}

type DrawCtx = { wb: Workbook; part: string; text: string[]; count: { n: number }; extent: { right: number; bottom: number } }

const missingBox = (box: Box, what: string, extra = '') =>
  `<div class="xl-draw xl-missing" style="inset-inline-start:${px(box.x)};top:${px(box.y)};width:${px(box.w)};height:${px(box.h)}${extra}">${escapeHtml(what)}</div>`

function renderDrawing(drawingPart: string, wb: Workbook, geo: Geometry, text: string[]): { html: string; right: number; bottom: number } {
  const xml = wb.pkg.xml(drawingPart)
  const ctx: DrawCtx = { wb, part: drawingPart, text, count: { n: 0 }, extent: { right: 0, bottom: 0 } }
  let html = ''
  // Objects placed on the frozen panes stay above them, as Excel draws them; the rest scroll under.
  let front = ''
  for (const anchor of xml?.children ?? []) {
    if (!['twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor'].includes(local(anchor.name))) continue
    if (ctx.count.n >= XLSX_LIMITS.drawings) {
      html += `<p class="xl-note xl-draw-cut">圖形太多,只顯示前 ${XLSX_LIMITS.drawings} 個。</p>`
      break
    }
    const box = anchorBox(anchor, geo)
    if (!box) continue
    ctx.extent.right = Math.max(ctx.extent.right, box.x + box.w)
    ctx.extent.bottom = Math.max(ctx.extent.bottom, box.y + box.h)
    const onFrozen = Boolean(geo.frozen && (box.y < geo.frozen.y || box.x < geo.frozen.x))
    // Each object stands alone: one that cannot be read becomes a labelled box.
    for (const item of anchor.children) {
      const name = local(item.name)
      if (['from', 'to', 'pos', 'ext', 'clientData'].includes(name)) continue
      let drawn: string
      try {
        drawn = renderObject(item, box, ctx)
      } catch (error) {
        drawn = missingBox(box, `圖形無法顯示(${reason(error)})`)
      }
      if (onFrozen) front += drawn
      else html += drawn
    }
  }
  if (front) html += `<div class="xl-front">${front}</div>`
  return { html, ...ctx.extent }
}

function renderObject(node: XmlNode, box: Box, ctx: DrawCtx): string {
  const name = local(node.name)
  ctx.count.n++
  if (name === 'AlternateContent') {
    const choice = child(node, 'Choice')
    const fallback = child(node, 'Fallback')
    const inner = choice?.children[0]
    if (inner && !(local(inner.name) === 'graphicFrame' && /chartex|cx/i.test(attr(path(inner, 'graphic', 'graphicData'), 'uri') ?? ''))) return renderObject(inner, box, ctx)
    if (inner) return missingBox(box, `圖表(新式圖表類型,${UNSUPPORTED})`)
    return fallback?.children[0] ? renderObject(fallback.children[0], box, ctx) : ''
  }
  if (name === 'pic') return renderPicture(node, box, ctx)
  if (name === 'graphicFrame') return renderGraphicFrame(node, box, ctx)
  if (name === 'sp' || name === 'cxnSp') return renderShape(node, box, ctx)
  if (name === 'grpSp') return renderGroup(node, box, ctx)
  return missingBox(box, `物件(${UNSUPPORTED})`)
}

type Xfrm = { x: number; y: number; w: number; h: number; rot: number; chX: number; chY: number; chW: number; chH: number }

function xfrmOf(node: XmlNode | null): Xfrm | null {
  if (!node) return null
  const off = child(node, 'off')
  const ext = child(node, 'ext')
  const chOff = child(node, 'chOff')
  const chExt = child(node, 'chExt')
  return {
    x: num(off, 'x') ?? 0, y: num(off, 'y') ?? 0, w: num(ext, 'cx') ?? 0, h: num(ext, 'cy') ?? 0, rot: (num(node, 'rot') ?? 0) / 60000,
    chX: num(chOff, 'x') ?? 0, chY: num(chOff, 'y') ?? 0, chW: num(chExt, 'cx') ?? 0, chH: num(chExt, 'cy') ?? 0
  }
}

function renderGroup(node: XmlNode, box: Box, ctx: DrawCtx): string {
  const xfrm = xfrmOf(path(node, 'grpSpPr', 'xfrm'))
  let html = ''
  for (const item of node.children) {
    const name = local(item.name)
    if (!['sp', 'pic', 'graphicFrame', 'grpSp', 'cxnSp', 'AlternateContent'].includes(name)) continue
    if (ctx.count.n >= XLSX_LIMITS.drawings) break
    const own = xfrmOf(name === 'grpSp' ? path(item, 'grpSpPr', 'xfrm') : name === 'graphicFrame' ? child(item, 'xfrm') : path(item, 'spPr', 'xfrm'))
    let inner = box
    if (xfrm && own && xfrm.chW > 0 && xfrm.chH > 0) {
      const sx = box.w / xfrm.chW
      const sy = box.h / xfrm.chH
      inner = { x: box.x + (own.x - xfrm.chX) * sx, y: box.y + (own.y - xfrm.chY) * sy, w: Math.max(1, own.w * sx), h: Math.max(1, own.h * sy) }
    }
    try {
      html += renderObject(item, inner, ctx)
    } catch (error) {
      html += missingBox(inner, `圖形無法顯示(${reason(error)})`)
    }
  }
  return html
}

const boxStyle = (box: Box, rot = 0) =>
  `inset-inline-start:${px(box.x)};top:${px(box.y)};width:${px(box.w)};height:${px(box.h)}${rot ? `;transform:rotate(${round(rot, 2)}deg)` : ''}`

function renderPicture(node: XmlNode, box: Box, ctx: DrawCtx): string {
  const blip = path(node, 'blipFill', 'blip')
  const label = attr(path(node, 'nvPicPr', 'cNvPr'), 'descr') ?? attr(path(node, 'nvPicPr', 'cNvPr'), 'name') ?? ''
  const rot = xfrmOf(path(node, 'spPr', 'xfrm'))?.rot ?? 0
  const svg = blip ? findDeep(blip, 'svgBlip') : null
  for (const id of [relAttr(svg, 'embed'), relAttr(blip, 'embed')]) {
    const rel = ctx.wb.pkg.rel(ctx.part, id)
    if (rel && !rel.external && isDrawablePart(rel.target) && ctx.wb.pkg.has(rel.target)) {
      if (label) ctx.text.push(label)
      return `<img class="xl-draw xl-pic" src="${escapeHtml(ctx.wb.media(rel.target))}" alt="${escapeHtml(label)}" loading="lazy" style="${boxStyle(box, rot)}">`
    }
  }
  return missingBox(box, pictureProblem(ctx.wb, ctx.part, relAttr(blip, 'embed'), relAttr(blip, 'link')))
}

/** Why a picture is not drawn, told apart: missing from the file, linked from outside it, or a format the browser cannot show. */
function pictureProblem(wb: Workbook, part: string, embed: string | undefined, link: string | undefined): string {
  const rel = wb.pkg.rel(part, embed)
  if ((!rel && link) || rel?.external) return '圖片(連結到檔案外的圖片,網頁版不會載入;請在 Excel 中開啟原始檔查看)'
  if (!rel || !wb.pkg.has(rel.target)) return '圖片(檔案裡找不到這張圖片,可能已被移除或檔案不完整)'
  const format = /\.([a-z0-9]{2,5})$/i.exec(rel.target)?.[1]?.toUpperCase()
  return `圖片(${format ? `${format} 格式,` : ''}${UNSUPPORTED})`
}

function findDeep(node: XmlNode, name: string): XmlNode | null {
  for (const candidate of node.children) {
    if (local(candidate.name) === name) return candidate
    const found = findDeep(candidate, name)
    if (found) return found
  }
  return null
}

let markerCount = 0

function renderShape(node: XmlNode, box: Box, ctx: DrawCtx): string {
  const spPr = child(node, 'spPr')
  const theme = ctx.wb.theme
  const xfrmNode = child(spPr, 'xfrm')
  const xfrm = xfrmOf(xfrmNode)
  const styleRef = child(node, 'style')
  const connector = local(node.name) === 'cxnSp'
  // A gradient shows as its first stop's colour.
  const fill = child(spPr, 'solidFill') ?? child(child(child(spPr, 'gradFill'), 'gsLst'), 'gs')
  const fillColor = connector ? null : fill ? drawingColor(fill, theme) : child(spPr, 'noFill') ? null : drawingColor(child(styleRef, 'fillRef'), theme)
  const ln = child(spPr, 'ln')
  const lineColor = child(ln, 'noFill') ? null : drawingColor(child(ln, 'solidFill'), theme) ?? drawingColor(child(styleRef, 'lnRef'), theme) ?? (connector ? { hex: '000000', alpha: 1 } : null)
  const lineWidth = Math.max(1, (num(ln, 'w') ?? 9525) / EMU_PER_PX)
  const geomNode = child(spPr, 'prstGeom') ?? child(spPr, 'custGeom')
  const prst = geomNode && local(geomNode.name) === 'prstGeom' ? attr(geomNode, 'prst') ?? 'rect' : null
  // A plain rectangle stays a CSS box; every other shape is its ECMA-376 geometry, evaluated for the box it is drawn in.
  const plain = !connector && (!geomNode || prst === 'rect')
  const wEmu = box.w * EMU_PER_PX
  const hEmu = box.h * EMU_PER_PX
  const shape = plain ? null : geomNode ? shapeGeometry(spPr, wEmu, hEmu) : presetGeometry('line', wEmu, hEmu, {})
  const unknown = !plain && !shape
  const txBody = child(node, 'txBody')
  const fontColor = drawingColor(child(styleRef, 'fontRef'), theme)
  const anchor = attr(child(txBody, 'bodyPr'), 'anchor') ?? 't'
  const textCss: string[] = [`justify-content:${anchor === 'ctr' ? 'center' : anchor === 'b' ? 'flex-end' : 'flex-start'}`]
  if (fontColor) textCss.push(`color:${cssDrawing(fontColor)}`)
  let body = ''
  for (const paragraph of children(txBody, 'p')) {
    const algn = attr(child(paragraph, 'pPr'), 'algn')
    const align = algn === 'ctr' ? 'center' : algn === 'r' ? 'right' : algn === 'just' ? 'justify' : ''
    let line = ''
    for (const run of paragraph.children) {
      const kind = local(run.name)
      if (kind === 'br') line += '<br>'
      if (kind !== 'r' && kind !== 'fld') continue
      const t = child(run, 't')?.text ?? ''
      ctx.text.push(t)
      const rPr = child(run, 'rPr')
      const style: string[] = []
      const sz = num(rPr, 'sz')
      if (sz) style.push(`font-size:${round(sz / 100, 2)}pt`)
      if (attr(rPr, 'b') === '1') style.push('font-weight:700')
      if (attr(rPr, 'i') === '1') style.push('font-style:italic')
      const runColor = drawingColor(child(rPr, 'solidFill'), theme)
      if (runColor) style.push(`color:${cssDrawing(runColor)}`)
      const typeface = attr(child(rPr, 'latin'), 'typeface') ?? attr(child(rPr, 'ea'), 'typeface')
      if (typeface && !typeface.startsWith('+')) style.push(`font-family:${fontStack(typeface)}`)
      line += style.length ? `<span style="${style.join(';')}">${escapeHtml(t)}</span>` : escapeHtml(t)
    }
    body += `<p${align ? ` style="text-align:${align}"` : ''}>${line || '&#8203;'}</p>`
  }
  if (!shape) {
    const css: string[] = [boxStyle(box, xfrm?.rot ?? 0)]
    if (fillColor) css.push(`background:${cssDrawing(fillColor)}`)
    if (lineColor) css.push(`border:${round(lineWidth, 1)}px solid ${cssDrawing(lineColor)}`)
    const label = unknown ? `<p class="xl-shape-note">${escapeHtml(`未支援的圖形:${prst}`)}</p>` : ''
    return `<div class="xl-draw xl-shape" style="${[...css, ...textCss].join(';')}">${body}${label}</div>`
  }
  // Every fill first, then every stroke. Presets fill each path as the standard says; an open path of a custom shape
  // is a line and takes no fill. Arrowheads go on open paths.
  const custom = local(geomNode?.name ?? '') === 'custGeom'
  const wPt = wEmu / EMU_PER_PT
  const hPt = hEmu / EMU_PER_PT
  const paint = fillColor ? cssDrawing(fillColor) : null
  let paths = ''
  if (paint) {
    for (const p of shape.paths) {
      if (p.fill === 'none' || (!p.closed && custom)) continue
      paths += `<path d="${p.d}" fill="${paint}"/>`
      const shade = fillShade(p.fill)
      if (shade) paths += `<path d="${p.d}" fill="${shade.color}" fill-opacity="${shade.opacity}"/>`
    }
  }
  let defs = ''
  if (lineColor) {
    const color = cssDrawing(lineColor)
    let ends = ''
    const end = (name: 'headEnd' | 'tailEnd') => {
      const type = attr(child(ln, name), 'type')
      if (!type || type === 'none') return
      const id = `xl-m${markerCount++}`
      const shapeOf = type === 'oval' ? '<circle cx="5" cy="5" r="4"/>' : type === 'diamond' ? '<path d="M5,0 L10,5 L5,10 L0,5 Z"/>' : '<path d="M0,0 L10,5 L0,10 Z"/>'
      defs += `<marker id="${id}" viewBox="0 0 10 10" refX="${type === 'oval' || type === 'diamond' ? 5 : 9}" refY="5" markerWidth="4" markerHeight="4" orient="${name === 'headEnd' ? 'auto-start-reverse' : 'auto'}" fill="${color}">${shapeOf}</marker>`
      ends += ` marker-${name === 'headEnd' ? 'start' : 'end'}="url(#${id})"`
    }
    if (shape.paths.some((p) => p.stroke && !p.closed)) {
      end('headEnd')
      end('tailEnd')
    }
    for (const p of shape.paths) {
      if (p.stroke) paths += `<path d="${p.d}" fill="none" stroke="${color}" stroke-width="${round(lineWidth, 1)}" vector-effect="non-scaling-stroke" stroke-linejoin="round"${p.closed ? '' : ends}/>`
    }
  }
  if (!paths && !body) return ''
  // A line with no height (or width) gets a box one pixel thick, centred on it.
  const flatW = box.w < 0.5
  const flatH = box.h < 0.5
  const viewBox = [flatW ? -0.375 : 0, flatH ? -0.375 : 0, flatW ? 0.75 : wPt, flatH ? 0.75 : hPt].map((v) => round(v, 3)).join(' ')
  const flip = [attr(xfrmNode, 'flipH') === '1' ? 'scaleX(-1)' : '', attr(xfrmNode, 'flipV') === '1' ? 'scaleY(-1)' : ''].filter(Boolean).join(' ')
  const svgStyle = [flatW ? 'left:-0.5px;width:1px' : '', flatH ? 'top:-0.5px;height:1px' : '', flip ? `transform:${flip}` : ''].filter(Boolean).join(';')
  const svg = paths ? `<svg class="xl-geom" viewBox="${viewBox}" preserveAspectRatio="none"${svgStyle ? ` style="${svgStyle}"` : ''} aria-hidden="true">${defs ? `<defs>${defs}</defs>` : ''}${paths}</svg>` : ''
  // Text sits in the geometry's text rectangle (a flip moves the rectangle, not the letters).
  let text = ''
  if (body) {
    const { l, t, r, b } = shape.textRect
    const pc = (value: number, of: number) => `${round(of > 0 ? (value / of) * 100 : 0, 3)}%`
    const left = flip.includes('scaleX') ? wPt - r : l
    const top = flip.includes('scaleY') ? hPt - b : t
    const rect = r > l && b > t ? `left:${pc(left, wPt)};top:${pc(top, hPt)};width:${pc(r - l, wPt)};height:${pc(b - t, hPt)}` : 'inset:0'
    text = `<div class="xl-shape xl-shape-tx" style="${[rect, ...textCss].join(';')}">${body}</div>`
  }
  return `<div class="xl-draw xl-geo" style="${boxStyle(box, xfrm?.rot ?? 0)}">${svg}${text}</div>`
}

function renderGraphicFrame(node: XmlNode, box: Box, ctx: DrawCtx): string {
  const data = path(node, 'graphic', 'graphicData')
  const uri = attr(data, 'uri') ?? ''
  if (uri.endsWith('/chart')) {
    const chartRef = child(data, 'chart')
    const rel = ctx.wb.pkg.rel(ctx.part, relAttr(chartRef, 'id'))
    const chart = rel && !rel.external ? ctx.wb.pkg.xml(rel.target) : null
    if (!chart) return missingBox(box, `圖表(找不到圖表內容)`)
    const drawn = drawChart(chart, box.w, box.h, ctx.wb)
    ctx.text.push(drawn?.text ?? chartTitle(chart))
    if (!drawn) return missingBox(box, `圖表${chartTitle(chart) ? `「${chartTitle(chart)}」` : ''}(此圖表類型${UNSUPPORTED})`)
    return `<div class="xl-draw xl-chart" style="${boxStyle(box)}">${drawn.svg}</div>`
  }
  return missingBox(box, `物件(${UNSUPPORTED})`)
}

/** A chart part drawn by the shared chart renderer at W × H (any unit: only the proportions and text scale follow). */
function drawChart(chart: XmlNode, W: number, H: number, wb: Workbook): ChartResult | null {
  const theme = wb.theme
  return renderChart(chart, {
    widthPt: W,
    heightPt: H,
    resolveColor: (holder) => {
      const color = drawingColor(holder, theme)
      return color ? cssDrawing(color) : null
    },
    themeColor: (i) => `#${theme.scheme[`accent${i + 1}`] ?? '4472C4'}`,
    fontFamily: fontStack(theme.minor, theme.minorEa),
    textColor: `#${theme.scheme.dk1 === '000000' ? '595959' : theme.scheme.dk1}`,
    // A tool that writes no cache (openpyxl) still names the cells: read them from this workbook.
    readCells: (formula) => lookupRange(wb, formula),
    date1904: wb.date1904
  })
}

// ---------------------------------------------------------------------------------------------
// Drawing one sheet

type SheetResult = { html: string; text: string[]; css: Set<number>; tabColor?: string | null }

const ROW_HEADER_MIN = 32
const HEADER_HEIGHT = 20

function sideCss(side: NonNullable<Side>): string {
  return `${BORDER_CSS[side.style]} #${side.color}`
}

const stronger = (a: Side, b: Side): Side => (!a ? b : !b ? a : b.weight > a.weight ? b : a)

/**
 * `cells`: this sheet's share of the workbook's cells (see shareCells); `spare`: what the shares left
 * unallocated, for a sheet that needs more than its data suggested (merges or a drawing far below it);
 * `chars`: text left for the whole page.
 */
type Budget = { cells: number; spare: number; chars: number }

/**
 * The workbook's cell budget shared fairly between its worksheets: sheets that need less than an even
 * share keep all they need, and what they leave is shared among the rest, so one large sheet early in
 * the workbook cannot starve the sheets after it.
 */
function shareCells(wb: Workbook, entries: SheetEntry[]): { shares: Map<SheetEntry, number>; spare: number } {
  const wanted = entries.map((entry) => {
    let cells = 0
    try {
      if (entry.part && isWorksheet(entry)) {
        const scanned = scannedSheet(wb, entry.part)
        cells = Math.min(XLSX_LIMITS.sheetCells, Math.min(XLSX_LIMITS.rows, Math.max(scanned.maxRow + 1, 30)) * Math.min(XLSX_LIMITS.cols, Math.max(scanned.maxCol + 1, 10)))
      }
    } catch {}
    return { entry, cells }
  })
  const shares = new Map<SheetEntry, number>()
  let left = XLSX_LIMITS.workbookCells
  const sheets = wanted.filter((item) => item.cells > 0).sort((a, b) => a.cells - b.cells)
  sheets.forEach((item, at) => {
    const share = Math.min(item.cells, Math.floor(left / (sheets.length - at)))
    shares.set(item.entry, share)
    left -= share
  })
  return { shares, spare: left }
}

const isWorksheet = (entry: SheetEntry) => entry.kind === 'worksheet' || (entry.kind === '' && Boolean(entry.part?.includes('worksheets/')))

function renderWorksheet(wb: Workbook, entry: SheetEntry, budget: Budget, wantHtml: boolean): SheetResult {
  const part = entry.part!
  const scanned = scannedSheet(wb, part)
  const { head, rows } = scanned
  const text: string[] = []
  const styles = wb.styles
  const base = styles.fonts[0] ?? DEFAULT_FONT
  const mdw = Math.max(5, Math.round(base.sz * (4 / 3) * 0.5))
  const widthPx = (chars: number) => Math.max(0, Math.trunc(((256 * chars + Math.trunc(128 / mdw)) / 256) * mdw))
  const format = child(head, 'sheetFormatPr')
  const defaultColPx = num(format, 'defaultColWidth') !== undefined
    ? widthPx(num(format, 'defaultColWidth')!)
    : Math.ceil(Math.trunc((num(format, 'baseColWidth') ?? 8) * mdw + 5) / 8) * 8
  const defaultRowPx = Math.max(4, Math.round((num(format, 'defaultRowHeight') ?? Math.max(15, base.sz * 1.36)) * (4 / 3)))
  const zeroHeight = ['1', 'true'].includes(attr(format, 'zeroHeight') ?? '')

  // Columns: widths, hidden flags and column styles.
  const colInfo = new Map<number, { px: number; hidden: boolean; style?: number }>()
  for (const col of children(child(head, 'cols'), 'col')) {
    const min = Math.max(1, num(col, 'min') ?? 1)
    const max = Math.min(XLSX_LIMITS.cols, num(col, 'max') ?? min)
    const width = num(col, 'width')
    const hidden = ['1', 'true'].includes(attr(col, 'hidden') ?? '') || width === 0
    for (let c = min; c <= max; c++) colInfo.set(c, { px: width !== undefined ? widthPx(width) : defaultColPx, hidden, style: num(col, 'style') })
  }
  const colPx = (c: number) => {
    const info = colInfo.get(c)
    return info?.hidden ? 0 : (info?.px ?? defaultColPx)
  }
  const rowPx = (r: number) => {
    const info = rows.get(r)
    // zeroHeight hides the rows the file does not list; listed rows show unless hidden themselves.
    if (info?.hidden || (zeroHeight && !info)) return 0
    return info?.ht !== undefined ? Math.round(info.ht * (4 / 3)) : defaultRowPx
  }

  const merges: Range[] = children(child(head, 'mergeCells'), 'mergeCell')
    .map((node) => parseRange(attr(node, 'ref')))
    .filter((range): range is Range => range !== null)
    .slice(0, 20000)

  // The drawing (pictures and charts) may reach past the last cell; the grid continues under it.
  const drawingRel = wb.pkg.rel(part, relAttr(child(head, 'drawing'), 'id'))
  let drawingXml: XmlNode | null = null
  let drawingError = ''
  try {
    drawingXml = drawingRel && !drawingRel.external ? wb.pkg.xml(drawingRel.target) : null
  } catch (error) {
    drawingError = reason(error)
  }
  let usedRow = scanned.maxRow
  let usedCol = scanned.maxCol
  for (const range of merges) {
    usedRow = Math.max(usedRow, Math.min(range.r2, XLSX_LIMITS.rows))
    usedCol = Math.max(usedCol, Math.min(range.c2, XLSX_LIMITS.cols))
  }
  for (const anchor of drawingXml?.children ?? []) {
    const end = anchorEnd(anchor, colPx, rowPx)
    if (end) {
      usedRow = Math.max(usedRow, Math.min(end.row + 2, XLSX_LIMITS.rows))
      usedCol = Math.max(usedCol, Math.min(end.col + 2, XLSX_LIMITS.cols))
    }
  }
  const empty = usedRow === 0 && usedCol === 0 && !drawingXml
  // A little empty grid past the last cell, as Excel shows it; never past the limits.
  const lastCol = Math.min(XLSX_LIMITS.cols, Math.max(usedCol + 1, 10))
  let lastRow = Math.min(XLSX_LIMITS.rows, Math.max(usedRow + 1, 30))
  const floor = Math.min(lastRow, XLSX_LIMITS.floorRows) * lastCol
  const cellBudget = Math.max(floor, Math.min(XLSX_LIMITS.sheetCells, budget.cells + budget.spare))
  if (lastRow * lastCol > cellBudget) lastRow = Math.max(1, Math.floor(cellBudget / lastCol))
  budget.spare = Math.max(0, budget.spare - Math.max(0, lastRow * lastCol - budget.cells))
  const totalRows = Math.max(scanned.totalRows, usedRow)
  const totalCols = Math.max(scanned.totalCols, usedCol)
  const cutRows = totalRows > lastRow
  const cutCols = totalCols > lastCol

  const visibleCols: number[] = []
  for (let c = 1; c <= lastCol; c++) if (colPx(c) > 0) visibleCols.push(c)
  const visibleRows: number[] = []
  for (let r = 1; r <= lastRow; r++) if (rowPx(r) > 0) visibleRows.push(r)

  const view = child(child(head, 'sheetViews'), 'sheetView')
  const showGrid = attr(view, 'showGridLines') !== '0' && attr(view, 'showGridLines') !== 'false'
  const rtl = attr(view, 'rightToLeft') === '1' || attr(view, 'rightToLeft') === 'true'
  const showHeaders = attr(view, 'showRowColHeaders') !== '0'
  const pane = child(view, 'pane')
  const frozen = pane && ['frozen', 'frozenSplit'].includes(attr(pane, 'state') ?? '')
  const freezeCols = frozen ? Math.min(Math.max(0, Math.floor(num(pane, 'xSplit') ?? 0)), lastCol) : 0
  const freezeRows = frozen ? Math.min(Math.max(0, Math.floor(num(pane, 'ySplit') ?? 0)), lastRow) : 0

  const rowHeaderPx = showHeaders ? Math.max(ROW_HEADER_MIN, String(lastRow).length * 8 + 12) : 0
  const headerPx = showHeaders ? HEADER_HEIGHT : 0
  const colLeftCache: number[] = [0]
  const geo: Geometry = {
    colLeft(col0) {
      const target = Math.min(col0, XLSX_LIMITS.cols + 50)
      for (let c = colLeftCache.length; c <= target; c++) colLeftCache[c] = colLeftCache[c - 1] + colPx(c)
      return rowHeaderPx + colLeftCache[target]
    },
    rowTop: (() => {
      const cache: number[] = [0]
      return (row0: number) => {
        const target = Math.min(row0, XLSX_LIMITS.rows + 50)
        for (let r = cache.length; r <= target; r++) cache[r] = cache[r - 1] + rowPx(r)
        return headerPx + cache[target]
      }
    })()
  }

  // Which xf applies where: the cell's own, else the row's (customFormat), else the column's.
  const styleAt = (r: number, c: number): number => rows.get(r)?.cells.get(c)?.s ?? rows.get(r)?.s ?? colInfo.get(c)?.style ?? 0
  const looks = tableLooks(wb, part, head)
  const { looks: conditional, skipped: cfSkipped } = conditionalLooks(wb, head, rows, XLSX_LIMITS.rows, XLSX_LIMITS.cols)
  const borderAt = (r: number, c: number): Border => {
    const own = styles.borders[styles.xfs[styleAt(r, c)]?.border ?? 0] ?? { left: null, right: null, top: null, bottom: null }
    const look = looks.get(r * 20000 + c)
    if (!look) return own
    return { left: own.left, right: own.right, top: own.top ?? look.top ?? null, bottom: own.bottom ?? look.bottom ?? null }
  }
  const fillAt = (r: number, c: number): string => styles.fills[styles.xfs[styleAt(r, c)]?.fill ?? 0] || looks.get(r * 20000 + c)?.fill || ''
  const hasValue = (r: number, c: number): boolean => {
    const cell = rows.get(r)?.cells.get(c)
    return Boolean(cell && (cell.v !== null || cell.is || (cell.f !== null && cell.f !== '')))
  }

  // Merged areas: the anchor spans the visible rows/columns it covers; the rest are skipped.
  const covered = new Set<number>()
  const spans = new Map<number, { rows: number; cols: number; r2: number; c2: number }>()
  const visibleColSet = new Set(visibleCols)
  const visibleRowSet = new Set(visibleRows)
  for (const range of merges) {
    if (range.r1 > lastRow || range.c1 > lastCol) continue
    const r2 = Math.min(range.r2, lastRow)
    const c2 = Math.min(range.c2, lastCol)
    let anchor: number | null = null
    let rowsSpan = 0
    let colsSpan = 0
    for (let c = range.c1; c <= c2; c++) if (visibleColSet.has(c)) colsSpan++
    for (let r = range.r1; r <= r2; r++) if (visibleRowSet.has(r)) rowsSpan++
    if (!rowsSpan || !colsSpan) continue
    for (let r = range.r1; r <= r2; r++) {
      for (let c = range.c1; c <= c2; c++) {
        if (!visibleRowSet.has(r) || !visibleColSet.has(c)) continue
        const key = r * 20000 + c
        if (anchor === null) {
          anchor = key
          spans.set(key, { rows: rowsSpan, cols: colsSpan, r2, c2 })
        } else covered.add(key)
      }
    }
  }

  // Hyperlinks and comments.
  const links = new Map<number, { title: string; href: string }>()
  const sheetIndex = (name: string) => wb.sheets.filter(isShown).findIndex((s) => s.name === name)
  for (const link of children(child(head, 'hyperlinks'), 'hyperlink').slice(0, 5000)) {
    const range = parseRange(attr(link, 'ref'))
    if (!range) continue
    const rel = wb.pkg.rel(part, relAttr(link, 'id'))
    const location = attr(link, 'location') ?? ''
    let title = attr(link, 'tooltip') ?? ''
    let href = ''
    if (rel?.external) title = title || rel.target
    else if (location) {
      const target = /^'?(.*?)'?!/.exec(location)?.[1]?.replace(/''/g, "'")
      const at = target !== undefined ? sheetIndex(target) : -1
      if (at >= 0) href = `#sheet-${at + 1}`
      title = title || location
    }
    for (let r = range.r1; r <= Math.min(range.r2, lastRow); r++) for (let c = range.c1; c <= Math.min(range.c2, lastCol); c++) links.set(r * 20000 + c, { title: title.slice(0, 500), href })
  }
  const comments = new Map<number, string>()
  try {
    for (const rel of wb.pkg.rels(part).values()) {
      if (!rel.type.endsWith('/comments') || rel.external) continue
      const xml = wb.pkg.xml(rel.target)
      const authors = children(child(xml, 'authors'), 'author').map((author) => author.text)
      for (const comment of children(child(xml, 'commentList'), 'comment').slice(0, 5000)) {
        const ref = parseRef(attr(comment, 'ref'))
        if (!ref) continue
        const author = authors[num(comment, 'authorId') ?? -1]
        const body = richText(child(comment, 'text')).text
        const note = `${author ? `${author}:\n` : ''}${body}`.slice(0, 1000)
        comments.set(ref.r * 20000 + ref.c, note)
        text.push(body)
      }
    }
  } catch {}

  /** A picture placed in the cell, fitted inside it keeping its shape; or a label saying why not. */
  const inCellPicture = (cell: Cell, r: number, c: number): Shown | null => {
    if (cell.vm === null) return null
    const picture = cellPicture(wb, cell)
    const label = (what: string): Shown => ({ text: '', html: wantHtml ? `<span class="xl-cellmiss">${escapeHtml(what)}</span>` : '', color: null, kind: 'picture' })
    if (!picture) return cell.t === 'e' ? label(`儲存格內的圖片或資料類型(${UNSUPPORTED})`) : null
    if ('web' in picture) return { ...label('儲存格內的網路圖片(網頁版不會載入)'), text: picture.alt }
    if (!wb.pkg.has(picture.target)) return { ...label('儲存格內的圖片(檔案裡找不到這張圖片)'), text: picture.alt }
    if (!isDrawablePart(picture.target)) {
      const format = /\.([a-z0-9]{2,5})$/i.exec(picture.target)?.[1]?.toUpperCase()
      return { ...label(`儲存格內的圖片(${format ? `${format} 格式,` : ''}${UNSUPPORTED})`), text: picture.alt }
    }
    const span = spans.get(r * 20000 + c)
    let w = 0
    let h = 0
    for (const col of visibleCols) if (col >= c && col <= (span?.c2 ?? c)) w += colPx(col)
    for (const row of visibleRows) if (row >= r && row <= (span?.r2 ?? r)) h += rowPx(row)
    const html = wantHtml
      ? `<img class="xl-cellpic" src="${escapeHtml(wb.media(picture.target))}" alt="${escapeHtml(picture.alt)}" loading="lazy" style="width:${px(Math.max(1, w - 4))};height:${px(Math.max(1, h - 2))}">`
      : ''
    return { text: picture.alt, html, color: null, kind: 'picture' }
  }

  const usedXfs = new Set<number>()
  const parts: string[] = []
  let textCut = false
  const cssOf = (r: number, c: number) => styles.xfs[styleAt(r, c)]
  const colLeftOf = new Map<number, number>()
  let left = rowHeaderPx
  for (const c of visibleCols) {
    colLeftOf.set(c, left)
    left += colPx(c)
  }
  const tableWidth = left
  const rowTopOf = new Map<number, number>()
  let topAt = headerPx
  for (const r of visibleRows) {
    rowTopOf.set(r, topAt)
    topAt += rowPx(r)
  }
  const freezeWidth = visibleCols.filter((c) => c <= freezeCols).reduce((sum, c) => sum + colPx(c), 0)
  const freezeHeight = visibleRows.filter((r) => r <= freezeRows).reduce((sum, r) => sum + rowPx(r), 0)
  if (freezeCols || freezeRows) geo.frozen = { x: freezeCols ? rowHeaderPx + freezeWidth : 0, y: freezeRows ? headerPx + freezeHeight : 0 }

  if (wantHtml) {
    parts.push(`<table class="xl-grid${showGrid ? '' : ' nogrid'}${showHeaders ? '' : ' noheads'}" style="width:${tableWidth}px">`)
    parts.push(`<colgroup>${showHeaders ? `<col style="width:${rowHeaderPx}px">` : ''}${visibleCols.map((c) => `<col style="width:${colPx(c)}px">`).join('')}</colgroup>`)
    if (showHeaders) {
      parts.push(`<thead><tr style="height:${HEADER_HEIGHT}px"><th class="xl-corner"></th>${visibleCols.map((c) => `<th${c <= freezeCols ? ` class="fc" style="inset-inline-start:${colLeftOf.get(c)}px"` : ''}>${colName(c)}</th>`).join('')}</tr></thead>`)
    }
    parts.push('<tbody>')
  }
  for (const r of visibleRows) {
    const row = rows.get(r)
    if (wantHtml) {
      const frozenRow = r <= freezeRows
      parts.push(`<tr style="height:${rowPx(r)}px">${showHeaders ? `<th${frozenRow ? ` class="fr" style="top:${rowTopOf.get(r)}px"` : ''}>${r}</th>` : ''}`)
    }
    for (let ci = 0; ci < visibleCols.length; ci++) {
      const c = visibleCols[ci]
      const key = r * 20000 + c
      if (covered.has(key)) continue
      const cell = row?.cells.get(c)
      const xfIndex = styleAt(r, c)
      const shown = cell ? (inCellPicture(cell, r, c) ?? showCell(cell, xfIndex, wb, wantHtml)) : null
      if (shown && shown.text && shown.kind !== 'formula') text.push(shown.text)
      if (!wantHtml) continue
      // Past the page's text budget, cells keep their look but not their text.
      if (shown?.html) {
        if (budget.chars <= 0) {
          shown.html = ''
          textCut = true
        } else budget.chars -= shown.html.length
      }
      const span = spans.get(key)
      const xf = cssOf(r, c)
      const classes: string[] = []
      const style: string[] = []
      if (xfIndex && styles.xfs[xfIndex]) {
        classes.push(`x${xfIndex}`)
        usedXfs.add(xfIndex)
      }
      const align = xf?.align
      const h = align?.h ?? 'general'
      if (shown && h === 'general') {
        if (shown.kind === 'num') classes.push('n')
        else if (shown.kind === 'bool' || shown.kind === 'err') classes.push('c')
      }
      if (shown?.kind === 'formula') classes.push('f')
      if (shown?.stale) classes.push('stale')
      const look = looks.get(key)
      if (look) {
        if (look.fill && !(styles.fills[xf?.fill ?? 0])) style.push(`background:${look.fill}`)
        const font = styles.fonts[xf?.font ?? 0]
        if (look.color && (font?.color ?? null) === (base.color ?? null)) style.push(`color:#${look.color}`)
        if (look.bold && !font?.b) style.push('font-weight:700')
      }
      // Conditional formats sit over the cell's own format, as in Excel.
      const cf = conditional.get(key)
      if (cf) {
        if (cf.fill) style.push(`background:${cf.fill}`)
        if (cf.bar) style.push(`background-image:${cf.bar}`)
        if (cf.color) style.push(`color:#${cf.color}`)
        if (cf.bold) style.push('font-weight:700')
        if (cf.italic) style.push('font-style:italic')
      }
      // Edges: this cell draws its end and bottom edges, settled against the neighbour's.
      const endC = span ? span.c2 : c
      const endR = span ? span.r2 : r
      const nextC = visibleCols[visibleCols.indexOf(endC) + 1]
      const nextR = visibleRows[visibleRows.indexOf(endR) + 1]
      const own = borderAt(r, c)
      const endSide = stronger(borderAt(r, endC).right, nextC !== undefined ? borderAt(r, nextC).left : null)
      const bottomSide = stronger(borderAt(endR, c).bottom, nextR !== undefined ? borderAt(nextR, c).top : null)
      const filled = Boolean(fillAt(r, c))
      // A fill hides the gridlines inside it; the xf's class already does that for its own fill.
      const tableFilled = filled && !styles.fills[styles.xfs[xfIndex]?.fill ?? 0]
      if (endSide) style.push(`border-inline-end:${sideCss(endSide)}`)
      else if (tableFilled) style.push('border-inline-end-color:transparent')
      if (bottomSide) style.push(`border-bottom:${sideCss(bottomSide)}`)
      else if (tableFilled) style.push('border-bottom-color:transparent')
      if (ci === 0 && own.left) style.push(`border-inline-start:${sideCss(own.left)}`)
      if (r === visibleRows[0] && own.top) style.push(`border-top:${sideCss(own.top)}`)
      // Text runs on into empty neighbours, as in Excel, unless it wraps, shrinks or turns, and stops
      // at the first cell that holds something (a value or a merged area): left-aligned text to the
      // end side, right-aligned text to the start side, centred text both ways around its own cell,
      // and centred-across-selection text over the empty cells after it that are set the same way.
      let runBack = 0
      let runOn = 0
      if (shown && shown.kind === 'text' && !align?.wrap && !span && !align?.shrink && !(align?.rot)) {
        const free = (k: number) => {
          const col = visibleCols[k]
          const colKey = r * 20000 + col
          return !hasValue(r, col) && !covered.has(colKey) && !spans.has(colKey)
        }
        const reach = (step: 1 | -1, more: (k: number) => boolean = free) => {
          let room = 0
          for (let k = ci + step, n = 0; k >= 0 && k < visibleCols.length && n < 64 && room < 4000 && more(k); k += step, n++) room += colPx(visibleCols[k])
          return room
        }
        if (h === 'general' || h === 'left') runOn = reach(1)
        else if (h === 'right') runBack = reach(-1)
        else if (h === 'center') {
          runBack = reach(-1)
          runOn = reach(1)
        } else if (h === 'centerContinuous') runOn = reach(1, (k) => free(k) && styles.xfs[styleAt(r, visibleCols[k])]?.align?.h === 'centerContinuous')
        if (runBack || runOn) classes.push('ov')
      }
      // Frozen panes stay put while the rest scrolls.
      const frozenRow = r <= freezeRows
      const frozenCol = c <= freezeCols
      if (frozenRow || frozenCol) {
        if (frozenRow) {
          classes.push('fr')
          style.push(`top:${rowTopOf.get(r)}px`)
        }
        if (frozenCol) {
          classes.push('fc')
          style.push(`inset-inline-start:${colLeftOf.get(c)}px`)
        }
        if (!filled && !look?.fill && !cf?.fill) style.push('background:#fff')
      }
      const link = links.get(key)
      const note = comments.get(key)
      if (note) classes.push('cm')
      if (shown?.color) style.push(`color:#${shown.color}`)
      let inner = shown ? shown.html : ''
      if (cf?.hideValue) inner = ''
      if (align && inner) {
        // Turned text is laid out of flow from the cell's corner, so it neither grows the row nor
        // swings outside the cell to be clipped away: rising text starts bottom-left, falling text
        // top-left, each moved over by its line height so the first letter stays inside.
        if (align.rot === 255) {
          classes.push('rc')
          inner = `<span class="vt">${inner}</span>`
        } else if (align.rot > 0 && align.rot <= 180) {
          classes.push('rc')
          const rising = align.rot <= 90
          const angle = rising ? align.rot : align.rot - 90
          const shift = round(1.2 * Math.sin((angle * Math.PI) / 180), 3)
          inner = `<span class="rot${rising ? '' : ' dn'}" style="transform:translateX(${shift}em) rotate(${rising ? -angle : angle}deg)">${inner}</span>`
        } else if (align.shrink && !align.wrap) {
          classes.push('sh')
          inner = `<span>${inner}</span>`
        }
      }
      if (link && inner) {
        inner = link.href
          ? `<a class="xl-link" href="${escapeHtml(link.href)}" title="${escapeHtml(link.title)}">${inner}</a>`
          : `<span class="xl-link" title="${escapeHtml(link.title)}">${inner}</span>`
      }
      if ((runOn || runBack) && inner) {
        const own = colPx(c)
        // Text longer than its run is cut where Excel cuts it: right-aligned text loses its start,
        // centred text both ends (the flex alignment overflows that way), left-aligned text its end.
        if (h === 'center') {
          // Centred on its own cell, then cut where an occupied neighbour begins on either side.
          const side = Math.max(runBack, runOn)
          const [physLeft, physRight] = rtl ? [side - runOn, side - runBack] : [side - runBack, side - runOn]
          const clip = physLeft || physRight ? `;clip-path:inset(0 ${px(physRight)} 0 ${px(physLeft)})` : ''
          inner = `<span class="ovs ovc" style="width:${px(Math.max(0, own + 2 * side - 4))};margin-inline-start:${px(-side)}${clip}"><span>${inner}</span></span>`
        } else if (runBack) inner = `<span class="ovs ovr" style="width:${px(Math.max(0, own + runBack - 4))};margin-inline-start:${px(-runBack)}"><span>${inner}</span></span>`
        else inner = `<span class="ovs" style="width:${px(Math.max(0, own + runOn - 4))}">${inner}</span>`
      }
      // An icon-set icon sits at the start of the cell, the value where its alignment puts it.
      if (cf?.icon) inner = `<span class="xl-ico" style="color:#${cf.icon.color}">${escapeHtml(cf.icon.glyph)}</span>${inner}`
      const attrs = [
        span && span.cols > 1 ? ` colspan="${span.cols}"` : '',
        span && span.rows > 1 ? ` rowspan="${span.rows}"` : '',
        classes.length ? ` class="${classes.join(' ')}"` : '',
        style.length ? ` style="${style.join(';')}"` : '',
        note ? ` title="${escapeHtml(note)}"` : shown?.stale ? ` title="${STALE_NOTE}"` : ''
      ].join('')
      parts.push(`<td${attrs}>${inner}</td>`)
    }
    if (wantHtml) parts.push('</tr>')
  }
  const tabColor = xColor(path(head, 'sheetPr', 'tabColor'), styles.palette)
  let drawings = { html: '', right: 0, bottom: 0 }
  if (drawingError) drawings.html = `<p class="xl-note">圖片與圖表無法顯示(${escapeHtml(drawingError)})</p>`
  else if (drawingXml && drawingRel) {
    try {
      drawings = renderDrawing(drawingRel.target, wb, geo, text)
    } catch (error) {
      drawings.html = `<p class="xl-note">圖片與圖表無法顯示(${escapeHtml(reason(error))})</p>`
    }
  }
  if (!wantHtml) return { html: '', text, css: usedXfs, tabColor }
  parts.push('</tbody></table>')
  const notes: string[] = []
  if (empty) notes.push('這個工作表是空的。')
  if (cutRows || cutCols) {
    const shownPart = [cutRows ? `前 ${lastRow.toLocaleString('en-US')} 列` : '', cutCols ? `前 ${lastCol.toLocaleString('en-US')} 欄(到 ${colName(lastCol)} 欄)` : ''].filter(Boolean).join('、')
    const size = `${totalRows.toLocaleString('en-US')} 列 × ${totalCols.toLocaleString('en-US')} 欄`
    notes.push(`此工作表${scanned.exact ? `共 ${size}` : `至少有 ${size}`},網頁版只顯示${shownPart};完整內容請下載原始檔查看。`)
  }
  if (scanned.cut) {
    const lead = cutRows || cutCols ? '這個工作表的檔案太大,網頁版只讀取了前段' : `這個工作表的檔案太大,網頁版只讀取了前段(前 ${scanned.maxRow.toLocaleString('en-US')} 列)`
    notes.push(`${lead};位於檔案後段的合併儲存格、超連結、條件式格式與圖片沒有讀到。`)
  }
  if (textCut) notes.push('這份活頁簿的文字太多,後面部分儲存格的內容沒有顯示;完整內容請下載原始檔查看。')
  if (cfSkipped.size) notes.push(`此工作表有部分條件式格式(${[...cfSkipped].join('、')})網頁版沒有套用,儲存格以原本的格式顯示。`)
  if (hasControls(head)) notes.push(`此工作表含有表單控制項或內嵌物件,${UNSUPPORTED}。`)
  const noteHtml = notes.map((note) => `<p class="xl-note">${escapeHtml(note)}</p>`).join('')
  const wide = freezeWidth > 200 ? ' data-wide-x' : ''
  const tall = freezeHeight > 200 ? ' data-tall-y' : ''
  const width = Math.ceil(Math.max(tableWidth, drawings.right))
  const height = Math.ceil(Math.max(topAt, drawings.bottom))
  const html = `${noteHtml}<div class="xl-scroll"${rtl ? ' dir="rtl"' : ''}${wide}${tall}><div class="xl-canvas" style="width:${width}px;min-height:${height}px">${parts.join('')}${drawings.html}</div></div>`
  return { html, text, css: usedXfs, tabColor }
}

/** Form controls and ActiveX live in VML and control parts this page does not draw. */
const hasControls = (head: XmlNode) => Boolean(child(head, 'controls') || child(head, 'oleObjects'))

function renderChartsheet(wb: Workbook, entry: SheetEntry, wantHtml: boolean): SheetResult {
  const part = entry.part!
  const xml = wb.pkg.xml(part)
  if (!xml) throw new Error(`找不到 ${part}`)
  const rel = wb.pkg.rel(part, relAttr(child(xml, 'drawing'), 'id'))
  const text: string[] = []
  if (!rel || rel.external) return { html: '<p class="xl-note">這張圖表工作表沒有內容。</p>', text, css: new Set() }
  const geo: Geometry = { colLeft: (col) => col * 64, rowTop: (row) => row * 20 }
  const drawing = wb.pkg.xml(rel.target)
  let html = ''
  for (const anchor of drawing?.children ?? []) {
    const frame = anchor.children.find((item) => local(item.name) === 'graphicFrame')
    if (!frame) continue
    const box = anchorBox(anchor, geo) ?? { x: 0, y: 0, w: 800, h: 500 }
    const chartRel = wb.pkg.rel(rel.target, relAttr(path(frame, 'graphic', 'graphicData', 'chart'), 'id'))
    const chart = chartRel && !chartRel.external ? wb.pkg.xml(chartRel.target) : null
    // Drawn at a reading size and scaled to the page, keeping the chart's own proportions.
    const aspect = Math.min(3, Math.max(0.5, box.w / Math.max(1, box.h)))
    const w = 720
    const h = Math.round(w / aspect)
    const drawn = chart ? drawChart(chart, w, h, wb) : null
    text.push(drawn?.text ?? (chart ? chartTitle(chart) : ''))
    html += drawn
      ? `<div class="xl-chartsheet" style="aspect-ratio:${round(w / h, 4)}">${drawn.svg}</div>`
      : `<p class="xl-note">圖表(此圖表類型${UNSUPPORTED})</p>`
  }
  return { html: wantHtml ? `<div class="xl-scroll xl-pad">${html || '<p class="xl-note">這張圖表工作表沒有內容。</p>'}</div>` : '', text, css: new Set() }
}

// ---------------------------------------------------------------------------------------------
// The page

function xfCss(wb: Workbook, index: number): string {
  const xf = wb.styles.xfs[index]
  if (!xf) return ''
  const base = wb.styles.fonts[0] ?? DEFAULT_FONT
  const font = wb.styles.fonts[xf.font] ?? base
  const rules: string[] = []
  const fontRules = fontCss(font, base, wb.theme)
  if (fontRules) rules.push(fontRules)
  const fill = wb.styles.fills[xf.fill]
  // Diagonal lines lie over the fill, as the first background layer.
  const diagonal = diagonalCss(wb.styles.borders[xf.border]?.diagonal)
  if (fill || diagonal) rules.push(`background:${[diagonal, fill].filter(Boolean).join(',')}`)
  if (fill) rules.push('border-inline-end-color:transparent;border-bottom-color:transparent')
  const align = xf.align
  if (align) {
    const h = { left: 'left', center: 'center', right: 'right', justify: 'justify', distributed: 'justify', centerContinuous: 'center', fill: 'left' }[align.h]
    if (h) rules.push(`text-align:${h}`)
    const v = { top: 'top', center: 'middle', justify: 'middle', distributed: 'middle' }[align.v]
    if (v) rules.push(`vertical-align:${v}`)
    if (align.wrap || align.h === 'justify' || align.h === 'distributed' || align.v === 'justify' || align.v === 'distributed') rules.push('white-space:pre-wrap;overflow-wrap:anywhere')
    if (align.indent) {
      const side = align.h === 'right' ? 'padding-inline-end' : 'padding-inline-start'
      rules.push(`${side}:${round(align.indent * 9 + 2, 1)}px`)
    }
  }
  return rules.length ? `.xl-grid td.x${index}{${rules.join(';')}}` : ''
}

const PAGE_CSS = `
:root{color-scheme:light}
html,body{margin:0;height:100%;background:#fff;color:#1f1f1f}
body{display:flex;flex-direction:column;font:14px/1.4 -apple-system,BlinkMacSystemFont,'Segoe UI','PingFang TC','Noto Sans TC','Microsoft JhengHei',sans-serif;overflow:hidden}
.xl-tabs{flex:none;display:flex;gap:2px;overflow-x:auto;background:#f3f3f3;border-bottom:1px solid #c8c8c8;padding:4px 6px 0;scrollbar-width:thin}
.xl-tab{flex:none;max-width:16em;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;padding:6px 14px 5px;color:#333;text-decoration:none;border:1px solid transparent;border-bottom:0;border-radius:6px 6px 0 0;box-shadow:inset 0 -3px 0 var(--tab,transparent)}
.xl-tab:hover{background:#e8e8e8}
.xl-tab[aria-selected=true]{background:#fff;border-color:#c8c8c8;color:#0b5e2b;font-weight:600;margin-bottom:-1px}
.xl-main{flex:1;min-height:0;position:relative}
.xl-sheet{position:absolute;inset:0;display:flex;flex-direction:column}
.xl-sheet[hidden]{display:none}
.xl-note{margin:0;padding:6px 12px;background:#fff8e1;border-bottom:1px solid #f0d58c;color:#5d4400;font-size:13px}
.xl-scroll{flex:1;min-height:0;overflow:auto;background:#fff;overscroll-behavior:contain}
.xl-pad{padding:12px}
.xl-canvas{position:relative}
.xl-grid{border-collapse:separate;border-spacing:0;table-layout:fixed;background:#fff;color:#000}
.xl-grid td{border:0;border-inline-end:1px solid #d4d4d4;border-bottom:1px solid #d4d4d4;padding:0 2px;white-space:pre;overflow:hidden;vertical-align:bottom;text-align:start;line-height:1.2;box-sizing:border-box}
.xl-grid.nogrid td{border-color:transparent}
.xl-grid td.n{text-align:end}
.xl-grid td.c{text-align:center}
.xl-grid td.f{color:#888;font-style:italic}.xl-grid td.stale{color:#888;font-style:italic}
.xl-grid th{position:sticky;background:#f3f3f3;color:#5f5f5f;font:12px/1 -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;font-weight:400;border:0;border-inline-end:1px solid #c8c8c8;border-bottom:1px solid #c8c8c8;padding:0 3px;text-align:center;box-sizing:border-box;overflow:hidden;white-space:nowrap}
.xl-grid thead th{top:0;z-index:4}
.xl-grid tbody th{inset-inline-start:0;z-index:3}
.xl-grid thead th.xl-corner{inset-inline-start:0;z-index:6}
.xl-grid td.fr,.xl-grid td.fc{position:sticky;z-index:2}
.xl-grid td.fr.fc{z-index:3}
.xl-grid th.fr{z-index:4}
.xl-grid td.fr{box-shadow:0 1px 0 #9e9e9e}
.xl-grid td.fc{box-shadow:1px 0 0 #9e9e9e}
.xl-grid td.cm{position:relative}
.xl-grid td.cm::after{content:'';position:absolute;top:0;inset-inline-end:0;border-style:solid;border-width:0 6px 6px 0;border-color:transparent #d32f2f transparent transparent}
.xl-grid td.fr.cm,.xl-grid td.fc.cm{position:sticky}
.xl-grid td.rc{position:relative}
.xl-grid td.fr.rc,.xl-grid td.fc.rc{position:sticky}
.xl-grid td.ov{overflow:visible}
.xl-grid td.ov>.ovs{display:inline-block;overflow:hidden;vertical-align:bottom}
.xl-grid td.ov>.ovr,.xl-grid td.ov>.ovc{display:inline-flex;justify-content:flex-end}
.xl-grid td.ov>.ovc{justify-content:center}
.xl-grid td.ov>.ovr>span,.xl-grid td.ov>.ovc>span{flex:none}
.xl-grid td .vt{position:absolute;top:0;left:50%;transform:translateX(-50%);writing-mode:vertical-rl;text-orientation:upright;white-space:pre}
.xl-grid td .rot{position:absolute;left:2px;bottom:0;transform-origin:left bottom;white-space:pre}
.xl-grid td .rot.dn{top:0;bottom:auto;transform-origin:left top}
.xl-grid td.sh>span{display:inline-block}
.xl-grid td img.xl-cellpic{display:block;object-fit:contain;margin:0 auto}
.xl-grid td .xl-cellmiss{display:inline-block;max-width:100%;overflow:hidden;text-overflow:ellipsis;vertical-align:middle;color:#666;font:italic 11px/1.2 -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif}
.xl-grid td .xl-ico{float:inline-start;margin-inline-end:3px}
.xl-link{text-decoration:underline;cursor:help}
a.xl-link{color:inherit;cursor:pointer}
.xl-draw{position:absolute;z-index:1;box-sizing:border-box}
.xl-draw svg,.xl-chartsheet svg{direction:ltr}
.xl-pic{object-fit:fill}
.xl-front{position:absolute;inset:0;z-index:3;pointer-events:none}
.xl-front>*{pointer-events:auto}
.xl-chart{background:#fff;border:1px solid #d9d9d9}
.xl-shape{display:flex;flex-direction:column;overflow:visible;padding:4px 7px;font-size:11pt;color:#000}
.xl-shape p{margin:0}
.xl-geo{overflow:visible}
.xl-geom{position:absolute;left:0;top:0;width:100%;height:100%;overflow:visible}
.xl-shape-tx{position:absolute;box-sizing:border-box}
.xl-shape-note{font-size:11px;color:#555}
.xl-missing{display:flex;align-items:center;justify-content:center;text-align:center;padding:6px;background:#f6f6f6;border:1px dashed #9e9e9e;color:#555;font-size:12px;overflow:hidden}
.xl-chartsheet{max-width:1100px;margin:0 auto;border:1px solid #d9d9d9;background:#fff}
.xl-chartsheet svg{display:block}
.xl-empty{padding:24px;color:#555}
.xl-fail{box-sizing:border-box;max-width:560px;margin:10vh auto;background:#fff;border:1px solid #dadce0;border-radius:10px;padding:20px 22px;font:15px/1.6 -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#202124}
.xl-fail h1{font-size:18px;margin:0 0 8px}.xl-fail p{margin:0 0 6px}
.xl-grid thead th.fc{z-index:5}
@media (max-width:700px){[data-wide-x] .xl-grid .fc{inset-inline-start:auto!important;box-shadow:none}[data-tall-y] .xl-grid .fr{top:auto!important;box-shadow:none}}
`

/**
 * A number too wide for its column shrinks a little or shows #, judged as Excel judges it: every digit as wide as
 * 0 (Office fonts draw digits one width), so values of one format in one column all fit or all do not.
 * Tabs live in the hash (#sheet-2) so the viewer's place report keeps the reader's sheet; the viewer
 * accepts only letters, digits, '-' and '_' there. Panels use other ids, so no hash scrolls the page.
 */
const PAGE_SCRIPT = `(()=>{const tabs=[...document.querySelectorAll('.xl-tab')];const sheets=[...document.querySelectorAll('.xl-sheet')];if(!sheets.length)return;const first=Number(document.body.dataset.active)||1;
function fit(root){if(root.dataset.fitted)return;root.dataset.fitted='1';const writes=[];
for(const td of root.querySelectorAll('td.sh')){const s=td.firstElementChild;if(!s)continue;const room=td.clientWidth-4,w=s.offsetWidth;if(room>0&&w>room)writes.push([s,(parseFloat(getComputedStyle(s).fontSize)*room/w).toFixed(2)+'px'])}
const ctx=document.createElement('canvas').getContext('2d'),hash={};
let last='';const use=f=>{if(f!==last){ctx.font=f;last=f}};
for(const td of root.querySelectorAll('td.n:not(.sh)')){let sw=td.scrollWidth;const cw=td.clientWidth,t=td.textContent,over=sw>cw+1;if(!over&&!/\\d/.test(t))continue;const cs=getComputedStyle(td),font=cs.fontStyle+' '+cs.fontWeight+' '+cs.fontSize+' '+cs.fontFamily;if(/\\d/.test(t)){use(font);const a=ctx.measureText(t).width,z=ctx.measureText(t.replace(/\\d/g,'0')).width;if(over){if(a>0)sw*=z/a}else sw=z+parseFloat(cs.paddingLeft)+parseFloat(cs.paddingRight)}if(sw<=cw+1)continue;const k=cw/sw;if(k>=0.7){writes.push([td,(parseFloat(cs.fontSize)*k*0.97).toFixed(2)+'px'])}else if(!td.querySelector('*')){if(!(font in hash)){use(font);hash[font]=ctx.measureText('#').width||7}writes.push([td,null,'#'.repeat(Math.max(1,Math.floor((cw-4)/hash[font])))])}}
for(const [el,size,text] of writes){if(size)el.style.fontSize=size;else{el.title=el.textContent;el.textContent=text}}}
function show(n){if(!(n>=1&&n<=sheets.length))n=first;sheets.forEach((s,i)=>{s.hidden=i!==n-1});tabs.forEach((t,i)=>{t.setAttribute('aria-selected',String(i===n-1));t.tabIndex=i===n-1?0:-1});const tab=tabs[n-1];if(tab){const bar=tab.parentElement;const l=tab.offsetLeft-bar.offsetLeft;if(l<bar.scrollLeft||l+tab.offsetWidth>bar.scrollLeft+bar.clientWidth)bar.scrollLeft=l-16}fit(sheets[n-1])}
function fromHash(){const m=/^#sheet-(\\d+)$/.exec(location.hash);show(m?Number(m[1]):first)}
addEventListener('hashchange',fromHash);fromHash()})()`

export type XlsxSheet = { name: string; html: string; text: string[]; tabColor: string | null }

/** Reads a workbook into one HTML block per visible sheet (or just its text, for search). */
export function readXlsx(data: Buffer, mediaBase: string, wantHtml = true): { sheets: XlsxSheet[]; css: string; warnings: string[]; active: number; baseFont: string } {
  const wb = openWorkbook(data, mediaBase)
  const visible = wb.sheets.filter(isShown)
  const { shares, spare } = shareCells(wb, visible)
  const budget: Budget = { cells: 0, spare, chars: XLSX_LIMITS.pageChars }
  const used = new Set<number>()
  const sheets = visible.map((entry): XlsxSheet => {
    try {
      if (!entry.part) throw new Error('找不到工作表內容')
      let result: SheetResult
      if (isWorksheet(entry)) {
        budget.cells = shares.get(entry) ?? 0
        result = renderWorksheet(wb, entry, budget, wantHtml)
      } else if (entry.kind === 'chartsheet') result = renderChartsheet(wb, entry, wantHtml)
      else result = { html: `<p class="xl-note">這種工作表(${escapeHtml(entry.kind || '未知')})${UNSUPPORTED}。</p>`, text: [], css: new Set() }
      for (const index of result.css) used.add(index)
      return { name: entry.name, html: result.html, text: result.text, tabColor: result.tabColor ?? null }
    } catch (error) {
      const message = `工作表「${entry.name}」無法顯示:${reason(error)}`
      return { name: entry.name, html: `<p class="xl-note xl-failed">${escapeHtml(message)}</p>${wantHtml ? failureScript(`Excel: ${message}`) : ''}`, text: [], tabColor: null }
    }
  })
  const active = Math.max(0, visible.findIndex((sheet) => sheet.tab === wb.activeTab)) + 1
  const base = wb.styles.fonts[0] ?? DEFAULT_FONT
  const baseFont = `font-family:${fontStack(base.name, wb.theme.minorEa)};font-size:${round(base.sz, 2)}pt${base.color ? `;color:#${base.color}` : ''}`
  const css = wantHtml ? [...used].sort((a, b) => a - b).map((index) => xfCss(wb, index)).filter(Boolean).join('\n') : ''
  return { sheets, css, warnings: wb.warnings, active, baseFont }
}

/** The whole page for an Excel file. Pictures go through officeMediaBase(fileName) (office.ts). */
export function renderXlsxPage(data: Buffer, fileName: string, title: string): string {
  const head = `<meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>${escapeHtml(title)}</title><style>${PAGE_CSS}</style>`
  let book: ReturnType<typeof readXlsx>
  try {
    book = readXlsx(data, officeMediaBase(fileName))
  } catch (error) {
    const message = `Excel: ${reason(error)}`
    return injectBridge(`<!doctype html><html lang="zh-Hant"><head>${head}</head><body><div class="xl-fail" role="alert"><h1>這份 Excel 檔無法顯示</h1><p>${escapeHtml(reason(error))}</p><p>可以下載原始檔,用 Excel 開啟查看。</p></div>${failureScript(message)}</body></html>`)
  }
  const warnings = book.warnings.map((warning) => `<p class="xl-note">${escapeHtml(warning)}</p>`).join('')
  if (!book.sheets.length) {
    return injectBridge(`<!doctype html><html lang="zh-Hant"><head>${head}</head><body>${warnings}<div class="xl-empty">這份活頁簿沒有可顯示的工作表(全部都已隱藏)。</div></body></html>`)
  }
  const tabs = book.sheets
    .map((sheet, at) => `<a class="xl-tab" role="tab" href="#sheet-${at + 1}" aria-controls="xs-${at + 1}" aria-selected="${at + 1 === book.active}"${sheet.tabColor ? ` style="--tab:#${sheet.tabColor}"` : ''}>${escapeHtml(sheet.name)}</a>`)
    .join('')
  const panels = book.sheets
    .map((sheet, at) => `<section class="xl-sheet" id="xs-${at + 1}" role="tabpanel" aria-label="${escapeHtml(sheet.name)}"${at + 1 === book.active ? '' : ' hidden'}>${sheet.html}</section>`)
    .join('\n')
  const style = `<style>.xl-grid{${book.baseFont}}\n${book.css}</style>`
  return injectBridge(`<!doctype html><html lang="zh-Hant"><head>${head}${style}</head><body data-active="${book.active}">${warnings}<nav class="xl-tabs" role="tablist" aria-label="工作表">${tabs}</nav><main class="xl-main">${panels}</main><script>${PAGE_SCRIPT}</script></body></html>`)
}

/** The searchable text of an Excel file: sheet names and what the cells show, bounded. */
export function xlsxText(data: Buffer): string {
  const book = readXlsx(data, '', false)
  const out: string[] = []
  let size = 0
  for (const sheet of book.sheets) {
    for (const piece of [sheet.name, ...sheet.text]) {
      if (!piece) continue
      out.push(piece)
      size += piece.length + 1
      if (size > XLSX_LIMITS.textChars) return out.join('\n').slice(0, XLSX_LIMITS.textChars)
    }
  }
  return out.join('\n')
}

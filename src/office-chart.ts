/**
 * One chart renderer for PowerPoint, Word and Excel files: a DrawingML chart part (`c:chartSpace`) drawn
 * as an SVG from the values the part caches (what Office drew when it last saved), or from the cells
 * its formulas name when a tool wrote no cache.
 *
 *   renderChart(chartXml, options) → { svg, text } | null
 *
 * - `chartXml`: the parsed chart part (`parseXml(part)`), its root `c:chartSpace`.
 * - `options.widthPt` / `heightPt`: the frame's size in points; the SVG's viewBox is that size, so the
 *   caller sizes it (`width="100%" height="100%"` is set) inside its own positioned box.
 * - `options.resolveColor(holder)`: the CSS colour of a DrawingML colour holder (`a:solidFill`, an
 *   `a:gs` gradient stop, `a:fgClr`) — the renderer's own theme and colour-map rules — or null.
 * - `options.themeColor(i)`: accent `i + 1` (0–5) of the theme, as `#RRGGBB`; series and points without
 *   their own colour cycle through these the way Office does (darker and lighter after the sixth).
 * - `options.fontFamily`, `options.textColor`: the theme's chart font stack and default text colour.
 * - `options.readCells(formula)`: optional; values of a range such as `'Sheet 1'!$A$2:$A$9`, for parts
 *   without caches (xlsx only — other formats have no cells to read).
 *
 * Null means "cannot draw this": the caller shows its labelled box (use `chartTitle` for its name).
 * Every chart group in the plot area is drawn (combo charts), each on its own value axis; a group of a
 * kind this cannot draw beside ones it can is reported inside the drawing. The SVG carries no ids, no
 * scripts and no external references; all text is escaped and every number written into it is finite.
 * The chart area's own fill is drawn only when the part sets one, so the caller picks the default
 * background (Excel: white; PowerPoint and Word: the page shows through).
 */
import { escapeHtml } from './content-render.ts'
import { attr, child, children, local, num, path, type XmlNode } from './xml.ts'

export type ChartOptions = {
  widthPt: number
  heightPt: number
  resolveColor(colorElementParent: XmlNode | null): string | null
  themeColor(index: number): string
  fontFamily: string
  textColor: string
  readCells?(formula: string): (string | number | null)[]
  /** The workbook counts dates from 1904 (a chart part may also say so itself with c:date1904). */
  date1904?: boolean
}

export type ChartResult = { svg: string; text: string }

export const CHART_LIMITS = { series: 255, points: 2000, cells: 50_000, ticks: 60, legend: 60, labels: 2000, markers: 5000, text: 200 }

// ---------------------------------------------------------------------------------------------
// Small helpers

const clamp = (value: number, lo: number, hi: number) => Math.min(hi, Math.max(lo, value))
/** A coordinate for the SVG: finite, bounded, two decimals. */
const r2 = (value: number) => (Number.isFinite(value) ? Math.round(clamp(value, -1e6, 1e6) * 100) / 100 : 0)
const finite = (value: number | null | undefined): value is number => typeof value === 'number' && Number.isFinite(value)

function findDeep(node: XmlNode | null | undefined, name: string, depth = 0): XmlNode | null {
  if (!node || depth > 12) return null
  for (const item of node.children) {
    if (local(item.name) === name) return item
    const found = findDeep(item, name, depth + 1)
    if (found) return found
  }
  return null
}

/** An OOXML boolean `val`: absent element → `absent`; element without val → true. */
function flag(node: XmlNode | null | undefined, absent: boolean): boolean {
  if (!node) return absent
  const val = attr(node, 'val')
  return val === undefined ? true : val === '1' || val === 'true'
}

const isCjk = (ch: string) => /[⺀-鿿가-힯＀-￯　-〿]/.test(ch)
/** Rough text width in points: CJK full width, Latin a little over half. */
const textWidth = (text: string, fs: number) => {
  let sum = 0
  for (const ch of text) sum += isCjk(ch) ? fs : /[A-Z0-9%$@#&mwMW]/.test(ch) ? fs * 0.66 : /[il.,:;!' ]/.test(ch) ? fs * 0.3 : fs * 0.56
  return sum
}

function ellipsize(text: string, maxW: number, fs: number): string {
  if (textWidth(text, fs) <= maxW + 0.5) return text
  const chars = [...text]
  let out = ''
  for (const ch of chars) {
    if (textWidth(`${out}${ch}…`, fs) > maxW) break
    out += ch
  }
  return `${out}…`
}

/** Text broken into lines no wider than `maxW` (words for Latin, characters for CJK); the last line ellipsized. */
function wrap(text: string, maxW: number, fs: number, maxLines: number): string[] {
  const lines: string[] = []
  for (const paragraph of text.split('\n')) {
    const tokens = paragraph.match(/[⺀-鿿가-힯＀-￯　-〿]|[^\s⺀-鿿가-힯＀-￯　-〿]+|\s+/g) ?? ['']
    let line = ''
    for (const token of tokens) {
      if (textWidth(line + token, fs) <= maxW || !line.trim()) {
        if (textWidth(line + token, fs) > maxW && !line.trim()) {
          // One word wider than the line: break it by characters.
          for (const ch of token) {
            if (line && textWidth(line + ch, fs) > maxW) {
              lines.push(line)
              line = ''
            }
            line += ch
          }
        } else line += token
      } else {
        lines.push(line.trimEnd())
        line = token.trim() ? token : ''
      }
    }
    lines.push(line.trimEnd())
  }
  const kept = lines.filter((line, at) => line || at === 0)
  if (kept.length <= maxLines) return kept
  const out = kept.slice(0, maxLines)
  out[maxLines - 1] = ellipsize(`${out[maxLines - 1]}${kept[maxLines]}`, maxW, fs)
  return out
}

const CSS_COLOR = /^(#[0-9a-fA-F]{3,8}|rgba?\(\s*[\d.]+\s*,\s*[\d.]+\s*,\s*[\d.]+\s*(,\s*[\d.]+\s*)?\)|none|transparent)$/
const safeColor = (value: string | null | undefined): string | null => (value && CSS_COLOR.test(value.trim()) ? value.trim() : null)

// HSL for the accent variations Office uses after the sixth series.
function hexToRgb(hex: string): [number, number, number] {
  const h = hex.replace('#', '')
  return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)]
}
function lum(hex: string, mod: number, off: number): string {
  const [r, g, b] = hexToRgb(hex).map((v) => v / 255)
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  let h = 0
  let s = 0
  let l = (max + min) / 2
  if (max !== min) {
    const d = max - min
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
    h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4
    h /= 6
  }
  l = clamp(l * mod + off, 0, 1)
  const hue = (p: number, q: number, t: number) => {
    const u = t < 0 ? t + 1 : t > 1 ? t - 1 : t
    return u < 1 / 6 ? p + (q - p) * 6 * u : u < 1 / 2 ? q : u < 2 / 3 ? p + (q - p) * (2 / 3 - u) * 6 : p
  }
  let rgb: number[]
  if (s === 0) rgb = [l, l, l]
  else {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s
    const p = 2 * l - q
    rgb = [hue(p, q, h + 1 / 3), hue(p, q, h), hue(p, q, h - 1 / 3)]
  }
  return `#${rgb.map((v) => Math.round(v * 255).toString(16).padStart(2, '0')).join('').toUpperCase()}`
}
const ACCENT_FALLBACK = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47']
// Office's order for series 7–12, 13–18, …: darker, lighter, and so on.
const VARIATIONS: [number, number][] = [[1, 0], [0.6, 0], [0.8, 0.2], [0.8, 0], [0.6, 0.4], [0.5, 0], [0.7, 0.3], [0.7, 0], [0.5, 0.5]]

// ---------------------------------------------------------------------------------------------
// Number formats: the common Excel codes, General for the rest

const DAY_MS = 86_400_000

/** Excel's General: up to ten significant digits, scientific beyond 1E+11. */
function general(value: number): string {
  if (!Number.isFinite(value)) return ''
  if (value === 0) return '0'
  const abs = Math.abs(value)
  if (abs >= 1e11 || abs < 1e-9) {
    const [mant, exp] = value.toExponential(5).split('e')
    const m = mant.replace(/\.?0+$/, '')
    const e = Number(exp)
    return `${m}E${e < 0 ? '-' : '+'}${String(Math.abs(e)).padStart(2, '0')}`
  }
  const digits = Math.max(1, Math.floor(Math.log10(abs)) + 1)
  const fixed = value.toFixed(clamp(10 - digits, 0, 10))
  return fixed.includes('.') ? fixed.replace(/\.?0+$/, '') : fixed
}

function splitSections(code: string): string[] {
  const out: string[] = []
  let current = ''
  let quoted = false
  for (let i = 0; i < code.length; i++) {
    const ch = code[i]
    if (ch === '"') quoted = !quoted
    if (ch === '\\' && !quoted && i + 1 < code.length) {
      current += ch + code[++i]
      continue
    }
    if (ch === ';' && !quoted) {
      out.push(current)
      current = ''
    } else current += ch
  }
  out.push(current)
  return out.slice(0, 4)
}

type Tok = { kind: 'lit'; text: string } | { kind: 'num'; text: string } | { kind: 'date'; text: string }

function tokenize(section: string): Tok[] | null {
  const toks: Tok[] = []
  const lit = (text: string) => toks.push({ kind: 'lit', text })
  for (let i = 0; i < section.length; i++) {
    const ch = section[i]
    if (ch === '"') {
      const end = section.indexOf('"', i + 1)
      lit(section.slice(i + 1, end < 0 ? section.length : end))
      i = end < 0 ? section.length : end
    } else if (ch === '\\') lit(section[++i] ?? '')
    else if (ch === '_') {
      i++
      lit(' ')
    } else if (ch === '*') i++
    else if (ch === '[') {
      const end = section.indexOf(']', i)
      if (end < 0) return null
      const inner = section.slice(i + 1, end)
      if (/^[<>=]/.test(inner)) return null // conditions: General
      if (inner.startsWith('$')) lit(inner.slice(1).split('-')[0])
      i = end
    } else if (/[0#?.,]/.test(ch)) {
      const last = toks[toks.length - 1]
      if (last?.kind === 'num') last.text += ch
      else toks.push({ kind: 'num', text: ch })
    } else if ((ch === 'E' || ch === 'e') && (section[i + 1] === '+' || section[i + 1] === '-') && toks[toks.length - 1]?.kind === 'num') {
      let j = i + 2
      while (section[j] === '0' || section[j] === '#') j++
      toks[toks.length - 1].text += section.slice(i, j)
      i = j - 1
    } else if (/[ymdhs]/i.test(ch)) {
      let j = i
      while (j < section.length && section[j].toLowerCase() === ch.toLowerCase()) j++
      toks.push({ kind: 'date', text: section.slice(i, j).toLowerCase() })
      i = j - 1
    } else if (section.slice(i, i + 5).toUpperCase() === 'AM/PM') {
      toks.push({ kind: 'date', text: 'ampm' })
      i += 4
    } else if (ch === '@') toks.push({ kind: 'num', text: '@' })
    else lit(ch)
  }
  return toks
}

function withThousands(int: string): string {
  return int.replace(/\B(?=(\d{3})+(?!\d))/g, ',')
}

function formatPlain(pattern: string, value: number): string {
  // value is already non-negative and scaled for %.
  const sci = /[Ee][+-]/.exec(pattern)
  const body = sci ? pattern.slice(0, sci.index) : pattern
  const point = body.indexOf('.')
  let int = point < 0 ? body : body.slice(0, point)
  const frac = point < 0 ? '' : body.slice(point + 1).replace(/,/g, '')
  // Trailing commas scale by thousands.
  let scaled = value
  while (int.endsWith(',') && point < 0) {
    int = int.slice(0, -1)
    scaled /= 1000
  }
  const comma = int.includes(',')
  const minInt = (int.match(/0/g) ?? []).length
  const decimals = (frac.match(/[0#?]/g) ?? []).length
  const minDecimals = (frac.match(/0/g) ?? []).length
  if (sci) {
    const expDigits = Math.max(1, pattern.length - sci.index - 2)
    let exp = scaled === 0 ? 0 : Math.floor(Math.log10(scaled))
    let mant = scaled === 0 ? 0 : scaled / 10 ** exp
    if (Number(mant.toFixed(decimals)) >= 10) {
      exp += 1
      mant /= 10
    }
    const sign = exp < 0 ? '-' : sci[0][1] === '+' ? '+' : ''
    return `${mant.toFixed(decimals)}E${sign}${String(Math.abs(exp)).padStart(expDigits, '0')}`
  }
  if (scaled >= 1e21) return general(scaled)
  let fixed = Number(scaled.toPrecision(15)).toFixed(decimals)
  let [i, f = ''] = fixed.split('.')
  while (f.length > minDecimals && f.endsWith('0')) f = f.slice(0, -1)
  if (i === '0' && minInt === 0) i = ''
  i = i.padStart(minInt, '0')
  if (comma) i = withThousands(i)
  fixed = f ? `${i}.${f}` : i
  return fixed
}

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']

// The date system of the chart being drawn: rendering is synchronous, so renderChart sets it for the call.
let date1904 = false

function formatDate(toks: Tok[], serial: number): string {
  // The 1904 system counts from 1 January 1904; the 1900 one has its fictitious 29 February 1900.
  const days = date1904 ? serial : serial < 61 ? serial + 1 : serial
  const date = new Date((date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30)) + Math.round(days * DAY_MS / 1000) * 1000)
  if (Number.isNaN(date.getTime())) return general(serial)
  const ampm = toks.some((t) => t.kind === 'date' && t.text === 'ampm')
  let out = ''
  toks.forEach((tok, at) => {
    if (tok.kind !== 'date') {
      out += tok.text
      return
    }
    const t = tok.text
    const near = (step: number) => {
      for (let j = at + step; j >= 0 && j < toks.length; j += step) if (toks[j].kind === 'date') return toks[j].text[0]
      return ''
    }
    const minute = t[0] === 'm' && t.length <= 2 && (near(-1) === 'h' || near(1) === 's')
    const hours = date.getUTCHours()
    if (t[0] === 'y') out += t.length <= 2 ? String(date.getUTCFullYear() % 100).padStart(2, '0') : String(date.getUTCFullYear())
    else if (minute) out += String(date.getUTCMinutes()).padStart(t.length, '0')
    else if (t[0] === 'm') out += t.length === 1 ? String(date.getUTCMonth() + 1) : t.length === 2 ? String(date.getUTCMonth() + 1).padStart(2, '0') : t.length === 3 ? MONTHS[date.getUTCMonth()].slice(0, 3) : t.length === 5 ? MONTHS[date.getUTCMonth()][0] : MONTHS[date.getUTCMonth()]
    else if (t[0] === 'd') out += t.length === 1 ? String(date.getUTCDate()) : t.length === 2 ? String(date.getUTCDate()).padStart(2, '0') : t.length === 3 ? DAYS[date.getUTCDay()].slice(0, 3) : DAYS[date.getUTCDay()]
    else if (t[0] === 'h') out += String(ampm ? hours % 12 || 12 : hours).padStart(Math.min(2, t.length), '0')
    else if (t[0] === 's') out += String(date.getUTCSeconds()).padStart(Math.min(2, t.length), '0')
    else if (t === 'ampm') out += hours < 12 ? 'AM' : 'PM'
  })
  return out
}

/** A number as an Excel number format shows it; codes this does not know show as General. */
export function formatChartNumber(code: string | null | undefined, value: number): string {
  if (!Number.isFinite(value)) return ''
  const trimmed = (code ?? '').trim()
  if (!trimmed || /^general$/i.test(trimmed)) return general(value)
  const sections = splitSections(trimmed)
  let section = sections[0]
  let negate = value < 0
  if (value < 0 && sections.length >= 2 && sections[1].trim()) {
    section = sections[1]
    negate = false
  } else if (value === 0 && sections.length >= 3 && sections[2].trim()) section = sections[2]
  const toks = tokenize(section)
  if (!toks) return general(value)
  const abs = Math.abs(value)
  if (toks.some((t) => t.kind === 'date')) return formatDate(toks, value)
  const numbers = toks.filter((t) => t.kind === 'num')
  if (!numbers.length || numbers.some((t) => t.text === '@')) {
    if (numbers.some((t) => t.text === '@') || !toks.length) return general(value)
    return `${negate ? '-' : ''}${toks.map((t) => t.text).join('')}`
  }
  const percents = toks.filter((t) => t.kind === 'lit' && t.text.includes('%')).reduce((sum, t) => sum + (t.text.match(/%/g) ?? []).length, 0)
  const pattern = numbers.map((t) => t.text).join('')
  const shown = formatPlain(pattern, abs * 100 ** Math.min(percents, 2))
  let out = ''
  let placed = false
  for (const tok of toks) {
    if (tok.kind === 'num') {
      if (!placed) out += shown
      placed = true
    } else out += tok.text
  }
  const zero = !/[1-9]/.test(shown)
  return `${negate && !zero ? '-' : ''}${out}`
}

// ---------------------------------------------------------------------------------------------
// Reading the part

/** The plain text of a title or other rich text holder: paragraphs on their own lines. */
function richText(holder: XmlNode | null, cells?: ChartOptions['readCells']): string {
  if (!holder) return ''
  const rich = findDeep(holder, 'rich')
  if (rich) {
    return children(rich, 'p')
      .map((p) => {
        let line = ''
        const walk = (node: XmlNode, depth: number) => {
          if (depth > 8) return
          for (const item of node.children) {
            const name = local(item.name)
            if (name === 't') line += item.text
            else if (name === 'br') line += '\n'
            else if (name !== 'rPr' && name !== 'pPr') walk(item, depth + 1)
          }
        }
        walk(p, 0)
        return line
      })
      .join('\n')
      .slice(0, CHART_LIMITS.text)
  }
  const tx = child(holder, 'tx')
  if (tx) return (readSeries(tx, cells).values[0] ?? '').slice(0, CHART_LIMITS.text)
  return ''
}

/** The chart's title as Office shows it, for the caller's labelled box and search text. */
export function chartTitle(chartXml: XmlNode): string {
  const space = local(chartXml.name) === 'chartSpace' ? chartXml : child(chartXml, 'chartSpace') ?? chartXml
  return richText(path(space, 'chart', 'title')).replace(/\n/g, ' ').trim()
}

type Read = { values: (string | null)[]; code: string | null; numeric: boolean }

/** The values a `c:tx`, `c:cat`, `c:val` … holds: its cache, its literal, or the cells its formula names. */
function readSeries(node: XmlNode | null, cells?: ChartOptions['readCells']): Read {
  if (!node) return { values: [], code: null, numeric: false }
  const literal = child(node, 'v')
  if (literal) return { values: [literal.text], code: null, numeric: false }
  const cache = findDeep(node, 'numCache') ?? findDeep(node, 'strCache') ?? findDeep(node, 'multiLvlStrCache') ?? findDeep(node, 'numLit') ?? findDeep(node, 'strLit')
  if (cache) {
    const name = local(cache.name)
    const numeric = name === 'numCache' || name === 'numLit'
    const code = child(cache, 'formatCode')?.text ?? null
    const points = name === 'multiLvlStrCache' ? children(child(cache, 'lvl'), 'pt') : children(cache, 'pt')
    const declared = num(child(cache, 'ptCount'), 'val')
    const highest = points.reduce((max, point) => Math.max(max, (num(point, 'idx') ?? -1) + 1), 0)
    const length = clamp(Math.max(finite(declared) ? declared : 0, highest), 0, CHART_LIMITS.points)
    const out: (string | null)[] = Array.from({ length }, () => null)
    points.forEach((point, at) => {
      const idx = num(point, 'idx') ?? at
      if (idx >= 0 && idx < length) out[idx] = child(point, 'v')?.text ?? null
    })
    return { values: out, code, numeric }
  }
  const formula = findDeep(node, 'f')?.text
  if (formula && cells) {
    try {
      const got = cells(formula.trim()).slice(0, CHART_LIMITS.points)
      return { values: got.map((v) => (v === null || v === undefined ? null : String(v))), code: null, numeric: got.some((v) => typeof v === 'number') }
    } catch {
      return { values: [], code: null, numeric: false }
    }
  }
  return { values: [], code: null, numeric: false }
}

const toNumber = (value: string | null): number | null => {
  if (value === null || value.trim() === '') return null
  const parsed = Number(value)
  return Number.isFinite(parsed) ? parsed : null
}

type Line = { color: string; width: number; dash: string }
type Marker = { symbol: string; size: number; fill: string; stroke: string }
type Labels = { val: boolean; cat: boolean; ser: boolean; pct: boolean; pos: string; code: string | null; fs: number; color: string; sep: string }

type Series = {
  index: number // position in the chart, for data-s
  idx: number // c:idx, for automatic colours
  name: string
  named: boolean
  cats: (string | null)[]
  catCode: string | null
  catNumeric: boolean
  values: (number | null)[]
  code: string | null
  xs: (number | null)[] | null
  sizes: (number | null)[] | null
  fill: string // area colour (bars, areas, slices, bubbles)
  line: Line | null // null: no line
  marker: Marker | null
  smooth: boolean
  ser: XmlNode
  labels: Labels | null
}

type GroupKind = 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter' | 'radar' | 'bubble' | 'stock'

type Group = { kind: GroupKind; node: XmlNode; series: Series[]; axIds: string[] }

const GROUP_KINDS: Record<string, GroupKind> = {
  barChart: 'bar', bar3DChart: 'bar', lineChart: 'line', line3DChart: 'line', areaChart: 'area', area3DChart: 'area',
  pieChart: 'pie', pie3DChart: 'pie', ofPieChart: 'pie', doughnutChart: 'doughnut', scatterChart: 'scatter',
  radarChart: 'radar', bubbleChart: 'bubble', stockChart: 'stock'
}

const DASHES: Record<string, number[]> = {
  dash: [4, 3], sysDash: [3, 1], dot: [1, 1], sysDot: [1, 1], lgDash: [8, 3], dashDot: [4, 3, 1, 3], sysDashDot: [3, 1, 1, 1],
  lgDashDot: [8, 3, 1, 3], lgDashDotDot: [8, 3, 1, 3, 1, 3], sysDashDotDot: [3, 1, 1, 1, 1, 1]
}
const AUTO_SYMBOLS = ['diamond', 'square', 'triangle', 'x', 'star', 'circle', 'plus', 'dash', 'dot']

/** Text size (pt) and colour a `c:txPr` sets for its paragraphs, if any. */
function textProps(holder: XmlNode | null, options: ChartOptions): { size: number | null; color: string | null; bold: boolean | null; rot: number | null } {
  const txPr = child(holder, 'txPr') ?? (holder && local(holder.name) === 'txPr' ? holder : null)
  const rich = txPr ?? findDeep(holder, 'rich')
  const def = findDeep(rich, 'defRPr') ?? findDeep(rich, 'rPr')
  const sz = num(def, 'sz')
  const b = attr(def, 'b')
  const rot = num(child(rich, 'bodyPr'), 'rot')
  return {
    size: finite(sz) && sz > 0 ? clamp(sz / 100, 4, 72) : null,
    color: safeColor(options.resolveColor(child(def, 'solidFill'))),
    bold: b === undefined ? null : b === '1' || b === 'true',
    rot: finite(rot) ? clamp(rot / 60000, -90, 90) : null
  }
}

// ---------------------------------------------------------------------------------------------
// Drawing context

type Ctx = {
  options: ChartOptions
  fs: number
  color: string
  W: number
  H: number
  out: string[]
  labelBudget: { left: number }
  markerBudget: { left: number }
}

const textEl = (x: number, y: number, body: string, fs: number, extra = '') =>
  `<text x="${r2(x)}" y="${r2(y)}" font-size="${r2(fs)}"${extra}>${escapeHtml(body)}</text>`

function markerSvg(m: Marker, x: number, y: number): string {
  const s = m.size
  const h = s / 2
  const paint = `fill="${m.fill}" stroke="${m.stroke}" stroke-width="0.75" data-marker=""`
  const p = (d: string) => `<path d="${d}" ${paint}/>`
  switch (m.symbol) {
    case 'circle':
      return `<circle cx="${r2(x)}" cy="${r2(y)}" r="${r2(h)}" ${paint}/>`
    case 'dot':
      return `<circle cx="${r2(x)}" cy="${r2(y)}" r="${r2(Math.max(1, h / 2))}" ${paint}/>`
    case 'square':
      return p(`M${r2(x - h)},${r2(y - h)} L${r2(x + h)},${r2(y - h)} L${r2(x + h)},${r2(y + h)} L${r2(x - h)},${r2(y + h)} Z`)
    case 'triangle':
      return p(`M${r2(x)},${r2(y - h)} L${r2(x + h)},${r2(y + h)} L${r2(x - h)},${r2(y + h)} Z`)
    case 'x':
      return `<path d="M${r2(x - h)},${r2(y - h)} L${r2(x + h)},${r2(y + h)} M${r2(x + h)},${r2(y - h)} L${r2(x - h)},${r2(y + h)}" fill="none" stroke="${m.stroke === 'none' ? m.fill : m.stroke}" stroke-width="1.25" data-marker=""/>`
    case 'plus':
      return `<path d="M${r2(x)},${r2(y - h)} L${r2(x)},${r2(y + h)} M${r2(x - h)},${r2(y)} L${r2(x + h)},${r2(y)}" fill="none" stroke="${m.stroke === 'none' ? m.fill : m.stroke}" stroke-width="1.25" data-marker=""/>`
    case 'star':
      return `<path d="M${r2(x - h)},${r2(y - h)} L${r2(x + h)},${r2(y + h)} M${r2(x + h)},${r2(y - h)} L${r2(x - h)},${r2(y + h)} M${r2(x)},${r2(y - h)} L${r2(x)},${r2(y + h)}" fill="none" stroke="${m.stroke === 'none' ? m.fill : m.stroke}" stroke-width="1" data-marker=""/>`
    case 'dash':
      return p(`M${r2(x - h)},${r2(y - s / 8)} L${r2(x + h)},${r2(y - s / 8)} L${r2(x + h)},${r2(y + s / 8)} L${r2(x - h)},${r2(y + s / 8)} Z`)
    default: // diamond, picture, auto
      return p(`M${r2(x)},${r2(y - h)} L${r2(x + h)},${r2(y)} L${r2(x)},${r2(y + h)} L${r2(x - h)},${r2(y)} Z`)
  }
}

function lineAttrs(line: Line): string {
  return `stroke="${line.color}" stroke-width="${r2(line.width)}" stroke-linejoin="round" stroke-linecap="round"${line.dash ? ` stroke-dasharray="${line.dash}"` : ''}`
}

/** A path through the points; smooth curves as Office draws them (Catmull-Rom as cubic Béziers). */
function polyline(points: [number, number][], smooth: boolean): string {
  if (!points.length) return ''
  const [x0, y0] = points[0]
  let d = `M${r2(x0)},${r2(y0)}`
  if (!smooth || points.length < 3) {
    for (const [x, y] of points.slice(1)) d += ` L${r2(x)},${r2(y)}`
    return d
  }
  for (let i = 0; i < points.length - 1; i++) {
    const p0 = points[Math.max(0, i - 1)]
    const p1 = points[i]
    const p2 = points[i + 1]
    const p3 = points[Math.min(points.length - 1, i + 2)]
    const c1 = [p1[0] + (p2[0] - p0[0]) / 6, p1[1] + (p2[1] - p0[1]) / 6]
    const c2 = [p2[0] - (p3[0] - p1[0]) / 6, p2[1] - (p3[1] - p1[1]) / 6]
    d += ` C${r2(c1[0])},${r2(c1[1])} ${r2(c2[0])},${r2(c2[1])} ${r2(p2[0])},${r2(p2[1])}`
  }
  return d
}

// ---------------------------------------------------------------------------------------------
// Value scales

type Scale = {
  min: number
  max: number
  ticks: number[]
  log: number | null
  reverse: boolean
  code: string
  /** 0 … length along the axis, first value at 0 (reversed axes flip it). */
  at(value: number): number
  length: number
}

/** Smallest of 1, 2, 5 × 10^k not below `raw`. */
function niceUp(raw: number): number {
  if (!(raw > 0) || !Number.isFinite(raw)) return 1
  const power = 10 ** Math.floor(Math.log10(raw))
  const unit = raw / power
  return (unit <= 1.000001 ? 1 : unit <= 2.000001 ? 2 : unit <= 5.000001 ? 5 : 10) * power
}

const tidy = (value: number) => Number(value.toPrecision(12))

/**
 * The scale of a value axis for data from `lo` to `hi`: the file's min, max, major unit, log base and
 * orientation where it sets them, Office's automatic choices otherwise. Null when nothing finite fits.
 */
function makeScale(axis: XmlNode | null, lo: number, hi: number, length: number, maxIntervals: number, code: string, percent: boolean): Scale | null {
  const scaling = child(axis, 'scaling')
  const reverse = attr(child(scaling, 'orientation'), 'val') === 'maxMin'
  const userMin = num(child(scaling, 'min'), 'val')
  const userMax = num(child(scaling, 'max'), 'val')
  const logBase = num(child(scaling, 'logBase'), 'val')
  const unit = num(child(axis, 'majorUnit'), 'val')
  const validUser = !(finite(userMin) && finite(userMax) && userMin >= userMax)
  const fixedMin = validUser && finite(userMin) ? userMin : null
  const fixedMax = validUser && finite(userMax) ? userMax : null
  if (!Number.isFinite(lo) || !Number.isFinite(hi)) return null
  if (lo > hi) [lo, hi] = [hi, lo]

  if (finite(logBase) && logBase >= 2 && logBase <= 1000) {
    const base = logBase
    const positiveLo = lo > 0 ? lo : hi > 0 ? Math.min(hi, 1) : 1
    let eMin = Math.floor(Math.log(fixedMin && fixedMin > 0 ? fixedMin : positiveLo) / Math.log(base) + 1e-9)
    let eMax = Math.ceil(Math.log(fixedMax && fixedMax > 0 ? fixedMax : Math.max(hi, positiveLo)) / Math.log(base) - 1e-9)
    if (!Number.isFinite(eMin) || !Number.isFinite(eMax)) return null
    if (eMax <= eMin) eMax = eMin + 1
    if (eMax - eMin > CHART_LIMITS.ticks) eMin = eMax - CHART_LIMITS.ticks
    const min = fixedMin && fixedMin > 0 ? fixedMin : base ** eMin
    const max = fixedMax && fixedMax > 0 ? fixedMax : base ** eMax
    if (!(max > min) || !Number.isFinite(max) || !(min > 0)) return null
    const ticks: number[] = []
    for (let e = eMin; e <= eMax && ticks.length <= CHART_LIMITS.ticks; e++) {
      const v = tidy(base ** e)
      if (v >= min * (1 - 1e-9) && v <= max * (1 + 1e-9)) ticks.push(v)
    }
    const lmin = Math.log(min)
    const span = Math.log(max) - lmin
    return {
      min, max, ticks, log: base, reverse, code, length,
      at(value) {
        const t = value > 0 ? clamp((Math.log(value) - lmin) / span, 0, 1) : 0
        return (reverse ? 1 - t : t) * length
      }
    }
  }

  let dMin = lo
  let dMax = hi
  if (percent && fixedMin === null && fixedMax === null) {
    // 100% stacked: the full share at the end, the data's own reach below zero.
    dMin = lo < -1e-9 ? Math.max(-1, lo) : 0
    dMax = hi > 1e-9 ? 1 : 0
    if (dMin === dMax) dMax = 1
  }
  if (dMin === dMax) {
    if (dMin === 0) dMax = 1
    else if (dMin > 0) dMin = 0
    else dMax = 0
  }
  // Office starts at zero unless the values sit far from it (all within the top sixth of the range).
  let min = fixedMin ?? (dMin >= 0 ? (dMax > 0 && dMin > (dMax * 5) / 6 ? dMin : 0) : dMin)
  let max = fixedMax ?? (dMax <= 0 ? (dMin < 0 && dMax < (dMin * 5) / 6 ? dMax : 0) : dMax)
  if (!Number.isFinite(max - min)) return null
  if (fixedMin === null && fixedMax === null && min > 0 && min === dMin) {
    // A non-zero start leaves a little room under the lowest value.
    min = dMin - (dMax - dMin) * 0.05
  }
  let step = finite(unit) && unit > 0 && (max - min) / unit <= CHART_LIMITS.ticks ? unit : 0
  if (!step) {
    const range = (max - min) * (fixedMax === null && !percent ? 1.05 : 1)
    step = niceUp(range / Math.max(1, maxIntervals))
  }
  if (!(step > 0) || !Number.isFinite(step)) return null
  if (fixedMin === null) min = min >= 0 && dMin >= 0 && min === 0 ? 0 : Math.floor(tidy(min / step)) * step
  if (fixedMax === null) {
    const headroom = percent || dMax <= 0 ? 0 : (dMax - min) * 0.05
    max = Math.ceil(tidy((dMax + headroom) / step)) * step
    if (dMax <= 0 && max < (fixedMax ?? 0)) max = 0
    if (max <= min) max = min + step
  }
  if (fixedMin === null && dMin < 0) min = Math.min(min, Math.floor(tidy((dMin - (percent ? 0 : (max - dMin) * 0.05)) / step)) * step)
  if (!(max > min) || !Number.isFinite(max) || !Number.isFinite(min)) return null
  if ((max - min) / step > CHART_LIMITS.ticks) step = niceUp((max - min) / CHART_LIMITS.ticks)
  const ticks: number[] = []
  for (let i = 0; i <= CHART_LIMITS.ticks; i++) {
    const v = tidy(min + i * step)
    if (v > max + step * 1e-9) break
    ticks.push(v)
  }
  const span = max - min
  return {
    min, max, ticks, log: null, reverse, code, length,
    at(value) {
      const t = clamp((value - min) / span, 0, 1)
      return (reverse ? 1 - t : t) * length
    }
  }
}

/** The value where the other axis crosses this one (bars grow from it). */
function crossing(perpendicular: XmlNode | null, scale: Scale): number {
  const at = num(child(perpendicular, 'crossesAt'), 'val')
  const crosses = attr(child(perpendicular, 'crosses'), 'val')
  if (finite(at)) return clamp(at, scale.min, scale.max)
  if (crosses === 'max') return scale.max
  if (crosses === 'min' || scale.log) return scale.min
  return clamp(0, scale.min, scale.max)
}

// ---------------------------------------------------------------------------------------------
// The renderer

export function renderChart(chartXml: XmlNode, options: ChartOptions): ChartResult | null {
  const space = local(chartXml.name) === 'chartSpace' ? chartXml : child(chartXml, 'chartSpace') ?? chartXml
  const own = attr(child(space, 'date1904'), 'val')
  date1904 = own !== undefined ? own === '1' || own === 'true' : Boolean(options.date1904)
  try {
    return draw(chartXml, options)
  } catch {
    return null
  } finally {
    date1904 = false
  }
}

function draw(chartXml: XmlNode, options: ChartOptions): ChartResult | null {
  const W = finite(options.widthPt) && options.widthPt > 0 ? clamp(options.widthPt, 20, 20000) : 480
  const H = finite(options.heightPt) && options.heightPt > 0 ? clamp(options.heightPt, 20, 20000) : 300
  const space = local(chartXml.name) === 'chartSpace' ? chartXml : child(chartXml, 'chartSpace') ?? chartXml
  const chartEl = child(space, 'chart')
  const plot = child(chartEl, 'plotArea')
  if (!plot) return null

  const base = textProps(space, options)
  const fs = base.size ?? 10
  const color = base.color ?? safeColor(options.textColor) ?? '#595959'
  const ctx: Ctx = { options, fs, color, W, H, out: [], labelBudget: { left: CHART_LIMITS.labels }, markerBudget: { left: CHART_LIMITS.markers } }

  const accents = Array.from({ length: 6 }, (_, i) => {
    let value = ''
    try {
      value = options.themeColor(i) ?? ''
    } catch {}
    const hex = value.startsWith('#') ? value : `#${value}`
    return /^#[0-9a-fA-F]{6}$/.test(hex) ? hex.toUpperCase() : ACCENT_FALLBACK[i]
  })
  const autoColor = (i: number) => {
    const n = Math.max(0, Math.floor(i))
    const [mod, off] = VARIATIONS[Math.floor(n / 6) % VARIATIONS.length]
    const accent = accents[n % 6]
    return mod === 1 && off === 0 ? accent : lum(accent, mod, off)
  }
  const fillOf = (spPr: XmlNode | null): string | null => {
    if (!spPr) return null
    if (child(spPr, 'noFill')) return 'none'
    const solid = child(spPr, 'solidFill')
    if (solid) return safeColor(options.resolveColor(solid))
    const stop = path(spPr, 'gradFill', 'gsLst', 'gs')
    if (stop) return safeColor(options.resolveColor(stop))
    const pattern = path(spPr, 'pattFill', 'fgClr')
    if (pattern) return safeColor(options.resolveColor(pattern))
    return null
  }
  const lineOf = (spPr: XmlNode | null, fallback: string | null, defaultWidth: number): Line | null => {
    const ln = child(spPr, 'ln')
    if (ln && child(ln, 'noFill')) return null
    const own = child(ln, 'solidFill') ? safeColor(options.resolveColor(child(ln, 'solidFill'))) : child(ln, 'gradFill') ? safeColor(options.resolveColor(path(ln, 'gradFill', 'gsLst', 'gs'))) : null
    const colour = own ?? fallback
    if (!colour) return null
    const w = num(ln, 'w')
    const width = finite(w) && w >= 0 ? clamp(w / 12700, 0.25, 20) : defaultWidth
    const dash = DASHES[attr(child(ln, 'prstDash'), 'val') ?? ''] ?? null
    return { color: colour, width, dash: dash ? dash.map((d) => r2(d * Math.max(1, width))).join(',') : '' }
  }
  const labelsOf = (holders: (XmlNode | null)[], defaultPos: string): Labels | null => {
    // Series settings override the group's.
    const pick = (name: string): XmlNode | null => {
      for (const holder of holders) {
        const found = child(holder, name)
        if (found) return found
      }
      return null
    }
    const own = holders.find((holder) => holder)
    if (!own) return null
    if (flag(pick('delete'), false)) return null
    const val = flag(pick('showVal'), false)
    const cat = flag(pick('showCatName'), false)
    const ser = flag(pick('showSerName'), false)
    const pct = flag(pick('showPercent'), false)
    if (!val && !cat && !ser && !pct) return null
    const fmt = pick('numFmt')
    const props = textProps(holders.find((holder) => child(holder, 'txPr')) ?? null, options)
    return {
      val, cat, ser, pct,
      pos: attr(pick('dLblPos'), 'val') ?? defaultPos,
      code: fmt && attr(fmt, 'sourceLinked') !== '1' ? attr(fmt, 'formatCode') ?? null : null,
      fs: props.size ?? fs,
      color: props.color ?? color,
      sep: (pick('separator')?.text ?? ', ').slice(0, 10)
    }
  }

  // Groups and their series, bounded as a whole.
  const groups: Group[] = []
  let missing = 0
  let seriesCount = 0
  let cells = 0
  for (const node of plot.children) {
    const name = local(node.name)
    if (!name.endsWith('Chart')) continue
    const kind = GROUP_KINDS[name]
    if (!kind) {
      missing++
      continue
    }
    const group: Group = { kind, node, series: [], axIds: children(node, 'axId').map((id) => attr(id, 'val') ?? '') }
    const grouping = attr(child(node, 'grouping'), 'val') ?? 'standard'
    const groupLabels = child(node, 'dLbls')
    const groupMarkers = flag(child(node, 'marker'), true)
    const radarStyle = attr(child(node, 'radarStyle'), 'val') ?? 'marker'
    const barDefault = kind === 'bar' ? (grouping === 'clustered' || grouping === 'standard' ? 'outEnd' : 'ctr') : kind === 'line' || kind === 'scatter' ? 'r' : kind === 'pie' || kind === 'doughnut' ? 'bestFit' : 'ctr'
    for (const ser of children(node, 'ser')) {
      if (seriesCount >= CHART_LIMITS.series || cells >= CHART_LIMITS.cells) break
      const idxRaw = num(child(ser, 'idx'), 'val')
      const idx = finite(idxRaw) ? clamp(Math.floor(idxRaw), 0, 10000) : seriesCount
      const xy = kind === 'scatter' || kind === 'bubble'
      const valRead = readSeries(child(ser, xy ? 'yVal' : 'val') ?? child(ser, 'val'), options.readCells)
      const catRead = readSeries(child(ser, xy ? 'xVal' : 'cat') ?? child(ser, 'cat'), options.readCells)
      const room = Math.max(0, CHART_LIMITS.cells - cells)
      const values = valRead.values.slice(0, room).map(toNumber)
      cells += values.length
      const ownName = readSeries(child(ser, 'tx'), options.readCells).values[0] ?? null
      const name2 = ownName ?? `數列 ${seriesCount + 1}`
      const spPr = child(ser, 'spPr')
      const auto = autoColor(idx)
      const fill = fillOf(spPr) ?? auto
      const ln = child(spPr, 'ln')
      const lnColoured = Boolean(child(ln, 'solidFill') || child(ln, 'gradFill'))
      let line: Line | null
      if (kind === 'line' || kind === 'scatter' || (kind === 'radar' && radarStyle !== 'filled')) {
        // Lines take the line colour, else a fill the series sets, else the automatic colour; width and dash apply either way.
        line = lineOf(spPr, fill === 'none' ? auto : fill, 2.25)
      } else line = lnColoured ? lineOf(spPr, null, 0.75) : null
      const markerNode = child(ser, 'marker')
      const symbol = attr(child(markerNode, 'symbol'), 'val')
      let marker: Marker | null = null
      const wantsMarker =
        symbol !== 'none' &&
        (kind === 'line' ? groupMarkers : kind === 'scatter' ? true : kind === 'radar' ? radarStyle === 'marker' || Boolean(symbol) : kind === 'stock' ? Boolean(markerNode) && Boolean(symbol) : false)
      if (wantsMarker) {
        const size = num(child(markerNode, 'size'), 'val')
        const mSpPr = child(markerNode, 'spPr')
        const lineColour = line?.color ?? fill
        const mFill = fillOf(mSpPr) ?? (lineColour === 'none' ? auto : lineColour)
        const mLine = lineOf(mSpPr, null, 0.75)
        marker = {
          symbol: symbol ?? AUTO_SYMBOLS[idx % AUTO_SYMBOLS.length],
          size: finite(size) ? clamp(size, 2, 72) : 5,
          fill: mFill,
          stroke: child(child(mSpPr, 'ln'), 'noFill') ? 'none' : mLine?.color ?? mFill
        }
      }
      const series: Series = {
        index: seriesCount,
        idx,
        name: name2.slice(0, CHART_LIMITS.text),
        named: ownName !== null,
        cats: catRead.values,
        catCode: catRead.code,
        catNumeric: catRead.numeric,
        values,
        code: valRead.code,
        xs: xy ? (catRead.values.length && catRead.values.every((v) => v === null || toNumber(v) !== null) ? catRead.values.map(toNumber) : null) : null,
        sizes: kind === 'bubble' ? readSeries(child(ser, 'bubbleSize'), options.readCells).values.slice(0, CHART_LIMITS.points).map(toNumber) : null,
        fill,
        line,
        marker,
        smooth: flag(child(ser, 'smooth'), false) && kind !== 'radar',
        ser,
        labels: labelsOf([child(ser, 'dLbls'), groupLabels], barDefault)
      }
      group.series.push(series)
      seriesCount++
    }
    if (group.series.some((s) => s.values.length)) groups.push(group)
  }
  if (!groups.length) return null

  // Axes by id.
  const axes = new Map<string, XmlNode>()
  for (const node of plot.children) {
    const name = local(node.name)
    if (name === 'catAx' || name === 'valAx' || name === 'dateAx' || name === 'serAx') axes.set(attr(child(node, 'axId'), 'val') ?? '', node)
  }

  // Which family draws: the first group decides; groups that cannot share its axes are reported missing.
  const first = groups[0].kind
  const family = first === 'pie' || first === 'doughnut' ? 'pie' : first === 'radar' ? 'radar' : first === 'scatter' || first === 'bubble' ? 'xy' : 'cat'
  const fits = (kind: GroupKind) =>
    family === 'pie' ? kind === first : family === 'radar' ? kind === 'radar' : family === 'xy' ? kind === 'scatter' || kind === 'bubble' : kind === 'bar' || kind === 'line' || kind === 'area' || kind === 'stock'
  const drawn = groups.filter((g) => fits(g.kind))
  missing += groups.length - drawn.length
  const allSeries = drawn.flatMap((g) => (family === 'pie' && g.kind === 'pie' ? g.series.slice(0, 1) : g.series))

  // Background of the chart area and the plot area, only where the part sets one.
  const svg: string[] = ctx.out
  const chartFill = fillOf(child(space, 'spPr'))
  const chartLine = lineOf(child(space, 'spPr'), null, 0.75)
  if ((chartFill && chartFill !== 'none') || chartLine) {
    svg.push(`<rect x="0" y="0" width="${r2(W)}" height="${r2(H)}" fill="${chartFill ?? 'none'}"${chartLine ? ` ${lineAttrs(chartLine)}` : ''}/>`)
  }

  // Title.
  const titleNode = child(chartEl, 'title')
  let title = titleNode ? richText(titleNode, options.readCells) : ''
  const autoDeleted = flag(child(chartEl, 'autoTitleDeleted'), false)
  if (titleNode && !title) title = allSeries.length === 1 ? allSeries[0].name : '圖表標題'
  // Office titles a one-series chart with the series' name unless the file turned that off.
  if (!titleNode && !autoDeleted && allSeries.length === 1 && allSeries[0].named) title = allSeries[0].name
  let top = 6
  if (title) {
    const props = textProps(titleNode, options)
    const tfs = props.size ?? r2(fs * 1.4)
    const lines = wrap(title, W * 0.9, tfs, 3)
    const bold = props.bold ?? true
    lines.forEach((line, at) => {
      svg.push(textEl(W / 2, top + tfs * (at + 0.9), line, tfs, ` text-anchor="middle"${bold ? ' font-weight="700"' : ''}${props.color ? ` fill="${props.color}"` : ''} data-title=""`))
    })
    top += lines.length * tfs * 1.2 + 4
  }

  // Legend.
  const legendNode = child(chartEl, 'legend')
  const area = { x: 6, y: top, w: W - 12, h: H - top - 6 }
  if (legendNode) {
    const items = legendItems(drawn, family, autoColor, fillOf, legendNode)
    const lprops = textProps(legendNode, options)
    const lfs = lprops.size ?? fs
    const pos = attr(child(legendNode, 'legendPos'), 'val') ?? 'r'
    drawLegend(ctx, items, pos, lfs, lprops.color, area)
  }
  if (area.w < 10 || area.h < 10) return finish(ctx, title, allSeries, [], missing)

  let categories: string[] = []
  if (family === 'pie') categories = drawPie(ctx, drawn, area, autoColor, fillOf, lineOf)
  else if (family === 'radar') categories = drawRadar(ctx, drawn, area, axes)
  else if (family === 'xy') {
    if (!drawXY(ctx, drawn, area, axes, autoColor)) return null
  } else {
    const blanks = attr(child(chartEl, 'dispBlanksAs'), 'val') ?? 'gap'
    const result = drawCategories(ctx, drawn, area, axes, plot, autoColor, fillOf, lineOf, blanks)
    if (!result) return null
    categories = result
  }
  return finish(ctx, title, allSeries, categories, missing)
}

function finish(ctx: Ctx, title: string, series: Series[], categories: string[], missing: number): ChartResult {
  const { W, H, fs, color, options } = ctx
  if (missing) ctx.out.push(textEl(4, H - 3, '(部分圖表類型未顯示)', Math.max(6, fs * 0.8), ' fill="#888"'))
  const font = escapeHtml(options.fontFamily.replace(/[^\p{L}\p{N} ,'._-]/gu, '').slice(0, 300) || 'sans-serif')
  const label = escapeHtml((title || '圖表').replace(/\n/g, ' '))
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${r2(W)} ${r2(H)}" width="100%" height="100%" role="img" aria-label="${label}" font-family="${font}" fill="${color}">${ctx.out.join('')}</svg>`
  const text = [title.replace(/\n/g, ' '), ...series.map((s) => s.name), ...categories.slice(0, CHART_LIMITS.text)].filter(Boolean).join(' ')
  return { svg, text }
}

// ---------------------------------------------------------------------------------------------
// Legend

type LegendItem = { name: string; fill: string; line: Line | null; marker: Marker | null; box: boolean }

function legendItems(
  groups: Group[],
  family: string,
  autoColor: (i: number) => string,
  fillOf: (spPr: XmlNode | null) => string | null,
  legendNode: XmlNode
): LegendItem[] {
  let items: LegendItem[] = []
  const g0 = groups[0]
  const varied = (g: Group) => g.series.length === 1 && flag(child(g.node, 'varyColors'), g.kind === 'pie' || g.kind === 'doughnut')
  if (family === 'pie' || (groups.length === 1 && varied(g0) && g0.kind !== 'line' && g0.kind !== 'scatter' && g0.kind !== 'area')) {
    const s = g0.series[0]
    const cats = s.cats.length ? s.cats : s.values.map((_, at) => String(at + 1))
    const vary = varied(g0) || family === 'pie'
    items = cats.slice(0, CHART_LIMITS.legend).map((name, at) => ({
      name: name ?? String(at + 1),
      fill: pointFill(s, at, vary && flag(child(g0.node, 'varyColors'), true), autoColor, fillOf),
      line: null,
      marker: null,
      box: true
    }))
  } else {
    for (const g of groups) {
      for (const s of g.series) {
        const box = g.kind === 'bar' || g.kind === 'area' || g.kind === 'bubble' || (g.kind === 'radar' && attr(child(g.node, 'radarStyle'), 'val') === 'filled')
        items.push({ name: s.name, fill: s.fill, line: box ? null : s.line, marker: box ? null : s.marker, box })
      }
    }
  }
  // Entries the file deletes.
  const deleted = new Set(children(legendNode, 'legendEntry').filter((e) => flag(child(e, 'delete'), false)).map((e) => num(child(e, 'idx'), 'val')))
  return items.filter((_, at) => !deleted.has(at)).slice(0, CHART_LIMITS.legend)
}

function legendKey(item: LegendItem, x: number, y: number, fs: number): string {
  const k = Math.max(5, fs * 0.7)
  if (!item.box && !item.line && !item.marker) return ''
  if (item.box) return `<rect x="${r2(x)}" y="${r2(y - k / 2)}" width="${r2(k)}" height="${r2(k)}" fill="${item.fill}"/>`
  let out = ''
  if (item.line) out += `<path d="M${r2(x - k * 0.4)},${r2(y)} L${r2(x + k * 1.4)},${r2(y)}" fill="none" ${lineAttrs({ ...item.line, width: Math.min(item.line.width, 3) })}/>`
  if (item.marker) out += markerSvg({ ...item.marker, size: Math.min(item.marker.size, k) }, x + k / 2, y).replace(' data-marker=""', '')
  return out
}

/** Draws the legend at its side of `area` and shrinks `area` to what is left for the plot. */
function drawLegend(ctx: Ctx, items: LegendItem[], pos: string, fs: number, colour: string | null, area: { x: number; y: number; w: number; h: number }): void {
  if (!items.length) return
  const keyW = fs * 1.6
  const fill = colour ? ` fill="${colour}"` : ''
  const lineH = fs * 1.25
  if (pos === 'r' || pos === 'l' || pos === 'tr') {
    const textMax = Math.max(fs * 3, area.w * 0.3 - keyW)
    const wrapped = items.map((item) => wrap(item.name, textMax, fs, 2))
    const widest = Math.max(...wrapped.flat().map((line) => textWidth(line, fs)))
    const legendW = Math.min(area.w * 0.35, keyW + widest + 6)
    let rows = wrapped.map((lines) => lines.length)
    let total = rows.reduce((a, b) => a + b, 0) * lineH + (items.length - 1) * fs * 0.3
    let count = items.length
    while (count > 1 && total > area.h) {
      count--
      rows = rows.slice(0, count)
      total = rows.reduce((a, b) => a + b, 0) * lineH + (count - 1) * fs * 0.3
    }
    const x0 = pos === 'l' ? area.x : area.x + area.w - legendW
    let y = pos === 'tr' ? area.y : area.y + Math.max(0, (area.h - total) / 2)
    for (let i = 0; i < count; i++) {
      const lines = i === count - 1 && count < items.length ? [ellipsize(`${wrapped[i][0]}…`, textMax, fs)] : wrapped[i]
      ctx.out.push(legendKey(items[i], x0 + 2, y + lineH * 0.5, fs))
      lines.forEach((line, at) => ctx.out.push(textEl(x0 + keyW, y + lineH * (at + 0.5) + fs * 0.35, line, fs, `${fill} data-legend=""`)))
      y += lines.length * lineH + fs * 0.3
    }
    if (pos === 'l') area.x += legendW + 6
    area.w -= legendW + 6
    return
  }
  // Top or bottom: entries in centred rows across the width.
  const maxItem = area.w - keyW - 8
  const entries = items.map((item) => ({ item, name: ellipsize(item.name, maxItem, fs) }))
  const rowsOut: (typeof entries)[] = []
  let row: typeof entries = []
  let used = 0
  for (const entry of entries) {
    const width = keyW + textWidth(entry.name, fs) + fs
    if (row.length && used + width > area.w) {
      rowsOut.push(row)
      row = []
      used = 0
    }
    row.push(entry)
    used += width
  }
  if (row.length) rowsOut.push(row)
  const rows2 = rowsOut.slice(0, Math.max(1, Math.min(6, Math.floor((area.h * 0.45) / lineH))))
  const height = rows2.length * lineH
  const y0 = pos === 't' ? area.y : area.y + area.h - height
  rows2.forEach((entriesRow, at) => {
    const widths = entriesRow.map((e) => keyW + textWidth(e.name, fs) + fs)
    let x = area.x + Math.max(0, (area.w - widths.reduce((a, b) => a + b, 0)) / 2)
    const y = y0 + lineH * (at + 0.5)
    entriesRow.forEach((e, i) => {
      ctx.out.push(legendKey(e.item, x, y, fs))
      ctx.out.push(textEl(x + keyW, y + fs * 0.35, e.name, fs, `${fill} data-legend=""`))
      x += widths[i]
    })
  })
  if (pos === 't') area.y += height + 4
  area.h -= height + 4
}

function pointFill(s: Series, at: number, vary: boolean, autoColor: (i: number) => string, fillOf: (spPr: XmlNode | null) => string | null): string {
  const dPt = children(s.ser, 'dPt').find((d) => num(child(d, 'idx'), 'val') === at)
  const own = dPt ? fillOf(child(dPt, 'spPr')) : null
  if (own) return own
  return vary ? autoColor(at) : s.fill
}

function labelText(labels: Labels, s: Series, value: number, percent: number | null, category: string): string {
  const parts: string[] = []
  if (labels.ser) parts.push(s.name)
  if (labels.cat) parts.push(category)
  if (labels.val) parts.push(formatChartNumber(labels.code ?? s.code, value))
  if (labels.pct && percent !== null) parts.push(formatChartNumber(labels.code && /%/.test(labels.code) ? labels.code : '0%', percent))
  return parts.filter((p) => p !== '').join(labels.sep)
}

// ---------------------------------------------------------------------------------------------
// Pie and doughnut

function drawPie(
  ctx: Ctx,
  groups: Group[],
  area: { x: number; y: number; w: number; h: number },
  autoColor: (i: number) => string,
  fillOf: (spPr: XmlNode | null) => string | null,
  lineOf: (spPr: XmlNode | null, fallback: string | null, width: number) => Line | null
): string[] {
  const g = groups[0]
  const doughnut = g.kind === 'doughnut'
  // Office puts a doughnut's first series in the middle ring.
  const rings = doughnut ? g.series.slice(0, 10).reverse() : g.series.slice(0, 1)
  const first = doughnut ? rings[rings.length - 1] : rings[0]
  const categories = (first.cats.length ? first.cats : first.values.map((_, at) => String(at + 1))).map((c, at) => c ?? String(at + 1))
  const vary = flag(child(g.node, 'varyColors'), true)
  const startDeg = clamp(num(child(g.node, 'firstSliceAng'), 'val') ?? 0, 0, 360)
  const hole = doughnut ? clamp(num(child(g.node, 'holeSize'), 'val') ?? 50, 10, 90) / 100 : 0
  const explosion = (s: Series, at: number) => {
    const dPt = children(s.ser, 'dPt').find((d) => num(child(d, 'idx'), 'val') === at)
    const value = num(child(dPt, 'explosion'), 'val') ?? num(child(s.ser, 'explosion'), 'val') ?? 0
    return clamp(value, 0, 400) / 100
  }
  const maxExplode = Math.max(0, ...rings.flatMap((s) => s.values.map((_, at) => explosion(s, at))))
  const anyOutside = rings.some((s) => s.labels && (s.labels.pos === 'outEnd' || s.labels.pos === 'bestFit') && !doughnut)
  const labelRoom = anyOutside ? ctx.fs * 2.2 : 0
  const cx = area.x + area.w / 2
  const cy = area.y + area.h / 2
  const R = Math.max(4, (Math.min(area.w, area.h) / 2 - 4 - labelRoom) / (1 + maxExplode))
  const ringW = (R * (1 - hole)) / rings.length
  rings.forEach((s, ri) => {
    const outer = R - ringW * ri
    const inner = doughnut ? outer - ringW : 0
    const raw = s.values.map((v) => (finite(v) ? Math.abs(v) : 0))
    const biggest = Math.max(0, ...raw)
    const values = biggest > 0 ? raw.map((v) => v / biggest) : raw
    const total = values.reduce((a, b) => a + b, 0)
    if (!(total > 0)) return
    const stroke = lineOf(child(s.ser, 'spPr'), '#FFFFFF', 0.75)
    let angle = (startDeg * Math.PI) / 180 - Math.PI / 2
    values.forEach((value, at) => {
      if (!(value > 0)) return
      const sweep = (value / total) * Math.PI * 2
      const end = angle + sweep
      const mid = angle + sweep / 2
      const shift = explosion(s, at) * R
      const ox = cx + Math.cos(mid) * shift
      const oy = cy + Math.sin(mid) * shift
      const dPt = children(s.ser, 'dPt').find((d) => num(child(d, 'idx'), 'val') === at)
      const fill = pointFill(s, at, vary, autoColor, fillOf)
      const edge = lineOf(child(dPt, 'spPr'), null, 0.75) ?? stroke
      const pt = (radius: number, a: number) => `${r2(ox + radius * Math.cos(a))},${r2(oy + radius * Math.sin(a))}`
      const large = sweep > Math.PI ? 1 : 0
      let d: string
      if (sweep >= Math.PI * 2 - 1e-6) {
        // A whole ring or disc: two half arcs (an arc cannot end where it starts).
        d = `M${pt(outer, angle)} A${r2(outer)},${r2(outer)} 0 1 1 ${pt(outer, angle + Math.PI)} A${r2(outer)},${r2(outer)} 0 1 1 ${pt(outer, angle)} Z`
        if (inner > 0) d += ` M${pt(inner, angle)} A${r2(inner)},${r2(inner)} 0 1 0 ${pt(inner, angle + Math.PI)} A${r2(inner)},${r2(inner)} 0 1 0 ${pt(inner, angle)} Z`
      } else if (inner > 0) {
        d = `M${pt(outer, angle)} A${r2(outer)},${r2(outer)} 0 ${large} 1 ${pt(outer, end)} L${pt(inner, end)} A${r2(inner)},${r2(inner)} 0 ${large} 0 ${pt(inner, angle)} Z`
      } else {
        d = `M${r2(ox)},${r2(oy)} L${pt(outer, angle)} A${r2(outer)},${r2(outer)} 0 ${large} 1 ${pt(outer, end)} Z`
      }
      ctx.out.push(`<path d="${d}" fill="${fill}"${edge ? ` ${lineAttrs(edge)}` : ''} fill-rule="evenodd" data-slice=""/>`)
      if (s.labels && ctx.labelBudget.left > 0) {
        ctx.labelBudget.left--
        const text = labelText(s.labels, s, s.values[at] ?? 0, value / total, categories[at] ?? '')
        if (text) {
          const lfs = s.labels.fs
          const outside = !doughnut && (s.labels.pos === 'outEnd' || (s.labels.pos === 'bestFit' && sweep * outer * 0.6 < textWidth(text, lfs) * 0.8))
          const radius = doughnut ? (outer + inner) / 2 : outside ? outer + lfs * 0.6 : s.labels.pos === 'ctr' ? outer * 0.5 : s.labels.pos === 'inEnd' ? outer * 0.78 : outer * 0.62
          const lx = ox + radius * Math.cos(mid)
          const ly = oy + radius * Math.sin(mid)
          const anchor = outside ? (Math.cos(mid) > 0.15 ? 'start' : Math.cos(mid) < -0.15 ? 'end' : 'middle') : 'middle'
          ctx.out.push(textEl(lx, ly + lfs * 0.35, text, lfs, ` text-anchor="${anchor}"${s.labels.color !== ctx.color ? ` fill="${s.labels.color}"` : ''}`))
        }
      }
      angle = end
    })
  })
  return categories
}

// ---------------------------------------------------------------------------------------------
// Category charts: bar, column, line, area and stock, in any combination

type Span = { from: number; to: number } | null

function drawCategories(
  ctx: Ctx,
  groups: Group[],
  area: { x: number; y: number; w: number; h: number },
  axes: Map<string, XmlNode>,
  plot: XmlNode,
  autoColor: (i: number) => string,
  fillOf: (spPr: XmlNode | null) => string | null,
  lineOf: (spPr: XmlNode | null, fallback: string | null, width: number) => Line | null,
  blanks: string
): string[] | null {
  const { fs, options } = ctx
  const horizontal = groups.some((g) => g.kind === 'bar' && attr(child(g.node, 'barDir'), 'val') === 'bar')
  const hasBars = groups.some((g) => g.kind === 'bar')
  const all = groups.flatMap((g) => g.series)
  const n = clamp(Math.max(1, ...all.map((s) => Math.max(s.values.length, s.cats.length))), 1, CHART_LIMITS.points)
  const catSource = all.find((s) => s.cats.some((c) => c !== null && c !== ''))
  // The secondary axis is the one the file places on the right (or top); the primary pair is the other.
  const farSide = (id: string) => ['r', 't'].includes(attr(child(axes.get(id), 'axPos'), 'val') ?? '')
  const primaryGroup = groups.find((g) => !farSide(g.axIds[1] ?? '')) ?? groups[0]
  const primaryCat = axes.get(primaryGroup.axIds[0] ?? '') ?? null
  const catAxisFor = (g: Group) => axes.get(g.axIds[0] ?? '') ?? primaryCat
  const catFmt = child(primaryCat, 'numFmt')
  const catCode = catFmt && attr(catFmt, 'sourceLinked') !== '1' ? attr(catFmt, 'formatCode') ?? catSource?.catCode ?? null : catSource?.catCode ?? null
  const categories = Array.from({ length: n }, (_, at) => {
    const raw = catSource?.cats[at] ?? null
    if (raw === null || raw === '') return catSource ? '' : String(at + 1)
    const value = toNumber(raw)
    return value !== null && (catSource?.catNumeric || local(primaryCat?.name ?? '') === 'dateAx') ? formatChartNumber(catCode, value) : raw
  }).map((c) => c.slice(0, CHART_LIMITS.text))

  // Spans per group: stacking within the group only.
  const spansOf = new Map<Group, Span[][]>()
  const extents = new Map<string, { lo: number; hi: number; percent: boolean; code: string | null }>()
  const valueAxisIds: string[] = []
  for (const g of groups) {
    const grouping = attr(child(g.node, 'grouping'), 'val') ?? 'standard'
    const percent = grouping === 'percentStacked'
    const stacked = percent || grouping === 'stacked'
    const up = new Array<number>(n).fill(0)
    const down = new Array<number>(n).fill(0)
    const totals = new Array<number>(n).fill(0)
    if (percent) for (const s of g.series) s.values.forEach((v, at) => at < n && finite(v) && (totals[at] += Math.abs(v)))
    const spans = g.series.map((s) =>
      Array.from({ length: n }, (_, at): Span => {
        const raw = s.values[at] ?? (blanks === 'zero' && at < s.values.length ? 0 : null)
        if (raw === null) return g.kind === 'area' && stacked ? { from: up[at], to: up[at] } : null
        let value = percent ? (totals[at] > 0 && Number.isFinite(totals[at]) ? raw / totals[at] : 0) : raw
        if (!Number.isFinite(value)) value = 0
        if (!stacked) return { from: 0, to: value }
        const base = g.kind === 'bar' && value < 0 ? down : up
        const from = base[at]
        const to = from + value
        if (!Number.isFinite(to)) return null
        base[at] = to
        return { from, to }
      })
    )
    spansOf.set(g, spans)
    const axisId = g.axIds[1] ?? ''
    if (!valueAxisIds.includes(axisId)) valueAxisIds.push(axisId)
    const ext = extents.get(axisId) ?? { lo: Infinity, hi: -Infinity, percent, code: null }
    // Stacked spans count from where they start; other values only where they end (the scale decides on zero).
    for (const list of spans) for (const span of list) if (span) {
      if (stacked) {
        ext.lo = Math.min(ext.lo, span.from)
        ext.hi = Math.max(ext.hi, span.from)
      }
      ext.lo = Math.min(ext.lo, span.to)
      ext.hi = Math.max(ext.hi, span.to)
    }
    ext.percent = ext.percent || percent
    ext.code ??= g.series.find((s) => s.code)?.code ?? null
    extents.set(axisId, ext)
  }

  if (valueAxisIds.length >= 2 && farSide(valueAxisIds[0]) && !farSide(valueAxisIds[1])) valueAxisIds.splice(0, 2, valueAxisIds[1], valueAxisIds[0])
  const valueAxes = valueAxisIds.slice(0, 2).map((id, at) => ({ id, node: axes.get(id) ?? null, secondary: at === 1 }))
  const shown = (axis: XmlNode | null) => Boolean(axis) && !flag(child(axis, 'delete'), false) && attr(child(axis, 'tickLblPos'), 'val') !== 'none'
  const axisTitle = (axis: XmlNode | null) => (axis ? richText(child(axis, 'title'), options.readCells).replace(/\n/g, ' ') : '')
  const axisFs = (axis: XmlNode | null) => textProps(axis, options).size ?? fs

  // Scales are laid out twice: once to measure their labels, once for the final plot size.
  const scaleFor = (id: string, node: XmlNode | null, length: number) => {
    const ext = extents.get(id)
    if (!ext || !Number.isFinite(ext.lo) || !Number.isFinite(ext.hi)) return makeScale(node, 0, 1, length, 5, 'General', false)
    const fmt = child(node, 'numFmt')
    const code = fmt && attr(fmt, 'sourceLinked') !== '1' ? attr(fmt, 'formatCode') ?? 'General' : ext.percent ? '0%' : ext.code ?? 'General'
    const afs = axisFs(node)
    const intervals = horizontal ? clamp(Math.floor(length / (afs * 5)), 2, 10) : clamp(Math.floor(length / (afs * 2.6)), 2, 12)
    return makeScale(node, ext.lo, ext.hi, length, intervals, code, ext.percent)
  }
  const valLabelSize = (scale: Scale | null, node: XmlNode | null) => (scale && shown(node) ? Math.max(0, ...scale.ticks.map((t) => textWidth(formatChartNumber(scale.code, t), axisFs(node)))) : 0)
  const catTitle = axisTitle(primaryCat)
  const valTitles = valueAxes.map((a) => axisTitle(a.node))

  // Margins.
  let left = 0
  let right = 0
  let bottom = 0
  let topPad = fs * 0.6
  const catFs = axisFs(primaryCat)
  const catShown = shown(primaryCat)
  const plotW0 = area.w
  const plotH0 = area.h - catFs * 2
  let scales = valueAxes.map((a) => scaleFor(a.id, a.node, horizontal ? plotW0 : plotH0))
  if (scales.some((s) => !s)) return null
  if (horizontal) {
    const widest = catShown ? Math.max(0, ...categories.map((c) => textWidth(c, catFs))) : 0
    left = Math.min(area.w * 0.35, widest + catFs * 0.8) + (catTitle ? catFs * 1.4 : 0)
    bottom = scales[0] && shown(valueAxes[0].node) ? axisFs(valueAxes[0].node) * 1.6 : 4
    if (valTitles[0]) bottom += fs * 1.4
    if (valueAxes[1]) topPad += shown(valueAxes[1].node) ? axisFs(valueAxes[1].node) * 1.6 : 0
  } else {
    left = valLabelSize(scales[0], valueAxes[0].node) + fs * 0.6 + (valTitles[0] ? fs * 1.4 : 0)
    right = valueAxes[1] ? valLabelSize(scales[1], valueAxes[1].node) + fs * 0.6 + (valTitles[1] ? fs * 1.4 : 0) : fs * 0.6
  }
  left = Math.max(left, 4)
  let plotW = Math.max(10, area.w - left - right)

  // Category labels (vertical charts): straight, wrapped, rotated, or thinned.
  const skipAttr = num(child(primaryCat, 'tickLblSkip'), 'val')
  let catMode: 'flat' | 'wrap' | 'rotate' = 'flat'
  let skip = finite(skipAttr) && skipAttr >= 1 ? Math.floor(skipAttr) : 1
  let catLines: string[][] = categories.map((c) => [c])
  const rotAttr = textProps(primaryCat, options).rot
  if (!horizontal && catShown) {
    const slot = plotW / n
    const widest = Math.max(0, ...categories.map((c) => textWidth(c, catFs)))
    if (rotAttr !== null && rotAttr !== 0) catMode = 'rotate'
    else if (widest * skip <= slot * 0.92) catMode = 'flat'
    else {
      const wrapped = categories.map((c) => wrap(c, slot * skip * 0.92, catFs, 2))
      const fitsWrap = wrapped.every((lines, at) => lines.join('') === categories[at].replace(/\s+/g, '') || lines.join(' ') === categories[at]) && !wrapped.some((lines) => lines.some((l) => l.endsWith('…'))) && slot * skip > catFs * 2.5
      if (fitsWrap) {
        catMode = 'wrap'
        catLines = wrapped
      } else catMode = 'rotate'
    }
    if (catMode === 'rotate') {
      const angle = rotAttr !== null && rotAttr !== 0 ? Math.abs(rotAttr) : 45
      const along = catFs * 1.15 / Math.sin((angle * Math.PI) / 180)
      if (!(finite(skipAttr) && skipAttr >= 1)) skip = Math.max(1, Math.ceil(along / slot))
      const maxLen = Math.max(catFs * 3, (area.h * 0.38) / Math.sin((angle * Math.PI) / 180))
      catLines = categories.map((c) => [ellipsize(c, maxLen, catFs)])
      const longest = Math.max(0, ...catLines.map((l) => textWidth(l[0], catFs)))
      bottom = longest * Math.sin((angle * Math.PI) / 180) + catFs * 1.2
      // Labels slant down to the left of their tick: the first ones need room before the plot.
      const reach = longest * Math.cos((angle * Math.PI) / 180) - slot / 2 + catFs * 0.6
      if (reach > left) {
        const room = Math.min(area.w * 0.3, reach + 2)
        if (room < reach) {
          const cut = maxLen - (reach - room) / Math.cos((angle * Math.PI) / 180)
          catLines = catLines.map((l) => [ellipsize(l[0], Math.max(catFs * 3, cut), catFs)])
        }
        left = room
      }
    } else {
      if (catMode === 'flat' && !(finite(skipAttr) && skipAttr >= 1)) skip = Math.max(1, Math.ceil((widest + catFs * 0.5) / Math.max(1, slot)))
      if (catMode === 'flat') catLines = categories.map((c) => [ellipsize(c, Math.max(slot * skip * 0.98, catFs * 2), catFs)])
      bottom = Math.max(...catLines.map((l) => l.length)) * catFs * 1.2 + catFs * 0.5
    }
    if (catTitle) bottom += fs * 1.4
  } else if (!horizontal) bottom = 4 + (catTitle ? fs * 1.4 : 0)
  if (horizontal && catShown) {
    const slot = Math.max(1, area.h - bottom - topPad) / n
    if (!(finite(skipAttr) && skipAttr >= 1)) skip = Math.max(1, Math.ceil((catFs * 1.15) / slot))
    catLines = categories.map((c) => [ellipsize(c, left - catFs * 0.6 - (catTitle ? catFs * 1.4 : 0), catFs)])
  }
  bottom = Math.min(bottom, area.h * 0.45)
  const plotX = area.x + left
  const plotY = area.y + topPad
  const plotH = Math.max(10, area.h - topPad - bottom)
  plotW = Math.max(10, area.w - left - right)
  scales = valueAxes.map((a) => scaleFor(a.id, a.node, horizontal ? plotW : plotH))
  if (scales.some((s) => !s)) return null
  const catLen = horizontal ? plotH : plotW
  const catReverse = attr(path(primaryCat, 'scaling', 'orientation'), 'val') === 'maxMin'
  const crossBetween = attr(child(valueAxes[0].node, 'crossBetween'), 'val')
  const between = hasBars || crossBetween !== 'midCat'
  const slot = catLen / n
  const catPos = (at: number) => {
    const t = between ? (at + 0.5) * slot : n > 1 ? (at / (n - 1)) * catLen : catLen / 2
    return catReverse ? catLen - t : t
  }
  const point = (c: number, v: number): [number, number] => (horizontal ? [plotX + v, plotY + plotH - c] : [plotX + c, plotY + plotH - v])

  // Plot area fill.
  const plotFill = fillOf(child(plot, 'spPr'))
  if (plotFill && plotFill !== 'none') ctx.out.push(`<rect x="${r2(plotX)}" y="${r2(plotY)}" width="${r2(plotW)}" height="${r2(plotH)}" fill="${plotFill}"/>`)

  // Gridlines and value labels.
  valueAxes.forEach((axis, ai) => {
    const scale = scales[ai]!
    const grid = child(axis.node, 'majorGridlines')
    const gridLine = grid ? lineOf(child(grid, 'spPr'), '#D9D9D9', 0.75) : null
    const afs = axisFs(axis.node)
    const acolor = textProps(axis.node, options).color
    const fillAttr = acolor ? ` fill="${acolor}"` : ''
    for (const tick of scale.ticks) {
      const v = scale.at(tick)
      if (gridLine) {
        const [x1, y1] = point(0, v)
        const [x2, y2] = point(catLen, v)
        ctx.out.push(`<line x1="${r2(x1)}" y1="${r2(y1)}" x2="${r2(x2)}" y2="${r2(y2)}" ${lineAttrs(gridLine)}/>`)
      }
      if (!shown(axis.node)) continue
      const label = formatChartNumber(scale.code, tick)
      if (horizontal) {
        const [x] = point(0, v)
        const y = axis.secondary ? plotY - afs * 0.5 : plotY + plotH + afs * 1.2
        ctx.out.push(textEl(x, y, label, afs, ` text-anchor="middle"${fillAttr}`))
      } else {
        const [, y] = point(0, v)
        const x = axis.secondary ? plotX + plotW + afs * 0.4 : plotX - afs * 0.4
        ctx.out.push(textEl(x, y + afs * 0.35, label, afs, ` text-anchor="${axis.secondary ? 'start' : 'end'}"${fillAttr}`))
      }
    }
    const axisLine = lineOf(child(axis.node, 'spPr'), null, 0.75)
    if (axisLine && !flag(child(axis.node, 'delete'), false)) {
      const c = axis.secondary ? catLen : 0
      const [x1, y1] = point(c, 0)
      const [x2, y2] = point(c, horizontal ? plotW : plotH)
      ctx.out.push(`<line x1="${r2(x1)}" y1="${r2(y1)}" x2="${r2(x2)}" y2="${r2(y2)}" ${lineAttrs(axisLine)}/>`)
    }
    const titleText = valTitles[ai]
    if (titleText) {
      if (horizontal) ctx.out.push(textEl(plotX + plotW / 2, axis.secondary ? area.y + fs : area.y + area.h - fs * 0.3, ellipsize(titleText, plotW, fs), fs, ' text-anchor="middle"'))
      else {
        const x = axis.secondary ? area.x + area.w - fs * 0.3 : area.x + fs
        const y = plotY + plotH / 2
        ctx.out.push(textEl(x, y, ellipsize(titleText, plotH, fs), fs, ` text-anchor="middle" transform="rotate(${axis.secondary ? 90 : -90} ${r2(x)} ${r2(y)})"`))
      }
    }
  })
  // Category gridlines.
  const catGrid = child(primaryCat, 'majorGridlines')
  if (catGrid) {
    const gl = lineOf(child(catGrid, 'spPr'), '#D9D9D9', 0.75)
    if (gl) for (let at = 0; at <= n && at <= CHART_LIMITS.labels; at += skip) {
      const c = between ? at * slot : catPos(Math.min(at, n - 1))
      const [x1, y1] = point(c, 0)
      const [x2, y2] = point(c, horizontal ? plotW : plotH)
      ctx.out.push(`<line x1="${r2(x1)}" y1="${r2(y1)}" x2="${r2(x2)}" y2="${r2(y2)}" ${lineAttrs(gl)}/>`)
    }
  }

  // Series, in plot-area order: later groups over earlier ones.
  const labelsOut: string[] = []
  for (const g of groups) {
    const ai = Math.max(0, valueAxes.findIndex((a) => a.id === (g.axIds[1] ?? '')))
    const scale = scales[ai]!
    const spans = spansOf.get(g)!
    const catAxis = catAxisFor(g)
    const baseValue = crossing(catAxis, scale)
    const baseAt = scale.at(baseValue)
    const grouping = attr(child(g.node, 'grouping'), 'val') ?? 'standard'
    const stacked = grouping === 'stacked' || grouping === 'percentStacked'
    const vary = g.series.length === 1 && flag(child(g.node, 'varyColors'), false)
    const valueAt = (span: NonNullable<Span>, which: 'from' | 'to') => (!stacked && which === 'from' ? baseAt : scale.at(span[which]))
    if (g.kind === 'bar') {
      const gap = clamp(num(child(g.node, 'gapWidth'), 'val') ?? 150, 0, 500) / 100
      const overlap = clamp(num(child(g.node, 'overlap'), 'val') ?? (stacked ? 100 : 0), -100, 100) / 100
      const k = stacked ? 1 : g.series.length
      const groupW = slot / (1 + gap)
      const barW = groupW / (k - (k - 1) * overlap || 1)
      const stepW = barW * (1 - overlap)
      const used = barW + stepW * (k - 1)
      g.series.forEach((s, si) => {
        spans[si].forEach((span, at) => {
          if (!span || at >= n) return
          const center = catPos(at)
          const offset = stacked ? 0 : si
          const c0 = center - used / 2 + offset * stepW
          const c1 = c0 + barW
          const v0 = valueAt(span, 'from')
          const v1 = valueAt(span, 'to')
          const [ax, ay] = point(horizontal ? c1 : c0, v0)
          const [bx, by] = point(horizontal ? c0 : c1, v1)
          const x = Math.min(ax, bx)
          const y = Math.min(ay, by)
          const w = Math.abs(bx - ax)
          const h = Math.abs(by - ay)
          const dPt = children(s.ser, 'dPt').find((d) => num(child(d, 'idx'), 'val') === at)
          const fill = pointFill(s, at, vary, autoColor, fillOf)
          const edge = lineOf(child(dPt, 'spPr'), null, 0.75) ?? s.line
          ctx.out.push(`<rect x="${r2(x)}" y="${r2(y)}" width="${r2(w)}" height="${r2(h)}" fill="${fill}"${edge ? ` ${lineAttrs(edge)}` : ''} data-s="${s.index}"/>`)
          if (s.labels && ctx.labelBudget.left > 0) {
            ctx.labelBudget.left--
            const raw = s.values[at] ?? 0
            const text = labelText(s.labels, s, raw, grouping === 'percentStacked' ? span.to - span.from : null, categories[at] ?? '')
            if (!text) return
            const lfs = s.labels.fs
            const positive = v1 >= v0
            const pos = s.labels.pos
            const mid = (c0 + c1) / 2
            const tw = textWidth(text, lfs)
            let along: number
            if (pos === 'ctr') along = (v0 + v1) / 2
            else if (pos === 'inBase') along = v0 + (positive ? 1 : -1) * (horizontal ? tw / 2 + 3 : lfs * 0.8)
            else if (pos === 'inEnd') along = v1 - (positive ? 1 : -1) * (horizontal ? tw / 2 + 3 : lfs * 0.8)
            else along = v1 + (positive ? 1 : -1) * (horizontal ? tw / 2 + 3 : lfs * 0.8)
            const [lx, ly] = point(mid, along)
            labelsOut.push(textEl(lx, ly + lfs * 0.35, text, lfs, ` text-anchor="middle"${s.labels.color !== ctx.color ? ` fill="${s.labels.color}"` : ''}`))
          }
        })
      })
    } else if (g.kind === 'area') {
      g.series.forEach((s, si) => {
        const list = spans[si]
        // Runs of present values; an area does not bridge a gap unless stacked.
        let run: { at: number; span: NonNullable<Span> }[] = []
        const flush = () => {
          if (run.length) {
            const upper = run.map(({ at, span }) => point(catPos(at), valueAt(span, 'to')))
            const lower = run.map(({ at, span }) => point(catPos(at), valueAt(span, 'from'))).reverse()
            if (run.length === 1) {
              const [x, y] = upper[0]
              upper.push([x + 0.01, y])
            }
            const d = `M${[...upper, ...lower].map(([x, y]) => `${r2(x)},${r2(y)}`).join(' L')} Z`
            ctx.out.push(`<path d="${d}" fill="${s.fill}"${s.line ? ` ${lineAttrs(s.line)}` : ''}/>`)
          }
          run = []
        }
        list.forEach((span, at) => {
          if (span && at < n) run.push({ at, span })
          else flush()
        })
        flush()
        if (s.labels) addPointLabels(ctx, s, list, categories, (at, span) => point(catPos(at), (valueAt(span, 'to') + valueAt(span, 'from')) / 2), labelsOut, 'ctr')
      })
    } else {
      // Lines and stock.
      if (g.kind === 'stock' && child(g.node, 'hiLowLines')) {
        const hl = lineOf(child(child(g.node, 'hiLowLines'), 'spPr'), ctx.color, 0.75)
        if (hl) for (let at = 0; at < n; at++) {
          const vals = g.series.map((s) => s.values[at]).filter(finite)
          if (vals.length < 2) continue
          const [x1, y1] = point(catPos(at), scale.at(Math.min(...vals)))
          const [x2, y2] = point(catPos(at), scale.at(Math.max(...vals)))
          ctx.out.push(`<line x1="${r2(x1)}" y1="${r2(y1)}" x2="${r2(x2)}" y2="${r2(y2)}" ${lineAttrs(hl)}/>`)
        }
      }
      g.series.forEach((s, si) => {
        const list = spans[si]
        if (s.line) {
          const pieces: [number, number][][] = [[]]
          list.forEach((span, at) => {
            if (at >= n) return
            if (span) pieces[pieces.length - 1].push(point(catPos(at), scale.at(span.to)))
            else if (blanks !== 'span' && pieces[pieces.length - 1].length) pieces.push([])
          })
          const d = pieces.filter((p) => p.length > 1).map((p) => polyline(p, s.smooth)).join(' ')
          if (d) ctx.out.push(`<path d="${d}" fill="none" ${lineAttrs(s.line)}/>`)
        }
        if (s.marker) {
          list.forEach((span, at) => {
            if (!span || at >= n) return
            if (ctx.markerBudget.left-- <= 0) return
            const [x, y] = point(catPos(at), scale.at(span.to))
            const dPt = children(s.ser, 'dPt').find((d) => num(child(d, 'idx'), 'val') === at)
            const own = dPt ? fillOf(path(dPt, 'marker', 'spPr')) ?? fillOf(child(dPt, 'spPr')) : null
            ctx.out.push(markerSvg(own ? { ...s.marker!, fill: own, stroke: own } : s.marker!, x, y))
          })
        }
        if (s.labels) addPointLabels(ctx, s, list, categories, (at, span) => point(catPos(at), scale.at(span.to)), labelsOut, s.labels.pos)
      })
    }
  }

  // The category axis line and labels.
  const catLine = lineOf(child(primaryCat, 'spPr'), primaryCat && !flag(child(primaryCat, 'delete'), false) ? '#BFBFBF' : null, 0.75)
  if (catLine) {
    const v = scales[0]!.at(crossing(primaryCat, scales[0]!))
    const [x1, y1] = point(0, v)
    const [x2, y2] = point(catLen, v)
    ctx.out.push(`<line x1="${r2(x1)}" y1="${r2(y1)}" x2="${r2(x2)}" y2="${r2(y2)}" ${lineAttrs(catLine)}/>`)
  }
  if (catShown) {
    const ccolor = textProps(primaryCat, options).color
    const fillAttr = ccolor ? ` fill="${ccolor}"` : ''
    const angle = rotAttr !== null && rotAttr !== 0 ? rotAttr : -45
    for (let at = 0; at < n && ctx.labelBudget.left > 0; at += skip) {
      const lines = catLines[at]
      if (!lines || !lines.join('')) continue
      ctx.labelBudget.left--
      if (horizontal) {
        const [, y] = point(catPos(at), 0)
        ctx.out.push(textEl(plotX - catFs * 0.4, y + catFs * 0.35, lines[0], catFs, ` text-anchor="end"${fillAttr}`))
      } else if (catMode === 'rotate') {
        const [x] = point(catPos(at), 0)
        const y = plotY + plotH + catFs * 0.9
        ctx.out.push(textEl(x, y, lines[0], catFs, ` text-anchor="${angle < 0 ? 'end' : 'start'}" transform="rotate(${r2(angle)} ${r2(x)} ${r2(y)})"${fillAttr}`))
      } else {
        const [x] = point(catPos(at), 0)
        lines.forEach((line, li) => ctx.out.push(textEl(x, plotY + plotH + catFs * (1.1 + li * 1.2), line, catFs, ` text-anchor="middle"${fillAttr}`)))
      }
    }
  }
  if (catTitle) {
    if (horizontal) {
      const x = area.x + fs
      const y = plotY + plotH / 2
      ctx.out.push(textEl(x, y, ellipsize(catTitle, plotH, fs), fs, ` text-anchor="middle" transform="rotate(-90 ${r2(x)} ${r2(y)})"`))
    } else ctx.out.push(textEl(plotX + plotW / 2, area.y + area.h - fs * 0.3, ellipsize(catTitle, plotW, fs), fs, ' text-anchor="middle"'))
  }
  ctx.out.push(...labelsOut)
  return categories
}

function addPointLabels(ctx: Ctx, s: Series, list: Span[], categories: string[], at: (at: number, span: NonNullable<Span>) => [number, number], out: string[], pos: string) {
  const labels = s.labels!
  list.forEach((span, i) => {
    if (!span || ctx.labelBudget.left <= 0 || s.values[i] === null || s.values[i] === undefined) return
    ctx.labelBudget.left--
    const text = labelText(labels, s, s.values[i]!, null, categories[i] ?? '')
    if (!text) return
    const [x, y] = at(i, span)
    const lfs = labels.fs
    const dx = pos === 'r' ? lfs * 0.5 : pos === 'l' ? -lfs * 0.5 : 0
    const dy = pos === 't' ? -lfs * 0.6 : pos === 'b' ? lfs * 1.2 : pos === 'r' || pos === 'l' || pos === 'ctr' ? lfs * 0.35 : -lfs * 0.6
    const anchor = pos === 'r' ? 'start' : pos === 'l' ? 'end' : 'middle'
    out.push(textEl(x + dx, y + dy, text, lfs, ` text-anchor="${anchor}"${labels.color !== ctx.color ? ` fill="${labels.color}"` : ''}`))
  })
}

// ---------------------------------------------------------------------------------------------
// Scatter and bubble: two value axes

function drawXY(ctx: Ctx, groups: Group[], area: { x: number; y: number; w: number; h: number }, axes: Map<string, XmlNode>, autoColor: (i: number) => string): boolean {
  const { fs, options } = ctx
  const g0 = groups[0]
  let xAxis = axes.get(g0.axIds[0] ?? '') ?? null
  let yAxis = axes.get(g0.axIds[1] ?? '') ?? null
  const posOf = (a: XmlNode | null) => attr(child(a, 'axPos'), 'val')
  const side = (a: XmlNode | null) => posOf(a) === 'l' || posOf(a) === 'r'
  // The first axis is X unless the file places it at a side and the other along the bottom or top.
  if ((side(xAxis) && !side(yAxis) && yAxis) || posOf(yAxis) === 'b' || posOf(yAxis) === 't') [xAxis, yAxis] = [yAxis, xAxis]
  const all = groups.flatMap((g) => g.series)
  const xsOf = (s: Series) => s.xs ?? s.values.map((_, at) => at + 1)
  let xlo = Infinity
  let xhi = -Infinity
  let ylo = Infinity
  let yhi = -Infinity
  for (const s of all) {
    const xs = xsOf(s)
    s.values.forEach((y, at) => {
      const x = xs[at]
      if (!finite(x) || !finite(y)) return
      xlo = Math.min(xlo, x)
      xhi = Math.max(xhi, x)
      ylo = Math.min(ylo, y)
      yhi = Math.max(yhi, y)
    })
  }
  if (!Number.isFinite(xlo)) {
    xlo = 0
    xhi = 1
    ylo = 0
    yhi = 1
  }
  const shown = (axis: XmlNode | null) => Boolean(axis) && !flag(child(axis, 'delete'), false) && attr(child(axis, 'tickLblPos'), 'val') !== 'none'
  const codeOf = (axis: XmlNode | null, fallback: string | null) => {
    const fmt = child(axis, 'numFmt')
    return fmt && attr(fmt, 'sourceLinked') !== '1' ? attr(fmt, 'formatCode') ?? 'General' : fallback ?? 'General'
  }
  const xfs = textProps(xAxis, options).size ?? fs
  const yfs = textProps(yAxis, options).size ?? fs
  const xCode = codeOf(xAxis, all.find((s) => s.catCode)?.catCode ?? null)
  const yCode = codeOf(yAxis, all.find((s) => s.code)?.code ?? null)
  const xTitle = richText(child(xAxis, 'title'), options.readCells).replace(/\n/g, ' ')
  const yTitle = richText(child(yAxis, 'title'), options.readCells).replace(/\n/g, ' ')
  let ys = makeScale(yAxis, ylo, yhi, area.h, clamp(Math.floor(area.h / (yfs * 2.6)), 2, 12), yCode, false)
  if (!ys) return false
  const left = (shown(yAxis) ? Math.max(0, ...ys.ticks.map((t) => textWidth(formatChartNumber(yCode, t), yfs))) + yfs * 0.6 : 4) + (yTitle ? fs * 1.4 : 0)
  const bottom = (shown(xAxis) ? xfs * 1.8 : 4) + (xTitle ? fs * 1.4 : 0)
  const plotX = area.x + left
  const plotY = area.y + fs * 0.6
  const plotW = Math.max(10, area.w - left - fs)
  const plotH = Math.max(10, area.h - bottom - fs * 0.6)
  const xs = makeScale(xAxis, xlo, xhi, plotW, clamp(Math.floor(plotW / (xfs * 5)), 2, 10), xCode, false)
  ys = makeScale(yAxis, ylo, yhi, plotH, clamp(Math.floor(plotH / (yfs * 2.6)), 2, 12), yCode, false)
  if (!xs || !ys) return false
  const px = (x: number) => plotX + xs.at(x)
  const py = (y: number) => plotY + plotH - ys!.at(y)
  const lineOf = (axis: XmlNode | null, name: string) => {
    const grid = child(axis, name)
    if (!grid) return null
    const ln = path(grid, 'spPr', 'ln')
    if (child(ln, 'noFill')) return null
    return safeColor(options.resolveColor(child(ln, 'solidFill'))) ?? '#D9D9D9'
  }
  const ygrid = lineOf(yAxis, 'majorGridlines')
  const xgrid = lineOf(xAxis, 'majorGridlines')
  for (const t of ys.ticks) {
    if (ygrid) ctx.out.push(`<line x1="${r2(plotX)}" y1="${r2(py(t))}" x2="${r2(plotX + plotW)}" y2="${r2(py(t))}" stroke="${ygrid}" stroke-width="0.75"/>`)
    if (shown(yAxis)) ctx.out.push(textEl(plotX - yfs * 0.4, py(t) + yfs * 0.35, formatChartNumber(yCode, t), yfs, ' text-anchor="end"'))
  }
  for (const t of xs.ticks) {
    if (xgrid) ctx.out.push(`<line x1="${r2(px(t))}" y1="${r2(plotY)}" x2="${r2(px(t))}" y2="${r2(plotY + plotH)}" stroke="${xgrid}" stroke-width="0.75"/>`)
    if (shown(xAxis)) ctx.out.push(textEl(px(t), plotY + plotH + xfs * 1.2, formatChartNumber(xCode, t), xfs, ' text-anchor="middle"'))
  }
  // Axis lines where each crosses the other.
  const yCross = crossing(xAxis, ys)
  const xCross = crossing(yAxis, xs)
  ctx.out.push(`<line x1="${r2(plotX)}" y1="${r2(py(yCross))}" x2="${r2(plotX + plotW)}" y2="${r2(py(yCross))}" stroke="#BFBFBF" stroke-width="0.75"/>`)
  ctx.out.push(`<line x1="${r2(px(xCross))}" y1="${r2(plotY)}" x2="${r2(px(xCross))}" y2="${r2(plotY + plotH)}" stroke="#BFBFBF" stroke-width="0.75"/>`)
  if (xTitle) ctx.out.push(textEl(plotX + plotW / 2, area.y + area.h - fs * 0.3, ellipsize(xTitle, plotW, fs), fs, ' text-anchor="middle"'))
  if (yTitle) {
    const x = area.x + fs
    const y = plotY + plotH / 2
    ctx.out.push(textEl(x, y, ellipsize(yTitle, plotH, fs), fs, ` text-anchor="middle" transform="rotate(-90 ${r2(x)} ${r2(y)})"`))
  }
  const labelsOut: string[] = []
  for (const g of groups) {
    if (g.kind === 'bubble') {
      const scale = clamp(num(child(g.node, 'bubbleScale'), 'val') ?? 100, 0, 300) / 100
      const byWidth = attr(child(g.node, 'sizeRepresents'), 'val') === 'w'
      const sizes = g.series.flatMap((s) => (s.sizes ?? []).filter((v): v is number => finite(v) && v > 0))
      const biggest = Math.max(0, ...sizes)
      const maxR = (Math.min(plotW, plotH) * 0.25 * scale) / 2
      const vary = flag(child(g.node, 'varyColors'), false) && g.series.length === 1
      const bubbles: string[] = []
      g.series.forEach((s) => {
        const xsList = xsOf(s)
        s.values.forEach((y, at) => {
          const x = xsList[at]
          const size = s.sizes?.[at]
          if (!finite(x) || !finite(y) || !finite(size) || size <= 0 || !(biggest > 0)) return
          const radius = byWidth ? (size / biggest) * maxR : Math.sqrt(size / biggest) * maxR
          const dPt = children(s.ser, 'dPt').find((d) => num(child(d, 'idx'), 'val') === at)
          const own = dPt ? safeColor(options.resolveColor(path(dPt, 'spPr', 'solidFill'))) : null
          const fill = own ?? (vary ? autoColor(at) : s.fill)
          bubbles.push(`<circle cx="${r2(px(x))}" cy="${r2(py(y))}" r="${r2(radius)}" fill="${fill}"${s.line ? ` ${lineAttrs(s.line)}` : ''} data-bubble=""/>`)
        })
        if (s.labels) addPointLabels(ctx, s, s.values.map((y) => (finite(y) ? { from: y, to: y } : null)), xsList.map((x) => (finite(x) ? formatChartNumber(xCode, x) : '')), (at, span) => [px(xsList[at] ?? 0), py(span.to)], labelsOut, 'ctr')
      })
      ctx.out.push(...bubbles.slice(0, CHART_LIMITS.cells))
      continue
    }
    for (const s of g.series) {
      const xsList = xsOf(s)
      const pts: [number, number][][] = [[]]
      s.values.forEach((y, at) => {
        const x = xsList[at]
        if (finite(x) && finite(y)) pts[pts.length - 1].push([px(x), py(y)])
        else if (pts[pts.length - 1].length) pts.push([])
      })
      if (s.line) {
        const d = pts.filter((p) => p.length > 1).map((p) => polyline(p, s.smooth)).join(' ')
        if (d) ctx.out.push(`<path d="${d}" fill="none" ${lineAttrs(s.line)}/>`)
      }
      if (s.marker) for (const [x, y] of pts.flat()) if (ctx.markerBudget.left-- > 0) ctx.out.push(markerSvg(s.marker, x, y))
      if (s.labels) addPointLabels(ctx, s, s.values.map((y) => (finite(y) ? { from: y, to: y } : null)), xsList.map((x) => (finite(x) ? formatChartNumber(xCode, x) : '')), (at, span) => [px(xsList[at] ?? 0), py(span.to)], labelsOut, s.labels.pos)
    }
  }
  ctx.out.push(...labelsOut)
  return true
}

// ---------------------------------------------------------------------------------------------
// Radar

function drawRadar(ctx: Ctx, groups: Group[], area: { x: number; y: number; w: number; h: number }, axes: Map<string, XmlNode>): string[] {
  const { fs, options } = ctx
  const g0 = groups[0]
  const catAxis = axes.get(g0.axIds[0] ?? '') ?? null
  const valAxis = axes.get(g0.axIds[1] ?? '') ?? null
  const all = groups.flatMap((g) => g.series)
  const n = clamp(Math.max(1, ...all.map((s) => Math.max(s.values.length, s.cats.length))), 1, 360)
  const source = all.find((s) => s.cats.some((c) => c))
  const categories = Array.from({ length: n }, (_, at) => (source?.cats[at] ?? String(at + 1)).slice(0, CHART_LIMITS.text))
  const values = all.flatMap((s) => s.values.filter(finite))
  const lo = values.length ? Math.min(...values) : 0
  const hi = values.length ? Math.max(...values) : 1
  const widest = Math.max(0, ...categories.map((c) => textWidth(c, fs)))
  const R = Math.max(10, Math.min(area.w / 2 - Math.min(widest, area.w * 0.25) - fs, area.h / 2 - fs * 1.6))
  const fmt = child(valAxis, 'numFmt')
  const code = fmt && attr(fmt, 'sourceLinked') !== '1' ? attr(fmt, 'formatCode') ?? 'General' : all.find((s) => s.code)?.code ?? 'General'
  const scale = makeScale(valAxis, lo, hi, R, clamp(Math.floor(R / (fs * 2.2)), 2, 8), code, false)
  if (!scale) return categories
  const cx = area.x + area.w / 2
  const cy = area.y + area.h / 2
  const angle = (at: number) => -Math.PI / 2 + (at / n) * Math.PI * 2
  const xy = (at: number, radius: number): [number, number] => [cx + radius * Math.cos(angle(at)), cy + radius * Math.sin(angle(at))]
  const grid = child(valAxis, 'majorGridlines') ? '#D9D9D9' : null
  if (grid) {
    for (const t of scale.ticks) {
      const radius = scale.at(t)
      const d = `M${Array.from({ length: n }, (_, at) => xy(at, radius).map(r2).join(',')).join(' L')} Z`
      ctx.out.push(`<path d="${d}" fill="none" stroke="${grid}" stroke-width="0.75"/>`)
    }
  }
  if (n <= CHART_LIMITS.ticks * 3) for (let at = 0; at < n; at++) {
    const [x, y] = xy(at, R)
    ctx.out.push(`<line x1="${r2(cx)}" y1="${r2(cy)}" x2="${r2(x)}" y2="${r2(y)}" stroke="#D9D9D9" stroke-width="0.75"/>`)
  }
  if (!flag(child(valAxis, 'delete'), false) && attr(child(valAxis, 'tickLblPos'), 'val') !== 'none') {
    for (const t of scale.ticks) ctx.out.push(textEl(cx + fs * 0.3, cy - scale.at(t) + fs * 0.35, formatChartNumber(code, t), fs * 0.9, ' text-anchor="start"'))
  }
  if (!flag(child(catAxis, 'delete'), false)) {
    const every = Math.max(1, Math.ceil((n * fs * 1.2) / (Math.PI * 2 * R)))
    for (let at = 0; at < n && at < CHART_LIMITS.labels; at += every) {
      const [x, y] = xy(at, R + fs * 0.7)
      const c = Math.cos(angle(at))
      ctx.out.push(textEl(x, y + fs * 0.35 + (Math.sin(angle(at)) > 0.5 ? fs * 0.4 : 0), ellipsize(categories[at], area.w * 0.25, fs), fs, ` text-anchor="${c > 0.2 ? 'start' : c < -0.2 ? 'end' : 'middle'}"`))
    }
  }
  for (const g of groups) {
    const filled = attr(child(g.node, 'radarStyle'), 'val') === 'filled'
    for (const s of g.series) {
      const pts = Array.from({ length: n }, (_, at) => {
        const v = s.values[at]
        return finite(v) ? xy(at, scale.at(v)) : null
      })
      const present = pts.filter((p): p is [number, number] => p !== null)
      if (!present.length) continue
      const d = `M${present.map((p) => p.map(r2).join(',')).join(' L')} Z`
      if (filled) ctx.out.push(`<path d="${d}" fill="${s.fill}"${s.line ? ` ${lineAttrs(s.line)}` : ''}/>`)
      else if (s.line) ctx.out.push(`<path d="${d}" fill="none" ${lineAttrs(s.line)}/>`)
      if (s.marker && !filled) for (const [x, y] of present) if (ctx.markerBudget.left-- > 0) ctx.out.push(markerSvg(s.marker, x, y))
    }
  }
  return categories
}

/**
 * DrawingML shape geometry for PowerPoint, Word and Excel drawings: the 187 preset shapes of ECMA-376
 * (`a:prstGeom`) and custom shapes (`a:custGeom`), evaluated by one engine into SVG path data.
 *
 * The preset table (`office-geometry-presets.json`) is the standard's presetShapeDefinitions.xml converted
 * to JSON; its `_meta` key carries the source and Ecma's notice. Each shape is adjust values, guide
 * formulas, a text rectangle and paths, all in EMU; this module evaluates them for one box and emits
 * points (EMU / 12700). Output is numbers only, so it can go straight into an SVG `d` attribute.
 */
import { readFileSync } from 'node:fs'
import { attr, child, children, local, type XmlNode } from './xml.ts'

export type GeometryFill = 'norm' | 'none' | 'lighten' | 'lightenLess' | 'darken' | 'darkenLess'
export type GeometryPath = {
  /** SVG path data in points, origin at the shape's top-left corner. */
  d: string
  /** How the path is filled: `none` is outline only; the shaded modes are the shape's fill darkened or lightened. */
  fill: GeometryFill
  stroke: boolean
  /** Whether the path closes (`Z`); an open path is a line and takes arrowheads. */
  closed: boolean
}
/**
 * Draw every path's fill first, in order, then every path's stroke: the standard lists chartPlus's lines
 * before its filled box, and PowerPoint still shows them on top.
 */
export type Geometry = { paths: GeometryPath[]; textRect: { l: number; t: number; r: number; b: number } }

type Arg = string | number
type Command = [string, ...Arg[]]
type PathDef = { w?: number; h?: number; fill?: GeometryFill; stroke?: boolean; c: Command[] }
type ShapeDef = { av?: [string, string][]; gd?: [string, string][]; rect?: [Arg, Arg, Arg, Arg]; p: PathDef[] }

const TABLE = JSON.parse(readFileSync(new URL('./office-geometry-presets.json', import.meta.url), 'utf8')) as { shapes: Record<string, ShapeDef> }
const PRESETS = new Map(Object.entries(TABLE.shapes))

/** Every preset shape name ECMA-376 defines (`ST_ShapeType`). */
export const PRESET_NAMES: readonly string[] = [...PRESETS.keys()]

const EMU_PER_PT = 12700
/** Sizes past this (about 79 km) are clamped so every intermediate stays finite. */
const MAX_EMU = 1e12
const MAX_COMMANDS = 20_000
const FILLS = new Set<GeometryFill>(['norm', 'none', 'lighten', 'lightenLess', 'darken', 'darkenLess'])

const ANGLE = Math.PI / 180 / 60000 // ECMA angles are 60000ths of a degree

/** A guide value: a finite number, or 0 (a division by zero or atan2(0, 0) yields 0, as PowerPoint draws). */
const safe = (n: number): number => (Number.isFinite(n) ? n : 0)

/** The guides every shape can use without defining them (ECMA-376 20.1.9.11, plus cd3 which presets use). */
function builtins(w: number, h: number): Map<string, number> {
  const ss = Math.min(w, h)
  const b = new Map<string, number>([
    ['w', w], ['h', h], ['l', 0], ['t', 0], ['r', w], ['b', h], ['hc', w / 2], ['vc', h / 2],
    ['ss', ss], ['ls', Math.max(w, h)],
    ['cd2', 10800000], ['cd3', 7200000], ['cd4', 5400000], ['cd8', 2700000],
    ['3cd4', 16200000], ['3cd8', 8100000], ['5cd8', 13500000], ['7cd8', 18900000]
  ])
  for (const n of [2, 3, 4, 5, 6, 8, 10, 12, 16, 32]) {
    b.set(`wd${n}`, w / n)
    b.set(`hd${n}`, h / n)
    b.set(`ssd${n}`, ss / n)
  }
  return b
}

/**
 * Guide evaluation for one box. Guides are evaluated in order and a later guide may reuse an earlier
 * name (gear6 redefines `a1`), so each formula sees the values defined before it. A name used before
 * any definition (custom shapes do this) resolves to its first definition further down; a guide that
 * depends on itself reads as 0.
 */
class Guides {
  private readonly values = new Map<string, number>()
  private pending = new Map<string, string | number>()
  private readonly busy = new Set<string>()
  private readonly builtin: Map<string, number>

  constructor(w: number, h: number) {
    this.builtin = builtins(w, h)
  }

  /** Defines the guides in order; a number is a fixed value (an adjust value the shape gives). */
  run(list: [string, string | number][]): void {
    this.pending = new Map()
    for (let i = list.length - 1; i >= 0; i--) this.pending.set(list[i][0], list[i][1])
    for (const [name, formula] of list) {
      this.pending.delete(name)
      this.values.set(name, typeof formula === 'number' ? safe(formula) : this.evaluate(formula))
    }
  }

  arg(token: Arg | undefined): number {
    if (typeof token === 'number') return safe(token)
    if (token === undefined || token === '') return 0
    const literal = Number(token)
    if (Number.isFinite(literal)) return literal
    const known = this.values.get(token)
    if (known !== undefined) return known
    const ahead = this.pending.get(token)
    if (ahead !== undefined && !this.busy.has(token)) {
      this.busy.add(token)
      const value = typeof ahead === 'number' ? safe(ahead) : this.evaluate(ahead)
      this.busy.delete(token)
      this.values.set(token, value) // until its own definition comes round and overwrites it
      return value
    }
    return this.builtin.get(token) ?? 0
  }

  evaluate(formula: string): number {
    const [op, ...rest] = formula.trim().split(/\s+/)
    const [x, y, z] = [0, 1, 2].map((i) => (i < rest.length ? this.arg(rest[i]) : 0))
    switch (op) {
      case 'val': return safe(x)
      case '*/': return safe((x * y) / z)
      case '+-': return safe(x + y - z)
      case '+/': return safe((x + y) / z)
      case '?:': return x > 0 ? y : z
      case 'abs': return Math.abs(x)
      case 'at2': return safe(Math.atan2(y, x) / ANGLE)
      case 'cat2': return safe(x * Math.cos(Math.atan2(z, y)))
      case 'sat2': return safe(x * Math.sin(Math.atan2(z, y)))
      case 'cos': return safe(x * Math.cos(y * ANGLE))
      case 'sin': return safe(x * Math.sin(y * ANGLE))
      case 'tan': return safe(x * Math.tan(y * ANGLE))
      case 'max': return Math.max(x, y)
      case 'min': return Math.min(x, y)
      case 'mod': return safe(Math.sqrt(x * x + y * y + z * z))
      case 'pin': return y < x ? x : y > z ? z : y
      case 'sqrt': return safe(Math.sqrt(x))
      default: return 0
    }
  }
}

const num = (n: number): string => {
  const rounded = Math.round(n * 100) / 100
  return String(Object.is(rounded, -0) ? 0 : rounded)
}
const pair = (x: number, y: number) => `${num(x)},${num(y)}`

/**
 * The parametric angle of the point an ECMA visual angle points at on an ellipse with radii a, b:
 * `stAng`/`swAng` are angles of the ray from the centre (ECMA-376 20.1.9.3), not of the parameter.
 * Continuous and monotonic in `visual`, so whole turns survive.
 */
function parametric(visual: number, a: number, b: number): number {
  if (a <= 0 || b <= 0) return visual
  const turns = Math.floor((visual + Math.PI) / (2 * Math.PI))
  const wrapped = visual - turns * 2 * Math.PI
  return Math.atan2(a * Math.sin(wrapped), b * Math.cos(wrapped)) + turns * 2 * Math.PI
}

function drawPath(def: PathDef, guides: Guides, w: number, h: number): GeometryPath {
  const pw = guides.arg(def.w)
  const ph = guides.arg(def.h)
  // Path coordinates are in the path's own w × h space when it gives one; scale to the box, then to points.
  const sx = (pw > 0 ? w / pw : 1) / EMU_PER_PT
  const sy = (ph > 0 ? h / ph : 1) / EMU_PER_PT
  const parts: string[] = []
  let cx = 0 // current point, in path units
  let cy = 0
  let closed = false
  const at = (x: number, y: number) => pair(x * sx, y * sy)
  for (const command of def.c.slice(0, MAX_COMMANDS)) {
    const [kind, ...args] = command
    const v = args.map((a) => guides.arg(a))
    if (kind !== 'M' && kind !== 'Z' && parts.length === 0) parts.push(`M${at(cx, cy)}`)
    if (kind === 'M' || kind === 'L') {
      ;[cx, cy] = [v[0] ?? 0, v[1] ?? 0]
      parts.push(`${kind}${at(cx, cy)}`)
    } else if (kind === 'Q' && v.length >= 4) {
      parts.push(`Q${at(v[0], v[1])} ${at(v[2], v[3])}`)
      ;[cx, cy] = [v[2], v[3]]
    } else if (kind === 'C' && v.length >= 6) {
      parts.push(`C${at(v[0], v[1])} ${at(v[2], v[3])} ${at(v[4], v[5])}`)
      ;[cx, cy] = [v[4], v[5]]
    } else if (kind === 'A') {
      const a = Math.abs(v[0] ?? 0)
      const b = Math.abs(v[1] ?? 0)
      const start = parametric((v[2] ?? 0) * ANGLE, a, b)
      const end = parametric(((v[2] ?? 0) + (v[3] ?? 0)) * ANGLE, a, b)
      const ox = cx - a * Math.cos(start)
      const oy = cy - b * Math.sin(start)
      const point = (t: number): [number, number] => [ox + a * Math.cos(t), oy + b * Math.sin(t)]
      const sweep = end - start
      const rx = num(a * sx)
      const ry = num(b * sy)
      if (a * sx === 0 || b * sy === 0 || sweep === 0) {
        ;[cx, cy] = point(end)
        parts.push(`L${at(cx, cy)}`)
        continue
      }
      // An SVG arc whose ends meet draws nothing: a whole turn goes in as two halves, then the remainder.
      const turn = 2 * Math.PI
      const whole = Math.abs(sweep) >= turn - 1e-9
      const rest = whole ? sweep - Math.sign(sweep) * turn * Math.floor(Math.abs(sweep) / turn + 1e-9) : sweep
      const steps: number[] = whole ? [start + Math.sign(sweep) * Math.PI, start + Math.sign(sweep) * turn] : []
      const from = whole ? start + Math.sign(sweep) * turn : start
      if (Math.abs(rest) > 1e-9) steps.push(from + rest)
      const flag = sweep > 0 ? 1 : 0
      let previous = start
      for (const t of steps) {
        const large = Math.abs(t - previous) > Math.PI + 1e-9 ? 1 : 0
        const [ex, ey] = point(t)
        parts.push(`A${rx},${ry} 0 ${large} ${flag} ${at(ex, ey)}`)
        previous = t
      }
      ;[cx, cy] = point(end)
    } else if (kind === 'Z') {
      parts.push('Z')
      closed = true
    }
  }
  const fill = def.fill && FILLS.has(def.fill) ? def.fill : 'norm'
  return { d: parts.join(' '), fill, stroke: def.stroke !== false, closed }
}

function evaluate(def: ShapeDef, widthEmu: number, heightEmu: number, adjust: Record<string, number>): Geometry | null {
  const w = Math.min(Math.max(safe(widthEmu), 0), MAX_EMU)
  const h = Math.min(Math.max(safe(heightEmu), 0), MAX_EMU)
  const guides = new Guides(w, h)
  const list: [string, string | number][] = (def.av ?? []).map(([name, formula]) => {
    const given = Object.hasOwn(adjust, name) ? adjust[name] : undefined
    return [name, typeof given === 'number' && Number.isFinite(given) ? given : formula]
  })
  guides.run([...list, ...(def.gd ?? [])])
  const paths = def.p.map((path) => drawPath(path, guides, w, h)).filter((path) => path.d !== '')
  if (paths.length === 0) return null
  const [l, t, r, b] = def.rect ?? ['l', 't', 'r', 'b']
  const edge = (arg: Arg, scale: number) => Math.round((guides.arg(arg) / scale) * 100) / 100
  return {
    paths,
    textRect: { l: edge(l, EMU_PER_PT), t: edge(t, EMU_PER_PT), r: edge(r, EMU_PER_PT), b: edge(b, EMU_PER_PT) }
  }
}

/**
 * A preset shape (`a:prstGeom prst`) in a box of widthEmu × heightEmu EMU, with the shape's adjust
 * values (from {@link adjustValues}) overriding the preset's defaults. Paths and text rectangle are in
 * points. Returns null for a name the standard does not define.
 */
export function presetGeometry(name: string, widthEmu: number, heightEmu: number, adjust: Record<string, number>): Geometry | null {
  const def = PRESETS.get(name)
  return def ? evaluate(def, widthEmu, heightEmu, adjust) : null
}

/** `<a:gd name fmla="val N">` entries of an `a:avLst` (a prstGeom's or custGeom's), by name. */
export function adjustValues(geometry: XmlNode | null | undefined): Record<string, number> {
  const out: Record<string, number> = {}
  for (const gd of children(child(geometry, 'avLst'), 'gd')) {
    const name = attr(gd, 'name')
    const match = /^\s*val\s+(-?\d+(?:\.\d+)?)\s*$/.exec(attr(gd, 'fmla') ?? '')
    if (name && match) out[name] = Number(match[1])
  }
  return out
}

const COMMANDS: Record<string, string> = { moveTo: 'M', lnTo: 'L', arcTo: 'A', quadBezTo: 'Q', cubicBezTo: 'C', close: 'Z' }

/**
 * A custom shape (`a:custGeom`) in a box of widthEmu × heightEmu EMU: its own adjust values, guides,
 * text rectangle and paths, evaluated like a preset. Null when it has no path to draw.
 */
export function customGeometry(custGeom: XmlNode, widthEmu: number, heightEmu: number): Geometry | null {
  const formulas = (list: string): [string, string][] =>
    children(child(custGeom, list), 'gd').flatMap((gd) => {
      const name = attr(gd, 'name')
      return name ? [[name, attr(gd, 'fmla') ?? '']] : []
    })
  const rectNode = child(custGeom, 'rect')
  const def: ShapeDef = {
    av: formulas('avLst'),
    gd: formulas('gdLst'),
    rect: rectNode ? [attr(rectNode, 'l') ?? 'l', attr(rectNode, 't') ?? 't', attr(rectNode, 'r') ?? 'r', attr(rectNode, 'b') ?? 'b'] : undefined,
    p: children(child(custGeom, 'pathLst'), 'path').map((path) => {
      const fill = attr(path, 'fill') as GeometryFill | undefined
      const stroke = attr(path, 'stroke')
      return {
        w: attr(path, 'w') !== undefined ? Number(attr(path, 'w')) : undefined,
        h: attr(path, 'h') !== undefined ? Number(attr(path, 'h')) : undefined,
        fill: fill && FILLS.has(fill) ? fill : undefined,
        stroke: stroke === 'false' || stroke === '0' ? false : undefined,
        c: path.children.flatMap((cmd): Command[] => {
          const kind = COMMANDS[local(cmd.name)]
          if (!kind) return []
          if (kind === 'A') return [[kind, attr(cmd, 'wR') ?? 0, attr(cmd, 'hR') ?? 0, attr(cmd, 'stAng') ?? 0, attr(cmd, 'swAng') ?? 0]]
          const pts = children(cmd, 'pt').flatMap((pt) => [attr(pt, 'x') ?? 0, attr(pt, 'y') ?? 0])
          const need = { M: 2, L: 2, Q: 4, C: 6, Z: 0 }[kind] ?? 0
          return pts.length >= need ? [[kind, ...pts.slice(0, need)]] : []
        })
      }
    })
  }
  return evaluate(def, widthEmu, heightEmu, {})
}

/** The geometry of a shape properties element (`spPr`): its prstGeom or custGeom; null when it has neither or an unknown preset. */
export function shapeGeometry(spPr: XmlNode | null | undefined, widthEmu: number, heightEmu: number): Geometry | null {
  const node = spPr?.children.find((c) => local(c.name) === 'prstGeom' || local(c.name) === 'custGeom')
  if (!node) return null
  if (local(node.name) === 'custGeom') return customGeometry(node, widthEmu, heightEmu)
  return presetGeometry(attr(node, 'prst') ?? 'rect', widthEmu, heightEmu, adjustValues(node))
}

/**
 * How a shaded path mode alters the fill: draw the path in the fill colour, then again in `color` at
 * `opacity` on top. Null for `norm` and `none`.
 */
export function fillShade(mode: GeometryFill): { color: '#000' | '#fff'; opacity: number } | null {
  switch (mode) {
    case 'darken': return { color: '#000', opacity: 0.4 }
    case 'darkenLess': return { color: '#000', opacity: 0.2 }
    case 'lighten': return { color: '#fff', opacity: 0.4 }
    case 'lightenLess': return { color: '#fff', opacity: 0.2 }
    default: return null
  }
}

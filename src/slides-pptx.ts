import { posix } from 'node:path'
import { lexer, type Token, type Tokens } from 'marked'
import { render as renderDiagram } from 'pele'
import pptxgen from 'pptxgenjs'
import { analyzeSlide, splitSlides, type SlideModel } from './slides.ts'

// Node loads the ES build, whose default export is the class; the types describe the CommonJS wrapper.
const PptxGenJS = ((pptxgen as unknown as { default?: typeof pptxgen.default }).default ?? pptxgen) as typeof pptxgen.default

/**
 * A deck as a .pptx: the same slides, as native PowerPoint text, tables and pictures, so the text stays
 * editable. PowerPoint does not flow layout the way the browser does, so block heights are estimated and
 * a slide's text is scaled down together when it would not fit. Close to the web deck, not identical.
 */

const W = 13.333
const H = 7.5
const MARGIN_X = 0.85
const FONT = 'Microsoft JhengHei'
const MONO = 'Consolas'
const C = { text: '1F2328', muted: '59636E', accent: '2563EB', accentSoft: 'DBE7FF', code: 'F6F8FA', border: 'D1D9E0' }
const SIZE = { title: 30, body: 20 }

type Run = { text: string; options?: pptxgen.default.TextPropsOptions }
type Box = { x: number; y: number; w: number; h: number }
type ReadFile = (path: string) => Buffer | null

// Text width in ems: CJK and full-width characters are one em, everything else about half.
const ems = (text: string) => [...text].reduce((sum, char) => sum + (/[⺀-￯]/.test(char) ? 1 : 0.55), 0)
const lineHeight = (pt: number) => (pt * 1.3) / 72

function estimateLines(text: string, pt: number, width: number): number {
  return text.split('\n').reduce((sum, line) => sum + Math.max(1, Math.ceil((ems(line) * pt) / 72 / Math.max(width, 0.5))), 0)
}

// Characters XML 1.0 cannot hold (control characters, lone surrogates) would make PowerPoint reject the file.
const xmlSafe = (text: string) => text.replace(/[^\t\n\r\u0020-\uD7FF\uE000-\uFFFD\u{10000}-\u{10FFFF}]/gu, '')

function decodeEntities(text: string): string {
  return xmlSafe(text).replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos|nbsp);/gi, (_, name: string) => {
    const named: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ' }
    if (name[0] !== '#') return named[name.toLowerCase()]
    const code = name[1].toLowerCase() === 'x' ? parseInt(name.slice(2), 16) : Number(name.slice(1))
    return code > 0 && code <= 0x10ffff ? xmlSafe(String.fromCodePoint(code)) : ''
  })
}

function inlineRuns(tokens: Token[] | undefined, style: pptxgen.default.TextPropsOptions = {}): Run[] {
  const runs: Run[] = []
  for (const token of tokens ?? []) {
    switch (token.type) {
      case 'strong':
        runs.push(...inlineRuns(token.tokens, { ...style, bold: true }))
        break
      case 'em':
        runs.push(...inlineRuns(token.tokens, { ...style, italic: true }))
        break
      case 'del':
        runs.push(...inlineRuns(token.tokens, { ...style, strike: 'sngStrike' }))
        break
      case 'codespan':
        runs.push({ text: decodeEntities(token.text), options: { ...style, fontFace: MONO, highlight: C.code } })
        break
      case 'link':
        // Only links that lead somewhere outside the deck; relative and empty ones stay plain text.
        runs.push(...inlineRuns(token.tokens, /^(https?:|mailto:)/i.test(token.href) ? { ...style, hyperlink: { url: token.href }, color: C.accent } : style))
        break
      case 'br':
        runs.push({ text: '\n', options: style })
        break
      case 'image':
        runs.push({ text: xmlSafe(token.text), options: style })
        break
      case 'html':
        break
      case 'text':
        if ('tokens' in token && token.tokens?.length) runs.push(...inlineRuns(token.tokens, style))
        else runs.push({ text: decodeEntities(token.text), options: style })
        break
      default:
        if ('text' in token && typeof token.text === 'string') runs.push({ text: decodeEntities(token.text), options: style })
    }
  }
  return runs
}

const runsText = (runs: Run[]) => runs.map((run) => run.text).join('')

/** One paragraph of a text box: its runs, paragraph options, and the font size for estimating height. */
type Paragraph = { runs: Run[]; options: pptxgen.default.TextPropsOptions; pt: number; indent: number }

type Block =
  | { kind: 'text'; paragraphs: Paragraph[] }
  | { kind: 'code'; text: string }
  | { kind: 'table'; header: Run[][]; rows: Run[][][] }
  | { kind: 'picture'; data: string; width: number; height: number }

function listParagraphs(list: Tokens.List, pt: number, level: number): Paragraph[] {
  const out: Paragraph[] = []
  list.items.forEach((item, at) => {
    const inline = item.tokens.filter((token) => token.type !== 'list')
    const runs = inline.flatMap((token) => ('tokens' in token && token.tokens ? inlineRuns(token.tokens) : inlineRuns([token])))
    const bullet: pptxgen.default.TextPropsOptions['bullet'] = list.ordered
      ? at === 0 && Number(list.start) > 1 ? { type: 'number', numberStartAt: Number(list.start) } : { type: 'number' }
      : { code: '2022' }
    const size = level > 0 ? Math.round(pt * 0.9) : pt
    out.push({ runs, options: { bullet, indentLevel: level, color: C.text }, pt: size, indent: 0.35 * (level + 1) })
    for (const nested of item.tokens.filter((token): token is Tokens.List => token.type === 'list')) {
      out.push(...listParagraphs(nested, pt, level + 1))
    }
  })
  return out
}

/** Picture size in pixels from the file header (PNG, GIF, JPEG, SVG); null when unknown. */
function pictureSize(data: Buffer, type: string): { width: number; height: number } | null {
  if (type === 'png' && data.length > 24) return { width: data.readUInt32BE(16), height: data.readUInt32BE(20) }
  if (type === 'gif' && data.length > 10) return { width: data.readUInt16LE(6), height: data.readUInt16LE(8) }
  if (type === 'jpeg') {
    let at = 2
    while (at + 9 < data.length) {
      if (data[at] !== 0xff) return null
      const marker = data[at + 1]
      if (marker >= 0xc0 && marker <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(marker)) {
        return { width: data.readUInt16BE(at + 7), height: data.readUInt16BE(at + 5) }
      }
      at += 2 + data.readUInt16BE(at + 2)
    }
    return null
  }
  if (type === 'svg') {
    const box = /viewBox="[\d.\s-]*?([\d.]+)\s+([\d.]+)"/.exec(data.toString('utf8', 0, 4000))
    return box ? { width: Number(box[1]), height: Number(box[2]) } : null
  }
  return null
}

const PICTURE_TYPES: Record<string, string> = { '.png': 'png', '.gif': 'gif', '.jpg': 'jpeg', '.jpeg': 'jpeg', '.svg': 'svg' }

/** pele themes diagrams with CSS variables, which PowerPoint's SVG renderer ignores: pin them to the light palette. */
function flattenDiagram(svg: string): string {
  const palette: Record<string, string> = {
    'pele-bg': '#ffffff', 'pele-fg': '#1f2328', 'pele-muted': '#59636e', 'pele-line': '#59636e', 'pele-surface': '#f6f8fa',
    'pele-surface-alt': '#eaeef2', 'pele-border': '#d1d9e0', 'pele-accent': '#2563eb'
  }
  const locals = new Map<string, string>()
  for (const match of svg.matchAll(/--(_\w+):var\(--(pele-[\w-]+),([^)]*)\)/g)) locals.set(match[1], palette[match[2]] ?? match[3])
  return svg
    .replace(/var\(--(_\w+)\)/g, (whole, name: string) => locals.get(name) ?? whole)
    .replace(/var\(--[\w-]+,([^()]*)\)/g, '$1')
    .replace(/--_\w+:[^;"]*;?/g, '')
}

function blocksOf(markdown: string, pt: number, read: ReadFile): Block[] {
  const blocks: Block[] = []
  let text: Paragraph[] = []
  const flush = () => {
    if (text.length) blocks.push({ kind: 'text', paragraphs: text })
    text = []
  }
  const picture = (data: Buffer, type: string) => {
    const size = pictureSize(data, type)
    if (!size || !size.width || !size.height) return
    flush()
    const mime = type === 'svg' ? 'image/svg+xml' : `image/${type}`
    blocks.push({ kind: 'picture', data: `${mime};base64,${data.toString('base64')}`, ...size })
  }
  for (const token of lexer(markdown)) {
    switch (token.type) {
      case 'heading':
        text.push({ runs: inlineRuns(token.tokens, { bold: true }), options: { color: token.depth >= 4 ? C.muted : C.text }, pt: token.depth <= 2 ? pt * 1.3 : pt * 1.12, indent: 0 })
        break
      case 'paragraph': {
        const images = (token.tokens ?? []).filter((inner) => inner.type === 'image') as Tokens.Image[]
        const onlyImages = images.length > 0 && (token.tokens ?? []).every((inner) => inner.type === 'image' || (inner.type === 'text' && !inner.text.trim()))
        if (onlyImages) {
          for (const image of images) {
            const type = PICTURE_TYPES[posix.extname(image.href.split(/[?#]/)[0]).toLowerCase()]
            const data = type ? read(image.href) : null
            if (data && type) picture(data, type)
          }
        } else {
          text.push({ runs: inlineRuns(token.tokens), options: { color: C.text }, pt, indent: 0 })
        }
        break
      }
      case 'list':
        text.push(...listParagraphs(token as Tokens.List, pt, 0))
        break
      case 'blockquote':
        for (const inner of (token as Tokens.Blockquote).tokens) {
          if (inner.type === 'paragraph') text.push({ runs: inlineRuns(inner.tokens, { italic: true }), options: { color: C.muted, indentLevel: 1 }, pt, indent: 0.4 })
        }
        break
      case 'code': {
        const language = (token.lang ?? '').trim().split(/\s+/)[0].toLowerCase()
        if (language === 'mermaid') {
          try {
            picture(Buffer.from(flattenDiagram(renderDiagram(token.text, { idPrefix: 'p-', fontFamily: FONT }).svg)), 'svg')
          } catch {
            flush()
            blocks.push({ kind: 'code', text: xmlSafe(token.text) })
          }
        } else {
          flush()
          blocks.push({ kind: 'code', text: xmlSafe(token.text) })
        }
        break
      }
      case 'table': {
        flush()
        const table = token as Tokens.Table
        blocks.push({ kind: 'table', header: table.header.map((cell) => inlineRuns(cell.tokens)), rows: table.rows.map((row) => row.map((cell) => inlineRuns(cell.tokens))) })
        break
      }
      case 'html':
        if (token.text.replace(/<[^>]*>/g, '').trim()) text.push({ runs: [{ text: decodeEntities(token.text.replace(/<[^>]*>/g, '').trim()) }], options: { color: C.text }, pt, indent: 0 })
        break
      default:
        break
    }
  }
  flush()
  return blocks
}

function blockHeight(block: Block, width: number, scale: number): number {
  switch (block.kind) {
    case 'text':
      return block.paragraphs.reduce((sum, p) => sum + estimateLines(runsText(p.runs), p.pt * scale, width - p.indent) * lineHeight(p.pt * scale) + (p.pt * scale * 0.45) / 72, 0.1)
    case 'code':
      return block.text.split('\n').length * lineHeight(SIZE.body * 0.68 * scale) + 0.3
    case 'table':
      return (block.rows.length + 1) * (lineHeight(SIZE.body * 0.8 * scale) + 0.16)
    case 'picture':
      return Math.min(H * 0.6, (width * block.height) / block.width)
  }
}

function drawBlocks(slide: pptxgen.default.Slide, blocks: Block[], box: Box, align: 'left' | 'center' = 'left', vcenter = false) {
  if (blocks.length === 0) return
  const gap = 0.12
  let scale = 1
  const total = () => blocks.reduce((sum, block) => sum + blockHeight(block, box.w, scale) + gap, -gap)
  while (total() > box.h && scale > 0.6) scale = Math.round((scale - 0.05) * 100) / 100
  // Pictures share what the text leaves over.
  const textHeight = blocks.filter((block) => block.kind !== 'picture').reduce((sum, block) => sum + blockHeight(block, box.w, scale) + gap, 0)
  const pictures = blocks.filter((block) => block.kind === 'picture').length
  const pictureRoom = pictures ? Math.max(1.2, (box.h - textHeight) / pictures) : 0
  let y = vcenter ? box.y + Math.max(0, (box.h - Math.min(total(), box.h)) / 2) : box.y
  for (const block of blocks) {
    let h = blockHeight(block, box.w, scale)
    if (block.kind === 'text') {
      const runs: pptxgen.default.TextProps[] = []
      block.paragraphs.forEach((p, at) => {
        const last = at === block.paragraphs.length - 1
        p.runs.forEach((run, i) => {
          runs.push({
            text: run.text,
            options: {
              ...p.options,
              ...run.options,
              fontSize: Math.round(p.pt * scale),
              align,
              paraSpaceAfter: Math.round(p.pt * scale * 0.45),
              breakLine: i === p.runs.length - 1 && !last
            }
          })
        })
        if (p.runs.length === 0) runs.push({ text: '', options: { breakLine: !last } })
      })
      slide.addText(runs, { x: box.x, y, w: box.w, h, valign: 'top', fontFace: FONT, margin: 0, fit: 'shrink' })
    } else if (block.kind === 'code') {
      slide.addText(block.text, {
        x: box.x, y, w: box.w, h, valign: 'top', fontFace: MONO, fontSize: Math.round(SIZE.body * 0.68 * scale), color: C.text,
        fill: { color: C.code }, margin: 0.12, fit: 'shrink'
      })
    } else if (block.kind === 'table') {
      const cell = (runs: Run[], header: boolean): pptxgen.default.TableCell => ({
        text: runs.length ? (runs as pptxgen.default.TextProps[]) : '',
        options: { bold: header || undefined, border: [{ type: 'none' }, { type: 'none' }, { pt: header ? 1.5 : 0.75, color: header ? C.accent : C.border }, { type: 'none' }] }
      })
      slide.addTable([block.header.map((runs) => cell(runs, true)), ...block.rows.map((row) => row.map((runs) => cell(runs, false)))], {
        x: box.x, y, w: box.w, fontFace: FONT, fontSize: Math.round(SIZE.body * 0.8 * scale), color: C.text, margin: [0.05, 0.12, 0.05, 0.12], valign: 'top'
      })
    } else {
      const maxH = Math.min(pictureRoom, box.y + box.h - y)
      let w = box.w
      h = (w * block.height) / block.width
      if (h > maxH) {
        h = maxH
        w = (h * block.width) / block.height
      }
      slide.addImage({ data: block.data, x: box.x + (box.w - w) / 2, y, w, h })
    }
    y += h + gap
  }
}

function drawTitle(slide: pptxgen.default.Slide, markdown: string, y = 0.55): number {
  const token = lexer(markdown).find((candidate) => candidate.type === 'heading') as Tokens.Heading | undefined
  if (!token) return y
  const runs = inlineRuns(token.tokens, { bold: true })
  const lines = estimateLines(runsText(runs), SIZE.title, W - 2 * MARGIN_X)
  const h = lines * lineHeight(SIZE.title)
  slide.addText(runs as pptxgen.default.TextProps[], { x: MARGIN_X, y, w: W - 2 * MARGIN_X, h, fontFace: FONT, fontSize: SIZE.title, color: C.text, valign: 'bottom', margin: 0, fit: 'shrink' })
  slide.addShape('rect', { x: MARGIN_X, y: y + h + 0.1, w: 0.8, h: 0.06, fill: { color: C.accent }, line: { type: 'none' } })
  return y + h + 0.45
}

/** The first heading of a slide body, taken off as its title. */
function splitTitle(markdown: string): { title: string | null; rest: string } {
  const match = /^\s*(#{1,2} [^\n]*)\n?/.exec(markdown)
  return match ? { title: match[1], rest: markdown.slice(match[0].length) } : { title: null, rest: markdown }
}

function notesText(markdown: string): string {
  return xmlSafe(markdown)
    .replace(/!?\[([^\]]*)\]\([^)]*\)/g, '$1')
    .replace(/^[ \t]*#+[ \t]+/gm, '')
    .replace(/(\*\*|__|`)/g, '')
    .trim()
}

function drawSlide(pptx: pptxgen.default, model: SlideModel, number: number, read: ReadFile) {
  const slide = pptx.addSlide()
  slide.background = { color: 'FFFFFF' }
  const fullWidth = { x: MARGIN_X, w: W - 2 * MARGIN_X }
  if (model.layout === 'cover') {
    slide.addShape('rect', { x: 0, y: 0, w: 0.22, h: H, fill: { color: C.accent }, line: { type: 'none' } })
    const blocks = blocksOf(model.body, SIZE.body, read)
    for (const block of blocks) {
      if (block.kind !== 'text') continue
      block.paragraphs.forEach((p, at) => {
        if (at === 0) p.pt = 44
        else if (p.runs.some((run) => run.options?.bold)) {
          p.pt = 24
          p.runs = p.runs.map((run) => ({ ...run, options: { ...run.options, bold: false, color: C.muted } }))
        } else p.options = { ...p.options, color: C.muted }
      })
    }
    drawBlocks(slide, blocks, { x: 1.2, y: 0.8, w: W - 2.4, h: H - 1.6 }, 'left', true)
  } else if (model.layout === 'section') {
    slide.background = { color: 'F3F7FF' }
    const blocks = blocksOf(model.body, SIZE.body, read)
    for (const block of blocks) if (block.kind === 'text') block.paragraphs.forEach((p, at) => (p.pt = at === 0 ? 40 : 22))
    drawBlocks(slide, blocks, { ...fullWidth, y: 0.8, h: H - 1.6 }, 'left', true)
  } else if (model.layout === 'center') {
    const blocks = blocksOf(model.body, SIZE.body, read)
    for (const block of blocks) if (block.kind === 'text') block.paragraphs.forEach((p) => { p.options = { ...p.options, indentLevel: undefined }; p.indent = 0; if (p.options.italic || p.runs.some((r) => r.options?.italic)) p.pt = 28 })
    drawBlocks(slide, blocks, { x: 1.5, y: 0.8, w: W - 3, h: H - 1.6 }, 'center', true)
  } else if (model.layout === 'columns' && model.columns) {
    const top = model.heading ? drawTitle(slide, model.heading) : 0.6
    const gap = 0.5
    const width = (W - 2 * MARGIN_X - gap * (model.columns.length - 1)) / model.columns.length
    model.columns.forEach((column, at) => {
      drawBlocks(slide, blocksOf(column, SIZE.body, read), { x: MARGIN_X + at * (width + gap), y: top, w: width, h: H - top - 0.7 })
    })
  } else {
    const { title, rest } = splitTitle(model.body)
    const top = title ? drawTitle(slide, title) : 0.6
    drawBlocks(slide, blocksOf(rest, SIZE.body, read), { ...fullWidth, y: top, h: H - top - 0.7 })
  }
  if (model.layout !== 'cover') {
    slide.addText(String(number), { x: W - 1.2, y: H - 0.55, w: 0.8, h: 0.3, align: 'right', fontFace: FONT, fontSize: 11, color: C.muted, margin: 0 })
  }
  if (model.notes) slide.addNotes(notesText(model.notes))
}

/** Builds the .pptx. `read` resolves an image path as written in the deck (relative to the entry). */
export async function deckToPptx(source: string, title: string, entry: string, readFile: (path: string) => Buffer | null): Promise<Buffer> {
  const pptx = new PptxGenJS()
  pptx.layout = 'LAYOUT_WIDE'
  pptx.title = xmlSafe(title)
  pptx.theme = { headFontFace: FONT, bodyFontFace: FONT }
  const read: ReadFile = (href) => {
    if (/^[a-z][a-z0-9+.-]*:/i.test(href) || href.startsWith('/')) return null
    let path: string
    try {
      path = posix.normalize(posix.join(posix.dirname(entry), decodeURI(href.split(/[?#]/)[0])))
    } catch {
      return null
    }
    return path.startsWith('..') ? null : readFile(path)
  }
  splitSlides(source).forEach((slide, index) => drawSlide(pptx, analyzeSlide(slide, index), index + 1, read))
  return (await pptx.write({ outputType: 'nodebuffer' })) as Buffer
}

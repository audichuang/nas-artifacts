import { lexer } from 'marked'
import { MARKDOWN_CSS_PATH, escapeHtml, injectBridge, markdownParser } from './content-render.ts'

/** A deck is a Markdown file named *.slides.md: one slide per `---` line, speaker notes after `???`. */
export const isSlidesPath = (path: string): boolean => /\.slides\.md$/i.test(path)

export const LAYOUTS = ['cover', 'section', 'center', 'columns', 'default'] as const
export type Layout = (typeof LAYOUTS)[number]

export type SlideSource = { body: string; notes: string }

const FENCE = /^ {0,3}(`{3,}|~{3,})/

/** Splits on lines matching `separator`, except inside fenced code. */
function splitLines(text: string, separator: RegExp, limit = Infinity): string[] {
  const parts: string[][] = [[]]
  let fence: string | null = null
  for (const line of text.split('\n')) {
    const marker = FENCE.exec(line)?.[1]
    if (fence) {
      if (marker && marker[0] === fence[0] && marker.length >= fence.length && line.trim() === marker) fence = null
    } else if (marker) {
      fence = marker
    } else if (parts.length < limit && separator.test(line)) {
      parts.push([])
      continue
    }
    parts[parts.length - 1].push(line)
  }
  return parts.map((lines) => lines.join('\n'))
}

export function splitSlides(source: string): SlideSource[] {
  return splitLines(source.replace(/\r\n?/g, '\n'), /^---\s*$/)
    .map((chunk) => {
      const [body, notes = ''] = splitLines(chunk, /^\?\?\?\s*$/, 2)
      return { body: body.trim(), notes: notes.trim() }
    })
    .filter((slide, index, all) => slide.body || slide.notes || all.length === 1)
}

const LAYOUT_COMMENT = /<!--\s*layout:\s*([a-z-]+)\s*-->/i

/** What a slide is made of: its layout, and for columns the heading above them and each column's Markdown. */
export type SlideModel = { layout: Layout; body: string; heading: string | null; columns: string[] | null; notes: string }

export function analyzeSlide(slide: SlideSource, index: number): SlideModel {
  const requested = LAYOUT_COMMENT.exec(slide.body)?.[1]?.toLowerCase()
  const body = slide.body.replace(LAYOUT_COMMENT, '').trim()
  const parts = splitLines(body, /^\|\|\|\s*$/)
  let layout: Layout
  let heading: string | null = null
  let columns: string[] | null = null
  if (parts.length > 1) {
    // The heading above the first column spans them all.
    const match = /^\s*(#{1,3} [^\n]*)\n?/.exec(parts[0])
    if (match) {
      heading = match[1]
      parts[0] = parts[0].slice(match[0].length)
    }
    columns = parts.filter((column, at) => at > 0 || column.trim())
    layout = 'columns'
  } else {
    const blocks = lexer(body).filter((token) => token.type !== 'space')
    if (index === 0 && blocks[0]?.type === 'heading' && blocks[0].depth === 1) layout = 'cover'
    else if (blocks.length > 0 && blocks.every((token) => token.type === 'heading')) layout = 'section'
    else layout = 'default'
  }
  if (requested && (LAYOUTS as readonly string[]).includes(requested) && requested !== 'columns') layout = requested as Layout
  return { layout, body, heading, columns, notes: slide.notes }
}

function renderSlide(slide: SlideSource, index: number): string {
  const model = analyzeSlide(slide, index)
  const marked = markdownParser({ anchors: false, diagramPrefix: `s${index}-` })
  const md = (text: string) => marked.parse(text, { async: false })
  const html = model.columns
    ? `${model.heading ? md(model.heading) : ''}<div class="columns" style="--cols:${model.columns.length}">${model.columns
        .map((column) => `<div class="column">${md(column)}</div>`)
        .join('')}</div>`
    : md(model.body)
  const notes = model.notes ? `<aside class="slide-notes" hidden>${md(model.notes)}</aside>` : ''
  const number = index + 1
  return `<section class="slide layout-${model.layout}" id="slide-${number}" data-number="${number}" aria-label="第 ${number} 頁"><div class="slide-inner">${html}</div>${notes}</section>\n`
}

/** The slides alone, as the deck page holds them; the editor's live preview swaps exactly this. */
export function renderSlides(source: string): { html: string; count: number } {
  const slides = splitSlides(source)
  return { html: slides.map(renderSlide).join(''), count: slides.length }
}

export type DeckAssets = { css: string; js: string }

export function renderDeck(source: string, title: string, assets: DeckAssets): string {
  return deckPage(renderSlides(source).html, title, assets)
}

/** The deck page around a set of slides; `head` adds what a kind of deck needs on top (a PowerPoint file's slide size). */
export function deckPage(html: string, title: string, assets: DeckAssets, head = ''): string {
  return injectBridge(`<!doctype html><html lang="zh-Hant"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml(title)}</title>
<link rel="stylesheet" href="${MARKDOWN_CSS_PATH}">
<style>${assets.css}</style>${head}
</head><body class="deck">
<main class="deck-stage" id="deck">${html}</main>
<aside class="deck-notes" id="deck-notes" aria-label="講稿" hidden></aside>
<nav class="deck-bar" aria-label="簡報">
<button type="button" data-deck="prev" aria-label="上一頁" title="上一頁(←)">‹</button>
<span class="deck-counter" id="deck-counter" aria-live="polite"></span>
<button type="button" data-deck="next" aria-label="下一頁" title="下一頁(→)">›</button>
<span class="deck-gap"></span>
<button type="button" data-deck="overview" aria-pressed="false" title="總覽(O)">總覽</button>
<button type="button" data-deck="notes" aria-pressed="false" title="講稿(N)">講稿</button>
<button type="button" data-deck="print" title="列印或存成 PDF">PDF</button>
<button type="button" data-deck="fullscreen" aria-label="全螢幕" title="全螢幕(F)">⛶</button>
</nav>
<script>${assets.js}</script>
</body></html>`)
}

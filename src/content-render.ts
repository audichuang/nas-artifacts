import { readFileSync } from 'node:fs'
import hljs from 'highlight.js/lib/common'
import { Marked, type Tokens } from 'marked'
import { render as renderDiagram } from 'pele'

export const BRIDGE_PATH = '/_/bridge.js'
export const MARKDOWN_CSS_PATH = '/_/markdown.css'

const packageVersion = (name: string) =>
  (JSON.parse(readFileSync(new URL(`../node_modules/${name}/package.json`, import.meta.url), 'utf8')) as { version: string }).version

/**
 * Part of every rendered page's ETag, so pages browsers already hold revalidate when the output changes.
 * Library upgrades count themselves; bump the leading tags when you change what renderMarkdown (md) or
 * the Office renderers in pptx.ts, docx.ts and xlsx.ts emit.
 */
export const RENDER_VERSION = ['md-2', 'pptx-1', 'docx-1', 'xlsx-1', ...['marked', 'highlight.js', 'pele'].map((name) => `${name}@${packageVersion(name)}`)].join(' ')

export function escapeHtml(value: string): string {
  return value.replace(/[&<>"']/g, (char) => `&#${char.charCodeAt(0)};`)
}

/** Puts the capability bridge first in <head>, so page scripts can use window.artifact immediately. */
export function injectBridge(html: string): string {
  const tag = `<script src="${BRIDGE_PATH}"></script>`
  const head = /<head(\s[^>]*)?>/i.exec(html)
  if (head) {
    const at = head.index + head[0].length
    return html.slice(0, at) + tag + html.slice(at)
  }
  const htmlTag = /<html(\s[^>]*)?>/i.exec(html)
  if (htmlTag) {
    const at = htmlTag.index + htmlTag[0].length
    return `${html.slice(0, at)}<head>${tag}</head>${html.slice(at)}`
  }
  // A fragment: give it the skeleton a full page would have declared itself.
  return `<!doctype html><html><head>${tag}<meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"></head><body>${html}</body></html>`
}

/** GitHub-style: lower case, punctuation dropped, letters of any script kept, spaces to dashes. */
function slugify(text: string): string {
  return text
    .toLowerCase()
    .trim()
    .replace(/[^\p{L}\p{M}\p{N}\s_-]/gu, '')
    .replace(/\s+/g, '-')
}

const FONT = 'ui-sans-serif, system-ui, -apple-system, "Segoe UI", "PingFang TC", "Noto Sans TC", sans-serif'

/**
 * Mermaid is drawn here, on the server (pele): the page gets plain SVG coloured by CSS variables, so it
 * needs no script and follows dark mode. A diagram that does not parse keeps its source on the page,
 * and an inline throw puts the reason in the viewer's error log.
 */
function diagram(source: string, idPrefix: string): string {
  try {
    const { svg } = renderDiagram(source, { idPrefix, fontFamily: FONT, responsive: true })
    return `<figure class="diagram">${svg}</figure>\n`
  } catch (error) {
    const message = `Mermaid: ${error instanceof Error ? error.message : String(error)}`
    const thrown = JSON.stringify(message).replace(/</g, '\\u003c')
    return `<div class="diagram-failed"><p>${escapeHtml(message)}</p><pre><code>${escapeHtml(source)}\n</code></pre></div><script>throw new Error(${thrown})</script>\n`
  }
}

export type TocEntry = { depth: number; id: string; html: string }

/**
 * The Marked setup every rendered page shares: code highlighted, mermaid drawn. Pages get heading anchors
 * and a table of contents (collected into `toc`); slides do not. Diagram ids start with `diagramPrefix`,
 * so several renders can share one document.
 */
export function markdownParser(options: { anchors: boolean; toc?: TocEntry[]; diagramPrefix?: string }): Marked {
  const slugs = new Map<string, number>()
  let diagrams = 0
  return new Marked({
    gfm: true,
    renderer: {
      heading({ tokens, depth, text }: Tokens.Heading) {
        const html = this.parser.parseInline(tokens)
        if (!options.anchors) return `<h${depth}>${html}</h${depth}>\n`
        const base = slugify(text.replace(/<[^>]*>/g, '')) || 'section'
        const seen = slugs.get(base) ?? 0
        slugs.set(base, seen + 1)
        const id = seen === 0 ? base : `${base}-${seen}`
        if (depth === 2 || depth === 3) options.toc?.push({ depth, id, html })
        return `<h${depth} id="${escapeHtml(id)}"><a class="anchor" href="#${escapeHtml(id)}" aria-hidden="true" tabindex="-1">#</a>${html}</h${depth}>\n`
      },
      code({ text, lang }: Tokens.Code) {
        const language = (lang ?? '').trim().split(/\s+/)[0].toLowerCase()
        if (language === 'mermaid') return diagram(text, `${options.diagramPrefix ?? ''}d${diagrams++}-`)
        if (language && hljs.getLanguage(language)) {
          const highlighted = hljs.highlight(text, { language, ignoreIllegals: true }).value
          return `<pre><code class="hljs language-${escapeHtml(language)}">${highlighted}\n</code></pre>\n`
        }
        return `<pre><code>${escapeHtml(text)}\n</code></pre>\n`
      }
    }
  })
}

export function renderMarkdown(markdown: string, title: string): string {
  const toc: TocEntry[] = []
  const marked = markdownParser({ anchors: true, toc })
  const body = marked.parse(markdown, { async: false })
  const list = `<ol>${toc
    .map((entry) => `<li class="toc-${entry.depth}"><a href="#${escapeHtml(entry.id)}">${entry.html.replace(/<a\b[^>]*>|<\/a>/g, '')}</a></li>`)
    .join('')}</ol>`
  // Wide screens keep the contents beside the text; narrow ones get it folded under the title.
  const nav = toc.length >= 3 ? `<nav class="toc" aria-label="目錄"><p class="toc-title">目錄</p>${list}</nav>` : ''
  const folded = toc.length >= 3 ? `<details class="toc-inline"><summary>目錄 · ${toc.length} 節</summary>${list}</details>\n` : ''
  const heading = /^<h1\b[^>]*>[\s\S]*?<\/h1>\n?/.exec(body)?.[0] ?? ''
  const main = heading + folded + body.slice(heading.length)
  return injectBridge(`<!doctype html><html lang="zh-Hant"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml(title)}</title>
<link rel="stylesheet" href="${MARKDOWN_CSS_PATH}">
</head><body><div class="markdown-layout${nav ? ' has-toc' : ''}">${nav}<main class="markdown-body">${main}</main></div></body></html>`)
}

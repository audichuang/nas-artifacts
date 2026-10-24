import { marked } from 'marked'

export const BRIDGE_PATH = '/_/bridge.js'
export const MARKDOWN_CSS_PATH = '/_/markdown.css'

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

export function renderMarkdown(markdown: string, title: string): string {
  const body = marked.parse(markdown, { async: false, gfm: true })
  return injectBridge(`<!doctype html><html lang="zh-Hant"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml(title)}</title>
<link rel="stylesheet" href="${MARKDOWN_CSS_PATH}">
</head><body><main class="markdown-body">${body}</main></body></html>`)
}

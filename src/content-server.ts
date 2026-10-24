import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { Hono } from 'hono'
import type { ArtifactStore, FileEntry } from './artifact-store.ts'
import { ArtifactError } from './artifact-store.ts'
import type { Config } from './config.ts'
import { BRIDGE_PATH, MARKDOWN_CSS_PATH, injectBridge, renderMarkdown } from './content-render.ts'
import { decodeRequestPath, isHtmlPath, isMarkdownPath, isServableType } from './file-paths.ts'

// Same CDN allowlist as Claude's artifacts; everything else must be inline or published with the page.
const SCRIPT_HOSTS = 'https://cdnjs.cloudflare.com https://cdn.jsdelivr.net https://unpkg.com https://cdn.tailwindcss.com https://code.jquery.com'

/**
 * `sandbox` without allow-same-origin gives every page an opaque origin: no cookies or storage on
 * this host, and no same-origin access to the app origin's API. connect-src 'none' keeps pages
 * from sending their data anywhere; page capabilities go through the viewer's postMessage bridge.
 * Own-file sources are scoped to the artifact's path so one artifact cannot load another's files.
 */
export function contentSecurityPolicy(appOrigin: string, contentOrigin: string, artifactId?: string): string {
  const own = artifactId ? `${contentOrigin}/a/${artifactId}/` : ''
  const sources = (...list: string[]) => list.filter(Boolean).join(' ')
  return [
    'sandbox allow-scripts allow-forms allow-modals allow-downloads',
    "default-src 'none'",
    `script-src ${sources(own, `${contentOrigin}${BRIDGE_PATH}`)} 'unsafe-inline' 'unsafe-eval' blob: ${SCRIPT_HOSTS}`,
    `style-src ${sources(own, `${contentOrigin}${MARKDOWN_CSS_PATH}`)} 'unsafe-inline' https://fonts.googleapis.com ${SCRIPT_HOSTS}`,
    `font-src ${sources(own)} data: https://fonts.gstatic.com ${SCRIPT_HOSTS}`,
    `img-src ${sources(own)} data: blob:`,
    `media-src ${sources(own)} data: blob:`,
    `worker-src ${sources(own)} blob:`,
    "connect-src 'none'",
    "form-action 'none'",
    `frame-ancestors ${appOrigin}`,
    "base-uri 'none'"
  ].join('; ')
}

const STATIC_ROOT = new URL('../content/', import.meta.url)
const staticFile = (name: string): string => readFileSync(new URL(name, STATIC_ROOT), 'utf8')

export function createContentServer(store: ArtifactStore, config: Config): Hono {
  const app = new Hono()
  const bridge = staticFile('bridge.js').replace('__APP_ORIGIN__', JSON.stringify(config.appOrigin))
  const markdownCss = staticFile('markdown.css')
  const renderTag = createHash('sha256').update(bridge).digest('hex').slice(0, 16)

  app.use('*', async (c, next) => {
    await next()
    const id = /^\/a\/([A-Za-z0-9_-]+)\//.exec(new URL(c.req.url).pathname)?.[1]
    c.header('Content-Security-Policy', contentSecurityPolicy(config.appOrigin, config.contentOrigin, id))
    c.header('X-Content-Type-Options', 'nosniff')
    c.header('Referrer-Policy', 'no-referrer')
  })

  app.get(BRIDGE_PATH, (c) => {
    c.header('Cache-Control', 'no-cache')
    return c.body(bridge, 200, { 'Content-Type': 'text/javascript; charset=utf-8' })
  })
  app.get(MARKDOWN_CSS_PATH, (c) => {
    c.header('Cache-Control', 'no-cache')
    return c.body(markdownCss, 200, { 'Content-Type': 'text/css; charset=utf-8' })
  })

  const etagFor = (entry: FileEntry, path: string, title: string): string => {
    if (isMarkdownPath(path)) {
      return `"${entry.sha256}-${renderTag}-${createHash('sha256').update(title).digest('hex').slice(0, 16)}"`
    }
    return isHtmlPath(path) ? `"${entry.sha256}-${renderTag}"` : `"${entry.sha256}"`
  }

  const matchesEtag = (header: string | undefined, etag: string): boolean =>
    header !== undefined && header.split(',').some((candidate) => candidate.trim().replace(/^W\//, '') === etag || candidate.trim() === '*')

  const serveFile = (read: () => Buffer, entry: FileEntry, path: string, title: string, immutable: boolean, ifNoneMatch: string | undefined) => {
    const etag = etagFor(entry, path, title)
    // Rendered pages embed the bridge, which changes across releases, so they revalidate (a cheap 304).
    const rendered = isMarkdownPath(path) || isHtmlPath(path)
    const headers: Record<string, string> = {
      'Cache-Control': immutable && !rendered ? 'public, max-age=31536000, immutable' : 'no-cache',
      ETag: etag
    }
    if (rendered) {
      headers.Vary = 'Sec-Fetch-Dest'
    }
    if (!isMarkdownPath(path) && !isHtmlPath(path) && !isServableType(entry.type)) {
      throw new ArtifactError(404, 'not_servable', `${path} is not a web media type.`)
    }
    if (matchesEtag(ifNoneMatch, etag)) {
      return new Response(null, { status: 304, headers })
    }
    const body = read()
    if (isMarkdownPath(path)) {
      return new Response(renderMarkdown(body.toString('utf8'), title), {
        headers: { ...headers, 'Content-Type': 'text/html; charset=utf-8' }
      })
    }
    if (isHtmlPath(path)) {
      return new Response(injectBridge(body.toString('utf8')), {
        headers: { ...headers, 'Content-Type': entry.type }
      })
    }
    return new Response(new Uint8Array(body), { headers: { ...headers, 'Content-Type': entry.type } })
  }

  const versionOf = (raw: string): number | null => {
    const version = Number(raw)
    return Number.isInteger(version) && version >= 1 ? version : null
  }
  const encodePath = (path: string): string => path.split('/').map(encodeURIComponent).join('/')

  // Relative URLs in a nested entry only resolve against its own directory.
  app.get('/a/:id/v/:version', (c) => {
    const id = c.req.param('id')
    const version = versionOf(c.req.param('version'))
    if (!version) {
      return c.text('Not found', 404)
    }
    const row = store.requireVersion(id, version)
    return c.redirect(`/a/${id}/v/${version}/${row.entry.includes('/') ? encodePath(row.entry) : ''}`, 302)
  })

  app.get('/a/:id/v/:version/*', (c) => {
    const id = c.req.param('id')
    const version = versionOf(c.req.param('version'))
    if (!version) {
      return c.text('Not found', 404)
    }
    const prefix = `/a/${id}/v/${version}/`
    const encoded = new URL(c.req.url).pathname.slice(prefix.length)
    const row = store.requireVersion(id, version)
    if (encoded === '') {
      return row.entry.includes('/')
        ? c.redirect(`${prefix}${encodePath(row.entry)}`, 302)
        : serveEntry(row.entry)
    }
    return serveEntry(decodeRequestPath(encoded))

    function serveEntry(path: string | null) {
      const entry = path ? row.files[path] : undefined
      if (!path || !entry) {
        return c.text('Not found', 404)
      }
      // A rendered page opened as its own document has no frame-src to confine it, so it goes back into the viewer.
      if ((isHtmlPath(path) || isMarkdownPath(path)) && c.req.header('sec-fetch-dest') === 'document') {
        const query = new URLSearchParams({ v: String(version) })
        if (path !== row.entry) {
          query.set('path', path)
        }
        c.header('Vary', 'Sec-Fetch-Dest')
        return c.redirect(`${config.appOrigin}/a/${id}?${query}`, 302)
      }
      return serveFile(() => store.blobs.read(entry.sha256), entry, path, row.title, true, c.req.header('if-none-match'))
    }
  })

  app.get('/a/:id/assets/:assetId/:name?', (c) => {
    const asset = store.getAsset(c.req.param('id'), c.req.param('assetId'))
    const entry: FileEntry = { sha256: asset.sha256, size: asset.size, type: asset.type }
    if (isHtmlPath(asset.name) || isMarkdownPath(asset.name)) {
      return c.body(new Uint8Array(store.blobs.read(asset.sha256)), 200, { 'Content-Type': 'text/plain; charset=utf-8' })
    }
    return serveFile(() => store.blobs.read(asset.sha256), entry, asset.name, asset.name, true, c.req.header('if-none-match'))
  })

  app.get('/healthz', (c) => c.text('ok'))

  app.onError((error, c) => {
    if (error instanceof ArtifactError) {
      return c.text(error.message, error.status)
    }
    console.error(error)
    return c.text('Internal error', 500)
  })
  app.notFound((c) => c.text('Not found', 404))
  return app
}

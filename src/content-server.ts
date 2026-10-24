import { readFileSync } from 'node:fs'
import { Hono } from 'hono'
import type { ArtifactStore, FileEntry } from './artifact-store.ts'
import { ArtifactError } from './artifact-store.ts'
import type { Config } from './config.ts'
import { BRIDGE_PATH, MARKDOWN_CSS_PATH, injectBridge, renderMarkdown } from './content-render.ts'
import { isHtmlPath, isMarkdownPath, isServableType, normalizeFilePath } from './file-paths.ts'

// Same CDN allowlist as Claude's artifacts; everything else must be inline or published with the page.
const SCRIPT_HOSTS = 'https://cdnjs.cloudflare.com https://cdn.jsdelivr.net https://unpkg.com https://cdn.tailwindcss.com https://code.jquery.com'

/**
 * `sandbox` without allow-same-origin gives every page an opaque origin: no cookies or storage on
 * this host, and no same-origin access to the app origin's API. connect-src 'none' keeps pages
 * from sending their data anywhere; page capabilities go through the viewer's postMessage bridge.
 */
export function contentSecurityPolicy(appOrigin: string): string {
  return [
    'sandbox allow-scripts allow-forms allow-popups allow-popups-to-escape-sandbox allow-modals allow-downloads',
    "default-src 'none'",
    `script-src 'self' 'unsafe-inline' 'unsafe-eval' blob: ${SCRIPT_HOSTS}`,
    `style-src 'self' 'unsafe-inline' https://fonts.googleapis.com ${SCRIPT_HOSTS}`,
    `font-src 'self' data: https://fonts.gstatic.com ${SCRIPT_HOSTS}`,
    "img-src 'self' data: blob:",
    "media-src 'self' data: blob:",
    "worker-src 'self' blob:",
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
  const csp = contentSecurityPolicy(config.appOrigin)
  const bridge = staticFile('bridge.js').replace('__APP_ORIGIN__', JSON.stringify(config.appOrigin))
  const markdownCss = staticFile('markdown.css')

  app.use('*', async (c, next) => {
    await next()
    c.header('Content-Security-Policy', csp)
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

  const serveFile = (body: Buffer, entry: FileEntry, path: string, title: string, immutable: boolean) => {
    const headers: Record<string, string> = {
      'Cache-Control': immutable ? 'public, max-age=31536000, immutable' : 'no-cache',
      ETag: `"${entry.sha256}"`
    }
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
    if (!isServableType(entry.type)) {
      throw new ArtifactError(404, 'not_servable', `${path} is not a web media type.`)
    }
    return new Response(new Uint8Array(body), { headers: { ...headers, 'Content-Type': entry.type } })
  }

  app.get('/a/:id/v/:version/*', (c) => {
    const id = c.req.param('id')
    const version = Number(c.req.param('version'))
    if (!Number.isInteger(version) || version < 1) {
      return c.text('Not found', 404)
    }
    const prefix = `/a/${id}/v/${version}/`
    const rawPath = decodeURIComponent(new URL(c.req.url).pathname.slice(prefix.length))
    const row = store.requireVersion(id, version)
    const path = rawPath === '' ? row.entry : normalizeFilePath(rawPath)
    const entry = path ? row.files[path] : undefined
    if (!path || !entry) {
      return c.text('Not found', 404)
    }
    return serveFile(store.blobs.read(entry.sha256), entry, path, row.title, true)
  })

  app.get('/a/:id/assets/:assetId/:name?', (c) => {
    const asset = store.getAsset(c.req.param('id'), c.req.param('assetId'))
    const entry: FileEntry = { sha256: asset.sha256, size: asset.size, type: asset.type }
    if (isHtmlPath(asset.name) || isMarkdownPath(asset.name)) {
      return c.body(new Uint8Array(store.blobs.read(asset.sha256)), 200, { 'Content-Type': 'text/plain; charset=utf-8' })
    }
    return serveFile(store.blobs.read(asset.sha256), entry, asset.name, asset.name, true)
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

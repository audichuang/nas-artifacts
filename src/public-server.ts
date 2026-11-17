import { readFileSync } from 'node:fs'
import { Hono } from 'hono'
import { ArtifactError, type ArtifactStore } from './artifact-store.ts'
import type { Config } from './config.ts'
import { escapeHtml } from './content-render.ts'
import { decodeRequestPath, isOfficePath } from './file-paths.ts'

/**
 * The internet-facing app: a read-only viewer for artifacts marked public, and their original files.
 * Everything else (gallery, list, search, comments, errors, diff, export, every write) is absent, not
 * guarded: a route that is not here cannot leak. The page runs no script of its own.
 */
export function createPublicAppServer(store: ArtifactStore, config: Config): Hono {
  const pair = config.public
  if (!pair) throw new Error('createPublicAppServer needs config.public.')
  const app = new Hono()
  const staticRoot = new URL('../public/', import.meta.url)

  app.use('*', async (c, next) => {
    await next()
    c.header('Cache-Control', 'private, no-cache')
    c.header('X-Robots-Tag', 'noindex, nofollow')
    c.header('X-Content-Type-Options', 'nosniff')
    c.header('Referrer-Policy', 'no-referrer')
    c.header('X-Frame-Options', 'DENY')
    if (!c.res.headers.has('Content-Security-Policy')) {
      c.header('Content-Security-Policy', "default-src 'none'; frame-ancestors 'none'; base-uri 'none'; form-action 'none'")
    }
  })

  // Only an artifact the publisher made public, and the version and file it actually has.
  const publicVersion = (id: string, raw: string | undefined) => {
    if (!store.isPublic(id)) throw new ArtifactError(404, 'not_found', 'Not found.')
    const summary = store.requireSummary(id)
    const version = raw === undefined ? summary.currentVersion : Number(raw)
    if (!Number.isInteger(version) || version < 1) throw new ArtifactError(404, 'not_found', 'Not found.')
    return { summary, row: store.requireVersion(id, version) }
  }

  app.get('/a/:id', (c) => {
    const id = c.req.param('id')
    const { summary, row } = publicVersion(id, c.req.query('v'))
    const versions = store.detail(id).versions.map((entry) => ({ version: entry.version, date: entry.createdAt.slice(0, 10) }))
    // A file of this version the reader was on (sent here by the content origin's redirect); anything else is ignored.
    const requested = c.req.query('path')
    const path = requested !== undefined && Object.hasOwn(row.files, requested) ? requested : row.entry
    const encode = (value: string) => value.split('/').map(encodeURIComponent).join('/')
    const frameSrc = `${pair.contentOrigin}/a/${id}/v/${row.version}/${path === row.entry && !path.includes('/') ? '' : encode(path)}`
    const original = /\.(md|markdown)$/i.test(row.entry) && !/\.slides\.md$/i.test(row.entry) ? 'MD' : isOfficePath(row.entry) ? row.entry.split('.').pop()!.toUpperCase() : null
    const current = versions.find((entry) => entry.version === row.version)
    const list = versions
      .slice()
      .reverse()
      .map((entry) => `<li><a href="/a/${id}?v=${entry.version}"${entry.version === row.version ? ' aria-current="page"' : ''}>v${entry.version} · ${entry.date}${entry.version === summary.currentVersion ? ' · 最新' : ''}</a></li>`)
      .join('')
    c.header('Content-Security-Policy', `default-src 'none'; style-src 'self'; img-src 'self'; frame-src ${pair.contentOrigin}/a/${id}/; script-src 'none'; frame-ancestors 'none'; base-uri 'none'; form-action 'none'`)
    return c.html(`<!doctype html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex, nofollow">
<title>${escapeHtml(row.title)}</title>
<link rel="icon" href="/static/favicon.svg">
<link rel="stylesheet" href="/static/app.css">
</head>
<body class="public-view">
<header class="topbar public-bar">
  <div class="viewer-title">
    <span class="artifact-icon" aria-hidden="true">${escapeHtml(summary.icon ?? '')}</span>
    <h1>${escapeHtml(row.title)}</h1>
  </div>
  <details class="public-versions">
    <summary class="button">v${row.version} · ${current?.date ?? ''}</summary>
    <ol>${list}</ol>
  </details>
  ${original ? `<a class="button" href="/api/artifacts/${id}/files/${encode(row.entry)}?v=${row.version}" download>${original}</a>` : ''}
</header>
<iframe class="public-frame" title="${escapeHtml(row.title)}" src="${escapeHtml(frameSrc)}"
  sandbox="allow-scripts allow-forms allow-modals allow-downloads" allow="fullscreen"></iframe>
</body>
</html>`)
  })

  app.get('/api/artifacts/:id/files/*', (c) => {
    const id = c.req.param('id')
    const { row } = publicVersion(id, c.req.query('v'))
    const path = decodeRequestPath(new URL(c.req.url).pathname.slice(`/api/artifacts/${id}/files/`.length))
    const entry = path ? row.files[path] : undefined
    if (!entry) throw new ArtifactError(404, 'not_found', 'Not found.')
    return c.body(new Uint8Array(store.blobs.read(entry.sha256)), 200, { 'Content-Type': 'application/octet-stream', 'Content-Disposition': 'attachment' })
  })

  const STATIC: Record<string, string> = { 'app.css': 'text/css; charset=utf-8', 'favicon.svg': 'image/svg+xml' }
  app.get('/static/:name', (c) => {
    const type = STATIC[c.req.param('name')]
    if (!type) return c.text('Not found', 404)
    return c.body(new Uint8Array(readFileSync(new URL(c.req.param('name'), staticRoot))), 200, { 'Content-Type': type })
  })

  app.get('/healthz', (c) => c.text('ok'))

  app.onError((error, c) => {
    if (error instanceof ArtifactError && error.status < 500) return c.text('Not found', 404)
    console.error(error)
    return c.text('Internal error', 500)
  })
  app.notFound((c) => c.text('Not found', 404))
  return app
}

/** The config the public content listener runs with: the same server, its own pair of origins. */
export function publicContentConfig(config: Config): Config {
  if (!config.public) throw new Error('publicContentConfig needs config.public.')
  return { ...config, appOrigin: config.public.appOrigin, contentOrigin: config.public.contentOrigin }
}

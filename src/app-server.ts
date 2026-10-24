import { createHash, timingSafeEqual } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { Hono, type Context } from 'hono'
import { bodyLimit } from 'hono/body-limit'
import { streamSSE } from 'hono/streaming'
import {
  ArtifactError,
  LIMITS,
  type ArtifactStore,
  type ArtifactSummary,
  type Capability,
  type IncomingFile
} from './artifact-store.ts'
import type { Config } from './config.ts'
import { decodeRequestPath, normalizeFilePath } from './file-paths.ts'
import { galleryPage, viewerPage } from './pages.ts'
import { isSlidesPath, renderSlides } from './slides.ts'
import { deckToPptx } from './slides-pptx.ts'

/** The most a deck's source may be when edited from the viewer. */
const SLIDES_SOURCE_BYTES = 1024 * 1024

/** Mutations from the viewer must carry this header: it forces a CORS preflight, which no other origin passes. */
export const VIEWER_HEADER = 'x-artifact-client'

type Caller = 'publisher' | 'viewer'

function digest(value: string): Buffer {
  return createHash('sha256').update(value).digest()
}

function parseFiles(value: unknown): Record<string, IncomingFile> {
  if (value === undefined) {
    return {}
  }
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new ArtifactError(400, 'invalid_files', 'files is a map of path → {content, encoding} or null.')
  }
  const files: Record<string, IncomingFile> = {}
  for (const [path, file] of Object.entries(value)) {
    if (file === null) {
      files[path] = null
      continue
    }
    const { content, encoding = 'utf8' } = (file ?? {}) as { content?: unknown; encoding?: unknown }
    if (typeof content !== 'string' || (encoding !== 'utf8' && encoding !== 'base64')) {
      throw new ArtifactError(400, 'invalid_files', `${path}: content must be a string with encoding utf8 or base64.`)
    }
    files[path] = { content: Buffer.from(content, encoding) }
  }
  return files
}

function optionalStringArray(value: unknown, field: string): string[] | undefined {
  if (value === undefined) {
    return undefined
  }
  if (!Array.isArray(value) || value.some((item) => typeof item !== 'string')) {
    throw new ArtifactError(400, 'invalid_field', `${field} must be an array of strings.`)
  }
  return value
}

async function jsonBody(c: Context): Promise<Record<string, unknown>> {
  const body: unknown = await c.req.json().catch(() => {
    throw new ArtifactError(400, 'invalid_json', 'The request body must be JSON.')
  })
  if (!body || typeof body !== 'object' || Array.isArray(body)) {
    throw new ArtifactError(400, 'invalid_json', 'The request body must be a JSON object.')
  }
  return body as Record<string, unknown>
}

export function createAppServer(store: ArtifactStore, config: Config): Hono {
  const app = new Hono()
  const tokenDigests = config.tokens.map(digest)
  const staticRoot = new URL('../public/', import.meta.url)

  const urls = (summary: ArtifactSummary) => ({
    url: `${config.appOrigin}/a/${summary.id}`,
    contentUrl: `${config.contentOrigin}/a/${summary.id}/v/${summary.currentVersion}/`
  })
  const present = (summary: ArtifactSummary) => ({ ...summary, ...urls(summary) })

  const isPublisher = (c: Context): boolean => {
    const header = c.req.header('authorization') ?? ''
    const match = /^Bearer\s+(.+)$/i.exec(header)
    if (!match) {
      return false
    }
    const candidate = digest(match[1].trim())
    return tokenDigests.some((known) => timingSafeEqual(known, candidate))
  }

  const requirePublisher = (c: Context): void => {
    if (!isPublisher(c)) {
      throw new ArtifactError(401, 'unauthorized', 'This call needs a publisher token (Authorization: Bearer …).')
    }
  }

  /** Viewer calls are any tailnet browser; they may use only what the artifact declared. */
  const requireCaller = (c: Context, id: string, capability?: Capability): Caller => {
    if (isPublisher(c)) {
      return 'publisher'
    }
    if (c.req.header(VIEWER_HEADER) !== 'viewer') {
      throw new ArtifactError(401, 'unauthorized', `Send a publisher token, or ${VIEWER_HEADER}: viewer from the viewer page.`)
    }
    if (capability && !store.requireSummary(id).capabilities.includes(capability)) {
      throw new ArtifactError(403, 'capability_not_declared', `This artifact did not declare the ${capability} capability.`)
    }
    return 'viewer'
  }

  app.use('*', async (c, next) => {
    await next()
    c.header('X-Content-Type-Options', 'nosniff')
    c.header('Referrer-Policy', 'same-origin')
    c.header('X-Frame-Options', 'DENY')
  })

  app.use(
    '/api/*',
    bodyLimit({
      maxSize: Math.ceil(LIMITS.publishBytes * 1.4) + 1024 * 1024,
      onError: (c) => c.json({ error: { code: 'body_too_large', message: 'Request body too large.' } }, 413)
    })
  )

  // Pages

  app.get('/', (c) => {
    c.header('Content-Security-Policy', "frame-src 'none'")
    return c.html(galleryPage(config))
  })
  app.get('/a/:id', (c) => {
    const detail = store.detail(c.req.param('id'))
    const requested = Number(c.req.query('v') ?? detail.currentVersion)
    const version = detail.versions.find((row) => row.version === requested)?.version ?? detail.currentVersion
    c.header('Content-Security-Policy', `frame-src ${config.contentOrigin}/a/${detail.id}/`)
    const requestedPath = normalizeFilePath(c.req.query('path') ?? '')
    const path = requestedPath && store.requireVersion(detail.id, version).files[requestedPath] ? requestedPath : undefined
    return c.html(viewerPage(config, detail, version, path))
  })
  app.get('/static/:name', (c) => {
    const name = c.req.param('name')
    if (!/^[a-z0-9-]+\.(css|js|svg)$/.test(name)) {
      return c.text('Not found', 404)
    }
    const type = name.endsWith('.css') ? 'text/css' : name.endsWith('.js') ? 'text/javascript' : 'image/svg+xml'
    let body: Buffer
    try {
      body = readFileSync(new URL(name, staticRoot))
    } catch {
      return c.text('Not found', 404)
    }
    c.header('Cache-Control', 'no-cache')
    return c.body(new Uint8Array(body), 200, { 'Content-Type': `${type}; charset=utf-8` })
  })
  app.get('/healthz', (c) => c.text('ok'))

  // Artifacts

  app.get('/api/artifacts', (c) => {
    const result = store.list({
      query: c.req.query('q'),
      limit: Number(c.req.query('limit') ?? 50),
      offset: Number(c.req.query('offset') ?? 0)
    })
    return c.json({ artifacts: result.artifacts.map(present), total: result.total })
  })

  app.post('/api/artifacts', async (c) => {
    requirePublisher(c)
    const body = await jsonBody(c)
    const summary = store.publish({
      title: body.title as string,
      description: body.description as string | undefined,
      icon: body.icon as string | undefined,
      entry: body.entry as string | undefined,
      label: body.label as string | undefined,
      capabilities: optionalStringArray(body.capabilities, 'capabilities'),
      source: body.source,
      files: parseFiles(body.files)
    })
    return c.json(present(summary), 201)
  })

  app.get('/api/artifacts/:id', (c) => {
    const detail = store.detail(c.req.param('id'))
    const version = Number(c.req.query('v') ?? detail.currentVersion)
    const files = store.requireVersion(detail.id, version).files
    return c.json({ ...detail, ...urls(detail), files })
  })

  app.put('/api/artifacts/:id', async (c) => {
    requirePublisher(c)
    const body = await jsonBody(c)
    const force = body.force === true
    if (typeof body.baseVersion !== 'number' && !force) {
      throw new ArtifactError(400, 'base_version_required', 'Pass baseVersion (the version you edited from), or force: true to overwrite.')
    }
    const summary = store.update(c.req.param('id'), {
      title: body.title as string | undefined,
      description: body.description as string | null | undefined,
      icon: body.icon as string | null | undefined,
      entry: body.entry as string | undefined,
      label: body.label as string | undefined,
      capabilities: optionalStringArray(body.capabilities, 'capabilities'),
      source: body.source,
      files: parseFiles(body.files),
      baseVersion: typeof body.baseVersion === 'number' ? body.baseVersion : undefined,
      force
    })
    return c.json(present(summary))
  })

  app.patch('/api/artifacts/:id', async (c) => {
    const id = c.req.param('id')
    requireCaller(c, id)
    const body = await jsonBody(c)
    if (typeof body.pinned !== 'boolean') {
      throw new ArtifactError(400, 'invalid_field', 'Only { pinned: boolean } can be patched.')
    }
    return c.json(present(store.setPinned(id, body.pinned)))
  })

  app.delete('/api/artifacts/:id', (c) => {
    requirePublisher(c)
    store.delete(c.req.param('id'))
    return c.body(null, 204)
  })

  app.get('/api/artifacts/:id/diff', (c) => {
    const number = (value: string | undefined) => (value === undefined ? undefined : Number(value))
    return c.json(store.diff(c.req.param('id'), number(c.req.query('to')), number(c.req.query('from'))))
  })

  app.post('/api/artifacts/:id/prune', async (c) => {
    requirePublisher(c)
    const body = await jsonBody(c)
    return c.json(store.prune(c.req.param('id'), typeof body.keep === 'number' ? body.keep : NaN))
  })

  // Slide decks edited in the viewer. A deck that declares no capabilities is the one artifact a viewer may
  // version: only its entry file changes, and its scripts reach nothing beyond the sandbox.

  const deckSource = (body: Record<string, unknown>): string => {
    const markdown = body.markdown
    if (typeof markdown !== 'string' || Buffer.byteLength(markdown) > SLIDES_SOURCE_BYTES) {
      throw new ArtifactError(400, 'invalid_field', `markdown is the deck source, at most ${SLIDES_SOURCE_BYTES} bytes.`)
    }
    return markdown
  }
  const requireEditableDeck = (c: Context, id: string) => {
    const caller = requireCaller(c, id)
    const summary = store.requireSummary(id)
    if (!isSlidesPath(summary.entry)) {
      throw new ArtifactError(409, 'not_a_deck', 'Only slide decks (*.slides.md) can be edited here.')
    }
    if (caller === 'viewer' && summary.capabilities.length > 0) {
      throw new ArtifactError(403, 'deck_not_editable', 'A deck that declares capabilities is edited by publishing, not in the viewer.')
    }
    return summary
  }

  app.post('/api/artifacts/:id/slides/preview', async (c) => {
    requireEditableDeck(c, c.req.param('id'))
    return c.json(renderSlides(deckSource(await jsonBody(c))))
  })

  app.put('/api/artifacts/:id/slides', async (c) => {
    const id = c.req.param('id')
    const summary = requireEditableDeck(c, id)
    const body = await jsonBody(c)
    const markdown = deckSource(body)
    if (typeof body.baseVersion !== 'number') {
      throw new ArtifactError(400, 'base_version_required', 'Pass baseVersion: the version the edit started from.')
    }
    const author = typeof body.author === 'string' && body.author.trim() ? body.author.trim().slice(0, 40) : null
    const saved = store.update(id, {
      files: { [summary.entry]: { content: Buffer.from(markdown) } },
      baseVersion: body.baseVersion,
      // No source: the artifact keeps showing the machine and folder it is published from; the label names the editor.
      label: author ? `瀏覽器編輯 · ${author}` : '瀏覽器編輯'
    })
    return c.json(present(saved))
  })

  // A deck as a PowerPoint file, for sending on; served as a download like every file from this origin.
  app.get('/api/artifacts/:id/export.pptx', async (c) => {
    const id = c.req.param('id')
    const summary = store.requireSummary(id)
    const row = store.requireVersion(id, Number(c.req.query('v') ?? summary.currentVersion))
    if (!isSlidesPath(row.entry)) {
      throw new ArtifactError(409, 'not_a_deck', 'Only slide decks (*.slides.md) export to .pptx.')
    }
    const read = (path: string) => (row.files[path] ? store.blobs.read(row.files[path].sha256) : null)
    const pptx = await deckToPptx(store.blobs.read(row.files[row.entry].sha256).toString('utf8'), row.title, row.entry, read)
    const name = `${row.title.replace(/[\\/:*?"<>|\u0000-\u001f]/g, '_').slice(0, 100) || 'deck'}.pptx`
    return c.body(new Uint8Array(pptx), 200, {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      'Content-Disposition': `attachment; filename="deck.pptx"; filename*=UTF-8''${encodeURIComponent(name)}`
    })
  })

  app.get('/api/artifacts/:id/files/*', (c) => {
    const id = c.req.param('id')
    const summary = store.requireSummary(id)
    const prefix = `/api/artifacts/${id}/files/`
    const requested = new URL(c.req.url).pathname.slice(prefix.length)
    const path = decodeRequestPath(requested)
    const row = store.requireVersion(id, Number(c.req.query('v') ?? summary.currentVersion))
    const entry = path ? row.files[path] : undefined
    if (!path || !entry) {
      throw new ArtifactError(404, 'file_not_found', `Version ${row.version} has no file ${path ?? requested}.`)
    }
    return c.body(new Uint8Array(store.blobs.read(entry.sha256)), 200, {
      'Content-Type': 'application/octet-stream',
      'Content-Disposition': 'attachment'
    })
  })

  // Assets

  const presentAsset = (id: string, asset: ReturnType<ArtifactStore['getAsset']>) => ({
    ...asset,
    url: `${config.contentOrigin}/a/${id}/assets/${asset.id}/${encodeURIComponent(asset.name)}`
  })

  app.get('/api/artifacts/:id/assets', (c) => {
    const id = c.req.param('id')
    return c.json({ assets: store.listAssets(id).map((asset) => presentAsset(id, asset)) })
  })

  app.post('/api/artifacts/:id/assets', async (c) => {
    const id = c.req.param('id')
    const caller = requireCaller(c, id, 'uploads')
    const name = c.req.query('name')
    if (!name) {
      throw new ArtifactError(400, 'invalid_field', 'Pass ?name=<file name>.')
    }
    const content = Buffer.from(await c.req.arrayBuffer())
    const asset = store.addAsset(id, name, content, caller === 'publisher' ? 'publisher' : 'page')
    return c.json(presentAsset(id, asset), 201)
  })

  app.delete('/api/artifacts/:id/assets/:assetId', (c) => {
    requirePublisher(c)
    store.deleteAsset(c.req.param('id'), c.req.param('assetId'))
    return c.body(null, 204)
  })

  // Shared database (db capability)

  app.post('/api/artifacts/:id/db/_batch', async (c) => {
    const id = c.req.param('id')
    requireCaller(c, id, 'db')
    const body = await jsonBody(c)
    if (!Array.isArray(body.ops) || body.ops.length === 0 || body.ops.length > 500) {
      throw new ArtifactError(400, 'invalid_batch', 'ops is an array of 1-500 operations.')
    }
    return c.json({ results: store.batch(id, body.ops) })
  })

  app.get('/api/artifacts/:id/db/:collection', (c) => {
    const id = c.req.param('id')
    requireCaller(c, id, 'db')
    return c.json(
      store.listDocs(id, c.req.param('collection'), {
        limit: Number(c.req.query('limit') ?? 100),
        after: c.req.query('after')
      })
    )
  })

  app.post('/api/artifacts/:id/db/:collection/_query', async (c) => {
    const id = c.req.param('id')
    requireCaller(c, id, 'db')
    const body = await jsonBody(c)
    const where = body.where ?? {}
    if (!where || typeof where !== 'object' || Array.isArray(where)) {
      throw new ArtifactError(400, 'invalid_query', 'where is an object of field → value.')
    }
    return c.json(
      store.listDocs(id, c.req.param('collection'), {
        where: where as Record<string, unknown>,
        limit: typeof body.limit === 'number' ? body.limit : undefined,
        after: typeof body.after === 'string' ? body.after : undefined
      })
    )
  })

  app.get('/api/artifacts/:id/db/:collection/:docId', (c) => {
    const id = c.req.param('id')
    requireCaller(c, id, 'db')
    const doc = store.getDoc(id, c.req.param('collection'), c.req.param('docId'))
    if (!doc) {
      throw new ArtifactError(404, 'doc_not_found', 'No such document.')
    }
    return c.json(doc)
  })

  for (const [method, merge] of [['put', false], ['patch', true]] as const) {
    app[method]('/api/artifacts/:id/db/:collection/:docId', async (c) => {
      const id = c.req.param('id')
      requireCaller(c, id, 'db')
      const body = await jsonBody(c)
      return c.json(store.setDoc(id, c.req.param('collection'), c.req.param('docId'), body.data, merge))
    })
  }

  app.delete('/api/artifacts/:id/db/:collection/:docId', (c) => {
    const id = c.req.param('id')
    requireCaller(c, id, 'db')
    return c.json({ deleted: store.deleteDoc(id, c.req.param('collection'), c.req.param('docId')) })
  })

  // Comments (every artifact)

  app.get('/api/artifacts/:id/comments', (c) => c.json({ comments: store.listComments(c.req.param('id')) }))

  app.post('/api/artifacts/:id/comments', async (c) => {
    const id = c.req.param('id')
    requireCaller(c, id)
    const body = await jsonBody(c)
    return c.json(store.addComment(id, body), 201)
  })

  app.patch('/api/artifacts/:id/comments/:commentId', async (c) => {
    const id = c.req.param('id')
    requireCaller(c, id)
    const body = await jsonBody(c)
    store.resolveComment(id, c.req.param('commentId'), body.resolved === true)
    return c.body(null, 204)
  })

  app.delete('/api/artifacts/:id/comments/:commentId', (c) => {
    const id = c.req.param('id')
    requireCaller(c, id)
    store.deleteComment(id, c.req.param('commentId'))
    return c.body(null, 204)
  })

  // Page errors: reported by the viewer while a page runs; any tailnet reader may read them.

  app.get('/api/artifacts/:id/errors', (c) => {
    const id = c.req.param('id')
    const version = Number(c.req.query('v') ?? store.requireSummary(id).currentVersion)
    return c.json({ errors: store.listErrors(id, version) })
  })

  app.post('/api/artifacts/:id/errors', async (c) => {
    const id = c.req.param('id')
    requireCaller(c, id)
    store.recordErrors(id, await jsonBody(c))
    return c.body(null, 204)
  })

  // Live changes for the viewer: new versions, db writes, comments, assets.

  app.get('/api/artifacts/:id/events', (c) => {
    const id = c.req.param('id')
    store.requireSummary(id)
    return streamSSE(c, async (stream) => {
      const queue: string[] = []
      let wake: (() => void) | null = null
      const unsubscribe = store.subscribe(id, (change) => {
        queue.push(JSON.stringify(change))
        wake?.()
      })
      stream.onAbort(() => {
        unsubscribe()
        wake?.()
      })
      await stream.writeSSE({ event: 'ready', data: '{}' })
      while (!stream.aborted) {
        if (queue.length === 0) {
          await new Promise<void>((resolve) => {
            wake = resolve
            setTimeout(resolve, 25_000)
          })
          wake = null
        }
        if (stream.aborted) {
          break
        }
        if (queue.length === 0) {
          await stream.writeSSE({ event: 'ping', data: '{}' })
          continue
        }
        for (const data of queue.splice(0)) {
          await stream.writeSSE({ event: 'change', data })
        }
      }
      unsubscribe()
    })
  })

  app.onError((error, c) => {
    if (error instanceof ArtifactError) {
      return c.json({ error: { code: error.code, message: error.message, ...(error.details ? { details: error.details } : {}) } }, error.status)
    }
    console.error(error)
    return c.json({ error: { code: 'internal', message: 'Internal error' } }, 500)
  })
  app.notFound((c) =>
    c.req.path.startsWith('/api/') ? c.json({ error: { code: 'not_found', message: 'No such route.' } }, 404) : c.text('Not found', 404)
  )
  return app
}

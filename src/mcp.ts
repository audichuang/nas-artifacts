import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { WebStandardStreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/webStandardStreamableHttp.js'
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js'
import type { Context, Hono } from 'hono'
import { z } from 'zod'
import { ArtifactError, LIMITS, type ArtifactStore, type ArtifactSummary, type IncomingFile } from './artifact-store.ts'
import type { Config } from './config.ts'
import type { OAuthServer } from './oauth.ts'

/** The most of one file `read` returns inline; larger files are for the original-file download. */
const READ_BYTES = 4 * 1024 * 1024

const fileSchema = z.object({
  path: z.string().describe('Path inside the artifact, e.g. "index.html" or "img/chart.png".'),
  content: z.string().optional().describe('File content; omit together with delete: true to remove the file (update only).'),
  encoding: z.enum(['utf8', 'base64']).optional().describe('utf8 (default) for text, base64 for pictures and Office files.'),
  delete: z.boolean().optional().describe('update only: remove this file from the new version.')
})

function toFiles(list: z.infer<typeof fileSchema>[] | undefined, allowDelete: boolean): Record<string, IncomingFile> {
  const files: Record<string, IncomingFile> = {}
  for (const file of list ?? []) {
    if (file.delete) {
      if (!allowDelete) throw new ArtifactError(400, 'invalid_files', 'delete is only for update.')
      files[file.path] = null
      continue
    }
    if (typeof file.content !== 'string') throw new ArtifactError(400, 'invalid_files', `${file.path}: content is required.`)
    files[file.path] = { content: Buffer.from(file.content, file.encoding ?? 'utf8') }
  }
  return files
}

const json = (value: unknown): CallToolResult => ({ content: [{ type: 'text', text: JSON.stringify(value, null, 2) }] })

/** Caller mistakes go back to the agent as a tool error it can act on (a conflict names the changed files). */
async function run(work: () => unknown | Promise<unknown>): Promise<CallToolResult> {
  try {
    return json(await work())
  } catch (error) {
    if (error instanceof ArtifactError) {
      return { isError: true, content: [{ type: 'text', text: JSON.stringify({ error: { code: error.code, message: error.message, ...(error.details ? { details: error.details } : {}) } }, null, 2) }] }
    }
    throw error
  }
}

function buildServer(store: ArtifactStore, config: Config): McpServer {
  const pair = config.public!
  const present = (summary: ArtifactSummary) => ({
    id: summary.id,
    title: summary.title,
    description: summary.description,
    icon: summary.icon,
    entry: summary.entry,
    currentVersion: summary.currentVersion,
    capabilities: summary.capabilities,
    visibility: summary.visibility,
    updatedAt: summary.updatedAt,
    // The owner opens it signed in; anyone may open the public one.
    url: `${pair.appOrigin}/a/${summary.id}`,
    publicUrl: summary.visibility === 'public' && summary.capabilities.length === 0 ? `${pair.appOrigin}/a/${summary.id}` : null
  })

  const server = new McpServer(
    { name: 'artifacts', version: '1.0.0' },
    {
      instructions:
        'Publish HTML or Markdown pages, slide decks (*.slides.md) and Office files (.pptx/.docx/.xlsx) for the owner to read on any device. ' +
        'New artifacts are private (only the signed-in owner can open them); visibility "public" lets anyone with the link read the current version. ' +
        'Pages run in a sandbox: no network requests (connect-src none), scripts only inline, from the artifact, or from cdnjs/jsdelivr/unpkg/tailwind/jquery CDNs. ' +
        'To change an artifact, call get for its currentVersion, then update with baseVersion set to it; a version_conflict error lists changedFiles to merge before retrying.'
    }
  )

  server.registerTool(
    'list',
    {
      title: 'List artifacts',
      description: 'List artifacts, newest first; query searches titles, page text and source folders.',
      inputSchema: { query: z.string().optional(), limit: z.number().int().min(1).max(200).optional(), offset: z.number().int().min(0).optional() },
      annotations: { readOnlyHint: true }
    },
    ({ query, limit, offset }) =>
      run(() => {
        const result = store.list({ query, limit, offset })
        return { total: result.total, artifacts: result.artifacts.map((row) => ({ ...present(row), ...(row.snippet ? { snippet: row.snippet } : {}) })) }
      })
  )

  server.registerTool(
    'get',
    {
      title: 'Get an artifact',
      description: 'Metadata, version history and the file list of a version (default: the current one).',
      inputSchema: { id: z.string(), version: z.number().int().min(1).optional() },
      annotations: { readOnlyHint: true }
    },
    ({ id, version }) =>
      run(() => {
        const detail = store.detail(id)
        const row = store.requireVersion(detail.id, version ?? detail.currentVersion)
        return {
          ...present(detail),
          versions: detail.versions.map((v) => ({ version: v.version, label: v.label, title: v.title, createdAt: v.createdAt })),
          version: row.version,
          files: Object.entries(row.files).map(([path, file]) => ({ path, size: file.size, type: file.type }))
        }
      })
  )

  server.registerTool(
    'read',
    {
      title: 'Read a file',
      description: 'The content of one file of a version (default: the entry of the current version). Text comes back as utf8, other files as base64.',
      inputSchema: { id: z.string(), path: z.string().optional(), version: z.number().int().min(1).optional() },
      annotations: { readOnlyHint: true }
    },
    ({ id, path, version }) =>
      run(() => {
        const summary = store.requireSummary(id)
        const row = store.requireVersion(id, version ?? summary.currentVersion)
        const name = path ?? row.entry
        const file = row.files[name]
        if (!file) throw new ArtifactError(404, 'file_not_found', `Version ${row.version} has no file ${name}.`)
        if (file.size > READ_BYTES) throw new ArtifactError(413, 'too_large', `${name} is ${file.size} bytes; download it from ${pair.appOrigin}/a/${id} instead.`)
        const bytes = store.blobs.read(file.sha256)
        const text = /^text\/|json|javascript|svg|xml|markdown/.test(file.type) && !/officedocument/.test(file.type)
        return { path: name, version: row.version, type: file.type, encoding: text ? 'utf8' : 'base64', content: bytes.toString(text ? 'utf8' : 'base64') }
      })
  )

  server.registerTool(
    'publish',
    {
      title: 'Publish a new artifact',
      description:
        'Create a new artifact from files. entry defaults to index.html / index.md / the only page. ' +
        `Limits: ${LIMITS.fileBytes / 1024 / 1024} MB per file, ${LIMITS.publishBytes / 1024 / 1024} MB per version. Returns the URL to give the owner.`,
      inputSchema: {
        title: z.string(),
        files: z.array(fileSchema).min(1),
        entry: z.string().optional(),
        description: z.string().optional(),
        icon: z.string().optional().describe('One emoji.'),
        label: z.string().optional().describe('A short note for this version.'),
        capabilities: z.array(z.enum(['db', 'uploads'])).optional().describe('db = shared database, uploads = files viewers add. Such pages cannot be public.'),
        visibility: z.enum(['private', 'public']).optional().describe('Default private. public = anyone with the link, read-only. Only when the owner asked for it.')
      },
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false }
    },
    (input) =>
      run(() => {
        if (input.visibility === 'public' && input.capabilities?.length) {
          throw new ArtifactError(409, 'has_capabilities', 'Public pages are read-only: a page that declares capabilities cannot be public.')
        }
        let summary = store.publish({
          title: input.title,
          description: input.description,
          icon: input.icon,
          entry: input.entry,
          label: input.label ?? 'MCP',
          capabilities: input.capabilities,
          files: toFiles(input.files, false)
        })
        if (input.visibility === 'public') summary = store.setVisibility(summary.id, 'public')
        return present(summary)
      })
  )

  server.registerTool(
    'update',
    {
      title: 'Publish a new version',
      description:
        'Add a version to an existing artifact. Send only changed or new files (delete: true removes one); the rest carry over. ' +
        'baseVersion is the version you read and edited; if someone published since, you get version_conflict with changedFiles: read those, merge, retry. force: true overwrites instead (only if the owner asked).',
      inputSchema: {
        id: z.string(),
        baseVersion: z.number().int().min(1).optional(),
        force: z.boolean().optional(),
        files: z.array(fileSchema).optional(),
        title: z.string().optional(),
        entry: z.string().optional(),
        description: z.string().nullable().optional(),
        icon: z.string().nullable().optional(),
        label: z.string().optional(),
        capabilities: z.array(z.enum(['db', 'uploads'])).optional()
      },
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false }
    },
    (input) =>
      run(() => {
        if (input.baseVersion === undefined && input.force !== true) {
          throw new ArtifactError(400, 'base_version_required', 'Pass baseVersion (the version you edited from), or force: true to overwrite.')
        }
        return present(
          store.update(input.id, {
            title: input.title,
            description: input.description,
            icon: input.icon,
            entry: input.entry,
            label: input.label ?? 'MCP',
            capabilities: input.capabilities,
            files: toFiles(input.files, true),
            baseVersion: input.baseVersion,
            force: input.force === true
          })
        )
      })
  )

  server.registerTool(
    'set_visibility',
    {
      title: 'Make an artifact public or private',
      description: 'public: anyone with the link may read the current version (read-only). private: only the signed-in owner. Pages with capabilities cannot be public.',
      inputSchema: { id: z.string(), visibility: z.enum(['private', 'public']) },
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true }
    },
    ({ id, visibility }) => run(() => present(store.setVisibility(id, visibility)))
  )

  server.registerTool(
    'delete',
    {
      title: 'Delete an artifact',
      description: 'Permanently delete an artifact with all its versions, data and comments. Only when the owner asked for it.',
      inputSchema: { id: z.string() },
      annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: true }
    },
    ({ id }) =>
      run(() => {
        store.delete(id)
        return { deleted: id }
      })
  )

  server.registerTool(
    'comments',
    {
      title: 'Read comments',
      description: 'Comments viewers left on an artifact, as threads (replies carry parentId).',
      inputSchema: { id: z.string(), open_only: z.boolean().optional().describe('Only threads not yet resolved.') },
      annotations: { readOnlyHint: true }
    },
    ({ id, open_only }) =>
      run(() => {
        const comments = store.listComments(id)
        if (!open_only) return { comments }
        const open = new Set(comments.filter((c) => c.parentId === null && !c.resolved).map((c) => c.id))
        return { comments: comments.filter((c) => open.has(c.parentId ?? c.id)) }
      })
  )

  server.registerTool(
    'reply',
    {
      title: 'Comment or reply',
      description: 'Add a comment, or reply to a thread (commentId); resolve: true also marks the thread resolved.',
      inputSchema: { id: z.string(), body: z.string(), commentId: z.string().optional(), resolve: z.boolean().optional(), author: z.string().optional() },
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false }
    },
    ({ id, body, commentId, resolve, author }) =>
      run(() => {
        const comment = store.addComment(id, { author: author ?? 'Agent', body, parentId: commentId })
        if (resolve) store.resolveComment(id, comment.parentId ?? comment.id, true)
        return comment
      })
  )

  server.registerTool(
    'errors',
    {
      title: 'Page errors',
      description: "Errors viewers' browsers hit on a version (default: the current one): script errors, failed loads, sandbox blocks.",
      inputSchema: { id: z.string(), version: z.number().int().min(1).optional() },
      annotations: { readOnlyHint: true }
    },
    ({ id, version }) => run(() => ({ errors: store.listErrors(id, version ?? store.requireSummary(id).currentVersion) }))
  )

  return server
}

/**
 * `/mcp` on the public app: MCP Streamable HTTP, stateless with JSON responses, behind OAuth access
 * tokens only. No cookies are read here, so a page that reaches this route from a browser has nothing to
 * ride on; that is also why any Origin may call it (DNS rebinding gains nothing without a token).
 */
export function mountMcp(app: Hono, store: ArtifactStore, config: Config, oauth: OAuthServer): void {
  const cors = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Expose-Headers': 'WWW-Authenticate, Mcp-Session-Id, Mcp-Protocol-Version'
  }
  app.options('/mcp', (c) =>
    c.body(null, 204, {
      ...cors,
      'Access-Control-Allow-Methods': 'POST, GET, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'authorization, content-type, accept, mcp-protocol-version, mcp-session-id, last-event-id, mcp-method',
      'Access-Control-Max-Age': '600'
    })
  )
  app.all('/mcp', async (c: Context) => {
    const match = /^Bearer\s+(.+)$/i.exec(c.req.header('authorization') ?? '')
    const caller = oauth.verifyAccessToken(match?.[1].trim())
    if (!caller) {
      return c.json({ error: 'invalid_token', error_description: 'Sign in through OAuth.' }, 401, { ...cors, 'WWW-Authenticate': oauth.challenge(Boolean(match)) })
    }
    if (c.req.method !== 'POST') {
      // Stateless: no server-initiated stream and no session to end.
      return c.body(null, 405, { ...cors, Allow: 'POST' })
    }
    const server = buildServer(store, config)
    // Files arrive base64 inside JSON: the same allowance as the publish API, not the SDK's 4 MiB default.
    const transport = new WebStandardStreamableHTTPServerTransport({
      sessionIdGenerator: undefined,
      enableJsonResponse: true,
      maxRequestBodySize: Math.ceil(LIMITS.publishBytes * 1.4) + 1024 * 1024
    })
    await server.connect(transport)
    try {
      const response = await transport.handleRequest(c.req.raw)
      const headers = new Headers(response.headers)
      for (const [key, value] of Object.entries(cors)) headers.set(key, value)
      return new Response(response.body, { status: response.status, headers })
    } finally {
      void server.close()
    }
  })
}


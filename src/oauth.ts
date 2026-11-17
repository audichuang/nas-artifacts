import { createHash, randomBytes } from 'node:crypto'
import type { DatabaseSync } from 'node:sqlite'
import { Hono, type Context } from 'hono'
import { bodyLimit } from 'hono/body-limit'
import { escapeHtml } from './content-render.ts'

/**
 * The OAuth 2.1 authorization server in front of the remote MCP (MCP authorization spec 2026-07-28):
 * protected-resource and authorization-server metadata, dynamic client registration (public clients
 * only), the authorization-code flow with PKCE S256 whose consent page needs the owner's passkey
 * session, rotating refresh tokens, and revocation. Every client signs in the same way; there is no
 * token path into the MCP besides the access tokens issued here.
 */

export const SCOPE = 'artifacts'
const ACCESS_SECONDS = 3600
const REFRESH_DAYS = 30
const CODE_SECONDS = 300
const PENDING_SECONDS = 600
const MAX_CLIENTS = 200
const UNUSED_CLIENT_DAYS = 1

const sha256 = (value: string): string => createHash('sha256').update(value).digest('hex')
const randomToken = (): string => randomBytes(32).toString('base64url')
const nowIso = (): string => new Date().toISOString()

export type ClientRow = { id: string; name: string; redirectUris: string[]; createdAt: string; lastUsedAt: string | null }
export type GrantedClient = ClientRow & { tokens: number }

class OAuthError extends Error {
  readonly status: 400 | 401
  readonly code: string
  constructor(status: 400 | 401, code: string, message: string) {
    super(message)
    this.status = status
    this.code = code
  }
}

const LOOPBACK = new Set(['localhost', '127.0.0.1', '[::1]'])

/** https anywhere, http only to this machine's loopback, or an app's private-use scheme; never a fragment. */
export function validRedirectUri(value: unknown): value is string {
  if (typeof value !== 'string' || value.length > 2000) return false
  let url: URL
  try {
    url = new URL(value)
  } catch {
    return false
  }
  if (url.hash || url.username || url.password) return false
  if (url.protocol === 'https:') return true
  if (url.protocol === 'http:') return LOOPBACK.has(url.hostname)
  // Private-use schemes (RFC 8252 7.1) name a native app; schemes a browser would run or read locally do not.
  return /^[a-z][a-z0-9+.-]*\.[a-z0-9+.-]+:$/.test(url.protocol) && !['javascript:', 'data:', 'file:', 'blob:', 'vbscript:'].includes(url.protocol)
}

/** Loopback redirects may use any port (RFC 8252 7.3): native clients pick a free one per sign-in. Nothing else may differ. */
function sameRedirect(registered: string, given: string): boolean {
  if (registered === given) return true
  if (!validRedirectUri(given)) return false
  const a = new URL(registered)
  const b = new URL(given)
  return a.protocol === 'http:' && b.protocol === 'http:' && LOOPBACK.has(a.hostname) && a.hostname === b.hostname && a.pathname === b.pathname && a.search === b.search
}

type Pending = { clientId: string; redirectUri: string; challenge: string; state: string | null; resource: string; scope: string; expires: number }

export class OAuthServer {
  readonly issuer: string
  readonly resource: string
  private readonly pending = new Map<string, Pending>()
  private readonly db: DatabaseSync

  constructor(db: DatabaseSync, publicAppOrigin: string) {
    this.db = db
    this.issuer = new URL(publicAppOrigin).origin
    this.resource = `${this.issuer}/mcp`
    db.exec(`
      CREATE TABLE IF NOT EXISTS oauth_clients (
        id TEXT PRIMARY KEY,
        name TEXT NOT NULL,
        redirect_uris TEXT NOT NULL,
        created_at TEXT NOT NULL,
        last_used_at TEXT
      );
      CREATE TABLE IF NOT EXISTS oauth_codes (
        hash TEXT PRIMARY KEY,
        client_id TEXT NOT NULL,
        redirect_uri TEXT NOT NULL,
        challenge TEXT NOT NULL,
        resource TEXT NOT NULL,
        scope TEXT NOT NULL,
        expires_at INTEGER NOT NULL
      );
      CREATE TABLE IF NOT EXISTS oauth_tokens (
        hash TEXT PRIMARY KEY,
        kind TEXT NOT NULL,
        client_id TEXT NOT NULL,
        family TEXT NOT NULL,
        resource TEXT NOT NULL,
        scope TEXT NOT NULL,
        expires_at INTEGER NOT NULL,
        used_at INTEGER
      );
      CREATE INDEX IF NOT EXISTS oauth_tokens_family ON oauth_tokens (family);
    `)
  }

  // Clients

  private client(id: unknown): ClientRow | null {
    if (typeof id !== 'string' || id.length > 100) return null
    const row = this.db.prepare('SELECT * FROM oauth_clients WHERE id = ?').get(id)
    return row
      ? { id: String(row.id), name: String(row.name), redirectUris: JSON.parse(String(row.redirect_uris)), createdAt: String(row.created_at), lastUsedAt: row.last_used_at === null ? null : String(row.last_used_at) }
      : null
  }

  register(input: Record<string, unknown>): Record<string, unknown> {
    const uris = input.redirect_uris
    if (!Array.isArray(uris) || uris.length === 0 || uris.length > 10 || !uris.every(validRedirectUri)) {
      throw new OAuthError(400, 'invalid_redirect_uri', 'redirect_uris: 1-10 https URLs, loopback http URLs, or private-use scheme URIs, without fragments.')
    }
    const method = input.token_endpoint_auth_method ?? 'none'
    if (method !== 'none') throw new OAuthError(400, 'invalid_client_metadata', 'Only public clients (token_endpoint_auth_method "none") are supported.')
    const grants = (input.grant_types as unknown) ?? ['authorization_code', 'refresh_token']
    if (!Array.isArray(grants) || !grants.every((grant) => grant === 'authorization_code' || grant === 'refresh_token')) {
      throw new OAuthError(400, 'invalid_client_metadata', 'grant_types: authorization_code and refresh_token only.')
    }
    const name = typeof input.client_name === 'string' && input.client_name.trim() ? input.client_name.trim().slice(0, 100) : '未命名的應用程式'
    this.forgetUnusedClients()
    if (Number(this.db.prepare('SELECT COUNT(*) AS n FROM oauth_clients').get()!.n) >= MAX_CLIENTS) {
      throw new OAuthError(400, 'invalid_client_metadata', 'Too many registered clients; revoke unused ones in the account page.')
    }
    const id = randomToken()
    const createdAt = nowIso()
    this.db.prepare('INSERT INTO oauth_clients (id, name, redirect_uris, created_at) VALUES (?, ?, ?, ?)').run(id, name, JSON.stringify(uris), createdAt)
    return {
      client_id: id,
      client_id_issued_at: Math.floor(Date.parse(createdAt) / 1000),
      client_name: name,
      redirect_uris: uris,
      grant_types: ['authorization_code', 'refresh_token'],
      response_types: ['code'],
      token_endpoint_auth_method: 'none',
      scope: SCOPE
    }
  }

  /**
   * Only registrations nobody ever finished signing in with go (after a day). A client that once had tokens
   * keeps its id: clients reuse it to sign in again, and an id that vanished would leave them stuck.
   */
  private forgetUnusedClients(): void {
    const before = new Date(Date.now() - UNUSED_CLIENT_DAYS * 86400_000).toISOString()
    this.db.prepare('DELETE FROM oauth_clients WHERE last_used_at IS NULL AND created_at < ?').run(before)
  }

  grantedClients(): GrantedClient[] {
    const rows = this.db
      .prepare(
        `SELECT c.*, COUNT(t.hash) AS tokens FROM oauth_clients c
         JOIN oauth_tokens t ON t.client_id = c.id AND t.expires_at > ? AND t.used_at IS NULL
         GROUP BY c.id ORDER BY c.last_used_at DESC`
      )
      .all(Date.now())
    return rows.map((row) => ({
      id: String(row.id),
      name: String(row.name),
      redirectUris: JSON.parse(String(row.redirect_uris)),
      createdAt: String(row.created_at),
      lastUsedAt: row.last_used_at === null ? null : String(row.last_used_at),
      tokens: Number(row.tokens)
    }))
  }

  revokeClient(id: string): void {
    this.db.prepare('DELETE FROM oauth_tokens WHERE client_id = ?').run(id)
    this.db.prepare('DELETE FROM oauth_codes WHERE client_id = ?').run(id)
  }

  // Authorization

  /** Checks an authorization request; errors the client can be told about come back as a redirect. */
  authorizationRequest(query: Record<string, string | undefined>): { ok: true; requestId: string; client: ClientRow; redirectUri: string } | { ok: false; message: string } {
    const client = this.client(query.client_id)
    if (!client) return { ok: false, message: '不認得這個應用程式（client_id）。請從應用程式重新連線。' }
    const redirectUri = query.redirect_uri ?? (client.redirectUris.length === 1 ? client.redirectUris[0] : undefined)
    if (!redirectUri || !client.redirectUris.some((registered) => sameRedirect(registered, redirectUri))) {
      return { ok: false, message: '回呼網址（redirect_uri）和這個應用程式註冊的不同。' }
    }
    // Shown here, not redirected: anyone may register a callback, and an error redirect would skip the consent page that names it.
    const fail = (error: string, description: string) => ({ ok: false as const, message: `${description}（${error}）` })
    if (query.response_type !== 'code') return fail('unsupported_response_type', 'response_type must be code.')
    if (!query.code_challenge || query.code_challenge_method !== 'S256' || !/^[A-Za-z0-9_-]{43}$/.test(query.code_challenge)) {
      return fail('invalid_request', 'PKCE with code_challenge_method S256 is required.')
    }
    const resource = this.resourceFor(query.resource)
    if (!resource) return fail('invalid_target', `resource must be ${this.resource}.`)
    const scopes = (query.scope ?? SCOPE).split(' ').filter(Boolean)
    if (scopes.some((scope) => scope !== SCOPE && scope !== 'offline_access')) return fail('invalid_scope', `The only scope is ${SCOPE}.`)
    const now = Date.now()
    for (const [key, value] of this.pending) if (value.expires <= now) this.pending.delete(key)
    if (this.pending.size >= 100) this.pending.delete(this.pending.keys().next().value!)
    const requestId = randomToken()
    this.pending.set(requestId, { clientId: client.id, redirectUri, challenge: query.code_challenge, state: query.state ?? null, resource, scope: SCOPE, expires: now + PENDING_SECONDS * 1000 })
    return { ok: true, requestId, client, redirectUri }
  }

  /** The owner's answer on the consent page; returns where to send the browser. */
  decide(requestId: unknown, allow: boolean): string | null {
    const pending = typeof requestId === 'string' ? this.pending.get(requestId) : undefined
    if (typeof requestId === 'string') this.pending.delete(requestId)
    if (!pending || pending.expires <= Date.now()) return null
    if (!allow) return this.redirectWith(pending.redirectUri, { error: 'access_denied', error_description: 'The owner declined.', state: pending.state ?? undefined })
    const code = randomToken()
    this.db
      .prepare('INSERT INTO oauth_codes (hash, client_id, redirect_uri, challenge, resource, scope, expires_at) VALUES (?, ?, ?, ?, ?, ?, ?)')
      .run(sha256(code), pending.clientId, pending.redirectUri, pending.challenge, pending.resource, pending.scope, Date.now() + CODE_SECONDS * 1000)
    return this.redirectWith(pending.redirectUri, { code, state: pending.state ?? undefined })
  }

  private redirectWith(uri: string, params: Record<string, string | undefined>): string {
    const url = new URL(uri)
    for (const [key, value] of Object.entries({ ...params, iss: this.issuer })) if (value !== undefined) url.searchParams.set(key, value)
    return url.toString()
  }

  /** The token is for this MCP server only; the origin alone names it too (some clients send that). */
  private resourceFor(value: string | undefined): string | null {
    if (value === undefined) return this.resource
    let url: URL
    try {
      url = new URL(value)
    } catch {
      return null
    }
    const normalized = `${url.origin}${url.pathname.replace(/\/$/, '')}`
    return url.search === '' && url.hash === '' && (normalized === this.resource || normalized === this.issuer) ? this.resource : null
  }

  // Tokens

  token(form: Record<string, string | undefined>): Record<string, unknown> {
    if (form.grant_type === 'authorization_code') return this.exchangeCode(form)
    if (form.grant_type === 'refresh_token') return this.refresh(form)
    throw new OAuthError(400, 'unsupported_grant_type', 'grant_type is authorization_code or refresh_token.')
  }

  private exchangeCode(form: Record<string, string | undefined>): Record<string, unknown> {
    const code = form.code ?? ''
    const row = this.db.prepare('SELECT * FROM oauth_codes WHERE hash = ?').get(sha256(code))
    // Single use: gone before anything else is checked.
    this.db.prepare('DELETE FROM oauth_codes WHERE hash = ? OR expires_at <= ?').run(sha256(code), Date.now())
    if (!row || Number(row.expires_at) <= Date.now()) throw new OAuthError(400, 'invalid_grant', 'The authorization code is invalid or expired.')
    if (form.client_id !== row.client_id) throw new OAuthError(400, 'invalid_grant', 'The code was issued to another client.')
    if (!form.redirect_uri || !sameRedirect(String(row.redirect_uri), form.redirect_uri)) throw new OAuthError(400, 'invalid_grant', 'redirect_uri does not match the authorization request.')
    const verifier = form.code_verifier ?? ''
    if (!/^[A-Za-z0-9._~-]{43,128}$/.test(verifier) || createHash('sha256').update(verifier).digest('base64url') !== row.challenge) {
      throw new OAuthError(400, 'invalid_grant', 'code_verifier does not match the code challenge.')
    }
    if (form.resource !== undefined && this.resourceFor(form.resource) !== row.resource) throw new OAuthError(400, 'invalid_target', `resource must be ${this.resource}.`)
    return this.issue(String(row.client_id), randomToken(), String(row.resource), String(row.scope))
  }

  private refresh(form: Record<string, string | undefined>): Record<string, unknown> {
    const hash = sha256(form.refresh_token ?? '')
    const row = this.db.prepare("SELECT * FROM oauth_tokens WHERE hash = ? AND kind = 'refresh'").get(hash)
    if (!row || Number(row.expires_at) <= Date.now()) throw new OAuthError(400, 'invalid_grant', 'The refresh token is invalid or expired.')
    if (row.used_at !== null) {
      // A rotated-out refresh token came back: someone holds a copy. End the whole sign-in.
      this.db.prepare('DELETE FROM oauth_tokens WHERE family = ?').run(row.family)
      throw new OAuthError(400, 'invalid_grant', 'The refresh token was already used; sign in again.')
    }
    if (form.client_id !== undefined && form.client_id !== row.client_id) throw new OAuthError(400, 'invalid_grant', 'The refresh token belongs to another client.')
    if (form.resource !== undefined && this.resourceFor(form.resource) !== row.resource) throw new OAuthError(400, 'invalid_target', `resource must be ${this.resource}.`)
    // No grace period: a second use is a replay. Claude Code serializes refreshes across its processes.
    this.db.prepare('UPDATE oauth_tokens SET used_at = ? WHERE hash = ?').run(Date.now(), hash)
    return this.issue(String(row.client_id), String(row.family), String(row.resource), String(row.scope))
  }

  private issue(clientId: string, family: string, resource: string, scope: string): Record<string, unknown> {
    if (!this.client(clientId)) throw new OAuthError(400, 'invalid_grant', 'The client is no longer registered; register again.')
    const access = randomToken()
    const refresh = randomToken()
    const now = Date.now()
    this.db.prepare('DELETE FROM oauth_tokens WHERE expires_at <= ?').run(now)
    const insert = this.db.prepare('INSERT INTO oauth_tokens (hash, kind, client_id, family, resource, scope, expires_at) VALUES (?, ?, ?, ?, ?, ?, ?)')
    insert.run(sha256(access), 'access', clientId, family, resource, scope, now + ACCESS_SECONDS * 1000)
    insert.run(sha256(refresh), 'refresh', clientId, family, resource, scope, now + REFRESH_DAYS * 86400_000)
    this.db.prepare('UPDATE oauth_clients SET last_used_at = ? WHERE id = ?').run(nowIso(), clientId)
    return { access_token: access, token_type: 'Bearer', expires_in: ACCESS_SECONDS, refresh_token: refresh, scope }
  }

  /** RFC 7009: revoking either token of a sign-in ends all of it. Unknown tokens are not an error. */
  revoke(token: string | undefined): void {
    const row = token ? this.db.prepare('SELECT family FROM oauth_tokens WHERE hash = ?').get(sha256(token)) : undefined
    if (row) this.db.prepare('DELETE FROM oauth_tokens WHERE family = ?').run(row.family)
  }

  /** A live access token issued for this MCP server, or null. */
  verifyAccessToken(token: string | undefined): { clientId: string } | null {
    if (!token || token.length > 200) return null
    const row = this.db.prepare("SELECT client_id, resource, expires_at FROM oauth_tokens WHERE hash = ? AND kind = 'access'").get(sha256(token))
    if (!row || Number(row.expires_at) <= Date.now() || row.resource !== this.resource) return null
    return { clientId: String(row.client_id) }
  }

  // Metadata

  protectedResourceMetadata() {
    return { resource: this.resource, authorization_servers: [this.issuer], scopes_supported: [SCOPE], bearer_methods_supported: ['header'], resource_name: 'Artifacts' }
  }

  authorizationServerMetadata() {
    return {
      issuer: this.issuer,
      authorization_endpoint: `${this.issuer}/oauth/authorize`,
      token_endpoint: `${this.issuer}/oauth/token`,
      registration_endpoint: `${this.issuer}/oauth/register`,
      revocation_endpoint: `${this.issuer}/oauth/revoke`,
      response_types_supported: ['code'],
      grant_types_supported: ['authorization_code', 'refresh_token'],
      code_challenge_methods_supported: ['S256'],
      token_endpoint_auth_methods_supported: ['none'],
      revocation_endpoint_auth_methods_supported: ['none'],
      scopes_supported: [SCOPE],
      authorization_response_iss_parameter_supported: true
    }
  }

  /** The 401 a request to the MCP without a valid token gets. */
  challenge(invalid: boolean): string {
    return `Bearer resource_metadata="${this.issuer}/.well-known/oauth-protected-resource/mcp", scope="${SCOPE}"${invalid ? ', error="invalid_token"' : ''}`
  }
}

const CONSENT_CSP = "default-src 'none'; script-src 'self'; style-src 'self'; img-src 'self'; connect-src 'self'; frame-ancestors 'none'; base-uri 'none'; form-action 'none'"

function consentPage(clientName: string, redirectUri: string, requestId: string): string {
  const target = new URL(redirectUri)
  const where = target.protocol === 'https:' || target.protocol === 'http:' ? target.host : target.protocol.replace(/:$/, '')
  return `<!doctype html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex, nofollow">
<title>授權 · Artifacts</title>
<link rel="icon" href="/static/favicon.svg">
<link rel="stylesheet" href="/static/app.css">
</head>
<body class="auth-page">
<main class="auth-card" data-request="${escapeHtml(requestId)}">
  <h1>允許「${escapeHtml(clientName)}」使用 Artifacts?</h1>
  <p class="muted">它將能發佈、更新、刪除 artifact、改可見性、讀留言與回覆。授權完成後會回到 <strong>${escapeHtml(where)}</strong>。</p>
  <p class="muted small">之後可在「帳號」頁撤銷。</p>
  <div class="row-buttons">
    <button id="allow" class="button primary" type="button">允許</button>
    <button id="deny" class="button" type="button">拒絕</button>
  </div>
  <p id="status" class="muted small" role="status"></p>
</main>
<script type="module" src="/static/consent.js"></script>
</body>
</html>`
}

function errorPage(message: string): string {
  return `<!doctype html>
<html lang="zh-Hant">
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>無法授權</title><link rel="stylesheet" href="/static/app.css"></head>
<body class="auth-page"><main class="auth-card"><h1>無法授權</h1><p class="muted">${escapeHtml(message)}</p></main></body>
</html>`
}

/**
 * The OAuth routes on the public app. `signedIn` is the owner session check; the consent page and the
 * decision need it, everything else is for clients.
 */
export function mountOAuth(app: Hono, oauth: OAuthServer, signedIn: (c: Context) => boolean): void {
  const routes = new Hono()
  const cors = async (c: Context, next: () => Promise<void>) => {
    if (c.req.method === 'OPTIONS') {
      return c.body(null, 204, {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
        'Access-Control-Allow-Headers': 'authorization, content-type, mcp-protocol-version',
        'Access-Control-Max-Age': '600'
      })
    }
    await next()
    c.header('Access-Control-Allow-Origin', '*')
  }
  // Clients in a browser read these; none uses cookies. The consent page and every other route get no CORS.
  for (const path of ['/.well-known/*', '/oauth/token', '/oauth/register', '/oauth/revoke']) routes.use(path, cors)
  routes.use('/oauth/*', bodyLimit({ maxSize: 64 * 1024, onError: (c) => c.json({ error: 'invalid_request', error_description: 'Request body too large.' }, 413) }))

  for (const path of ['/.well-known/oauth-protected-resource', '/.well-known/oauth-protected-resource/mcp']) {
    routes.get(path, (c) => c.json(oauth.protectedResourceMetadata()))
  }
  for (const path of ['/.well-known/oauth-authorization-server', '/.well-known/oauth-authorization-server/mcp', '/.well-known/openid-configuration']) {
    routes.get(path, (c) => c.json(oauth.authorizationServerMetadata()))
  }

  const fields = async (c: Context): Promise<Record<string, string | undefined>> => {
    const type = c.req.header('content-type') ?? ''
    if (type.includes('application/json')) {
      const value: unknown = await c.req.json().catch(() => null)
      if (!value || typeof value !== 'object') return {}
      return Object.fromEntries(Object.entries(value).map(([key, item]) => [key, typeof item === 'string' ? item : undefined]))
    }
    const form = await c.req.parseBody().catch(() => ({}))
    return Object.fromEntries(Object.entries(form).map(([key, item]) => [key, typeof item === 'string' ? item : undefined]))
  }

  routes.post('/oauth/register', async (c) => {
    const value: unknown = await c.req.json().catch(() => null)
    if (!value || typeof value !== 'object' || Array.isArray(value)) return c.json({ error: 'invalid_client_metadata', error_description: 'JSON body required.' }, 400)
    return c.json(oauth.register(value as Record<string, unknown>), 201)
  })

  routes.get('/oauth/authorize', (c) => {
    if (!signedIn(c)) {
      const here = new URL(c.req.url)
      return c.redirect(`/login?next=${encodeURIComponent(here.pathname + here.search)}`, 302)
    }
    const request = oauth.authorizationRequest(c.req.query())
    if (!request.ok) return c.html(errorPage(request.message), 400)
    c.header('Content-Security-Policy', CONSENT_CSP)
    return c.html(consentPage(request.client.name, request.redirectUri, request.requestId))
  })

  routes.post('/oauth/authorize/decision', async (c) => {
    // The viewer header forces a preflight no other origin passes: a cross-site page cannot press "allow".
    if (c.req.header('x-artifact-client') !== 'viewer' || !signedIn(c)) return c.json({ error: 'access_denied' }, 403)
    const body: unknown = await c.req.json().catch(() => null)
    const { request, allow } = (body ?? {}) as { request?: unknown; allow?: unknown }
    const redirect = oauth.decide(request, allow === true)
    if (!redirect) return c.json({ error: 'invalid_request', error_description: '這個授權請求已過期，請從應用程式重新連線。' }, 400)
    return c.json({ redirect })
  })

  routes.post('/oauth/token', async (c) => {
    c.header('Cache-Control', 'no-store')
    return c.json(oauth.token(await fields(c)))
  })

  routes.post('/oauth/revoke', async (c) => {
    oauth.revoke((await fields(c)).token)
    return c.body(null, 200)
  })

  routes.onError((error, c) => {
    if (error instanceof OAuthError) {
      c.header('Cache-Control', 'no-store')
      return c.json({ error: error.code, error_description: error.message }, error.status)
    }
    console.error(error)
    return c.json({ error: 'server_error' }, 500)
  })

  app.route('/', routes)
}

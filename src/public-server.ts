import { readFileSync } from 'node:fs'
import { Hono, type Context } from 'hono'
import { bodyLimit } from 'hono/body-limit'
import { createAppServer, VIEWER_HEADER } from './app-server.ts'
import { ArtifactError, type ArtifactStore } from './artifact-store.ts'
import type { Config } from './config.ts'
import { escapeHtml } from './content-render.ts'
import { decodeRequestPath, isOfficePath } from './file-paths.ts'
import { mountMcp } from './mcp.ts'
import { mountOAuth, OAuthServer } from './oauth.ts'
import { readCookie, sessionCookie, SESSION_COOKIE, type OwnerAuth } from './owner-auth.ts'

/** Pages of the public app that run a script of their own (sign-in, account): own files only, no frames. */
const SCRIPT_PAGE_CSP = "default-src 'none'; script-src 'self'; style-src 'self'; img-src 'self'; connect-src 'self'; frame-ancestors 'none'; base-uri 'none'; form-action 'none'"

/** Where to go after signing in: a path on this app only, never another host. */
export function safeNext(value: string | undefined): string {
  // No backslash or control character anywhere: URL parsers drop tabs and newlines and read "\\" as "/".
  if (!value || value.length >= 2000 || !value.startsWith('/') || value.startsWith('//') || /[\\\u0000-\u001f\u007f]/.test(value)) return '/'
  const base = 'https://next.invalid'
  return new URL(value, base).origin === base ? value : '/'
}

function scriptPage(title: string, body: string, script: string): string {
  return `<!doctype html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex, nofollow">
<title>${escapeHtml(title)}</title>
<link rel="icon" href="/static/favicon.svg">
<link rel="stylesheet" href="/static/app.css">
</head>
<body class="auth-page">
${body}
<script type="module" src="/static/${script}"></script>
</body>
</html>`
}

/** Not found and not shown to a visitor who is signed out look the same, so neither tells which ids exist. */
function notFoundPage(next: string): string {
  return `<!doctype html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex, nofollow">
<title>找不到頁面</title>
<link rel="icon" href="/static/favicon.svg">
<link rel="stylesheet" href="/static/app.css">
</head>
<body class="auth-page">
<main class="auth-card">
  <h1>找不到這個頁面</h1>
  <p class="muted">它不存在、不公開，或需要登入才能看。</p>
  <a class="button primary" href="/login?next=${encodeURIComponent(next)}">擁有者登入</a>
</main>
</body>
</html>`
}

/**
 * The internet-facing app. Signed out, it is a read-only viewer for artifacts marked public and their
 * original files; everything else (gallery, list, search, comments, errors, diff, export, every write) is
 * absent, not guarded, and that page runs no script of its own. With `auth`, a valid owner session passes
 * one gate at the mount point into the full app (built for the public origins, with no publisher-token
 * path), and the sign-in routes are added. No route below the gate checks the session again.
 */
export function createPublicAppServer(store: ArtifactStore, config: Config, auth?: OwnerAuth): Hono {
  const pair = config.public
  if (!pair) throw new Error('createPublicAppServer needs config.public.')
  const app = new Hono()
  const staticRoot = new URL('../public/', import.meta.url)
  const ownerConfig: Config = { ...config, appOrigin: pair.appOrigin, contentOrigin: pair.contentOrigin, tokens: [], office: undefined }
  const ownerApp = auth ? createAppServer(store, ownerConfig, { owner: { auth } }) : null
  const signedIn = (c: Context): boolean => Boolean(auth?.isSession(readCookie(c, SESSION_COOKIE)))

  app.use('*', async (c, next) => {
    await next()
    // Never cacheable by Cloudflare; token responses keep their stricter no-store.
    if (!c.res.headers.get('Cache-Control')?.includes('no-store')) c.header('Cache-Control', 'private, no-cache')
    c.header('X-Robots-Tag', 'noindex, nofollow')
    c.header('X-Content-Type-Options', 'nosniff')
    c.header('Referrer-Policy', 'no-referrer')
    c.header('X-Frame-Options', 'DENY')
    if (!c.res.headers.has('Content-Security-Policy')) {
      c.header('Content-Security-Policy', "default-src 'none'; frame-ancestors 'none'; base-uri 'none'; form-action 'none'")
    }
  })

  app.get('/healthz', (c) => c.text('ok'))

  if (auth && ownerApp) {
    // The remote MCP and its OAuth server: bearer tokens only, so they sit before the session gate and read no cookie
    // (the consent page asks for the session itself).
    const oauth = new OAuthServer(auth.db, pair.appOrigin)
    mountOAuth(app, oauth, signedIn)
    mountMcp(app, store, config, oauth)

    // Every call that changes sign-in state needs the viewer header: a CORS preflight no other origin passes.
    const requireClient = (c: Context) => {
      if (c.req.header(VIEWER_HEADER) !== 'viewer') throw new ArtifactError(403, 'forbidden', 'Missing client header.')
    }
    const body = async (c: Context): Promise<Record<string, unknown>> => {
      const value: unknown = await c.req.json().catch(() => null)
      if (!value || typeof value !== 'object' || Array.isArray(value)) throw new ArtifactError(400, 'invalid_json', 'The request body must be a JSON object.')
      return value as Record<string, unknown>
    }
    app.use('/auth/*', bodyLimit({ maxSize: 64 * 1024, onError: (c) => c.json({ error: { code: 'body_too_large', message: 'Request body too large.' } }, 413) }))
    const startSession = (c: Context) => {
      c.header('Set-Cookie', sessionCookie(auth.createSession()))
    }

    app.get('/login', (c) => {
      const next = safeNext(c.req.query('next'))
      if (signedIn(c)) return c.redirect(next, 302)
      c.header('Content-Security-Policy', SCRIPT_PAGE_CSP)
      return c.html(scriptPage('登入 · Artifacts', `<main class="auth-card" data-mode="login" data-next="${escapeHtml(next)}">
  <h1>Artifacts</h1>
  <p class="muted">用你註冊過的 passkey 登入。</p>
  <button id="go" class="button primary" type="button">用 passkey 登入</button>
  <p id="status" class="muted small" role="status"></p>
</main>`, 'login.js'))
    })
    app.get('/enroll', (c) => {
      c.header('Content-Security-Policy', SCRIPT_PAGE_CSP)
      return c.html(scriptPage('註冊 passkey · Artifacts', `<main class="auth-card" data-mode="enroll" data-next="/">
  <h1>註冊 passkey</h1>
  <p class="muted">這台裝置會存一把 passkey，之後用它登入（手機上註冊的可以掃碼在其他電腦登入）。</p>
  <label class="field">名稱<input id="name" type="text" maxlength="60" placeholder="例如 iPhone" autocomplete="off"></label>
  <button id="go" class="button primary" type="button">註冊</button>
  <p id="status" class="muted small" role="status"></p>
</main>`, 'login.js'))
    })
    app.post('/auth/login/options', async (c) => {
      requireClient(c)
      return c.json(await auth.loginOptions())
    })
    app.post('/auth/login/verify', async (c) => {
      requireClient(c)
      const input = await body(c)
      await auth.login(input.flow, input.response)
      startSession(c)
      return c.json({ ok: true })
    })
    app.post('/auth/register/options', async (c) => {
      requireClient(c)
      const input = await body(c)
      return c.json(await auth.registrationOptions({ code: typeof input.code === 'string' ? input.code : undefined, signedIn: signedIn(c) }))
    })
    app.post('/auth/register/verify', async (c) => {
      requireClient(c)
      const input = await body(c)
      const already = signedIn(c)
      await auth.register(input.flow, input.response, input.name, already)
      if (!already) startSession(c)
      return c.json({ ok: true })
    })
    app.post('/auth/logout', (c) => {
      requireClient(c)
      auth.endSession(readCookie(c, SESSION_COOKIE))
      c.header('Set-Cookie', sessionCookie('', 0))
      return c.json({ ok: true })
    })

    // The owner's own pages and calls: account settings, then everything of the full app.
    const owner = new Hono()
    owner.get('/account', (c) => {
      c.header('Content-Security-Policy', SCRIPT_PAGE_CSP)
      return c.html(scriptPage('帳號 · Artifacts', `<main class="auth-card account">
  <p><a href="/">← 所有 Artifacts</a></p>
  <h1>帳號</h1>
  <h2>Passkey</h2>
  <ul id="passkeys" class="plain-list"></ul>
  <button id="add-passkey" class="button" type="button">在這台裝置新增 passkey</button>
  <h2>已授權的應用程式（MCP）</h2>
  <ul id="clients" class="plain-list"></ul>
  <h2>登入中的裝置</h2>
  <p class="muted"><span id="sessions"></span> 個工作階段</p>
  <div class="row-buttons">
    <button id="logout" class="button" type="button">登出</button>
    <button id="logout-all" class="button danger" type="button">登出所有裝置</button>
  </div>
  <p id="status" class="muted small" role="status"></p>
</main>`, 'account.js'))
    })
    owner.get('/api/account', (c) => c.json({ passkeys: auth.listPasskeys(), sessions: auth.sessionCount(), clients: oauth.grantedClients() }))
    owner.delete('/api/account/clients/:id', (c) => {
      requireClient(c)
      oauth.revokeClient(c.req.param('id'))
      return c.body(null, 204)
    })
    owner.delete('/api/account/passkeys/:id', (c) => {
      requireClient(c)
      if (auth.listPasskeys().length <= 1) throw new ArtifactError(409, 'last_passkey', '不能刪掉最後一把 passkey。')
      auth.removePasskey(c.req.param('id'))
      return c.body(null, 204)
    })
    owner.post('/auth/logout-all', (c) => {
      requireClient(c)
      auth.endAllSessions()
      c.header('Set-Cookie', sessionCookie('', 0))
      return c.json({ ok: true })
    })
    owner.onError((error, c) => {
      if (error instanceof ArtifactError) return c.json({ error: { code: error.code, message: error.message } }, error.status)
      console.error(error)
      return c.json({ error: { code: 'internal', message: 'Internal error' } }, 500)
    })
    owner.all('*', (c) => ownerApp.fetch(c.req.raw))

    // The one gate: a valid session reaches the owner's routes; anything else falls through to the anonymous ones.
    app.all('*', async (c, next) => {
      if (signedIn(c)) return owner.fetch(c.req.raw)
      await next()
    })
  }

  // Only an artifact the publisher made public, and only its current version (drafts stay on the tailnet).
  const publicVersion = (id: string, raw: string | undefined) => {
    if (!store.isPublic(id)) throw new ArtifactError(404, 'not_found', 'Not found.')
    const summary = store.requireSummary(id)
    if (raw !== undefined && raw !== String(summary.currentVersion)) throw new ArtifactError(404, 'not_found', 'Not found.')
    return { summary, row: store.requireVersion(id, summary.currentVersion) }
  }

  app.get('/a/:id', (c) => {
    const id = c.req.param('id')
    const { summary, row } = publicVersion(id, c.req.query('v'))
    // A file of this version the reader was on (sent here by the content origin's redirect); anything else is ignored.
    const requested = c.req.query('path')
    const path = requested !== undefined && Object.hasOwn(row.files, requested) ? requested : row.entry
    const encode = (value: string) => value.split('/').map(encodeURIComponent).join('/')
    const frameSrc = `${pair.contentOrigin}/a/${id}/v/${row.version}/${path === row.entry && !path.includes('/') ? '' : encode(path)}`
    const original = /\.(md|markdown)$/i.test(row.entry) && !/\.slides\.md$/i.test(row.entry) ? 'MD' : isOfficePath(row.entry) ? row.entry.split('.').pop()!.toUpperCase() : null
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
  <span class="public-updated">更新於 ${row.createdAt.slice(0, 10)}</span>
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

  const STATIC: Record<string, string> = {
    'app.css': 'text/css; charset=utf-8',
    'favicon.svg': 'image/svg+xml',
    ...(auth ? { 'login.js': 'text/javascript; charset=utf-8', 'webauthn.js': 'text/javascript; charset=utf-8', 'consent.js': 'text/javascript; charset=utf-8' } : {})
  }
  app.get('/static/:name', (c) => {
    const type = STATIC[c.req.param('name')]
    if (!type) return c.text('Not found', 404)
    return c.body(new Uint8Array(readFileSync(new URL(c.req.param('name'), staticRoot))), 200, { 'Content-Type': type })
  })

  app.onError((error, c) => {
    // Sign-in calls answer in JSON so the page can say what went wrong; nothing else says more than "not found".
    if (error instanceof ArtifactError && c.req.path.startsWith('/auth/')) return c.json({ error: { code: error.code, message: error.message } }, error.status)
    if (error instanceof ArtifactError && error.status < 500) return notFound(c)
    console.error(error)
    return c.text('Internal error', 500)
  })
  const notFound = (c: Context) => {
    const path = c.req.path
    if (auth && c.req.method === 'GET' && (path === '/' || path.startsWith('/a/'))) {
      return c.html(notFoundPage(new URL(c.req.url).pathname + new URL(c.req.url).search), 404)
    }
    return c.text('Not found', 404)
  }
  app.notFound(notFound)
  return app
}

/** The config the public content listener runs with: the same server, its own pair of origins. */
export function publicContentConfig(config: Config): Config {
  if (!config.public) throw new Error('publicContentConfig needs config.public.')
  return { ...config, appOrigin: config.public.appOrigin, contentOrigin: config.public.contentOrigin }
}

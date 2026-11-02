import { createHash, createHmac, randomBytes, timingSafeEqual } from 'node:crypto'
import { existsSync, mkdirSync, readFileSync, readdirSync, renameSync, rmSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import type { Context, Hono } from 'hono'
import { ArtifactError, LIMITS, type ArtifactStore } from './artifact-store.ts'
import type { Config } from './config.ts'
import { officeEditorPage } from './pages.ts'
import { ZipReader } from './zip.ts'
import { parseXml } from './xml.ts'

type Session = {
  id: string; artifactId: string; entry: string; title: string; baseVersion: number
  status: 'editing' | 'saved' | 'conflict' | 'error' | 'closed'; closed: boolean
  savedVersion?: number; resultHash?: string; message?: string; createdAt: number
  /** Set once Office fetched the source; a session it never opened holds no edits. */
  opened?: boolean
}
type Settings = NonNullable<Config['office']>
const DAY = 86_400_000
const MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex')
const object = (value: unknown): value is Record<string, unknown> => !!value && typeof value === 'object' && !Array.isArray(value)

function sign(value: unknown, secret: string): string {
  const unsigned = `${Buffer.from('{"alg":"HS256","typ":"JWT"}').toString('base64url')}.${Buffer.from(JSON.stringify(value)).toString('base64url')}`
  return `${unsigned}.${createHmac('sha256', secret).update(unsigned).digest('base64url')}`
}

function verify(token: string, secret: string): Record<string, unknown> {
  try {
    const [header, body, signature, extra] = token.split('.')
    if (extra || !header || !body || !signature || token.length > 65_536) throw new Error()
    if (JSON.parse(Buffer.from(header, 'base64url').toString()).alg !== 'HS256') throw new Error()
    const expected = createHmac('sha256', secret).update(`${header}.${body}`).digest()
    const actual = Buffer.from(signature, 'base64url')
    if (actual.length !== expected.length || !timingSafeEqual(actual, expected)) throw new Error()
    const value: unknown = JSON.parse(Buffer.from(body, 'base64url').toString())
    if (!object(value) || (value.exp !== undefined && (typeof value.exp !== 'number' || value.exp * 1000 <= Date.now()))) throw new Error()
    return value
  } catch { throw new ArtifactError(401, 'office_unauthorized', 'Office 授權無效或已到期。') }
}

/** Durable working copies are separate from immutable blobs; only store.update publishes their result. */
class OfficeEdits {
  private readonly root: string
  private readonly saves = new Map<string, Promise<void>>()
  constructor(privateStore: ArtifactStore, settings: Settings, dataDir: string) {
    this.store = privateStore
    this.settings = settings
    this.root = join(dataDir, 'office-edits')
    mkdirSync(this.root, { recursive: true, mode: 0o700 })
  }
  private readonly store: ArtifactStore
  private readonly settings: Settings
  private file(id: string, name: string): string {
    if (!/^[A-Za-z0-9_-]{43}$/.test(id)) throw new ArtifactError(404, 'office_session_not_found', '找不到這次編輯。')
    return join(this.root, id, name)
  }
  read(id: string, artifactId?: string): Session {
    const file = this.file(id, 'session.json')
    if (!existsSync(file)) throw new ArtifactError(404, 'office_session_not_found', '找不到這次編輯。')
    const session = JSON.parse(readFileSync(file, 'utf8')) as Session
    if (artifactId && session.artifactId !== artifactId) throw new ArtifactError(404, 'office_session_not_found', '找不到這次編輯。')
    this.store.requireSummary(session.artifactId)
    return session
  }
  private write(id: string, name: string, bytes: string | Buffer): void {
    const file = this.file(id, name)
    writeFileSync(`${file}.tmp`, bytes, { mode: 0o600 })
    renameSync(`${file}.tmp`, file)
  }
  private remember(session: Session): void { this.write(session.id, 'session.json', JSON.stringify(session)) }
  start(artifactId: string, baseVersion: unknown): Session {
    const summary = this.store.requireSummary(artifactId)
    if (summary.capabilities.length) throw new ArtifactError(403, 'office_not_editable', '宣告互動能力的簡報只能由發佈端更新。')
    if (!/\.pptx$/i.test(summary.entry)) throw new ArtifactError(409, 'office_not_editable', '網頁物件編輯只支援 .pptx。')
    if (baseVersion !== summary.currentVersion) throw new ArtifactError(409, 'version_conflict', '簡報已有新版，請回檢視頁重新開啟編輯。')
    this.prune()
    if (this.sessionDirs().length >= 200) throw new ArtifactError(409, 'office_sessions_full', '編輯副本已達上限，請先由管理者取回並清理未保存的結果。')
    const row = this.store.requireVersion(artifactId, summary.currentVersion)
    const id = randomBytes(32).toString('base64url')
    const session: Session = { id, artifactId, entry: row.entry, title: row.title, baseVersion: row.version, status: 'editing', closed: false, createdAt: Date.now() }
    mkdirSync(join(this.root, id), { mode: 0o700 })
    try {
      this.write(id, 'source.pptx', this.store.blobs.read(row.files[row.entry].sha256))
      this.remember(session)
    } catch (error) { rmSync(join(this.root, id), { recursive: true }); throw error }
    return session
  }
  private sessionDirs(): string[] {
    return readdirSync(this.root, { withFileTypes: true }).filter((entry) => entry.isDirectory() && /^[A-Za-z0-9_-]{43}$/.test(entry.name)).map((entry) => entry.name)
  }
  /** Copies may go after a day once nothing in them is unpublished: saved, closed unchanged, or never opened by Office. */
  private prune(): void {
    for (const id of this.sessionDirs()) {
      const dir = join(this.root, id)
      if (!existsSync(join(dir, 'session.json'))) {
        // start() writes session.json last and before any result, so this is a start interrupted mid-copy.
        if (!existsSync(join(dir, 'result.pptx'))) rmSync(dir, { recursive: true, force: true })
        continue
      }
      let session: Session
      try { session = JSON.parse(readFileSync(join(dir, 'session.json'), 'utf8')) as Session } catch { continue }
      if (Date.now() - session.createdAt <= DAY) continue
      const done = session.closed && ['saved', 'closed'].includes(session.status)
      const neverOpened = session.status === 'editing' && !session.opened && !session.resultHash
      if (done || neverOpened) rmSync(dir, { recursive: true, force: true })
    }
  }
  source(id: string): Buffer {
    const session = this.read(id)
    if (!session.opened) { session.opened = true; this.remember(session) }
    return readFileSync(this.file(id, 'source.pptx'))
  }
  result(id: string): Buffer {
    const session = this.read(id)
    if (!session.resultHash) throw new ArtifactError(404, 'office_result_not_found', '尚未收到編輯結果。')
    return readFileSync(this.file(id, 'result.pptx'))
  }
  config(session: Session): Record<string, unknown> {
    if (session.closed) throw new ArtifactError(409, 'office_session_closed', '這次編輯已結束，請從檢視頁重新開啟。')
    const sourceToken = sign({ scope: 'source', sessionId: session.id, exp: Math.floor(Date.now() / 1000) + 86400 }, this.settings.secret)
    const config = {
      documentType: 'slide', type: 'desktop', width: '100%', height: '100%',
      document: { fileType: 'pptx', key: session.id, title: session.entry.split('/').pop(),
        url: `${this.settings.callbackOrigin}/api/office/source/${session.id}?token=${sourceToken}`,
        permissions: { edit: true, download: true, print: true } },
      editorConfig: { mode: 'edit', lang: 'zh-TW', callbackUrl: `${this.settings.callbackOrigin}/api/office/callback/${session.id}`,
        customization: { forcesave: true, autosave: true, macrosMode: 'disable', plugins: false, anonymous: { request: false }, goback: { requestClose: true, text: '返回檢視' } } }
    }
    return { ...config, token: sign(config, this.settings.secret) }
  }
  async callback(id: string, payload: Record<string, unknown>): Promise<void> {
    // Force-save and close callbacks can overlap. Fetch and commit each session in arrival order.
    const previous = this.saves.get(id) ?? Promise.resolve()
    const next = previous.catch(() => {}).then(() => this.applyCallback(id, payload))
    this.saves.set(id, next)
    try { await next } finally { if (this.saves.get(id) === next) this.saves.delete(id) }
  }
  private async applyCallback(id: string, payload: Record<string, unknown>): Promise<void> {
    const session = this.read(id)
    if (payload.key !== id || ![1, 2, 3, 4, 6, 7].includes(Number(payload.status))) throw new ArtifactError(400, 'office_callback_invalid', 'Office 回呼內容無效。')
    const status = Number(payload.status)
    if (status === 1) {
      // ONLYOFFICE sends 4 on a no-change disconnect, then 1 if that client reconnects.
      if (session.status === 'closed') { session.closed = false; session.status = 'editing'; this.remember(session) }
      return
    }
    if (status === 4) { session.closed = true; if (!['conflict', 'error'].includes(session.status)) session.status = 'closed'; this.remember(session); return }
    const saveError = status === 3 || status === 7
    if (saveError) {
      session.status = 'error'
      session.closed ||= status === 3
      session.message = 'Office 回報保存失敗，請取回編輯結果並保留編輯頁。'
      this.remember(session)
      if (typeof payload.url !== 'string') return
    }
    let url: URL
    try { url = new URL(String(payload.url)) } catch { throw new ArtifactError(400, 'office_callback_invalid', 'Office 未提供有效的保存網址。') }
    if (![this.settings.url, this.settings.internalUrl].includes(url.origin) || url.username || url.password) throw new ArtifactError(403, 'office_download_forbidden', '保存結果必須來自設定的 Office 服務。')
    const target = `${this.settings.internalUrl}${url.pathname}${url.search}`
    let bytes: Buffer
    try {
      const response = await fetch(target, { redirect: 'error', signal: AbortSignal.timeout(30_000) })
      if (!response.ok || !response.body || Number(response.headers.get('content-length')) > LIMITS.publishBytes) throw new Error()
      const reader = response.body.getReader()
      const parts: Uint8Array[] = []
      let size = 0
      for (;;) {
        const { value, done } = await reader.read()
        if (done) break
        size += value.length
        if (size > LIMITS.publishBytes) { await reader.cancel(); throw new Error() }
        parts.push(value)
      }
      bytes = Buffer.concat(parts)
      const pkg = new ZipReader(bytes)
      const presentation = pkg.text('ppt/presentation.xml')
      if (!presentation) throw new Error()
      parseXml(presentation)
    } catch {
      session.status = 'error'
      session.message = '尚未成功取回有效的 PPTX，請保留編輯頁與 Office 副本。'
      this.remember(session)
      throw new ArtifactError(400, 'office_download_failed', session.message)
    }
    const resultHash = hash(bytes)
    // Preserve the result before committing, including conflicts and a crash between the two writes.
    this.write(id, 'result.pptx', bytes)
    session.closed ||= status === 2
    if (saveError) { session.resultHash = resultHash; this.remember(session); return }
    try {
      const current = this.store.requireSummary(session.artifactId)
      if (session.status === 'conflict') throw new ArtifactError(409, 'version_conflict', '這次編輯已有版本衝突。')
      if (current.capabilities.length || current.entry !== session.entry) throw new ArtifactError(409, 'office_not_editable', '簡報的能力或入口已變更。')
      const latest = this.store.requireVersion(session.artifactId, current.currentVersion)
      if (latest.files[session.entry]?.sha256 !== resultHash) {
        const saved = this.store.update(session.artifactId, { baseVersion: session.baseVersion, files: { [session.entry]: { content: bytes } }, label: 'PowerPoint 網頁編輯' })
        session.baseVersion = saved.currentVersion
      } else session.baseVersion = current.currentVersion
      session.savedVersion = session.baseVersion
      session.status = 'saved'
      delete session.message
    } catch (error) {
      if (!(error instanceof ArtifactError)) throw error
      session.status = 'conflict'
      session.message = '已有其他變更，編輯結果已保留，請下載後合併再發佈。'
    }
    session.resultHash = resultHash
    this.remember(session)
  }
}

export function mountOfficeEditor(app: Hono, store: ArtifactStore, config: Config, requireCaller: (c: Context, id: string) => unknown): void {
  if (!config.office) return
  const settings = config.office
  const edits = new OfficeEdits(store, settings, config.dataDir)
  const publicState = (session: Session) => ({ status: session.status, closed: session.closed, savedVersion: session.savedVersion, message: session.message,
    ...(session.resultHash ? { downloadUrl: `/api/artifacts/${session.artifactId}/office-edits/${session.id}/result` } : {}) })
  const download = (c: Context, bytes: Buffer) => c.body(new Uint8Array(bytes), 200, { 'Content-Type': MIME, 'Content-Disposition': 'attachment; filename="edited.pptx"', 'Cache-Control': 'no-store' })
  app.use('/api/office/*', async (c, next) => { c.header('Cache-Control', 'no-store'); await next() })
  app.use('/api/artifacts/:id/office-edits/*', async (c, next) => { c.header('Cache-Control', 'no-store'); await next() })
  app.post('/api/artifacts/:id/office-edits', async (c) => {
    const body: unknown = await c.req.json().catch(() => null)
    if (!object(body)) throw new ArtifactError(400, 'invalid_json', '請提供有效的 JSON 物件。')
    const id = c.req.param('id')
    requireCaller(c, id)
    const session = edits.start(id, body.baseVersion)
    return c.json({ sessionId: session.id, editorUrl: `/a/${id}/edit/${session.id}` }, 201)
  })
  app.get('/a/:id/edit/:session', (c) => {
    const session = edits.read(c.req.param('session'), c.req.param('id'))
    c.header('Cache-Control', 'no-store')
    c.header('Content-Security-Policy', `default-src 'self'; script-src 'self' ${settings.url}; frame-src ${settings.url}; style-src 'self' 'unsafe-inline'; img-src 'self' data:; object-src 'none'; base-uri 'none'; frame-ancestors 'none'`)
    return c.html(officeEditorPage(session, settings.url))
  })
  app.get('/api/artifacts/:id/office-edits/:session', (c) => c.json(publicState(edits.read(c.req.param('session'), c.req.param('id')))))
  app.get('/api/artifacts/:id/office-edits/:session/config', (c) => c.json(edits.config(edits.read(c.req.param('session'), c.req.param('id')))))
  app.get('/api/artifacts/:id/office-edits/:session/result', (c) => { edits.read(c.req.param('session'), c.req.param('id')); return download(c, edits.result(c.req.param('session'))) })
  app.get('/api/office/source/:session', (c) => {
    const value = verify(c.req.query('token') ?? '', settings.secret)
    if (value.scope !== 'source' || value.sessionId !== c.req.param('session')) throw new ArtifactError(401, 'office_unauthorized', 'Office 來源授權無效。')
    return download(c, edits.source(c.req.param('session')))
  })
  app.post('/api/office/callback/:session', async (c) => {
    const body: unknown = await c.req.json().catch(() => null)
    const token = c.req.header('authorization')?.replace(/^Bearer\s+/i, '') ?? (object(body) && typeof body.token === 'string' ? body.token : '')
    const verified = verify(token, settings.secret)
    if (verified.operation !== undefined && verified.operation !== 'callback') throw new ArtifactError(401, 'office_unauthorized', 'Office 回呼授權用途不符。')
    const payload = object(verified.payload) ? verified.payload : verified
    await edits.callback(c.req.param('session'), payload)
    return c.json({ error: 0 })
  })
}

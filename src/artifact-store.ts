import { randomBytes } from 'node:crypto'
import { mkdirSync } from 'node:fs'
import { join } from 'node:path'
import { DatabaseSync } from 'node:sqlite'
import { BlobStore } from './blob-store.ts'
import { contentTypeFor, isHtmlPath, isMarkdownPath, normalizeFilePath } from './file-paths.ts'
import { diffLines, hunks, type ShownLine } from './text-diff.ts'

export const CAPABILITIES = ['db', 'uploads'] as const
export type Capability = (typeof CAPABILITIES)[number]

export const LIMITS = {
  fileBytes: 16 * 1024 * 1024,
  publishBytes: 64 * 1024 * 1024,
  filesPerVersion: 511,
  assetBytes: 15 * 1024 * 1024,
  docBytes: 256 * 1024,
  docsPerArtifact: 50_000,
  commentChars: 10_000,
  errorsPerReport: 20,
  errorsPerVersion: 100,
  errorChars: 500
}

export type FileEntry = { sha256: string; size: number; type: string }
export type FileMap = Record<string, FileEntry>

/** Where a version was published from: the machine and the folder (home shown as ~). */
export type Source = { host: string; path: string }

export type VersionRow = {
  version: number
  label: string | null
  source: Source | null
  title: string
  description: string | null
  entry: string
  files: FileMap
  createdAt: string
}

export type ArtifactSummary = {
  id: string
  title: string
  description: string | null
  icon: string | null
  entry: string
  currentVersion: number
  capabilities: Capability[]
  pinned: boolean
  source: Source | null
  createdAt: string
  updatedAt: string
}

export type FileChange = {
  path: string
  status: 'added' | 'removed' | 'modified'
  /** Unified lines; absent when the file is not text, too large, or rewritten past showing (see note). */
  lines?: ShownLine[]
  note?: 'binary' | 'too_large' | 'rewritten'
}

export type VersionDiff = {
  from: number | null
  to: number
  label: string | null
  source: Source | null
  createdAt: string
  files: FileChange[]
}

const DIFF_LIMITS = { bytes: 1024 * 1024, lines: 20_000, edits: 2000, shown: 1000, lineChars: 2000 }
const isTextType = (type: string) => /^text\/|json|javascript|xml|svg/.test(type)

export type ArtifactListing = ArtifactSummary & { snippet?: string | null }

export type ArtifactDetail = ArtifactSummary & { versions: Omit<VersionRow, 'files'>[] }

export type AssetRow = {
  id: string
  name: string
  sha256: string
  size: number
  type: string
  source: 'publisher' | 'page'
  createdAt: string
}

export type DocRow = { id: string; data: unknown; createdAt: string; updatedAt: string }

export const PAGE_ERROR_KINDS = ['error', 'rejection', 'resource', 'csp'] as const

export type PageErrorRow = {
  version: number
  path: string
  kind: (typeof PAGE_ERROR_KINDS)[number]
  message: string
  source: string | null
  count: number
  firstAt: string
  lastAt: string
}

export type CommentRow = {
  id: string
  parentId: string | null
  version: number
  author: string
  body: string
  resolved: boolean
  createdAt: string
}

/** Thrown for caller mistakes; the HTTP layer maps `status` straight onto the response. */
export class ArtifactError extends Error {
  readonly status: 400 | 401 | 403 | 404 | 409 | 413
  readonly code: string
  readonly details?: Record<string, unknown>

  constructor(status: ArtifactError['status'], code: string, message: string, details?: Record<string, unknown>) {
    super(message)
    this.status = status
    this.code = code
    this.details = details
  }
}

export type IncomingFile = { content: Buffer } | null

export type PublishInput = {
  title: string
  description?: string | null
  icon?: string | null
  entry?: string
  label?: string | null
  capabilities?: string[]
  source?: unknown
  files: Record<string, IncomingFile>
}

export type UpdateInput = Partial<Omit<PublishInput, 'files'>> & {
  files?: Record<string, IncomingFile>
  baseVersion?: number
  force?: boolean
}

export type ArtifactChange =
  | { kind: 'version'; version: number }
  | { kind: 'deleted' }
  | { kind: 'db'; collection: string; docId: string; deleted: boolean }
  | { kind: 'comments' }
  | { kind: 'assets' }
  | { kind: 'errors'; version: number }
  | { kind: 'pruned'; versions: number[] }

const ID_ALPHABET = 'abcdefghijkmnopqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789'

/** Unguessable, URL-safe id; rejection sampling keeps every character uniform. */
function newId(length: number): string {
  let id = ''
  while (id.length < length) {
    for (const byte of randomBytes(length * 2)) {
      if (byte < 256 - (256 % ID_ALPHABET.length) && id.length < length) {
        id += ID_ALPHABET[byte % ID_ALPHABET.length]
      }
    }
  }
  return id
}

function now(): string {
  return new Date().toISOString()
}

function parseCapabilities(input: string[] | undefined): Capability[] {
  const values = input ?? []
  const unknown = values.filter((value) => !(CAPABILITIES as readonly string[]).includes(value))
  if (unknown.length > 0) {
    throw new ArtifactError(400, 'unknown_capability', `Unknown capabilities: ${unknown.join(', ')}. Known: ${CAPABILITIES.join(', ')}.`)
  }
  return [...new Set(values)].sort() as Capability[]
}

function requireText(value: unknown, field: string, max: number): string {
  if (typeof value !== 'string' || value.trim().length === 0) {
    throw new ArtifactError(400, 'invalid_field', `${field} is required.`)
  }
  if (value.length > max) {
    throw new ArtifactError(400, 'invalid_field', `${field} is longer than ${max} characters.`)
  }
  return value.trim()
}

function optionalText(value: unknown, field: string, max: number): string | null {
  if (value === undefined || value === null || value === '') {
    return null
  }
  return requireText(value, field, max)
}

/** Truncates to an integer in [min, max]; anything non-finite falls back to the default. */
function clampInt(value: number | undefined, fallback: number, min: number, max: number): number {
  const number = typeof value === 'number' && Number.isFinite(value) ? Math.trunc(value) : fallback
  return Math.min(Math.max(number, min), max)
}

function parseSource(value: unknown): Source | null {
  if (value === undefined || value === null) {
    return null
  }
  const { host, path } = value as { host?: unknown; path?: unknown }
  if (typeof value !== 'object' || typeof host !== 'string' || typeof path !== 'string' || !host || host.length > 100 || path.length > 500) {
    throw new ArtifactError(400, 'invalid_field', 'source is { host, path }: strings of at most 100 and 500 characters.')
  }
  return { host, path }
}

const ENTITIES: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ' }
const SEARCH_CHARS = 500_000

/** The words a reader sees: markup, scripts, styles and link targets dropped, entities decoded. */
function pageText(path: string, content: string): string {
  // Markdown: line-leading markers and emphasis go, but - and _ inside words stay (nas-artifacts, search_text).
  const text = isMarkdownPath(path)
    ? content
        .replace(/!?\[([^\]]*)\]\([^)]*\)/g, '$1')
        .replace(/<[^>]+>/g, ' ')
        .replace(/^[ \t]*(?:[#>*+-]+|\d+\.)[ \t]+/gm, ' ')
        .replace(/[*`~]+|(?<![\p{L}\p{N}])[_-]+|[_-]+(?![\p{L}\p{N}])/gu, '')
        .replace(/\|/g, ' ')
    : content
        .replace(/<(script|style|template)\b[\s\S]*?<\/\1\s*>/gi, ' ')
        .replace(/<!--[\s\S]*?-->/g, ' ')
        .replace(/<[^>]+>/g, ' ')
  return text
    .replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (match, name: string) => {
      if (name[0] === '#') {
        const code = name[1] === 'x' || name[1] === 'X' ? parseInt(name.slice(2), 16) : Number(name.slice(1))
        return code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : ' '
      }
      return ENTITIES[name.toLowerCase()] ?? match
    })
    .replace(/\s+/g, ' ')
    .trim()
}

/** About 40 characters either side of the first hit. */
function snippetAround(text: string, query: string): string | null {
  const at = text.toLowerCase().indexOf(query.toLowerCase())
  if (at < 0) {
    return null
  }
  const start = Math.max(0, at - 40)
  const end = Math.min(text.length, at + query.length + 40)
  return `${start > 0 ? '…' : ''}${text.slice(start, end)}${end < text.length ? '…' : ''}`
}

const sourceFromRow = (row: Record<string, unknown>): Source | null =>
  row.source_host === null || row.source_host === undefined ? null : { host: String(row.source_host), path: String(row.source_path ?? '') }

function chooseEntry(files: FileMap, requested: string | undefined): string {
  if (requested !== undefined) {
    const entry = normalizeFilePath(requested)
    if (!entry || !files[entry]) {
      throw new ArtifactError(400, 'invalid_entry', `Entry "${requested}" is not one of the published files.`)
    }
    if (!isHtmlPath(entry) && !isMarkdownPath(entry)) {
      throw new ArtifactError(400, 'invalid_entry', 'The entry must be an .html or .md file.')
    }
    return entry
  }
  for (const candidate of ['index.html', 'index.md', 'index.slides.md', 'README.md']) {
    if (files[candidate]) {
      return candidate
    }
  }
  const pages = Object.keys(files).filter((path) => isHtmlPath(path) || isMarkdownPath(path))
  if (pages.length === 1) {
    return pages[0]
  }
  throw new ArtifactError(400, 'invalid_entry', 'Pass entry: no index.html/index.md and not exactly one page among the files.')
}

export class ArtifactStore {
  readonly blobs: BlobStore
  private readonly db: DatabaseSync
  private readonly listeners = new Map<string, Set<(change: ArtifactChange) => void>>()
  private pendingChanges: [string, ArtifactChange][] | null = null

  constructor(dataDir: string) {
    mkdirSync(dataDir, { recursive: true })
    this.blobs = new BlobStore(dataDir)
    this.db = new DatabaseSync(join(dataDir, 'artifacts.db'))
    this.db.exec(`
      PRAGMA journal_mode = WAL;
      PRAGMA foreign_keys = ON;
      CREATE TABLE IF NOT EXISTS artifacts (
        id TEXT PRIMARY KEY,
        title TEXT NOT NULL,
        description TEXT,
        icon TEXT,
        entry TEXT NOT NULL,
        current_version INTEGER NOT NULL,
        capabilities TEXT NOT NULL DEFAULT '[]',
        pinned INTEGER NOT NULL DEFAULT 0,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL
      );
      CREATE TABLE IF NOT EXISTS versions (
        artifact_id TEXT NOT NULL REFERENCES artifacts(id) ON DELETE CASCADE,
        version INTEGER NOT NULL,
        label TEXT,
        title TEXT NOT NULL,
        description TEXT,
        entry TEXT NOT NULL,
        files TEXT NOT NULL,
        created_at TEXT NOT NULL,
        PRIMARY KEY (artifact_id, version)
      );
      CREATE TABLE IF NOT EXISTS assets (
        artifact_id TEXT NOT NULL REFERENCES artifacts(id) ON DELETE CASCADE,
        id TEXT NOT NULL,
        name TEXT NOT NULL,
        sha256 TEXT NOT NULL,
        size INTEGER NOT NULL,
        type TEXT NOT NULL,
        source TEXT NOT NULL,
        created_at TEXT NOT NULL,
        PRIMARY KEY (artifact_id, id)
      );
      CREATE TABLE IF NOT EXISTS docs (
        artifact_id TEXT NOT NULL REFERENCES artifacts(id) ON DELETE CASCADE,
        collection TEXT NOT NULL,
        doc_id TEXT NOT NULL,
        data TEXT NOT NULL,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL,
        PRIMARY KEY (artifact_id, collection, doc_id)
      );
      CREATE TABLE IF NOT EXISTS comments (
        id TEXT PRIMARY KEY,
        artifact_id TEXT NOT NULL REFERENCES artifacts(id) ON DELETE CASCADE,
        parent_id TEXT,
        version INTEGER NOT NULL,
        author TEXT NOT NULL,
        body TEXT NOT NULL,
        resolved INTEGER NOT NULL DEFAULT 0,
        created_at TEXT NOT NULL
      );
      CREATE INDEX IF NOT EXISTS comments_by_artifact ON comments(artifact_id, created_at);
      CREATE TABLE IF NOT EXISTS page_errors (
        artifact_id TEXT NOT NULL REFERENCES artifacts(id) ON DELETE CASCADE,
        version INTEGER NOT NULL,
        path TEXT NOT NULL,
        kind TEXT NOT NULL,
        message TEXT NOT NULL,
        source TEXT NOT NULL,
        count INTEGER NOT NULL,
        first_at TEXT NOT NULL,
        last_at TEXT NOT NULL,
        PRIMARY KEY (artifact_id, version, path, kind, message, source)
      );
    `)
    for (const table of ['artifacts', 'versions']) {
      this.addColumn(table, 'source_host', 'TEXT')
      this.addColumn(table, 'source_path', 'TEXT')
    }
    this.db.exec(`CREATE TABLE IF NOT EXISTS search_text (
      artifact_id TEXT PRIMARY KEY REFERENCES artifacts(id) ON DELETE CASCADE,
      version INTEGER NOT NULL,
      text TEXT NOT NULL
    )`)
    // Artifacts published before search existed, or whose index fell behind, are indexed now.
    for (const row of this.db
      .prepare('SELECT a.id, a.current_version FROM artifacts a LEFT JOIN search_text t ON t.artifact_id = a.id WHERE t.version IS NOT a.current_version')
      .all()) {
      // Search is a convenience: an artifact it cannot read stays unsearchable rather than stopping the server.
      try {
        this.indexText(String(row.id), Number(row.current_version))
      } catch (error) {
        console.error(`search: cannot index ${String(row.id)}:`, error instanceof Error ? error.message : error)
      }
    }
  }

  /** Columns added after the first release; CREATE TABLE IF NOT EXISTS leaves existing tables as they were. */
  private addColumn(table: string, column: string, type: string): void {
    const columns = this.db.prepare(`PRAGMA table_info(${table})`).all().map((row) => String(row.name))
    if (!columns.includes(column)) {
      this.db.exec(`ALTER TABLE ${table} ADD COLUMN ${column} ${type}`)
    }
  }

  /** Keeps the searchable text of the artifact's current version: its pages, entry first. */
  private indexText(id: string, version: number): void {
    const row = this.requireVersion(id, version)
    const pages = Object.keys(row.files)
      .filter((path) => isHtmlPath(path) || isMarkdownPath(path))
      .sort((a, b) => Number(b === row.entry) - Number(a === row.entry) || a.localeCompare(b))
    let text = ''
    for (const path of pages) {
      if (text.length >= SEARCH_CHARS) break
      text += ` ${pageText(path, this.blobs.read(row.files[path].sha256).toString('utf8'))}`
    }
    this.db
      .prepare('INSERT INTO search_text (artifact_id, version, text) VALUES (?, ?, ?) ON CONFLICT(artifact_id) DO UPDATE SET version = excluded.version, text = excluded.text')
      .run(id, version, text.trim().slice(0, SEARCH_CHARS))
  }

  close(): void {
    this.db.close()
  }

  subscribe(id: string, listener: (change: ArtifactChange) => void): () => void {
    const set = this.listeners.get(id) ?? new Set()
    set.add(listener)
    this.listeners.set(id, set)
    return () => {
      set.delete(listener)
      if (set.size === 0) {
        this.listeners.delete(id)
      }
    }
  }

  private emit(id: string, change: ArtifactChange): void {
    if (this.pendingChanges) {
      this.pendingChanges.push([id, change])
      return
    }
    for (const listener of this.listeners.get(id) ?? []) {
      listener(change)
    }
  }

  /** Changes made inside it are announced only once it commits; a rollback announces nothing. */
  private transaction<T>(work: () => T): T {
    this.db.exec('BEGIN IMMEDIATE')
    this.pendingChanges = []
    let changes: [string, ArtifactChange][] = []
    try {
      const result = work()
      this.db.exec('COMMIT')
      changes = this.pendingChanges
      return result
    } catch (error) {
      this.db.exec('ROLLBACK')
      throw error
    } finally {
      this.pendingChanges = null
      for (const [id, change] of changes) {
        this.emit(id, change)
      }
    }
  }

  /** Validates and hashes only; bodies reach disk through writeBlobs once the version is going to be committed. */
  private planFiles(base: FileMap, incoming: Record<string, IncomingFile>): { files: FileMap; bodies: Map<string, Buffer> } {
    const files: FileMap = { ...base }
    const bodies = new Map<string, Buffer>()
    let total = 0
    for (const [rawPath, file] of Object.entries(incoming)) {
      const path = normalizeFilePath(rawPath)
      if (!path) {
        throw new ArtifactError(400, 'invalid_path', `Unsafe file path: ${JSON.stringify(rawPath)}.`)
      }
      if (file === null) {
        delete files[path]
        continue
      }
      if (file.content.byteLength > LIMITS.fileBytes) {
        throw new ArtifactError(413, 'file_too_large', `${path} is larger than ${LIMITS.fileBytes} bytes.`)
      }
      total += file.content.byteLength
      if (total > LIMITS.publishBytes) {
        throw new ArtifactError(413, 'publish_too_large', `One publish may send at most ${LIMITS.publishBytes} bytes.`)
      }
      const sha256 = BlobStore.hash(file.content)
      bodies.set(sha256, file.content)
      files[path] = { sha256, size: file.content.byteLength, type: contentTypeFor(path) }
    }
    if (Object.keys(files).length === 0) {
      throw new ArtifactError(400, 'no_files', 'An artifact needs at least one file.')
    }
    if (Object.keys(files).length > LIMITS.filesPerVersion) {
      throw new ArtifactError(413, 'too_many_files', `A version may hold at most ${LIMITS.filesPerVersion} files.`)
    }
    return { files, bodies }
  }

  /** Runs the commit after writing the bodies; if it throws, blobs this call created and nothing references are removed. */
  private commitWithBlobs<T>(bodies: Map<string, Buffer>, commit: () => T): T {
    const created: string[] = []
    try {
      for (const [sha256, body] of bodies) {
        if (!this.blobs.has(sha256)) {
          created.push(sha256)
        }
        this.blobs.put(body)
      }
      return commit()
    } catch (error) {
      this.removeUnreferenced(created)
      throw error
    }
  }

  publish(input: PublishInput): ArtifactSummary {
    const title = requireText(input.title, 'title', 200)
    const description = optionalText(input.description, 'description', 1000)
    const icon = optionalText(input.icon, 'icon', 40)
    const label = optionalText(input.label, 'label', 60)
    const capabilities = parseCapabilities(input.capabilities)
    const source = parseSource(input.source)
    const { files, bodies } = this.planFiles({}, input.files ?? {})
    const entry = chooseEntry(files, input.entry)
    const id = newId(12)
    const at = now()
    this.commitWithBlobs(bodies, () => this.transaction(() => {
      this.db
        .prepare(`INSERT INTO artifacts (id, title, description, icon, entry, current_version, capabilities, source_host, source_path, created_at, updated_at)
          VALUES (?, ?, ?, ?, ?, 1, ?, ?, ?, ?, ?)`)
        .run(id, title, description, icon, entry, JSON.stringify(capabilities), source?.host ?? null, source?.path ?? null, at, at)
      this.db
        .prepare(`INSERT INTO versions (artifact_id, version, label, title, description, entry, files, source_host, source_path, created_at)
          VALUES (?, 1, ?, ?, ?, ?, ?, ?, ?, ?)`)
        .run(id, label, title, description, entry, JSON.stringify(files), source?.host ?? null, source?.path ?? null, at)
      this.indexText(id, 1)
    }))
    return this.requireSummary(id)
  }

  update(id: string, input: UpdateInput): ArtifactSummary {
    const current = this.requireSummary(id)
    if (input.baseVersion !== undefined && input.baseVersion !== current.currentVersion && !input.force) {
      throw new ArtifactError(
        409,
        'version_conflict',
        `Version ${current.currentVersion} is newer than baseVersion ${input.baseVersion}; re-read it or pass force.`,
        {
          currentVersion: current.currentVersion,
          baseVersion: input.baseVersion,
          changedFiles: this.changedFilesSince(id, input.baseVersion, current.currentVersion)
        }
      )
    }
    const previous = this.requireVersion(id, current.currentVersion)
    const title = input.title === undefined ? current.title : requireText(input.title, 'title', 200)
    const description = input.description === undefined ? current.description : optionalText(input.description, 'description', 1000)
    const icon = input.icon === undefined ? current.icon : optionalText(input.icon, 'icon', 40)
    const label = optionalText(input.label, 'label', 60)
    const capabilities = input.capabilities === undefined ? current.capabilities : parseCapabilities(input.capabilities)
    const source = parseSource(input.source)
    const { files, bodies } = this.planFiles(previous.files, input.files ?? {})
    const entry = input.entry !== undefined ? chooseEntry(files, input.entry) : files[current.entry] ? current.entry : chooseEntry(files, undefined)
    const version = current.currentVersion + 1
    const at = now()
    this.commitWithBlobs(bodies, () => this.transaction(() => {
      this.db
        .prepare(`INSERT INTO versions (artifact_id, version, label, title, description, entry, files, source_host, source_path, created_at)
          VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`)
        .run(id, version, label, title, description, entry, JSON.stringify(files), source?.host ?? null, source?.path ?? null, at)
      // The artifact shows the latest known source; a publish from an older CLI that sends none keeps it.
      this.db
        .prepare(`UPDATE artifacts SET title = ?, description = ?, icon = ?, entry = ?, current_version = ?, capabilities = ?, updated_at = ?,
          source_host = COALESCE(?, source_host), source_path = COALESCE(?, source_path) WHERE id = ?`)
        .run(title, description, icon, entry, version, JSON.stringify(capabilities), at, source?.host ?? null, source?.path ?? null, id)
      this.indexText(id, version)
    }))
    this.emit(id, { kind: 'version', version })
    return this.requireSummary(id)
  }

  /** Paths added, removed or changed between two versions; a base that no longer exists counts as empty. */
  private changedFilesSince(id: string, baseVersion: number, currentVersion: number): string[] {
    let base: FileMap = {}
    try {
      base = this.requireVersion(id, baseVersion).files
    } catch (error) {
      if (!(error instanceof ArtifactError)) throw error
    }
    const latest = this.requireVersion(id, currentVersion).files
    return [...new Set([...Object.keys(base), ...Object.keys(latest)])]
      .filter((path) => base[path]?.sha256 !== latest[path]?.sha256)
      .sort()
  }

  /** What changed from `from` (default: the nearest older version still kept) to `to` (default: current). */
  diff(id: string, to?: number, from?: number): VersionDiff {
    const summary = this.requireSummary(id)
    const target = this.requireVersion(id, to ?? summary.currentVersion)
    const previous = from ?? Number(this.db.prepare('SELECT MAX(version) AS v FROM versions WHERE artifact_id = ? AND version < ?').get(id, target.version)?.v ?? NaN)
    const base = Number.isInteger(previous) ? this.requireVersion(id, previous) : null
    const before = base?.files ?? {}
    const after = target.files
    const read = (entry: FileEntry) => this.blobs.read(entry.sha256).toString('utf8').split(/\r?\n/)
    const files = [...new Set([...Object.keys(before), ...Object.keys(after)])]
      .sort()
      .filter((path) => before[path]?.sha256 !== after[path]?.sha256)
      .map((path): FileChange => {
        const old = before[path]
        const now = after[path]
        const change: FileChange = { path, status: !old ? 'added' : !now ? 'removed' : 'modified' }
        const entries = [old, now].filter((entry): entry is FileEntry => Boolean(entry))
        if (!entries.every((entry) => isTextType(entry.type))) return { ...change, note: 'binary' }
        if (entries.some((entry) => entry.size > DIFF_LIMITS.bytes)) return { ...change, note: 'too_large' }
        const a = old ? read(old) : []
        const b = now ? read(now) : []
        if (a.length > DIFF_LIMITS.lines || b.length > DIFF_LIMITS.lines) return { ...change, note: 'too_large' }
        const ops = diffLines(a, b, DIFF_LIMITS.edits)
        if (!ops) return { ...change, note: 'rewritten' }
        const shown = hunks(ops).map((line) => ('text' in line && line.text.length > DIFF_LIMITS.lineChars ? { ...line, text: `${line.text.slice(0, DIFF_LIMITS.lineChars)}…` } : line))
        if (shown.length > DIFF_LIMITS.shown) {
          shown.splice(DIFF_LIMITS.shown, shown.length, { op: '…', skipped: shown.length - DIFF_LIMITS.shown })
        }
        return { ...change, lines: shown }
      })
    return { from: base?.version ?? null, to: target.version, label: target.label, source: target.source, createdAt: target.createdAt, files }
  }

  setPinned(id: string, pinned: boolean): ArtifactSummary {
    this.requireSummary(id)
    this.db.prepare('UPDATE artifacts SET pinned = ? WHERE id = ?').run(pinned ? 1 : 0, id)
    return this.requireSummary(id)
  }

  delete(id: string): void {
    this.requireSummary(id)
    const candidates = this.db
      .prepare(`SELECT json_extract(f.value, '$.sha256') AS sha256 FROM versions v, json_each(v.files) f WHERE v.artifact_id = ?
        UNION SELECT sha256 FROM assets WHERE artifact_id = ?`)
      .all(id, id)
      .map((row) => String(row.sha256))
    this.db.prepare('DELETE FROM artifacts WHERE id = ?').run(id)
    this.emit(id, { kind: 'deleted' })
    this.removeUnreferenced(candidates)
  }

  /** Removes the given blobs unless a remaining version or asset still references them; returns the ones removed. */
  private removeUnreferenced(candidates: string[]): string[] {
    const removed: string[] = []
    for (let start = 0; start < candidates.length; start += 500) {
      const chunk = candidates.slice(start, start + 500)
      const marks = chunk.map(() => '?').join(',')
      const referenced = new Set(
        this.db
          .prepare(`SELECT sha256 FROM (
              SELECT json_extract(f.value, '$.sha256') AS sha256 FROM versions v, json_each(v.files) f
              UNION SELECT sha256 FROM assets
            ) WHERE sha256 IN (${marks})`)
          .all(...chunk)
          .map((row) => String(row.sha256))
      )
      for (const sha256 of chunk) {
        if (!referenced.has(sha256)) {
          this.blobs.remove(sha256)
          removed.push(sha256)
        }
      }
    }
    return removed
  }

  /**
   * Drops all but the newest `keep` versions (the current one always stays). Comments keep their version
   * number; a publish based on a dropped version is still a 409 that names every file (changedFilesSince).
   */
  prune(id: string, keep: number): { removed: number[]; freedBytes: number } {
    if (!Number.isInteger(keep) || keep < 1) {
      throw new ArtifactError(400, 'invalid_field', 'keep is a whole number of versions, at least 1.')
    }
    this.requireSummary(id)
    const rows = this.db.prepare('SELECT version, files FROM versions WHERE artifact_id = ? ORDER BY version DESC').all(id).slice(keep)
    if (rows.length === 0) {
      return { removed: [], freedBytes: 0 }
    }
    const removed = rows.map((row) => Number(row.version))
    const sizes = new Map<string, number>()
    for (const row of rows) {
      for (const file of Object.values(JSON.parse(String(row.files)) as FileMap)) {
        sizes.set(file.sha256, file.size)
      }
    }
    const marks = removed.map(() => '?').join(',')
    this.transaction(() => {
      this.db.prepare(`DELETE FROM versions WHERE artifact_id = ? AND version IN (${marks})`).run(id, ...removed)
      this.db.prepare(`DELETE FROM page_errors WHERE artifact_id = ? AND version IN (${marks})`).run(id, ...removed)
      this.emit(id, { kind: 'pruned', versions: removed })
    })
    const freed = this.removeUnreferenced([...sizes.keys()])
    return { removed, freedBytes: freed.reduce((total, sha256) => total + (sizes.get(sha256) ?? 0), 0) }
  }

  /** Maintenance sweep over the whole blob directory; delete paths use removeUnreferenced instead. */
  collectGarbage(): number {
    const referenced = new Set<string>()
    for (const row of this.db.prepare('SELECT files FROM versions').all()) {
      for (const file of Object.values(JSON.parse(String(row.files)) as FileMap)) {
        referenced.add(file.sha256)
      }
    }
    for (const row of this.db.prepare('SELECT sha256 FROM assets').all()) {
      referenced.add(String(row.sha256))
    }
    let removed = 0
    for (const sha256 of this.blobs.list()) {
      if (!referenced.has(sha256)) {
        this.blobs.remove(sha256)
        removed += 1
      }
    }
    return removed
  }

  private summaryFromRow(row: Record<string, unknown>): ArtifactSummary {
    return {
      id: String(row.id),
      title: String(row.title),
      description: row.description === null ? null : String(row.description),
      icon: row.icon === null ? null : String(row.icon),
      entry: String(row.entry),
      currentVersion: Number(row.current_version),
      capabilities: JSON.parse(String(row.capabilities)) as Capability[],
      pinned: Number(row.pinned) === 1,
      source: sourceFromRow(row),
      createdAt: String(row.created_at),
      updatedAt: String(row.updated_at)
    }
  }

  getSummary(id: string): ArtifactSummary | null {
    const row = this.db.prepare('SELECT * FROM artifacts WHERE id = ?').get(id)
    return row ? this.summaryFromRow(row) : null
  }

  requireSummary(id: string): ArtifactSummary {
    const summary = this.getSummary(id)
    if (!summary) {
      throw new ArtifactError(404, 'artifact_not_found', `No artifact ${id}.`)
    }
    return summary
  }

  detail(id: string): ArtifactDetail {
    const summary = this.requireSummary(id)
    const versions = this.db
      .prepare('SELECT version, label, title, description, entry, source_host, source_path, created_at FROM versions WHERE artifact_id = ? ORDER BY version DESC')
      .all(id)
      .map((row) => ({
        version: Number(row.version),
        label: row.label === null ? null : String(row.label),
        source: sourceFromRow(row),
        title: String(row.title),
        description: row.description === null ? null : String(row.description),
        entry: String(row.entry),
        createdAt: String(row.created_at)
      }))
    return { ...summary, versions }
  }

  requireVersion(id: string, version: number): VersionRow {
    const row = this.db.prepare('SELECT * FROM versions WHERE artifact_id = ? AND version = ?').get(id, version)
    if (!row) {
      throw new ArtifactError(404, 'version_not_found', `Artifact ${id} has no version ${version}.`)
    }
    return {
      version: Number(row.version),
      label: row.label === null ? null : String(row.label),
      source: sourceFromRow(row),
      title: String(row.title),
      description: row.description === null ? null : String(row.description),
      entry: String(row.entry),
      files: JSON.parse(String(row.files)) as FileMap,
      createdAt: String(row.created_at)
    }
  }

  /** A query matches the title, description, id, source machine or folder, or the current version's page text. */
  list(options: { query?: string; limit?: number; offset?: number } = {}): { artifacts: ArtifactListing[]; total: number } {
    const limit = clampInt(options.limit, 50, 1, 200)
    const offset = clampInt(options.offset, 0, 0, Number.MAX_SAFE_INTEGER)
    const query = (options.query ?? '').trim()
    const pattern = `%${query.replace(/[\\%_]/g, (char) => `\\${char}`)}%`
    const like = (column: string) => `${column} LIKE ? ESCAPE '\\'`
    const where = query
      ? `WHERE ${[like('a.title'), like("IFNULL(a.description, '')"), 'a.id = ?', like("IFNULL(a.source_host, '')"), like("IFNULL(a.source_path, '')"), like("IFNULL(t.text, '')")].join(' OR ')}`
      : ''
    const params = query ? [pattern, pattern, query, pattern, pattern, pattern] : []
    // The page text is read only to search it; a plain gallery listing never carries it.
    const from = query ? 'FROM artifacts a LEFT JOIN search_text t ON t.artifact_id = a.id' : 'FROM artifacts a'
    const rows = this.db
      .prepare(`SELECT a.*, ${query ? 't.text' : 'NULL'} AS search_text ${from} ${where} ORDER BY a.pinned DESC, a.updated_at DESC LIMIT ? OFFSET ?`)
      .all(...params, limit, offset)
    const total = this.db.prepare(`SELECT COUNT(*) AS n ${from} ${where}`).get(...params)
    return {
      artifacts: rows.map((row) => {
        const summary: ArtifactListing = this.summaryFromRow(row)
        // A snippet only when the words are not already in what the card shows.
        const shown = `${summary.title} ${summary.description ?? ''}`.toLowerCase()
        if (query && !shown.includes(query.toLowerCase())) {
          summary.snippet = snippetAround(String(row.search_text ?? ''), query)
        }
        return summary
      }),
      total: Number(total?.n ?? 0)
    }
  }

  // Assets: files that outlive versions, uploaded by the publisher or by a page with the uploads capability.

  addAsset(id: string, name: string, content: Buffer, source: AssetRow['source']): AssetRow {
    this.requireSummary(id)
    const safeName = normalizeFilePath(name.split('/').pop() ?? '') ?? 'file'
    if (content.byteLength > LIMITS.assetBytes) {
      throw new ArtifactError(413, 'asset_too_large', `Assets are limited to ${LIMITS.assetBytes} bytes.`)
    }
    const asset: AssetRow = {
      id: randomBytes(16).toString('hex'),
      name: safeName,
      sha256: BlobStore.hash(content),
      size: content.byteLength,
      type: contentTypeFor(safeName),
      source,
      createdAt: now()
    }
    this.commitWithBlobs(new Map([[asset.sha256, content]]), () =>
      this.db
        .prepare('INSERT INTO assets (artifact_id, id, name, sha256, size, type, source, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)')
        .run(id, asset.id, asset.name, asset.sha256, asset.size, asset.type, asset.source, asset.createdAt)
    )
    this.emit(id, { kind: 'assets' })
    return asset
  }

  private assetFromRow(row: Record<string, unknown>): AssetRow {
    return {
      id: String(row.id),
      name: String(row.name),
      sha256: String(row.sha256),
      size: Number(row.size),
      type: String(row.type),
      source: String(row.source) === 'page' ? 'page' : 'publisher',
      createdAt: String(row.created_at)
    }
  }

  listAssets(id: string): AssetRow[] {
    this.requireSummary(id)
    return this.db
      .prepare('SELECT * FROM assets WHERE artifact_id = ? ORDER BY created_at DESC')
      .all(id)
      .map((row) => this.assetFromRow(row))
  }

  getAsset(id: string, assetId: string): AssetRow {
    this.requireSummary(id)
    const row = this.db.prepare('SELECT * FROM assets WHERE artifact_id = ? AND id = ?').get(id, assetId)
    if (!row) {
      throw new ArtifactError(404, 'asset_not_found', `No asset ${assetId}.`)
    }
    return this.assetFromRow(row)
  }

  deleteAsset(id: string, assetId: string): void {
    const asset = this.getAsset(id, assetId)
    this.db.prepare('DELETE FROM assets WHERE artifact_id = ? AND id = ?').run(id, assetId)
    this.emit(id, { kind: 'assets' })
    this.removeUnreferenced([asset.sha256])
  }

  // Docs: the per-artifact shared database behind the db capability.

  private docKey(value: string, field: string): string {
    if (typeof value !== 'string' || !/^[A-Za-z0-9_.:-]{1,128}$/.test(value)) {
      throw new ArtifactError(400, 'invalid_key', `${field} must be 1-128 characters of A-Z a-z 0-9 _ . : -`)
    }
    return value
  }

  private docFromRow(row: Record<string, unknown>): DocRow {
    return {
      id: String(row.doc_id),
      data: JSON.parse(String(row.data)),
      createdAt: String(row.created_at),
      updatedAt: String(row.updated_at)
    }
  }

  getDoc(id: string, collection: string, docId: string): DocRow | null {
    this.requireSummary(id)
    const row = this.db
      .prepare('SELECT * FROM docs WHERE artifact_id = ? AND collection = ? AND doc_id = ?')
      .get(id, this.docKey(collection, 'collection'), this.docKey(docId, 'doc id'))
    return row ? this.docFromRow(row) : null
  }

  setDoc(id: string, collection: string, docId: string, data: unknown, merge: boolean): DocRow {
    this.requireSummary(id)
    this.docKey(collection, 'collection')
    this.docKey(docId, 'doc id')
    const existing = this.getDoc(id, collection, docId)
    let next = data
    if (merge) {
      if (data === null || typeof data !== 'object' || Array.isArray(data)) {
        throw new ArtifactError(400, 'invalid_doc', 'update merges an object into the document.')
      }
      const base = existing && existing.data && typeof existing.data === 'object' && !Array.isArray(existing.data) ? existing.data : {}
      next = { ...base, ...data }
    }
    const serialized = JSON.stringify(next ?? null)
    if (Buffer.byteLength(serialized) > LIMITS.docBytes) {
      throw new ArtifactError(413, 'doc_too_large', `Documents are limited to ${LIMITS.docBytes} bytes.`)
    }
    if (!existing) {
      const count = this.db.prepare('SELECT COUNT(*) AS n FROM docs WHERE artifact_id = ?').get(id)
      if (Number(count?.n ?? 0) >= LIMITS.docsPerArtifact) {
        throw new ArtifactError(413, 'too_many_docs', `An artifact may hold at most ${LIMITS.docsPerArtifact} documents.`)
      }
    }
    const at = now()
    this.db
      .prepare(`INSERT INTO docs (artifact_id, collection, doc_id, data, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?)
        ON CONFLICT (artifact_id, collection, doc_id) DO UPDATE SET data = excluded.data, updated_at = excluded.updated_at`)
      .run(id, collection, docId, serialized, at, at)
    this.emit(id, { kind: 'db', collection, docId, deleted: false })
    return this.getDoc(id, collection, docId) as DocRow
  }

  /** Applies every write or none; ops are {op: 'set' | 'update' | 'delete', collection, id, data?}. */
  batch(id: string, ops: unknown[]): unknown[] {
    this.requireSummary(id)
    return this.transaction(() =>
      ops.map((raw, index) => {
        const op = (raw ?? {}) as { op?: unknown; collection?: unknown; id?: unknown; data?: unknown }
        const collection = String(op.collection ?? '')
        const docId = String(op.id ?? '')
        if (op.op === 'set' || op.op === 'update') {
          return this.setDoc(id, collection, docId, op.data, op.op === 'update')
        }
        if (op.op === 'delete') {
          return { deleted: this.deleteDoc(id, collection, docId) }
        }
        throw new ArtifactError(400, 'invalid_batch', `ops[${index}].op must be set, update or delete.`)
      })
    )
  }

  deleteDoc(id: string, collection: string, docId: string): boolean {
    this.requireSummary(id)
    const result = this.db
      .prepare('DELETE FROM docs WHERE artifact_id = ? AND collection = ? AND doc_id = ?')
      .run(id, this.docKey(collection, 'collection'), this.docKey(docId, 'doc id'))
    if (Number(result.changes) > 0) {
      this.emit(id, { kind: 'db', collection, docId, deleted: true })
      return true
    }
    return false
  }

  listDocs(id: string, collection: string, options: { limit?: number; after?: string; where?: Record<string, unknown> } = {}): { docs: DocRow[]; next: string | null } {
    this.requireSummary(id)
    this.docKey(collection, 'collection')
    const limit = clampInt(options.limit, 100, 1, 500)
    const clauses = ['artifact_id = ?', 'collection = ?']
    const params: (string | number | null)[] = [id, collection]
    if (options.after) {
      clauses.push('doc_id > ?')
      params.push(options.after)
    }
    for (const [field, value] of Object.entries(options.where ?? {})) {
      if (!/^[A-Za-z0-9_]{1,64}$/.test(field)) {
        throw new ArtifactError(400, 'invalid_query', `Query fields must be plain top-level names; got ${JSON.stringify(field)}.`)
      }
      if (value !== null && !['string', 'number', 'boolean'].includes(typeof value)) {
        throw new ArtifactError(400, 'invalid_query', 'Query values must be strings, numbers, booleans or null.')
      }
      if (value === null) {
        clauses.push(`json_extract(data, '$.${field}') IS NULL`)
      } else {
        clauses.push(`json_extract(data, '$.${field}') = ?`)
        params.push(typeof value === 'boolean' ? (value ? 1 : 0) : (value as string | number))
      }
    }
    const rows = this.db
      .prepare(`SELECT * FROM docs WHERE ${clauses.join(' AND ')} ORDER BY doc_id LIMIT ?`)
      .all(...params, limit + 1)
    const docs = rows.slice(0, limit).map((row) => this.docFromRow(row))
    return { docs, next: rows.length > limit ? docs[docs.length - 1].id : null }
  }

  // Comments: threads on the viewer, available on every artifact.

  addComment(id: string, input: { author?: unknown; body?: unknown; parentId?: unknown; version?: unknown }): CommentRow {
    const summary = this.requireSummary(id)
    const author = requireText(input.author, 'author', 80)
    const body = requireText(input.body, 'body', LIMITS.commentChars)
    let parentId: string | null = null
    if (input.parentId !== undefined && input.parentId !== null) {
      const parent = this.db.prepare('SELECT id, parent_id FROM comments WHERE id = ? AND artifact_id = ?').get(String(input.parentId), id)
      if (!parent) {
        throw new ArtifactError(404, 'comment_not_found', 'The comment being answered does not exist.')
      }
      parentId = parent.parent_id === null ? String(parent.id) : String(parent.parent_id)
    }
    const version = typeof input.version === 'number' && Number.isInteger(input.version) ? input.version : summary.currentVersion
    const comment: CommentRow = { id: newId(16), parentId, version, author, body, resolved: false, createdAt: now() }
    this.db
      .prepare('INSERT INTO comments (id, artifact_id, parent_id, version, author, body, created_at) VALUES (?, ?, ?, ?, ?, ?, ?)')
      .run(comment.id, id, parentId, version, author, body, comment.createdAt)
    this.emit(id, { kind: 'comments' })
    return comment
  }

  listComments(id: string): CommentRow[] {
    this.requireSummary(id)
    return this.db
      .prepare('SELECT * FROM comments WHERE artifact_id = ? ORDER BY created_at')
      .all(id)
      .map((row) => ({
        id: String(row.id),
        parentId: row.parent_id === null ? null : String(row.parent_id),
        version: Number(row.version),
        author: String(row.author),
        body: String(row.body),
        resolved: Number(row.resolved) === 1,
        createdAt: String(row.created_at)
      }))
  }

  resolveComment(id: string, commentId: string, resolved: boolean): void {
    const result = this.db.prepare('UPDATE comments SET resolved = ? WHERE id = ? AND artifact_id = ? AND parent_id IS NULL').run(resolved ? 1 : 0, commentId, id)
    if (Number(result.changes) === 0) {
      throw new ArtifactError(404, 'comment_not_found', 'Only an existing thread root can be resolved.')
    }
    this.emit(id, { kind: 'comments' })
  }

  deleteComment(id: string, commentId: string): void {
    const result = this.db.prepare('DELETE FROM comments WHERE artifact_id = ? AND (id = ? OR parent_id = ?)').run(id, commentId, commentId)
    if (Number(result.changes) === 0) {
      throw new ArtifactError(404, 'comment_not_found', `No comment ${commentId}.`)
    }
    this.emit(id, { kind: 'comments' })
  }

  // Page errors: what broke while someone viewed a page, reported through the viewer. Diagnostic only;
  // the text comes from the page, so it is data, never instructions.

  recordErrors(id: string, input: { version?: unknown; path?: unknown; errors?: unknown }): void {
    const version = typeof input.version === 'number' ? input.version : NaN
    const row = this.requireVersion(id, version)
    const path = typeof input.path === 'string' ? input.path : ''
    if (!row.files[path]) {
      throw new ArtifactError(400, 'invalid_field', `Version ${version} has no file ${path}.`)
    }
    if (!Array.isArray(input.errors) || input.errors.length > LIMITS.errorsPerReport) {
      throw new ArtifactError(400, 'invalid_field', `errors is an array of at most ${LIMITS.errorsPerReport} reports.`)
    }
    const errors = input.errors.map((raw) => {
      const { kind, message, source } = (raw ?? {}) as { kind?: unknown; message?: unknown; source?: unknown }
      if (!(PAGE_ERROR_KINDS as readonly unknown[]).includes(kind) || typeof message !== 'string') {
        throw new ArtifactError(400, 'invalid_field', `Each error has a kind (${PAGE_ERROR_KINDS.join(', ')}) and a message.`)
      }
      return {
        kind: kind as string,
        message: message.slice(0, LIMITS.errorChars),
        source: typeof source === 'string' ? source.slice(0, LIMITS.errorChars) : ''
      }
    })
    if (errors.length === 0) {
      return
    }
    this.transaction(() => {
      const at = now()
      const key = [id, version, path]
      const exists = this.db.prepare('SELECT 1 FROM page_errors WHERE artifact_id = ? AND version = ? AND path = ? AND kind = ? AND message = ? AND source = ?')
      const stored = this.db.prepare('SELECT COUNT(*) AS n FROM page_errors WHERE artifact_id = ? AND version = ?')
      const bump = this.db.prepare('UPDATE page_errors SET count = count + 1, last_at = ? WHERE artifact_id = ? AND version = ? AND path = ? AND kind = ? AND message = ? AND source = ?')
      const insert = this.db.prepare('INSERT INTO page_errors (artifact_id, version, path, kind, message, source, count, first_at, last_at) VALUES (?, ?, ?, ?, ?, ?, 1, ?, ?)')
      for (const error of errors) {
        const fields = [...key, error.kind, error.message, error.source]
        if (exists.get(...fields)) {
          bump.run(at, ...fields)
        } else if (Number(stored.get(id, version)?.n ?? 0) < LIMITS.errorsPerVersion) {
          insert.run(...fields, at, at)
        }
      }
      this.emit(id, { kind: 'errors', version })
    })
  }

  listErrors(id: string, version: number): PageErrorRow[] {
    this.requireVersion(id, version)
    return this.db
      .prepare('SELECT * FROM page_errors WHERE artifact_id = ? AND version = ? ORDER BY last_at DESC, path, message')
      .all(id, version)
      .map((row) => ({
        version: Number(row.version),
        path: String(row.path),
        kind: String(row.kind) as PageErrorRow['kind'],
        message: String(row.message),
        source: row.source === '' ? null : String(row.source),
        count: Number(row.count),
        firstAt: String(row.first_at),
        lastAt: String(row.last_at)
      }))
  }
}

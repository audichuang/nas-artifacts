import { createHash, createHmac, randomBytes, timingSafeEqual } from 'node:crypto'
import { join } from 'node:path'
import { DatabaseSync } from 'node:sqlite'
import {
  generateAuthenticationOptions,
  generateRegistrationOptions,
  verifyAuthenticationResponse,
  verifyRegistrationResponse,
  type AuthenticationResponseJSON,
  type RegistrationResponseJSON
} from '@simplewebauthn/server'
import type { Context } from 'hono'
import { ArtifactError } from './artifact-store.ts'

/** `__Host-`: Secure, Path=/, no Domain, so no other audichuang.app host can set or overwrite it. */
export const SESSION_COOKIE = '__Host-artifacts_session'
const SESSION_DAYS = 30
const LOGIN_LINK_MINUTES = 10
const CHALLENGE_MINUTES = 5
/** How long a content token lets the public content listener serve one private artifact. */
export const CONTENT_TOKEN_HOURS = 12
/** A content token as it appears in a path: base-36 expiry, a dot, the base64url signature. */
export const CONTENT_TOKEN_PATTERN = /^[0-9a-z]{1,12}\.[A-Za-z0-9_-]{22}$/

const sha256 = (value: string): string => createHash('sha256').update(value).digest('hex')
const randomToken = (): string => randomBytes(32).toString('base64url')
const nowIso = (): string => new Date().toISOString()

/** The one account every passkey is registered to, as WebAuthn's user handle. */
export const OWNER_USER_ID = Buffer.from('artifacts-owner').toString('base64url')

/**
 * Default passkey names by AAGUID, from the community registry (passkeydeveloper/passkey-authenticator-aaguids).
 * A UX hint only: the AAGUID of a platform passkey is not attested, so nothing may depend on it.
 */
const PROVIDERS: Record<string, string> = {
  'fbfc3007-154e-4ecc-8c0b-6e020557d7bd': 'Apple 密碼',
  'dd4ec289-e01d-41c9-bb89-70fa845d4bf2': 'Apple 密碼',
  'ea9b8d66-4d01-1d21-3ce4-b6b48cb575d4': 'Google 密碼管理工具',
  'adce0002-35bc-c60a-648b-0b25f1f05503': 'Chrome on Mac',
  '771b48fd-d3d4-4f74-9232-fc157ab0507a': 'Edge on Mac',
  'd3452668-01fd-4c12-926c-83a4204853aa': 'Microsoft 密碼管理工具',
  '08987058-cadc-4b81-b6e1-30de50dcbe96': 'Windows Hello',
  '9ddd1817-af5a-4672-a2b9-3e3dd95000a9': 'Windows Hello',
  '6028b017-b1d4-4c02-b4b3-afcdafc96bb2': 'Windows Hello',
  '53414d53-554e-4700-0000-000000000000': 'Samsung Pass',
  'bada5566-a7aa-401f-bd96-45619a55120d': '1Password',
  'd548826e-79b4-db40-a3d8-11116f7e8349': 'Bitwarden',
  '531126d6-e717-415c-9320-3d9aa6981239': 'Dashlane',
  '50726f74-6f6e-5061-7373-50726f746f6e': 'Proton Pass',
  'b84e4048-15dc-4dd0-8640-f4f60813c8af': 'NordPass',
  'f3809540-7f14-49c1-a8b3-8f813b225541': 'Enpass',
  'fdb141b2-5d84-443e-8a35-4698c205a502': 'KeePassXC'
}

export function passkeyProvider(aaguid: string | undefined): string | undefined {
  return aaguid ? PROVIDERS[aaguid.toLowerCase()] : undefined
}

export type PasskeyRow = { id: string; name: string; createdAt: string; lastUsedAt: string | null }

type Challenge = { challenge: string; kind: 'register' | 'login'; expires: number; linkHash?: string }

/**
 * The owner's sign-in for the public listeners: passkeys, 30-day sessions, one-time enrolment links and
 * the short-lived content tokens that let the public content listener serve a private page to the owner.
 * Kept in its own database file: nothing here touches artifacts, and nothing in artifacts depends on it.
 */
export class OwnerAuth {
  readonly db: DatabaseSync
  readonly rpID: string
  readonly origin: string
  private contentKey: Buffer
  private readonly challenges = new Map<string, Challenge>()

  constructor(dataDir: string, publicAppOrigin: string) {
    this.origin = new URL(publicAppOrigin).origin
    this.rpID = new URL(publicAppOrigin).hostname
    this.db = new DatabaseSync(join(dataDir, 'auth.db'))
    this.db.exec(`
      PRAGMA journal_mode = WAL;
      CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT NOT NULL);
      CREATE TABLE IF NOT EXISTS passkeys (
        id TEXT PRIMARY KEY,
        public_key BLOB NOT NULL,
        counter INTEGER NOT NULL,
        transports TEXT,
        name TEXT NOT NULL,
        created_at TEXT NOT NULL,
        last_used_at TEXT
      );
      CREATE TABLE IF NOT EXISTS sessions (
        hash TEXT PRIMARY KEY,
        created_at TEXT NOT NULL,
        expires_at INTEGER NOT NULL
      );
      CREATE TABLE IF NOT EXISTS login_links (
        hash TEXT PRIMARY KEY,
        expires_at INTEGER NOT NULL
      );
    `)
    let key = this.db.prepare("SELECT value FROM settings WHERE key = 'content_token_key'").get()?.value as string | undefined
    if (!key) {
      key = randomBytes(32).toString('base64')
      this.db.prepare("INSERT OR IGNORE INTO settings (key, value) VALUES ('content_token_key', ?)").run(key)
      key = this.db.prepare("SELECT value FROM settings WHERE key = 'content_token_key'").get()!.value as string
    }
    this.contentKey = Buffer.from(key, 'base64')
  }


  close(): void {
    this.db.close()
  }

  // Content tokens: HMAC(artifact id, expiry). Path-borne, so a page's relative links carry it.

  private sign(id: string, expires: string): string {
    return createHmac('sha256', this.contentKey).update(`content\n${id}\n${expires}`).digest().subarray(0, 16).toString('base64url')
  }

  contentToken(id: string, now = Date.now()): { token: string; expiresAt: number } {
    // Whole minutes, so tokens issued within the same minute are the same path and the browser cache keeps working.
    const expiresAt = Math.ceil((now + CONTENT_TOKEN_HOURS * 3600_000) / 60_000) * 60_000
    const expires = expiresAt.toString(36)
    return { token: `${expires}.${this.sign(id, expires)}`, expiresAt }
  }

  verifyContentToken(token: string, id: string, now = Date.now()): boolean {
    if (!CONTENT_TOKEN_PATTERN.test(token)) return false
    const [expires, signature] = token.split('.')
    if (parseInt(expires, 36) <= now) return false
    const expected = Buffer.from(this.sign(id, expires))
    const given = Buffer.from(signature)
    return expected.length === given.length && timingSafeEqual(expected, given)
  }

  // Sessions

  createSession(): string {
    const raw = randomToken()
    this.db.prepare('DELETE FROM sessions WHERE expires_at <= ?').run(Date.now())
    this.db.prepare('INSERT INTO sessions (hash, created_at, expires_at) VALUES (?, ?, ?)').run(sha256(raw), nowIso(), Date.now() + SESSION_DAYS * 86400_000)
    return raw
  }

  isSession(raw: string | undefined): boolean {
    if (!raw || raw.length > 100) return false
    const row = this.db.prepare('SELECT expires_at FROM sessions WHERE hash = ?').get(sha256(raw))
    return Boolean(row && Number(row.expires_at) > Date.now())
  }

  endSession(raw: string | undefined): void {
    if (raw) this.db.prepare('DELETE FROM sessions WHERE hash = ?').run(sha256(raw))
  }

  /** Ends every session and, with a new content key, every content token: no 12-hour tail after signing out everywhere. */
  endAllSessions(): void {
    // One transaction: a crash in between must not end the sessions yet keep the old content tokens valid.
    const key = randomBytes(32)
    this.db.exec('BEGIN IMMEDIATE')
    try {
      this.db.prepare('DELETE FROM sessions').run()
      this.db.prepare("INSERT OR REPLACE INTO settings (key, value) VALUES ('content_token_key', ?)").run(key.toString('base64'))
      this.db.exec('COMMIT')
    } catch (error) {
      this.db.exec('ROLLBACK')
      throw error
    }
    this.contentKey = key
  }

  sessionCount(): number {
    return Number(this.db.prepare('SELECT COUNT(*) AS n FROM sessions WHERE expires_at > ?').get(Date.now())!.n)
  }

  // One-time enrolment links, issued with the publish token on the tailnet.

  createLoginLink(): { code: string; expiresAt: string } {
    const code = randomToken()
    const expires = Date.now() + LOGIN_LINK_MINUTES * 60_000
    this.db.prepare('DELETE FROM login_links WHERE expires_at <= ?').run(Date.now())
    this.db.prepare('INSERT INTO login_links (hash, expires_at) VALUES (?, ?)').run(sha256(code), expires)
    return { code, expiresAt: new Date(expires).toISOString() }
  }

  isLoginLink(code: string | undefined): boolean {
    if (!code || code.length > 100) return false
    const row = this.db.prepare('SELECT expires_at FROM login_links WHERE hash = ?').get(sha256(code))
    return Boolean(row && Number(row.expires_at) > Date.now())
  }

  // Passkeys

  listPasskeys(): PasskeyRow[] {
    return this.db
      .prepare('SELECT id, name, created_at, last_used_at FROM passkeys ORDER BY created_at')
      .all()
      .map((row) => ({ id: String(row.id), name: String(row.name), createdAt: String(row.created_at), lastUsedAt: row.last_used_at === null ? null : String(row.last_used_at) }))
  }

  removePasskey(id: string): void {
    const result = this.db.prepare('DELETE FROM passkeys WHERE id = ?').run(id)
    if (result.changes === 0) throw new ArtifactError(404, 'not_found', 'No such passkey.')
    // A removed passkey may belong to a lost device: whatever it opened ends with it.
    this.endAllSessions()
  }

  private takeChallenge(flow: unknown, kind: Challenge['kind']): Challenge {
    const now = Date.now()
    for (const [key, value] of this.challenges) if (value.expires <= now) this.challenges.delete(key)
    const challenge = typeof flow === 'string' ? this.challenges.get(flow) : undefined
    if (typeof flow === 'string') this.challenges.delete(flow)
    if (!challenge || challenge.kind !== kind) throw new ArtifactError(400, 'challenge_expired', '登入逾時，請重新整理再試一次。')
    return challenge
  }

  private putChallenge(challenge: Omit<Challenge, 'expires'>): string {
    // A bounded number of pending ceremonies: each costs a little memory and anyone may start one.
    if (this.challenges.size >= 200) this.challenges.delete(this.challenges.keys().next().value!)
    const flow = randomToken()
    this.challenges.set(flow, { ...challenge, expires: Date.now() + CHALLENGE_MINUTES * 60_000 })
    return flow
  }

  /** Registration needs a live enrolment link or a signed-in session; the caller checks the session. */
  async registrationOptions(link: { code?: string; signedIn: boolean }) {
    let linkHash: string | undefined
    // A link is spent by the passkey it enrols even on a signed-in device; only the account page enrols without one.
    if (link.code !== undefined || !link.signedIn) {
      if (!this.isLoginLink(link.code)) throw new ArtifactError(403, 'link_expired', '這個註冊連結無效或已過期，請重新產生。')
      linkHash = sha256(link.code!)
    }
    const existing = this.db.prepare('SELECT id, transports FROM passkeys').all()
    const options = await generateRegistrationOptions({
      rpName: 'Artifacts',
      rpID: this.rpID,
      userName: 'owner',
      userDisplayName: 'Artifacts 擁有者',
      userID: Buffer.from(OWNER_USER_ID, 'base64url'),
      attestationType: 'none',
      excludeCredentials: existing.map((row) => ({ id: String(row.id), transports: row.transports ? JSON.parse(String(row.transports)) : undefined })),
      authenticatorSelection: { residentKey: 'required', userVerification: 'required' }
    })
    return { flow: this.putChallenge({ challenge: options.challenge, kind: 'register', linkHash }), options }
  }

  async register(flow: unknown, response: unknown, name: unknown, signedIn: boolean): Promise<void> {
    const challenge = this.takeChallenge(flow, 'register')
    if (!signedIn && !challenge.linkHash) throw new ArtifactError(403, 'link_expired', '這個註冊連結無效或已過期，請重新產生。')
    if (challenge.linkHash && !this.db.prepare('SELECT 1 FROM login_links WHERE hash = ? AND expires_at > ?').get(challenge.linkHash, Date.now())) {
      throw new ArtifactError(403, 'link_expired', '這個註冊連結無效或已過期，請重新產生。')
    }
    const verification = await verifyRegistrationResponse({
      response: response as RegistrationResponseJSON,
      expectedChallenge: challenge.challenge,
      expectedOrigin: this.origin,
      expectedRPID: this.rpID,
      requireUserVerification: true
    }).catch(() => ({ verified: false as const }))
    if (!verification.verified) throw new ArtifactError(400, 'passkey_rejected', '無法驗證這把 passkey。')
    const { credential, aaguid } = verification.registrationInfo
    if (challenge.linkHash) {
      // The link is spent by the first passkey it enrols.
      const spent = this.db.prepare('DELETE FROM login_links WHERE hash = ? AND expires_at > ?').run(challenge.linkHash, Date.now())
      if (spent.changes === 0) throw new ArtifactError(403, 'link_expired', '這個註冊連結無效或已過期，請重新產生。')
    }
    const label = typeof name === 'string' && name.trim() ? name.trim().slice(0, 60) : (passkeyProvider(aaguid) ?? `passkey ${this.listPasskeys().length + 1}`)
    this.db
      .prepare('INSERT INTO passkeys (id, public_key, counter, transports, name, created_at) VALUES (?, ?, ?, ?, ?, ?)')
      .run(credential.id, credential.publicKey, credential.counter, credential.transports ? JSON.stringify(credential.transports) : null, label, nowIso())
  }

  async loginOptions() {
    // Discoverable credentials: the authenticator offers the passkey, so the page names no account.
    const options = await generateAuthenticationOptions({ rpID: this.rpID, userVerification: 'required' })
    return { flow: this.putChallenge({ challenge: options.challenge, kind: 'login' }), options }
  }

  async login(flow: unknown, response: unknown): Promise<void> {
    const challenge = this.takeChallenge(flow, 'login')
    const id = (response as { id?: unknown } | null)?.id
    const row = typeof id === 'string' ? this.db.prepare('SELECT * FROM passkeys WHERE id = ?').get(id) : undefined
    // 404, not 400: the page then tells the passkey provider to forget the credential (WebAuthn Signal API).
    // Not when no passkey is registered at all: that is a server started on the wrong or an empty auth.db,
    // and the owner's only way in must not be hidden from their password manager because of it.
    if (!row) {
      if (this.listPasskeys().length === 0) throw new ArtifactError(400, 'passkey_rejected', '這把 passkey 沒有註冊在這裡。')
      throw new ArtifactError(404, 'unknown_passkey', '這把 passkey 沒有註冊在這裡。')
    }
    const verification = await verifyAuthenticationResponse({
      response: response as AuthenticationResponseJSON,
      expectedChallenge: challenge.challenge,
      expectedOrigin: this.origin,
      expectedRPID: this.rpID,
      credential: {
        id: String(row.id),
        publicKey: new Uint8Array(row.public_key as Uint8Array),
        counter: Number(row.counter),
        transports: row.transports ? JSON.parse(String(row.transports)) : undefined
      },
      requireUserVerification: true
    }).catch(() => ({ verified: false, authenticationInfo: null }))
    if (!verification.verified || !verification.authenticationInfo) throw new ArtifactError(400, 'passkey_rejected', '無法驗證這把 passkey。')
    // Verification awaited: the passkey may have been removed meanwhile, and a removed passkey signs nobody in.
    const still = this.db.prepare('UPDATE passkeys SET counter = ?, last_used_at = ? WHERE id = ?').run(verification.authenticationInfo.newCounter, nowIso(), row.id)
    if (still.changes === 0) throw new ArtifactError(404, 'unknown_passkey', '這把 passkey 已被移除。')
  }
}

export function readCookie(c: Context, name: string): string | undefined {
  const header = c.req.header('cookie') ?? ''
  for (const part of header.split(';')) {
    const at = part.indexOf('=')
    if (at > 0 && part.slice(0, at).trim() === name) return part.slice(at + 1).trim()
  }
  return undefined
}

export function sessionCookie(value: string, maxAgeSeconds = SESSION_DAYS * 86400): string {
  return `${SESSION_COOKIE}=${value}; Path=/; Secure; HttpOnly; SameSite=Lax; Max-Age=${maxAgeSeconds}`
}

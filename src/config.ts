import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'

export type Config = {
  dataDir: string
  appPort: number
  contentPort: number
  /** Public origin of the gallery, viewer and API. */
  appOrigin: string
  /** Public origin that serves artifact files; must differ from appOrigin (stored-XSS isolation). */
  contentOrigin: string
  tokens: string[]
  /**
   * The internet-facing pair, reached through the Cloudflare tunnel: read-only, and only artifacts marked
   * public. Separate listeners with their own origins, so nothing on the tailnet pair is exposed by accident.
   */
  public?: { appPort: number; contentPort: number; appOrigin: string; contentOrigin: string }
  office?: { url: string; internalUrl: string; callbackOrigin: string; secret: string }
}

function readTokens(env: NodeJS.ProcessEnv): string[] {
  const fromFile = env.ARTIFACTS_TOKEN_FILE ? readFileSync(env.ARTIFACTS_TOKEN_FILE, 'utf8') : ''
  const all = `${fromFile}\n${env.ARTIFACTS_TOKENS ?? ''}`
    .split(/[\n,]/)
    .map((token) => token.trim())
    .filter((token) => token.length > 0 && !token.startsWith('#'))
  if (all.length === 0) {
    throw new Error('Set ARTIFACTS_TOKEN_FILE or ARTIFACTS_TOKENS: publishing needs at least one token.')
  }
  if (all.some((token) => token.length < 16) && env.NODE_ENV === 'production') {
    throw new Error('Tokens must be at least 16 characters in production.')
  }
  return all
}

function origin(value: string): string {
  return new URL(value).origin
}

export function loadConfig(env: NodeJS.ProcessEnv = process.env): Config {
  const appPort = Number(env.APP_PORT ?? 8790)
  const contentPort = Number(env.CONTENT_PORT ?? 8791)
  const appOrigin = origin(env.APP_ORIGIN ?? `http://localhost:${appPort}`)
  const contentOrigin = origin(env.CONTENT_ORIGIN ?? `http://127.0.0.1:${contentPort}`)
  if (appOrigin === contentOrigin) {
    throw new Error('APP_ORIGIN and CONTENT_ORIGIN must differ.')
  }
  let office: Config['office']
  if (env.ONLYOFFICE_URL) {
    const url = origin(env.ONLYOFFICE_URL)
    const internalUrl = origin(env.ONLYOFFICE_INTERNAL_URL ?? url)
    const callbackOrigin = origin(env.ONLYOFFICE_CALLBACK_ORIGIN ?? appOrigin)
    if ([appOrigin, contentOrigin].includes(url)) throw new Error('ONLYOFFICE_URL must use its own origin.')
    if (!env.ONLYOFFICE_JWT_SECRET_FILE) throw new Error('Set ONLYOFFICE_JWT_SECRET_FILE when ONLYOFFICE_URL is enabled.')
    const secret = readFileSync(env.ONLYOFFICE_JWT_SECRET_FILE, 'utf8').trim()
    if (secret.length < 32) throw new Error('The ONLYOFFICE JWT secret must be at least 32 characters.')
    office = { url, internalUrl, callbackOrigin, secret }
  }
  let publicPair: Config['public']
  if (env.PUBLIC_APP_ORIGIN || env.PUBLIC_CONTENT_ORIGIN) {
    if (!env.PUBLIC_APP_ORIGIN || !env.PUBLIC_CONTENT_ORIGIN) throw new Error('Set both PUBLIC_APP_ORIGIN and PUBLIC_CONTENT_ORIGIN, or neither.')
    publicPair = {
      appPort: Number(env.PUBLIC_APP_PORT ?? 8792),
      contentPort: Number(env.PUBLIC_CONTENT_PORT ?? 8793),
      appOrigin: origin(env.PUBLIC_APP_ORIGIN),
      contentOrigin: origin(env.PUBLIC_CONTENT_ORIGIN)
    }
    if (new Set([appOrigin, contentOrigin, publicPair.appOrigin, publicPair.contentOrigin]).size !== 4) {
      throw new Error('The tailnet and public app and content origins must all differ.')
    }
  }
  return {
    dataDir: resolve(env.DATA_DIR ?? '/data'),
    appPort,
    contentPort,
    appOrigin,
    contentOrigin,
    tokens: readTokens(env),
    ...(publicPair ? { public: publicPair } : {}),
    ...(office ? { office } : {})
  }
}

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
  return {
    dataDir: resolve(env.DATA_DIR ?? '/data'),
    appPort,
    contentPort,
    appOrigin,
    contentOrigin,
    tokens: readTokens(env)
  }
}

import { createHash } from 'node:crypto'
import { readFileSync, statSync } from 'node:fs'
import type { Context } from 'hono'

export const hashOf = (body: string | Uint8Array): string => createHash('sha256').update(body).digest('hex').slice(0, 12)

/**
 * Sends a static asset with a validator, so a reload costs a 304 instead of the bytes. A request that
 * names the asset's current hash (`?v=`) may be kept for good: the next change gets a new URL.
 */
export function sendAsset(c: Context, body: string | Uint8Array, type: string, hash: string = hashOf(body)): Response {
  const etag = `"${hash}"`
  c.header('Cache-Control', c.req.query('v') === hash ? 'public, max-age=31536000, immutable' : 'no-cache')
  c.header('ETag', etag)
  const ifNoneMatch = c.req.header('if-none-match')
  if (ifNoneMatch?.split(',').some((tag) => tag.trim().replace(/^W\//, '') === etag)) return c.body(null, 304)
  return c.body(typeof body === 'string' ? body : new Uint8Array(body), 200, { 'Content-Type': type })
}

const publicHashes = new Map<string, { mtimeMs: number; hash: string }>()
/** `/static/<name>?v=<hash>`, rehashed when the file changes on disk (development edits public/ live). */
export function staticUrl(name: string): string {
  const file = new URL(`../public/${name}`, import.meta.url)
  const { mtimeMs } = statSync(file)
  let entry = publicHashes.get(name)
  if (!entry || entry.mtimeMs !== mtimeMs) publicHashes.set(name, (entry = { mtimeMs, hash: hashOf(readFileSync(file)) }))
  return `/static/${name}?v=${entry.hash}`
}

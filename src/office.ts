import { contentTypeFor, isServableType } from './file-paths.ts'
import { ZipReader } from './zip.ts'

/**
 * Office files (PowerPoint, Word, Excel) are published as they are and drawn on the server. Their
 * pictures stay inside the file and are read through `~office/<file name>/<part>`, a path under the
 * version's own directory, so they share the page's CSP scope and no new source is needed.
 */
export const OFFICE_MEDIA_SEGMENT = '~office'

/** Parts a page may load from inside an Office file: pictures and media anywhere in the document's own folder, never XML. */
const MEDIA_PART = /^(ppt|word|xl)\/(?:[^/]+\/)*[^/]+$/

/** A URL path segment that is also safe inside CSS url('…') and an HTML attribute. */
export const encodeSegment = (segment: string): string =>
  encodeURIComponent(segment).replace(/['()*!]/g, (char) => `%${char.charCodeAt(0).toString(16).toUpperCase()}`)

/** Where a page drawn from `fileName` finds the parts inside it, relative to the page's own URL. */
export const officeMediaBase = (fileName: string): string => `${OFFICE_MEDIA_SEGMENT}/${encodeSegment(fileName)}/`

/** The URL of one part, for a page whose media base is `mediaBase`. */
export const officeMediaUrl = (mediaBase: string, part: string): string => `${mediaBase}${part.split('/').map(encodeSegment).join('/')}`

/** True when a part is a picture or media file a browser can show (what the media route will serve). */
export function isDrawablePart(part: string): boolean {
  if (!MEDIA_PART.test(part)) return false
  const type = contentTypeFor(part)
  return isServableType(type) && !type.startsWith('text/')
}

/** One picture inside an Office file, for the media route; null for anything else. */
export function officeMedia(data: Buffer, part: string): { bytes: Buffer; type: string } | null {
  if (!isDrawablePart(part)) return null
  try {
    const bytes = new ZipReader(data).read(part)
    return bytes ? { bytes, type: contentTypeFor(part) } : null
  } catch {
    return null
  }
}

/** Escaped, as the error pages of every Office renderer show it, with a throw that lands in the viewer's error log. */
export function failureScript(message: string): string {
  return `<script>throw new Error(${JSON.stringify(message).replace(/</g, '\\u003c')})</script>`
}

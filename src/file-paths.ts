import { extname } from 'node:path'

const CONTENT_TYPES: Record<string, string> = {
  '.html': 'text/html; charset=utf-8',
  '.htm': 'text/html; charset=utf-8',
  '.md': 'text/markdown; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.txt': 'text/plain; charset=utf-8',
  '.csv': 'text/csv; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.webp': 'image/webp',
  '.avif': 'image/avif',
  '.ico': 'image/x-icon',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
  '.ttf': 'font/ttf',
  '.otf': 'font/otf',
  '.mp4': 'video/mp4',
  '.webm': 'video/webm',
  '.mp3': 'audio/mpeg',
  '.wav': 'audio/wav',
  '.pdf': 'application/pdf',
  '.wasm': 'application/wasm',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
}

export function contentTypeFor(path: string): string {
  return CONTENT_TYPES[extname(path).toLowerCase()] ?? 'application/octet-stream'
}

export function isServableType(type: string): boolean {
  return type !== 'application/octet-stream'
}

/** Relative, forward-slash path with no traversal; returns null when unsafe. */
export function normalizeFilePath(input: string): string | null {
  if (typeof input !== 'string' || input.length === 0 || input.length > 512) {
    return null
  }
  if (input.includes('\\') || input.includes('\0') || input.startsWith('/')) {
    return null
  }
  const segments = input.split('/')
  if (segments.some((segment) => segment === '' || segment === '.' || segment === '..')) {
    return null
  }
  return segments.join('/')
}

export function isMarkdownPath(path: string): boolean {
  return extname(path).toLowerCase() === '.md'
}

export function isPptxPath(path: string): boolean {
  return extname(path).toLowerCase() === '.pptx'
}

export function isDocxPath(path: string): boolean {
  return extname(path).toLowerCase() === '.docx'
}

export function isXlsxPath(path: string): boolean {
  return extname(path).toLowerCase() === '.xlsx'
}

/** PowerPoint, Word and Excel files: published as they are, drawn on the server. */
export function isOfficePath(path: string): boolean {
  return isPptxPath(path) || isDocxPath(path) || isXlsxPath(path)
}

/** A file the content origin renders into a page: HTML as is, Markdown and Office files drawn on the server. */
export function isPagePath(path: string): boolean {
  return isHtmlPath(path) || isMarkdownPath(path) || isOfficePath(path)
}

export function isHtmlPath(path: string): boolean {
  return ['.html', '.htm'].includes(extname(path).toLowerCase())
}

/** Decodes and normalizes a path taken from a request URL; malformed escapes are null, not a throw. */
export function decodeRequestPath(encoded: string): string | null {
  try {
    return normalizeFilePath(decodeURIComponent(encoded))
  } catch {
    return null
  }
}

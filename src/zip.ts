import { inflateRawSync } from 'node:zlib'

/** A .pptx is a zip of untrusted parts: bounded in count and inflated size, read only by name. */
export const ZIP_LIMITS = { entries: 5000, entryBytes: 64 * 1024 * 1024, totalBytes: 256 * 1024 * 1024 }

export class ZipError extends Error {}

type Entry = { method: number; compressedSize: number; size: number; offset: number }

export class ZipReader {
  private readonly entries = new Map<string, Entry>()
  private inflated = 0
  private readonly data: Buffer

  constructor(data: Buffer) {
    this.data = data
    // The end-of-central-directory record sits in the last 64 KiB + 22 bytes.
    let end = -1
    for (let at = data.length - 22; at >= Math.max(0, data.length - 65557); at--) {
      if (data.readUInt32LE(at) === 0x06054b50) {
        end = at
        break
      }
    }
    if (end < 0) throw new ZipError('not a zip file')
    const count = data.readUInt16LE(end + 10)
    let at = data.readUInt32LE(end + 16)
    if (count > ZIP_LIMITS.entries) throw new ZipError('too many zip entries')
    for (let i = 0; i < count; i++) {
      if (at + 46 > data.length || data.readUInt32LE(at) !== 0x02014b50) throw new ZipError('broken zip directory')
      const method = data.readUInt16LE(at + 10)
      const compressedSize = data.readUInt32LE(at + 20)
      const size = data.readUInt32LE(at + 24)
      const nameLength = data.readUInt16LE(at + 28)
      const extraLength = data.readUInt16LE(at + 30)
      const commentLength = data.readUInt16LE(at + 32)
      const offset = data.readUInt32LE(at + 42)
      const name = data.subarray(at + 46, at + 46 + nameLength).toString('utf8')
      at += 46 + nameLength + extraLength + commentLength
      if (name.endsWith('/')) continue
      this.entries.set(name.replace(/^\/+/, ''), { method, compressedSize, size, offset })
    }
  }

  has(name: string): boolean {
    return this.entries.has(name)
  }

  names(): string[] {
    return [...this.entries.keys()]
  }

  read(name: string): Buffer | null {
    const entry = this.entries.get(name)
    if (!entry) return null
    if (entry.size > ZIP_LIMITS.entryBytes) throw new ZipError(`${name} is too large`)
    const at = entry.offset
    if (at + 30 > this.data.length || this.data.readUInt32LE(at) !== 0x04034b50) throw new ZipError('broken zip entry')
    const start = at + 30 + this.data.readUInt16LE(at + 26) + this.data.readUInt16LE(at + 28)
    const raw = this.data.subarray(start, start + entry.compressedSize)
    let out: Buffer
    if (entry.method === 0) out = Buffer.from(raw)
    else if (entry.method === 8) {
      try {
        out = inflateRawSync(raw, { maxOutputLength: Math.max(1, entry.size) })
      } catch {
        throw new ZipError(`${name} does not inflate`)
      }
    } else throw new ZipError(`${name} uses an unsupported compression method`)
    this.inflated += out.length
    if (this.inflated > ZIP_LIMITS.totalBytes) throw new ZipError('the zip inflates past its limit')
    return out
  }

  text(name: string): string | null {
    return this.read(name)?.toString('utf8') ?? null
  }
}

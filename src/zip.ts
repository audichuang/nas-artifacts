import { constants, inflateRawSync } from 'node:zlib'

/**
 * An Office file is a zip of untrusted parts: bounded in count and inflated size, read only by name. 20,000 entries
 * (a few MB of directory) fits a deck of about 5000 slides with notes; a 64 MB upload could otherwise list close to
 * a million.
 */
export const ZIP_LIMITS = { entries: 20_000, entryBytes: 64 * 1024 * 1024, totalBytes: 256 * 1024 * 1024 }

/** `limit`: the file is readable but past one of ZIP_LIMITS (too big for this page), not broken. */
export class ZipError extends Error {
  readonly limit: boolean
  constructor(message: string, limit = false) {
    super(message)
    this.limit = limit
  }
}

const fold = (name: string) => name.replace(/\\/g, '/').replace(/^\/+/, '').toLowerCase()

type Entry = { method: number; compressedSize: number; size: number; offset: number }

export class ZipReader {
  private readonly entries = new Map<string, Entry>()
  // Some writers store part names with backslashes or another case; Office still opens them.
  private readonly folded = new Map<string, string>()
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
    if (end < 0) throw new ZipError('不是 zip 檔')
    const count = data.readUInt16LE(end + 10)
    let at = data.readUInt32LE(end + 16)
    if (count > ZIP_LIMITS.entries) throw new ZipError(`檔案太大:裡面有 ${count.toLocaleString('en-US')} 個項目,超過網頁版能讀取的 ${ZIP_LIMITS.entries.toLocaleString('en-US')} 個`, true)
    for (let i = 0; i < count; i++) {
      if (at + 46 > data.length || data.readUInt32LE(at) !== 0x02014b50) throw new ZipError('zip 檔的目錄已損壞')
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
      const key = name.replace(/^\/+/, '')
      this.entries.set(key, { method, compressedSize, size, offset })
      if (!this.folded.has(fold(key))) this.folded.set(fold(key), key)
    }
  }

  private entry(name: string): Entry | undefined {
    return this.entries.get(name) ?? this.entries.get(this.folded.get(fold(name)) ?? '')
  }

  has(name: string): boolean {
    return this.entry(name) !== undefined
  }

  names(): string[] {
    return [...this.entries.keys()]
  }

  /** The compressed bytes of an entry, found through its local header. */
  private raw(name: string, entry: Entry): Buffer {
    const at = entry.offset
    if (at + 30 > this.data.length || this.data.readUInt32LE(at) !== 0x04034b50) throw new ZipError('zip 檔內的項目已損壞')
    const start = at + 30 + this.data.readUInt16LE(at + 26) + this.data.readUInt16LE(at + 28)
    if (entry.method !== 0 && entry.method !== 8) throw new ZipError(`${name} 使用不支援的壓縮方式`)
    return this.data.subarray(start, start + entry.compressedSize)
  }

  private count(bytes: number): void {
    this.inflated += bytes
    if (this.inflated > ZIP_LIMITS.totalBytes) throw new ZipError('檔案太大:解壓縮後超過網頁版能讀取的大小上限', true)
  }

  read(name: string): Buffer | null {
    const entry = this.entry(name)
    if (!entry) return null
    if (entry.size > ZIP_LIMITS.entryBytes) throw new ZipError(`檔案太大:${name} 超過網頁版能讀取的單一項目大小上限`, true)
    const raw = this.raw(name, entry)
    let out: Buffer
    if (entry.method === 0) out = Buffer.from(raw)
    else {
      try {
        out = inflateRawSync(raw, { maxOutputLength: Math.max(1, entry.size) })
      } catch {
        throw new ZipError(`${name} 無法解壓縮`)
      }
    }
    this.count(out.length)
    return out
  }

  /**
   * The first `limit` bytes of an entry, for parts too large to read whole: `cut` says whether more
   * followed. Deflate cannot start mid-stream, so ever longer slices of the compressed bytes are
   * inflated (a cut stream flushes what it has) until the output is close to the limit; a slice that
   * would pass the limit gives no output at all, so the search then narrows between the last two.
   */
  readPrefix(name: string, limit: number = ZIP_LIMITS.entryBytes): { data: Buffer; cut: boolean } | null {
    const entry = this.entry(name)
    if (!entry) return null
    const max = Math.max(1, Math.min(limit, ZIP_LIMITS.entryBytes))
    if (entry.size <= max) return { data: this.read(name)!, cut: false }
    const raw = this.raw(name, entry)
    if (entry.method === 0) {
      const data = Buffer.from(raw.subarray(0, max))
      this.count(data.length)
      return { data, cut: raw.length > max }
    }
    let best = Buffer.alloc(0)
    let good = 0
    let bad = raw.length + 1
    let take = Math.min(raw.length, Math.max(1024, Math.floor(max / 64)))
    for (let tries = 0; tries < 60 && take > good && take < bad; tries++) {
      try {
        best = inflateRawSync(raw.subarray(0, take), { finishFlush: constants.Z_SYNC_FLUSH, maxOutputLength: max })
        good = take
        if (take >= raw.length || best.length >= max * 0.9) break
        const grow = best.length ? Math.min(8, (max * 0.95) / best.length) : 8
        take = Math.min(raw.length, bad - 1, Math.max(take + 64, Math.floor(take * grow)))
      } catch (error) {
        if (!(error instanceof RangeError)) throw new ZipError(`${name} 無法解壓縮`)
        bad = take
        if (bad - good <= 64) break
        take = Math.floor((good + bad) / 2)
      }
    }
    this.count(best.length)
    return { data: best, cut: good < raw.length }
  }

  text(name: string): string | null {
    return this.read(name)?.toString('utf8') ?? null
  }
}

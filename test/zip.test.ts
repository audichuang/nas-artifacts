import assert from 'node:assert/strict'
import { describe, test } from 'node:test'
import { crc32, deflateRawSync } from 'node:zlib'
import { ZIP_LIMITS, ZipError, ZipReader } from '../src/zip.ts'

/** A zip whose entries are deflated (method 8), as every Office tool writes them. */
function deflatedZip(files: Record<string, Buffer | string>): Buffer {
  const locals: Buffer[] = []
  const central: Buffer[] = []
  let offset = 0
  for (const [name, content] of Object.entries(files)) {
    const data = Buffer.from(content)
    const packed = deflateRawSync(data)
    const nameBytes = Buffer.from(name)
    const local = Buffer.alloc(30)
    local.writeUInt32LE(0x04034b50, 0)
    local.writeUInt16LE(8, 8)
    local.writeUInt32LE(crc32(data), 14)
    local.writeUInt32LE(packed.length, 18)
    local.writeUInt32LE(data.length, 22)
    local.writeUInt16LE(nameBytes.length, 26)
    const entry = Buffer.alloc(46)
    entry.writeUInt32LE(0x02014b50, 0)
    entry.writeUInt16LE(8, 10)
    entry.writeUInt32LE(crc32(data), 16)
    entry.writeUInt32LE(packed.length, 20)
    entry.writeUInt32LE(data.length, 24)
    entry.writeUInt16LE(nameBytes.length, 28)
    entry.writeUInt32LE(offset, 42)
    locals.push(local, nameBytes, packed)
    central.push(entry, nameBytes)
    offset += 30 + nameBytes.length + packed.length
  }
  const size = central.reduce((sum, part) => sum + part.length, 0)
  const end = Buffer.alloc(22)
  end.writeUInt32LE(0x06054b50, 0)
  end.writeUInt16LE(Object.keys(files).length, 8)
  end.writeUInt16LE(Object.keys(files).length, 10)
  end.writeUInt32LE(size, 12)
  end.writeUInt32LE(offset, 16)
  return Buffer.concat([...locals, ...central, end])
}

/** Rows of varied text, so the deflate stream is not one long repeat. */
const rows = (count: number) => Array.from({ length: count }, (_, at) => `<row r="${at + 1}"><c><v>${(at * 7919) % 100003}</v></c></row>`).join('')

describe('reading a zip', () => {
  test('a whole entry reads as before; a prefix of a small entry is the whole entry', () => {
    const text = rows(500)
    const reader = new ZipReader(deflatedZip({ 'a.xml': text }))
    assert.equal(reader.text('a.xml'), text)
    assert.deepEqual(reader.readPrefix('a.xml', text.length), { data: Buffer.from(text), cut: false })
    assert.equal(reader.readPrefix('missing.xml', 10), null)
  })

  test('a deflated entry past the limit gives its first bytes, exactly, close to the limit', () => {
    const text = rows(200_000)
    const reader = new ZipReader(deflatedZip({ 'big.xml': text }))
    for (const limit of [5_000, 100_000, 1_000_000]) {
      const prefix = reader.readPrefix('big.xml', limit)!
      assert.equal(prefix.cut, true)
      assert.ok(prefix.data.length <= limit, `${prefix.data.length} <= ${limit}`)
      assert.ok(prefix.data.length >= limit * 0.85, `${prefix.data.length} near ${limit}`)
      assert.equal(prefix.data.toString(), text.slice(0, prefix.data.length))
    }
  })

  test('a prefix never goes past the per-entry limit, and whole reads past it still refuse', () => {
    const saved = ZIP_LIMITS.entryBytes
    ZIP_LIMITS.entryBytes = 50_000
    try {
      const text = rows(20_000)
      const reader = new ZipReader(deflatedZip({ 'big.xml': text }))
      assert.throws(() => reader.read('big.xml'), (error: Error) => error instanceof ZipError && /太大/.test(error.message))
      const prefix = reader.readPrefix('big.xml')!
      assert.ok(prefix.cut && prefix.data.length <= 50_000 && prefix.data.length > 40_000)
      assert.ok(reader.readPrefix('big.xml', 10_000_000)!.data.length <= 50_000)
    } finally {
      ZIP_LIMITS.entryBytes = saved
    }
  })

  test('what prefixes inflate counts toward the whole-file budget', () => {
    const saved = { ...ZIP_LIMITS }
    ZIP_LIMITS.entryBytes = 40_000
    ZIP_LIMITS.totalBytes = 100_000
    try {
      const reader = new ZipReader(deflatedZip({ 'big.xml': rows(20_000) }))
      reader.readPrefix('big.xml')
      reader.readPrefix('big.xml')
      assert.throws(() => reader.readPrefix('big.xml'), (error: Error) => error instanceof ZipError && /上限/.test(error.message))
    } finally {
      Object.assign(ZIP_LIMITS, saved)
    }
  })

  test('a stored entry past the limit gives its first bytes', () => {
    const saved = ZIP_LIMITS.entryBytes
    ZIP_LIMITS.entryBytes = 1_000
    try {
      const text = rows(1_000)
      const data = Buffer.from(text)
      const reader = new ZipReader(storedZip('a.xml', data))
      const prefix = reader.readPrefix('a.xml')!
      assert.deepEqual(prefix, { data: data.subarray(0, 1_000), cut: true })
    } finally {
      ZIP_LIMITS.entryBytes = saved
    }
  })

  test('a broken deflate stream is a ZipError in Traditional Chinese, as are the other zip errors', () => {
    const good = deflatedZip({ 'a.xml': rows(50_000) })
    const start = 30 + 'a.xml'.length
    const broken = Buffer.from(good)
    broken.fill(0xff, start + 10, start + 400)
    const saved = ZIP_LIMITS.entryBytes
    ZIP_LIMITS.entryBytes = 10_000
    try {
      assert.throws(() => new ZipReader(broken).readPrefix('a.xml'), (error: Error) => error instanceof ZipError && /無法解壓縮/.test(error.message))
    } finally {
      ZIP_LIMITS.entryBytes = saved
    }
    assert.throws(() => new ZipReader(Buffer.from('not a zip at all, just text')), /不是 zip 檔/)
  })
})

function storedZip(name: string, data: Buffer): Buffer {
  const nameBytes = Buffer.from(name)
  const local = Buffer.alloc(30)
  local.writeUInt32LE(0x04034b50, 0)
  local.writeUInt32LE(data.length, 18)
  local.writeUInt32LE(data.length, 22)
  local.writeUInt16LE(nameBytes.length, 26)
  const entry = Buffer.alloc(46)
  entry.writeUInt32LE(0x02014b50, 0)
  entry.writeUInt32LE(data.length, 20)
  entry.writeUInt32LE(data.length, 24)
  entry.writeUInt16LE(nameBytes.length, 28)
  const end = Buffer.alloc(22)
  end.writeUInt32LE(0x06054b50, 0)
  end.writeUInt16LE(1, 8)
  end.writeUInt16LE(1, 10)
  end.writeUInt32LE(46 + nameBytes.length, 12)
  end.writeUInt32LE(30 + nameBytes.length + data.length, 16)
  return Buffer.concat([local, nameBytes, data, entry, nameBytes, end])
}

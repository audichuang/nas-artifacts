import { createHash, randomUUID } from 'node:crypto'
import { existsSync, mkdirSync, readdirSync, readFileSync, renameSync, rmSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

/** Content-addressed file bodies shared by every version and asset, so a republish only stores what changed. */
export class BlobStore {
  private readonly root: string

  constructor(dataDir: string) {
    this.root = join(dataDir, 'blobs')
    mkdirSync(this.root, { recursive: true })
  }

  private path(sha256: string): string {
    return join(this.root, sha256.slice(0, 2), sha256.slice(2))
  }

  static hash(body: Buffer): string {
    return createHash('sha256').update(body).digest('hex')
  }

  has(sha256: string): boolean {
    return existsSync(this.path(sha256))
  }

  put(body: Buffer): string {
    const sha256 = BlobStore.hash(body)
    const target = this.path(sha256)
    if (!existsSync(target)) {
      mkdirSync(join(this.root, sha256.slice(0, 2)), { recursive: true })
      const temp = `${target}.${randomUUID()}.tmp`
      writeFileSync(temp, body)
      renameSync(temp, target)
    }
    return sha256
  }

  read(sha256: string): Buffer {
    return readFileSync(this.path(sha256))
  }

  list(): string[] {
    return readdirSync(this.root).flatMap((prefix) =>
      /^[0-9a-f]{2}$/.test(prefix)
        ? readdirSync(join(this.root, prefix))
            .filter((rest) => /^[0-9a-f]{62}$/.test(rest))
            .map((rest) => prefix + rest)
        : []
    )
  }

  remove(sha256: string): void {
    rmSync(this.path(sha256), { force: true })
  }
}

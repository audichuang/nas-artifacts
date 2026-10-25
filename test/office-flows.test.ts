import assert from 'node:assert/strict'
import { execFile } from 'node:child_process'
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { after, before, describe, test } from 'node:test'
import { serve, type ServerType } from '@hono/node-server'
import { createAppServer } from '../src/app-server.ts'
import { ArtifactStore } from '../src/artifact-store.ts'
import { isDrawablePart } from '../src/office.ts'
import { zip } from './helpers/zip.ts'

const CLI = new URL('../skill/nas-artifacts/scripts/artifact.mjs', import.meta.url).pathname
const TOKEN = 'office-flow-token-0123456789'
const officeFile = (marker: string) => zip({ '[Content_Types].xml': `<Types>${marker}</Types>`, 'docProps/app.xml': '<Properties/>' })

let dir: string
let store: ArtifactStore
let server: ServerType
let url: string

function cli(args: string[]): Promise<{ code: number; stdout: Buffer; stderr: string }> {
  return new Promise((resolve) => {
    execFile(
      process.execPath,
      [CLI, ...args],
      { env: { ...process.env, ARTIFACTS_URL: url, ARTIFACTS_TOKEN: TOKEN, ARTIFACTS_CONFIG_DIR: join(dir, 'config') }, cwd: dir, encoding: 'buffer' },
      (error, stdout, stderr) => resolve({ code: error ? Number(error.code ?? 1) : 0, stdout, stderr: stderr.toString() })
    )
  })
}

before(async () => {
  dir = mkdtempSync(join(tmpdir(), 'office-flows-'))
  store = new ArtifactStore(join(dir, 'data'))
  const app = createAppServer(store, { dataDir: dir, appPort: 0, contentPort: 0, appOrigin: 'http://app.test', contentOrigin: 'http://content.test', tokens: [TOKEN] })
  await new Promise<void>((resolve) => {
    server = serve({ fetch: app.fetch, port: 0, hostname: '127.0.0.1' }, (info) => {
      url = `http://127.0.0.1:${info.port}`
      resolve()
    })
  })
})

after(() => {
  server.close()
  store.close()
  rmSync(dir, { recursive: true, force: true })
})

describe('Office files next to pages', () => {
  test('a page with an Office file beside it is still the entry, as before Office files were pages', () => {
    const { entry } = store.publish({ title: 'R', files: { 'report.md': { content: Buffer.from('# R') }, 'data.xlsx': { content: officeFile('a') } } })
    assert.equal(entry, 'report.md')
    // Office files alone: one of them is the entry.
    assert.equal(store.publish({ title: 'D', files: { 'deck.pptx': { content: officeFile('b') }, 'a.png': { content: Buffer.from('png') } } }).entry, 'deck.pptx')
  })

  test('macro-enabled files, slide shows, templates and the old binary formats are Office pages too', () => {
    for (const name of ['a.pptm', 'a.ppsx', 'a.potx', 'a.docm', 'a.dotx', 'a.xlsm', 'a.xltx', 'a.ppt', 'a.doc', 'a.xls', 'UPPER.PPTX']) {
      assert.equal(store.publish({ title: name, files: { [name]: { content: officeFile(name) } } }).entry, name)
    }
  })

  test('the entry error names every kind of page', () => {
    assert.throws(
      () => store.publish({ title: 'X', entry: 'a.png', files: { 'a.png': { content: Buffer.from('png') } } }),
      /\.html, \.md, \.pptx, \.docx or \.xlsx/
    )
  })

  test('a version diff treats Office files as binary, not as lines of zip bytes', () => {
    const { id } = store.publish({ title: 'D', files: { 'r.docx': { content: officeFile('one') } } })
    store.update(id, { files: { 'r.docx': { content: officeFile('two') } } })
    const diff = store.diff(id, 2)
    assert.equal(diff.files[0].note, 'binary')
  })
})

describe('the CLI with Office files', () => {
  test('the default title drops the extension of every Office file', async () => {
    for (const name of ['Q3 報告.docx', 'budget.xlsx', 'talk.pptx', 'macro.xlsm', 'show.ppsx', 'old.doc']) {
      const file = join(dir, name)
      writeFileSync(file, officeFile(name))
      const run = await cli(['publish', file, '--new', '--json'])
      assert.equal(run.code, 0, run.stderr)
      assert.equal(JSON.parse(run.stdout.toString()).title, name.replace(/\.[a-z]+$/, ''))
    }
  })

  test('read does not pour a binary file into the terminal; --out saves it', async () => {
    const bytes = officeFile('read')
    const { id } = store.publish({ title: 'B', files: { 'r.docx': { content: bytes } } })
    const run = await cli(['read', id])
    assert.notEqual(run.code, 0)
    assert.equal(run.stdout.length, 0)
    assert.match(run.stderr, /binary.*--out/)
    mkdirSync(join(dir, 'out'), { recursive: true })
    const out = join(dir, 'out', 'r.docx')
    assert.equal((await cli(['read', id, '--out', out])).code, 0)
    assert.ok(readFileSync(out).equals(bytes))
  })
})

describe('pictures inside Office files', () => {
  test('BMP and JFIF pictures are served, and pictures kept outside a media folder too; XML never is', () => {
    for (const part of ['ppt/media/a.bmp', 'word/media/b.jfif', 'xl/media/c.png', 'ppt/images/d.jpeg', 'word/drawings/e.gif', 'ppt/slides/media/f.png']) assert.equal(isDrawablePart(part), true, part)
    for (const part of ['ppt/slides/slide1.xml', 'word/document.xml', 'xl/media/x.svgz', 'docProps/thumbnail.jpeg', 'ppt/media/a.txt', 'customXml/a.png']) assert.equal(isDrawablePart(part), false, part)
  })
})

describe('export.pptx', () => {
  test('hands back only a real .pptx; a template, macro or old deck is not relabelled as one', async () => {
    const app = createAppServer(store, { dataDir: dir, appPort: 0, contentPort: 0, appOrigin: 'http://app.test', contentOrigin: 'http://content.test', tokens: [TOKEN] })
    const real = store.publish({ title: 'Real', files: { 'a.pptx': { content: officeFile('a') } } })
    assert.equal((await app.request(`http://app.test/api/artifacts/${real.id}/export.pptx`)).status, 200)
    for (const name of ['Template.potx', 'Macro.pptm', 'Show.ppsx', 'Old.ppt']) {
      const { id } = store.publish({ title: name, files: { [name]: { content: officeFile(name) } } })
      const response = await app.request(`http://app.test/api/artifacts/${id}/export.pptx`)
      assert.equal(response.status, 409, name)
      assert.match(((await response.json()) as { error: { message: string } }).error.message, /read \S+ --out/)
    }
  })
})

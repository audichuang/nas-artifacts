import assert from 'node:assert/strict'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { inflateRawSync } from 'node:zlib'
import { after, before, describe, test } from 'node:test'
import { createAppServer } from '../src/app-server.ts'
import { ArtifactStore } from '../src/artifact-store.ts'
import { createContentServer } from '../src/content-server.ts'
import { renderSlides, splitSlides } from '../src/slides.ts'

describe('slide source', () => {
  test('splits on --- lines, keeps fenced code whole, and takes notes after ???', () => {
    const slides = splitSlides(['# A', '', '???', 'say hi', '---', '## B', '```yaml', 'a: 1', '---', 'b: 2', '```', '???', 'note', '', '---', ''].join('\n'))
    assert.deepEqual(slides, [
      { body: '# A', notes: 'say hi' },
      { body: '## B\n```yaml\na: 1\n---\nb: 2\n```', notes: 'note' }
    ])
  })

  test('CRLF line endings split the same', () => {
    assert.equal(splitSlides('# A\r\n---\r\n# B\r\n').length, 2)
  })

  test('layouts follow the content unless a slide names one', () => {
    const { html, count } = renderSlides(
      ['# Deck', 'sub', '---', '## Part two', '---', '## List', '- a', '---', '## Two', 'left', '|||', 'right', '---', '<!-- layout: center -->', '> quote'].join('\n')
    )
    assert.equal(count, 5)
    assert.deepEqual([...html.matchAll(/class="slide layout-([a-z]+)"/g)].map((match) => match[1]), ['cover', 'section', 'default', 'columns', 'center'])
    assert.match(html, /<h2>Two<\/h2>\n<div class="columns" style="--cols:2"><div class="column"><p>left<\/p>\n<\/div><div class="column"><p>right<\/p>/)
    assert.doesNotMatch(html, /layout: center/)
  })

  test('notes are kept out of the slide face; mermaid is drawn on the slide', () => {
    const { html } = renderSlides('## One\n```mermaid\nflowchart LR\n  A --> B\n```\n???\nsecret')
    assert.match(html, /<aside class="slide-notes" hidden><p>secret<\/p>/)
    assert.match(html, /<figure class="diagram"><svg/)
  })
})

describe('serving decks', () => {
  let dir: string
  let store: ArtifactStore
  let content: ReturnType<typeof createContentServer>
  before(() => {
    dir = mkdtempSync(join(tmpdir(), 'slides-test-'))
    store = new ArtifactStore(dir)
    content = createContentServer(store, {
      dataDir: dir,
      appPort: 0,
      contentPort: 0,
      appOrigin: 'http://app.test',
      contentOrigin: 'http://content.test',
      tokens: ['t']
    })
  })
  after(() => {
    store.close()
    rmSync(dir, { recursive: true, force: true })
  })

  test('a *.slides.md entry renders as a deck with its runtime inline; plain .md stays a document', async () => {
    const { id } = store.publish({
      title: 'Deck <1>',
      entry: 'talk.slides.md',
      files: { 'talk.slides.md': { content: Buffer.from('# Hi\n---\n## Two') }, 'notes.md': { content: Buffer.from('# Doc') } }
    })
    const response = await content.request(`http://content.test/a/${id}/v/1/`)
    assert.equal(response.status, 200)
    const html = await response.text()
    assert.match(html, /<body class="deck">/)
    assert.equal(html.match(/<section class="slide /g)?.length, 2)
    assert.match(html, /<title>Deck &#60;1&#62;<\/title>/)
    assert.match(html, /addEventListener\('artifact-deck'/)
    // The CSP is the same one every page gets: the deck needs no new source.
    assert.doesNotMatch(response.headers.get('content-security-policy') ?? '', /deck/)
    const doc = await (await content.request(`http://content.test/a/${id}/v/1/notes.md`)).text()
    assert.doesNotMatch(doc, /class="deck"/)
    assert.match(doc, /markdown-body/)
  })
  test('index.slides.md is picked as the entry like index.md', () => {
    const { entry } = store.publish({ title: 'T', files: { 'index.slides.md': { content: Buffer.from('# A') }, 'other.md': { content: Buffer.from('b') } } })
    assert.equal(entry, 'index.slides.md')
  })
})

describe('editing decks from the viewer', () => {
  let dir: string
  let store: ArtifactStore
  let app: ReturnType<typeof createAppServer>
  before(() => {
    dir = mkdtempSync(join(tmpdir(), 'slides-edit-test-'))
    store = new ArtifactStore(dir)
    app = createAppServer(store, { dataDir: dir, appPort: 0, contentPort: 0, appOrigin: 'http://app.test', contentOrigin: 'http://content.test', tokens: ['t'] })
  })
  after(() => {
    store.close()
    rmSync(dir, { recursive: true, force: true })
  })
  const viewerCall = (path: string, method: string, json: unknown, viewer = true) =>
    app.request(`http://app.test${path}`, {
      method,
      headers: { 'content-type': 'application/json', ...(viewer ? { 'x-artifact-client': 'viewer' } : {}) },
      body: JSON.stringify(json)
    })
  const deck = (extra: Record<string, unknown> = {}) =>
    store.publish({
      title: 'D',
      files: { 'index.slides.md': { content: Buffer.from('# One') }, 'img/a.png': { content: Buffer.from('png') } },
      ...extra
    })

  test('a viewer saves a new version that changes only the deck source', async () => {
    const { id } = deck()
    const response = await viewerCall(`/api/artifacts/${id}/slides`, 'PUT', { markdown: '# One\n---\n## Two', baseVersion: 1, author: 'Ann' })
    assert.equal(response.status, 200, await response.clone().text())
    const v1 = store.requireVersion(id, 1)
    const v2 = store.requireVersion(id, 2)
    assert.deepEqual(Object.keys(v2.files).sort(), ['img/a.png', 'index.slides.md'])
    assert.equal(v2.files['img/a.png'].sha256, v1.files['img/a.png'].sha256)
    assert.equal(store.blobs.read(v2.files['index.slides.md'].sha256).toString(), '# One\n---\n## Two')
    assert.equal(v2.label, '瀏覽器編輯 · Ann')
    assert.equal(v2.source, null)
  })

  test('a stale base is a conflict that names the changed file, not an overwrite', async () => {
    const { id } = deck()
    store.update(id, { files: { 'index.slides.md': { content: Buffer.from('# Someone else') } } })
    const response = await viewerCall(`/api/artifacts/${id}/slides`, 'PUT', { markdown: '# Mine', baseVersion: 1 })
    assert.equal(response.status, 409)
    const body = (await response.json()) as { error: { code: string; details: { changedFiles: string[] } } }
    assert.equal(body.error.code, 'version_conflict')
    assert.deepEqual(body.error.details.changedFiles, ['index.slides.md'])
    assert.equal(store.requireSummary(id).currentVersion, 2)
  })

  test('the viewer cannot edit a deck with capabilities, a page that is not a deck, or without the viewer header', async () => {
    const withDb = deck({ capabilities: ['db'] })
    assert.equal((await viewerCall(`/api/artifacts/${withDb.id}/slides`, 'PUT', { markdown: 'x', baseVersion: 1 })).status, 403)
    assert.equal((await viewerCall(`/api/artifacts/${withDb.id}/slides/preview`, 'POST', { markdown: 'x' })).status, 403)
    const page = store.publish({ title: 'P', files: { 'index.html': { content: Buffer.from('<p>x</p>') } } })
    assert.equal((await viewerCall(`/api/artifacts/${page.id}/slides`, 'PUT', { markdown: 'x', baseVersion: 1 })).status, 409)
    const plain = deck()
    assert.equal((await viewerCall(`/api/artifacts/${plain.id}/slides`, 'PUT', { markdown: 'x', baseVersion: 1 }, false)).status, 401)
    assert.equal((await viewerCall(`/api/artifacts/${plain.id}/slides`, 'PUT', { markdown: 'x' })).status, 400)
    assert.equal(store.requireSummary(plain.id).currentVersion, 1)
  })

  test('a deck that gains capabilities or loses its slides entry while a save is still arriving is not written', async () => {
    for (const change of ['capabilities', 'entry'] as const) {
      const { id } = deck()
      let send!: (text: string) => void
      let finish!: () => void
      const body = new ReadableStream<Uint8Array>({
        start(controller) {
          send = (text) => controller.enqueue(new TextEncoder().encode(text))
          finish = () => controller.close()
        }
      })
      const pending = app.request(`http://app.test/api/artifacts/${id}/slides`, {
        method: 'PUT',
        headers: { 'content-type': 'application/json', 'x-artifact-client': 'viewer' },
        body,
        // @ts-expect-error Node's fetch needs this for a streamed body
        duplex: 'half'
      })
      send('{"markdown":"# viewer text",')
      await new Promise((resolve) => setTimeout(resolve, 20))
      if (change === 'capabilities') store.update(id, { capabilities: ['db'] })
      else store.update(id, { files: { 'index.html': { content: Buffer.from('<p>page</p>') } }, entry: 'index.html' })
      send('"baseVersion":2}')
      finish()
      const response = await pending
      assert.ok([403, 409].includes(response.status), `${change}: ${response.status}`)
      assert.equal(store.requireSummary(id).currentVersion, 2, change)
    }
  })

  test('preview renders a draft without saving it', async () => {
    const { id } = deck()
    const response = await viewerCall(`/api/artifacts/${id}/slides/preview`, 'POST', { markdown: '# A\n---\n## B' })
    const body = (await response.json()) as { html: string; count: number }
    assert.equal(body.count, 2)
    assert.match(body.html, /<h2>B<\/h2>/)
    assert.equal(store.requireSummary(id).currentVersion, 1)
  })
})

/** The entries of a zip file by name, inflated (enough of the format for a .pptx). */
function unzip(zip: Buffer): Map<string, string> {
  const entries = new Map<string, string>()
  const end = zip.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]))
  let at = zip.readUInt32LE(end + 16)
  for (let n = zip.readUInt16LE(end + 10); n > 0; n--) {
    const method = zip.readUInt16LE(at + 10)
    const size = zip.readUInt32LE(at + 20)
    const nameLength = zip.readUInt16LE(at + 28)
    const name = zip.toString('utf8', at + 46, at + 46 + nameLength)
    const local = zip.readUInt32LE(at + 42)
    const start = local + 30 + zip.readUInt16LE(local + 26) + zip.readUInt16LE(local + 28)
    const data = zip.subarray(start, start + size)
    entries.set(name, (method === 8 ? inflateRawSync(data) : data).toString('utf8'))
    at += 46 + nameLength + zip.readUInt16LE(at + 30) + zip.readUInt16LE(at + 32)
  }
  return entries
}

describe('exporting decks', () => {
  let dir: string
  let store: ArtifactStore
  let app: ReturnType<typeof createAppServer>
  before(() => {
    dir = mkdtempSync(join(tmpdir(), 'slides-export-test-'))
    store = new ArtifactStore(dir)
    app = createAppServer(store, { dataDir: dir, appPort: 0, contentPort: 0, appOrigin: 'http://app.test', contentOrigin: 'http://content.test', tokens: ['t'] })
  })
  after(() => {
    store.close()
    rmSync(dir, { recursive: true, force: true })
  })

  // A 1×1 PNG.
  const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64')

  test('a deck downloads as a .pptx with its text, tables, pictures, diagrams and speaker notes', async () => {
    const source = [
      '# 年度報告', '副標', '???', '開場白', '---',
      '## 清單', '- **粗體** 與 `code`', '- 第二點', '---',
      '## 表格', '| A | B |', '|---|---|', '| 1 | 2 |', '---',
      '## 圖', '![logo](img/logo.png)', '```mermaid', 'flowchart LR', '  A --> B', '```'
    ].join('\n')
    const { id } = store.publish({ title: '報告 <v1>', files: { 'deck/index.slides.md': { content: Buffer.from(source) }, 'deck/img/logo.png': { content: PNG } }, entry: 'deck/index.slides.md' })
    const response = await app.request(`http://app.test/api/artifacts/${id}/export.pptx`)
    assert.equal(response.status, 200)
    assert.match(response.headers.get('content-disposition') ?? '', /^attachment; .*filename\*=UTF-8''%E5%A0%B1%E5%91%8A%20_v1_\.pptx$/)
    const entries = unzip(Buffer.from(await response.arrayBuffer()))
    const slides = [...entries.keys()].filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    assert.equal(slides.length, 4)
    assert.match(entries.get('ppt/slides/slide1.xml')!, /年度報告/)
    assert.match(entries.get('ppt/slides/slide2.xml')!, /<a:rPr[^>]* b="1"[^>]*>(?:(?!<\/a:r>).)*<a:t>粗體<\/a:t>/)
    assert.match(entries.get('ppt/slides/slide3.xml')!, /<a:tbl>/)
    const media = [...entries.keys()].filter((name) => name.startsWith('ppt/media/'))
    assert.ok(media.some((name) => name.endsWith('.svg')), media.join(' '))
    // The logo, plus the fallback picture PowerPoint keeps beside every SVG.
    assert.equal(media.filter((name) => name.endsWith('.png')).length, 2, media.join(' '))
    const diagram = [...entries].find(([name]) => name.endsWith('.svg'))![1]
    assert.doesNotMatch(diagram, /var\(/)
    assert.ok([...entries].some(([name, xml]) => name.startsWith('ppt/notesSlides/') && xml.includes('開場白')))
  })

  test('odd Markdown still exports: empty and relative links, control characters, lone surrogates', async () => {
    const source = '## A [empty]() [rel](other.md) [web](https://example.com)\n\nbell\u0007 &#xD800; esc\u001b\n???\nnote\u0001'
    const { id } = store.publish({ title: 'Odd\u0002', files: { 'index.slides.md': { content: Buffer.from(source) } } })
    const response = await app.request(`http://app.test/api/artifacts/${id}/export.pptx`)
    assert.equal(response.status, 200)
    const entries = unzip(Buffer.from(await response.arrayBuffer()))
    for (const [name, xml] of entries) {
      if (name.endsWith('.xml')) assert.doesNotMatch(xml, /[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffd]/, name)
    }
    const slide = entries.get('ppt/slides/slide1.xml')!
    assert.equal((slide.match(/<a:hlinkClick /g) ?? []).length, 1)
  })

  test('crowded slides still export positive sizes, and link targets cannot break the package XML', async () => {
    const code = '```\n' + Array.from({ length: 50 }, (_, i) => `line ${i}`).join('\n') + '\n```'
    const source = [
      `## Code then picture\n${code}\n\n![a](a.png)`,
      `## Eight pictures\n${Array.from({ length: 8 }, () => '![a](a.png)').join('\n\n')}`,
      Array(25).fill('text').join('\n|||\n'),
      '## Link\n[x](https://e.com/a\u0007b)'
    ].join('\n---\n')
    const { id } = store.publish({ title: 'Crowded', files: { 'index.slides.md': { content: Buffer.from(source) }, 'a.png': { content: PNG } } })
    const response = await app.request(`http://app.test/api/artifacts/${id}/export.pptx`)
    assert.equal(response.status, 200)
    for (const [name, xml] of unzip(Buffer.from(await response.arrayBuffer()))) {
      if (!/\.(xml|rels)$/.test(name)) continue
      if (/^ppt\/slides\/slide\d+\.xml$/.test(name)) assert.doesNotMatch(xml, /\b(cx|cy)="-/, `${name} has a negative size`)
      assert.doesNotMatch(xml, /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/, `${name} has a character XML cannot hold`)
    }
  })

  test('a page that is not a deck does not export', async () => {
    const { id } = store.publish({ title: 'P', files: { 'index.md': { content: Buffer.from('# x') } } })
    assert.equal((await app.request(`http://app.test/api/artifacts/${id}/export.pptx`)).status, 409)
  })
})

import assert from 'node:assert/strict'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { after, before, describe, test } from 'node:test'
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

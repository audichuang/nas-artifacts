import assert from 'node:assert/strict'
import { execFile, spawn } from 'node:child_process'
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { after, before, describe, test } from 'node:test'
import { serve, type ServerType } from '@hono/node-server'
import { createAppServer } from '../src/app-server.ts'
import { ArtifactStore } from '../src/artifact-store.ts'

const CLI = new URL('../skill/nas-artifacts/scripts/artifact.mjs', import.meta.url).pathname
const TOKEN = 'cli-test-token-0123456789'
let dir: string
let store: ArtifactStore
let server: ServerType
let url: string

type Run = { code: number; stdout: string; stderr: string }

function cli(args: string[], env: Record<string, string> = {}): Promise<Run> {
  return new Promise((resolve) => {
    execFile(
      process.execPath,
      [CLI, ...args],
      { env: { ...process.env, ARTIFACTS_URL: url, ARTIFACTS_TOKEN: TOKEN, ARTIFACTS_CONFIG_DIR: join(dir, 'config'), ...env }, cwd: dir },
      (error, stdout, stderr) => resolve({ code: error ? Number(error.code ?? 1) : 0, stdout, stderr })
    )
  })
}

function page(html = '<!doctype html><title>x</title><p>x</p>') {
  return store.publish({ title: 'T', files: { 'index.html': { content: Buffer.from(html) } } })
}

before(async () => {
  dir = mkdtempSync(join(tmpdir(), 'artifacts-cli-'))
  store = new ArtifactStore(join(dir, 'data'))
  const app = createAppServer(store, {
    dataDir: dir,
    appPort: 0,
    contentPort: 0,
    appOrigin: 'http://app.test',
    contentOrigin: 'http://content.test',
    tokens: [TOKEN]
  })
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

describe('comments', () => {
  test('--open lists only unresolved threads, with their replies', async () => {
    const { id } = page()
    const open = store.addComment(id, { author: 'u', body: 'fix the chart' })
    store.addComment(id, { author: 'u', body: 'which one?', parentId: open.id })
    const done = store.addComment(id, { author: 'u', body: 'typo in title' })
    store.resolveComment(id, done.id, true)
    const run = await cli(['comments', id, '--open'])
    assert.equal(run.code, 0, run.stderr)
    assert.match(run.stdout, /fix the chart/)
    assert.match(run.stdout, /which one\?/)
    assert.doesNotMatch(run.stdout, /typo in title/)
  })

  test('reply answers in the thread and --resolve closes it', async () => {
    const { id } = page()
    const root = store.addComment(id, { author: 'u', body: 'please add totals' })
    const run = await cli(['comment', 'reply', id, root.id, 'Added a totals row in v2.', '--resolve', '--author', 'claude'])
    assert.equal(run.code, 0, run.stderr)
    const comments = store.listComments(id)
    const reply = comments.find((comment) => comment.parentId === root.id)
    assert.equal(reply?.author, 'claude')
    assert.equal(reply?.body, 'Added a totals row in v2.')
    assert.equal(comments.find((comment) => comment.id === root.id)?.resolved, true)
  })

  test('the author falls back to the configured name, then to agent@host', async () => {
    const { id } = page()
    const root = store.addComment(id, { author: 'u', body: 'q' })
    assert.equal((await cli(['comment', 'reply', id, root.id, 'a'])).code, 0)
    assert.match(store.listComments(id).at(-1)?.author ?? '', /^agent@/)
    assert.equal((await cli(['config', '--author', 'codex'])).code, 0)
    assert.equal((await cli(['comment', 'reply', id, root.id, 'b'])).code, 0)
    assert.equal(store.listComments(id).at(-1)?.author, 'codex')
    rmSync(join(dir, 'config'), { recursive: true, force: true })
  })

  test('resolve and reopen work on a thread root and explain a reply id', async () => {
    const { id } = page()
    const root = store.addComment(id, { author: 'u', body: 'q' })
    const reply = store.addComment(id, { author: 'u', body: 'r', parentId: root.id })
    assert.equal((await cli(['comment', 'resolve', id, root.id])).code, 0)
    assert.equal(store.listComments(id)[0].resolved, true)
    assert.equal((await cli(['comment', 'reopen', id, root.id])).code, 0)
    assert.equal(store.listComments(id)[0].resolved, false)
    const wrong = await cli(['comment', 'resolve', id, reply.id])
    assert.equal(wrong.code, 1)
    assert.match(wrong.stderr, /thread root/)
  })

  test('--wait returns the comments that arrive after it starts', async () => {
    const { id } = page()
    store.addComment(id, { author: 'u', body: 'old feedback' })
    const child = spawn(process.execPath, [CLI, 'comments', id, '--wait', '--timeout', '20'], {
      env: { ...process.env, ARTIFACTS_URL: url, ARTIFACTS_CONFIG_DIR: join(dir, 'config') }
    })
    let stdout = ''
    let stderr = ''
    child.stdout.on('data', (chunk) => { stdout += chunk })
    const exited = new Promise<number>((resolve) => child.on('close', (code) => resolve(code ?? 1)))
    await new Promise<void>((resolve) => child.stderr.on('data', (chunk) => {
      stderr += chunk
      if (stderr.includes('waiting')) resolve()
    }))
    store.addComment(id, { author: 'u', body: 'new feedback' })
    assert.equal(await exited, 0, stderr)
    assert.match(stdout, /new feedback/)
    assert.doesNotMatch(stdout, /old feedback/)
  })

  test('--wait gives up after --timeout with exit code 2', async () => {
    const { id } = page()
    const run = await cli(['comments', id, '--wait', '--timeout', '1'])
    assert.equal(run.code, 2)
    assert.match(run.stderr, /no new comments/)
  })
})

describe('page checks', () => {
  const BAD = `<!doctype html><html><head>
<script src="https://evil.example/lib.js"></script>
<script src="https://unpkg.com/react/umd/react.production.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/d3/7.9.0/d3.min.js"></script>
<link rel="stylesheet" href="/site.css">
</head><body>
<img src="img/missing.png"><img src="https://images.example/a.png"><img src="data:image/png;base64,AAAA">
<a href="#top">top</a><a href="https://example.com" target="_blank">out</a>
<script src="data.js"></script>
<pre><code>fetch('/api') is how you would do it elsewhere</code></pre>
<script>fetch('/api/x'); localStorage.setItem('k', 'v')</script>
</body></html>`

  function site(files: Record<string, string>) {
    const root = mkdtempSync(join(dir, 'site-'))
    for (const [name, text] of Object.entries(files)) writeFileSync(join(root, name), text)
    return root
  }

  test('check names what will break in the sandbox and passes what will not', async () => {
    const root = site({ 'index.html': BAD, 'data.js': 'window.DATA = 1' })
    const run = await cli(['check', root])
    assert.equal(run.code, 1)
    const out = run.stdout
    assert.match(out, /evil\.example/)
    assert.match(out, /unpkg\.com\/react\/.*exact version/)
    assert.doesNotMatch(out, /d3\.min\.js/)
    assert.match(out, /\/site\.css/)
    assert.match(out, /img\/missing\.png/)
    assert.match(out, /images\.example/)
    assert.match(out, /_blank/)
    assert.match(out, /fetch/)
    assert.match(out, /localStorage/)
    assert.doesNotMatch(out, /data\.js|data:image|#top/)
    assert.equal(out.match(/fetch/g)?.length, 1)
  })

  test('the bundled templates pass the check', async () => {
    for (const name of ['report', 'dashboard', 'checklist']) {
      const run = await cli(['check', new URL(`../skill/nas-artifacts/templates/${name}`, import.meta.url).pathname])
      assert.equal(run.code, 0, `${name}: ${run.stdout}`)
    }
  })

  test('publish warns about the same problems but still publishes', async () => {
    const root = site({ 'index.html': BAD, 'data.js': 'window.DATA = 1' })
    const run = await cli(['publish', root])
    assert.equal(run.code, 0, run.stderr)
    assert.match(run.stderr, /img\/missing\.png/)
    assert.match(run.stdout, /\/a\//)
  })

  test('errors prints what viewers hit on the current version', async () => {
    const { id } = page()
    store.recordErrors(id, { version: 1, path: 'index.html', errors: [{ kind: 'csp', message: 'connect-src blocked https://x.test/', source: null }] })
    const run = await cli(['errors', id])
    assert.equal(run.code, 0, run.stderr)
    assert.match(run.stdout, /\[csp\] index\.html ×1 {2}connect-src blocked/)
  })
})

describe('source', () => {
  test('publish says which machine and folder it came from, with the home folder as ~', async () => {
    const root = mkdtempSync(join(dir, 'src-'))
    writeFileSync(join(root, 'index.html'), '<!doctype html><title>s</title><p>s</p>')
    const run = await cli(['publish', root, '--json'], { HOME: dir })
    assert.equal(run.code, 0, run.stderr)
    const { id } = JSON.parse(run.stdout) as { id: string }
    const source = store.requireSummary(id).source
    assert.equal(source?.path, `~/${root.slice(dir.length + 1)}`)
    assert.ok(source?.host)
    const list = await cli(['list'])
    assert.match(list.stdout, new RegExp(`${id}.*${source.host}`))
  })
})

#!/usr/bin/env node
// artifact — publish and manage pages on a nas-artifacts server. Zero dependencies; Node 20+.
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, readFileSync, readdirSync, renameSync, statSync, writeFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { basename, dirname, join, relative, resolve, sep } from 'node:path'
import { parseArgs } from 'node:util'

const CONFIG_DIR = process.env.ARTIFACTS_CONFIG_DIR ?? join(homedir(), '.config', 'nas-artifacts')
const CONFIG_FILE = join(CONFIG_DIR, 'config.json')
const STATE_FILE = join(CONFIG_DIR, 'state.json')
const DEFAULT_URL = 'https://artifacts.ttnas.audichuang.app'
const SKIP_DIRS = new Set(['node_modules', '.git'])

const USAGE = `artifact <command> [options]

  publish <file|dir>     Publish, or republish to the same URL when this path was published before
      --title T --description D --icon EMOJI --entry PATH --label L
      --cap db,uploads   page capabilities (db = shared database, uploads = files people add)
      --id ID            update this artifact instead of the one remembered for the path
      --new              always create a new artifact
      --force            overwrite newer remote changes (default: stop on a conflict)
  checkout <id|url> [dir]  download the current version into dir (default ./<id>) and remember it,
                         so the next publish from dir sends only your changes; --force reuses a non-empty dir
  list [--q TEXT]        list artifacts
  get <id>               metadata and version history
  read <id> [path] [--v N]          print a published file (default: the entry)
  comments <id>          print the viewer comments
  delete <id> --yes      permanently delete
  db list <id> <collection> [--limit N]
  db get|delete <id> <collection> <docId>
  db set|update <id> <collection> <docId> <json>
  db query <id> <collection> <json-where>
  asset upload <id> <file> | asset list <id>
  config [--url URL] [--token-stdin]  save server URL / token to ${CONFIG_FILE}

Common: --json for machine output. Env: ARTIFACTS_URL, ARTIFACTS_TOKEN.`

function fail(message) {
  console.error(`artifact: ${message}`)
  process.exit(1)
}

function readJson(path, fallback) {
  try {
    return JSON.parse(readFileSync(path, 'utf8'))
  } catch {
    return fallback
  }
}

function writePrivateJson(path, value) {
  mkdirSync(dirname(path), { recursive: true, mode: 0o700 })
  const temp = `${path}.${process.pid}.tmp`
  writeFileSync(temp, `${JSON.stringify(value, null, 2)}\n`, { mode: 0o600 })
  renameSync(temp, path)
}

// Re-reads state.json so concurrent publishes keep each other's entries.
function updateState(mutate) {
  const state = readJson(STATE_FILE, {})
  mutate(state)
  writePrivateJson(STATE_FILE, state)
}

function settings() {
  const file = readJson(CONFIG_FILE, {})
  return {
    url: (process.env.ARTIFACTS_URL || file.url || DEFAULT_URL).replace(/\/+$/, ''),
    token: process.env.ARTIFACTS_TOKEN || file.token || ''
  }
}

async function api(path, { method = 'GET', json, body, auth = false, raw = false, headers = {}, allowNotFound = false, allowConflict = false } = {}) {
  const { url, token } = settings()
  if (auth && !token) {
    fail(`no token: set ARTIFACTS_TOKEN or run "artifact config --token-stdin".`)
  }
  const response = await fetch(`${url}${path}`, {
    method,
    headers: {
      ...(auth ? { authorization: `Bearer ${token}` } : {}),
      ...(json !== undefined ? { 'content-type': 'application/json' } : {}),
      ...headers
    },
    body: json !== undefined ? JSON.stringify(json) : body
  }).catch((error) => fail(`cannot reach ${url}: ${error.cause?.code ?? error.message}`))
  if (raw && response.ok) {
    return Buffer.from(await response.arrayBuffer())
  }
  if (response.status === 204) {
    return null
  }
  const payload = await response.json().catch(() => null)
  if (!response.ok) {
    if (allowNotFound && response.status === 404 && payload?.error?.code === 'artifact_not_found') return null
    if (allowConflict && response.status === 409 && payload?.error?.code === 'version_conflict') return { conflict: payload.error }
    fail(`${method} ${path} → ${response.status} ${payload?.error?.code ?? ''}: ${payload?.error?.message ?? response.statusText}`)
  }
  return payload
}

const isSkipped = (name) => name.startsWith('.') || SKIP_DIRS.has(name)

const sha256 = (bytes) => createHash('sha256').update(bytes).digest('hex')

function hashFiles(paths) {
  return Object.fromEntries(Object.entries(paths).map(([path, full]) => [path, sha256(readFileSync(full))]))
}

function collectFiles(root) {
  const stat = statSync(root)
  if (stat.isFile()) {
    return { [basename(root)]: root }
  }
  const files = {}
  const walk = (dir) => {
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
      if (isSkipped(entry.name)) continue
      const full = join(dir, entry.name)
      if (entry.isDirectory()) walk(full)
      else if (entry.isFile()) files[relative(root, full).split(sep).join('/')] = full
    }
  }
  walk(root)
  return files
}

const TEXT_EXTENSIONS = /\.(html?|md|css|m?js|json|txt|csv|svg|xml)$/i

function encodeFile(path) {
  const buffer = readFileSync(path)
  return TEXT_EXTENSIONS.test(path)
    ? { content: buffer.toString('utf8') }
    : { content: buffer.toString('base64'), encoding: 'base64' }
}

function guessTitle(files, entry) {
  const page = entry ? files[entry] : Object.values(files).find((path) => /\.(html?|md)$/i.test(path))
  if (!page) return null
  const text = readFileSync(page, 'utf8')
  const fromHtml = /<title[^>]*>([^<]{1,200})<\/title>/i.exec(text)?.[1]
  const fromMd = /^#\s+(.{1,200})$/m.exec(text)?.[1]
  return (fromHtml ?? fromMd ?? '').trim() || null
}

function print(value, json, human) {
  if (json) console.log(JSON.stringify(value, null, 2))
  else console.log(human(value))
}

function reportConflict(conflict, id, recordedVersion) {
  const details = conflict.details ?? {}
  console.error(`artifact: ${id} is at v${details.currentVersion ?? '?'}, but your copy was based on v${recordedVersion}. Nothing was uploaded.`)
  const changed = details.changedFiles ?? []
  console.error(`changed remotely since v${recordedVersion}: ${changed.length ? changed.join(', ') : '(metadata only)'}`)
  console.error(`  1. artifact checkout ${id} <newdir>   fetch the current version, merge your edits by hand, publish from <newdir>`)
  console.error('  2. rerun the same publish with --force   overwrite those remote changes with your files')
  process.exit(1)
}

async function publish(target, values) {
  const source = resolve(target)
  if (!existsSync(source)) fail(`${target} does not exist`)
  const paths = collectFiles(source)
  if (Object.keys(paths).length === 0) fail(`${target} has no files to publish`)
  const state = readJson(STATE_FILE, {})
  const knownId = values.new ? undefined : values.id ?? state[source]?.id
  const hashes = hashFiles(paths)
  const capabilities = values.cap === undefined ? undefined : values.cap.split(',').map((cap) => cap.trim()).filter(Boolean)
  const common = {
    description: values.description,
    icon: values.icon,
    entry: values.entry,
    label: values.label,
    capabilities
  }
  const encode = (changed) => Object.fromEntries(changed.map((path) => [path, encodeFile(paths[path])]))
  const entry = state[source]
  const recorded = entry && entry.id === knownId && Number.isInteger(entry.version) && entry.files ? entry : null
  let result
  if (knownId) {
    const id = encodeURIComponent(knownId)
    let body
    if (recorded) {
      const files = encode(Object.keys(paths).filter((path) => recorded.files[path] !== hashes[path]))
      for (const path of Object.keys(recorded.files)) if (!(path in paths)) files[path] = null
      body = { ...common, title: values.title, files, baseVersion: recorded.version }
    } else {
      const current = await api(`/api/artifacts/${id}`, { allowNotFound: true })
      if (current) {
        console.error(`artifact: no recorded version for ${target}; publishing a full snapshot on top of v${current.currentVersion}.`)
        const files = encode(Object.keys(paths))
        for (const stale of Object.keys(current.files)) if (!(stale in files)) files[stale] = null
        body = { ...common, title: values.title, files, baseVersion: current.currentVersion }
      }
    }
    // A path that no longer exists on the server falls through to a fresh publish.
    if (body) {
      if (values.force) body.force = true
      result = await api(`/api/artifacts/${id}`, { method: 'PUT', auth: true, json: body, allowNotFound: true, allowConflict: true })
      if (result?.conflict) reportConflict(result.conflict, knownId, recorded?.version ?? body.baseVersion)
    }
  }
  if (!result) {
    const title = values.title ?? guessTitle(paths, values.entry) ?? basename(source)
    result = await api('/api/artifacts', { method: 'POST', auth: true, json: { ...common, title, files: encode(Object.keys(paths)) } })
  }
  const recordedNow = { id: result.id, url: result.url, version: result.currentVersion, files: hashes }
  updateState((latest) => {
    latest[source] = recordedNow
  })
  print(result, values.json, (r) => `${r.url}  (v${r.currentVersion}${r.currentVersion > 1 ? ', same URL' : ''})`)
}

async function checkout(target, dirArg, values) {
  const id = /^https?:\/\//.test(target) ? /\/a\/([^/?#]+)/.exec(new URL(target).pathname)?.[1] : target
  if (!id) fail(`cannot find an artifact id in ${target}`)
  const enc = encodeURIComponent
  const detail = await api(`/api/artifacts/${enc(id)}`)
  const dir = resolve(dirArg ?? id)
  if (existsSync(dir) && readdirSync(dir).length > 0 && !values.force) fail(`${dirArg ?? id} is not empty; pass --force to write into it`)
  const hashes = {}
  for (const path of Object.keys(detail.files)) {
    const full = resolve(dir, ...path.split('/'))
    if (!full.startsWith(dir + sep)) fail(`refusing unsafe path ${path}`)
    const bytes = await api(`/api/artifacts/${enc(id)}/files/${path.split('/').map(enc).join('/')}?v=${detail.currentVersion}`, { raw: true })
    mkdirSync(dirname(full), { recursive: true })
    writeFileSync(full, bytes)
    // Files publish skips stay out of the record so a later publish never deletes them remotely.
    if (!path.split('/').some(isSkipped)) hashes[path] = sha256(bytes)
  }
  updateState((latest) => {
    latest[dir] = { id: detail.id, url: detail.url, version: detail.currentVersion, files: hashes }
  })
  print({ id: detail.id, version: detail.currentVersion, dir, files: Object.keys(detail.files) }, values.json, (r) => `checked out ${r.id} v${r.version} (${r.files.length} files) into ${r.dir}`)
}

function parseJsonArg(text, what) {
  try {
    return JSON.parse(text)
  } catch {
    fail(`${what} must be JSON`)
  }
}

async function main() {
  const { values, positionals } = parseArgs({
    allowPositionals: true,
    options: {
      title: { type: 'string' },
      description: { type: 'string' },
      icon: { type: 'string' },
      entry: { type: 'string' },
      label: { type: 'string' },
      cap: { type: 'string' },
      id: { type: 'string' },
      new: { type: 'boolean' },
      force: { type: 'boolean' },
      q: { type: 'string' },
      v: { type: 'string' },
      limit: { type: 'string' },
      yes: { type: 'boolean' },
      json: { type: 'boolean' },
      url: { type: 'string' },
      'token-stdin': { type: 'boolean' },
      help: { type: 'boolean', short: 'h' }
    }
  })
  const [command, ...args] = positionals
  const json = values.json === true
  const enc = encodeURIComponent

  switch (command) {
    case 'publish':
      if (!args[0]) fail('publish needs a file or directory')
      return publish(args[0], values)
    case 'checkout':
      if (!args[0]) fail('checkout needs an id or viewer URL')
      return checkout(args[0], args[1], values)
    case 'list': {
      const params = new URLSearchParams({ limit: '200' })
      if (values.q) params.set('q', values.q)
      const data = await api(`/api/artifacts?${params}`)
      return print(data, json, (d) =>
        d.artifacts.map((a) => `${a.id}  v${a.currentVersion}  ${a.updatedAt.slice(0, 16).replace('T', ' ')}  ${a.icon ?? ''} ${a.title}`).join('\n') || '(none)'
      )
    }
    case 'get': {
      const data = await api(`/api/artifacts/${enc(args[0] ?? fail('get needs an id'))}`)
      return print(data, json, (d) =>
        [`${d.icon ?? ''} ${d.title}`.trim(), d.url, `entry ${d.entry} · capabilities ${d.capabilities.join(', ') || 'none'}`, ...d.versions.map((v) => `  v${v.version}  ${v.createdAt.slice(0, 16).replace('T', ' ')}  ${v.label ?? ''}`)].join('\n')
      )
    }
    case 'read': {
      const id = args[0] ?? fail('read needs an id')
      const detail = await api(`/api/artifacts/${enc(id)}${values.v ? `?v=${enc(values.v)}` : ''}`)
      const path = args[1] ?? detail.entry
      const query = values.v ? `?v=${enc(values.v)}` : ''
      const bytes = await api(`/api/artifacts/${enc(id)}/files/${path.split('/').map(enc).join('/')}${query}`, { raw: true })
      process.stdout.write(bytes)
      return
    }
    case 'comments': {
      const data = await api(`/api/artifacts/${enc(args[0] ?? fail('comments needs an id'))}/comments`)
      return print(data, json, (d) =>
        d.comments.map((c) => `${c.parentId ? '    ↳ ' : ''}[${c.id}${c.resolved ? ' resolved' : ''}] v${c.version} ${c.author}: ${c.body}`).join('\n') || '(no comments)'
      )
    }
    case 'delete': {
      const id = args[0] ?? fail('delete needs an id')
      if (!values.yes) fail(`deleting ${id} is permanent; pass --yes`)
      await api(`/api/artifacts/${enc(id)}`, { method: 'DELETE', auth: true })
      updateState((state) => {
        for (const [path, entry] of Object.entries(state)) if (entry.id === id) delete state[path]
      })
      return console.log(`deleted ${id}`)
    }
    case 'db': {
      const [action, id, collection, docId, payload] = args
      if (!action || !id || !collection) fail('db <action> <id> <collection> …')
      const base = `/api/artifacts/${enc(id)}/db/${enc(collection)}`
      let data
      if (action === 'list') data = await api(`${base}?limit=${enc(values.limit ?? '100')}`, { auth: true })
      else if (action === 'query') data = await api(`${base}/_query`, { method: 'POST', auth: true, json: { where: parseJsonArg(docId ?? '{}', 'where') } })
      else if (!docId) fail(`db ${action} needs a document id`)
      else if (action === 'get') data = await api(`${base}/${enc(docId)}`, { auth: true })
      else if (action === 'delete') data = await api(`${base}/${enc(docId)}`, { method: 'DELETE', auth: true })
      else if (action === 'set' || action === 'update') {
        data = await api(`${base}/${enc(docId)}`, { method: action === 'set' ? 'PUT' : 'PATCH', auth: true, json: { data: parseJsonArg(payload ?? fail('missing json'), 'data') } })
      } else fail(`unknown db action ${action}`)
      return console.log(JSON.stringify(data, null, 2))
    }
    case 'asset': {
      const [action, id, file] = args
      if (action === 'list') {
        const data = await api(`/api/artifacts/${enc(id ?? fail('asset list needs an id'))}/assets`)
        return print(data, json, (d) => d.assets.map((a) => `${a.id}  ${a.size}B  ${a.source}  ${a.url}`).join('\n') || '(no assets)')
      }
      if (action === 'upload') {
        if (!id || !file) fail('asset upload <id> <file>')
        const data = await api(`/api/artifacts/${enc(id)}/assets?name=${enc(basename(file))}`, { method: 'POST', auth: true, body: readFileSync(file) })
        return print(data, json, (d) => d.url)
      }
      return fail('asset upload|list')
    }
    case 'config': {
      const file = readJson(CONFIG_FILE, {})
      if (values.url) file.url = values.url.replace(/\/+$/, '')
      if (values['token-stdin']) file.token = readFileSync(0, 'utf8').trim()
      writePrivateJson(CONFIG_FILE, file)
      return console.log(`url ${file.url ?? DEFAULT_URL} · token ${file.token ? 'saved' : 'not set'}`)
    }
    default:
      console.log(USAGE)
      process.exit(command && !values.help ? 1 : 0)
  }
}

await main()

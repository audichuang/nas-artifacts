#!/usr/bin/env node
// artifact — publish and manage pages on a nas-artifacts server. Zero dependencies; Node 20+.
import { execFileSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, readFileSync, readdirSync, renameSync, statSync, writeFileSync } from 'node:fs'
import { homedir, hostname } from 'node:os'
import { basename, dirname, join, posix, relative, resolve, sep } from 'node:path'
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
  read <id> [path] [--v N] [--out FILE]  print a published file (default: the entry); a binary one (an Office file, a picture) needs --out
  check <file|dir>       report what the page sandbox will block or cannot find (publish runs it too)
  errors <id> [--v N]    errors viewers' browsers hit on a version (default: the current one)
  diff <id> [--from N] [--to M]  what a version changed (default: current against the one before)
  export <id> [--v N] [--out FILE]  save a slide deck as .pptx (default ./<id>.pptx); a published .pptx comes back as it was
  comments <id>          print the viewer comments
      --open             only threads not yet resolved
      --wait [--timeout S]  block until new comments arrive (default 540s; exit 2 on timeout)
  comment reply <id> <commentId> <text> [--resolve] [--author NAME]
  comment resolve|reopen <id> <commentId>
  prune <id>|--all --keep N  drop all but the newest N versions (frees the files only they used)
  delete <id> --yes      permanently delete
  db list <id> <collection> [--limit N]
  db get|delete <id> <collection> <docId>
  db set|update <id> <collection> <docId> <json>
  db query <id> <collection> <json-where>
  asset upload <id> <file> | asset list <id>
  config [--url URL] [--token-stdin] [--author NAME] [--machine NAME]
                         save server URL / token / comment author / machine name to ${CONFIG_FILE}

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

async function api(path, { method = 'GET', json, body, auth = false, raw = false, headers = {}, allowNotFound = false, allowConflict = false, hints = {} } = {}) {
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
    if (Object.hasOwn(hints, payload?.error?.code)) fail(hints[payload.error.code])
    fail(`${method} ${path} → ${response.status} ${payload?.error?.code ?? ''}: ${payload?.error?.message ?? response.statusText}`)
  }
  return payload
}

const kb = (bytes) => (bytes < 1024 ? `${bytes} bytes` : `${(bytes / 1024).toFixed(bytes < 10240 ? 1 : 0)} KB`)

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
  if (!page || /\.(pptx|pptm|ppsx|ppsm|potx|potm|ppt|docx|docm|dotx|dotm|doc|xlsx|xlsm|xltx|xltm|xls)$/i.test(page)) return null
  const text = readFileSync(page, 'utf8')
  const fromHtml = /<title[^>]*>([^<]{1,200})<\/title>/i.exec(text)?.[1]
  const fromMd = /^#\s+(.{1,200})$/m.exec(text)?.[1]
  return (fromHtml ?? fromMd ?? '').trim() || null
}

// What the content sandbox allows (src/content-server.ts), repeated so the CLI stays one file.
const SCRIPT_HOSTS = ['cdnjs.cloudflare.com', 'cdn.jsdelivr.net', 'unpkg.com', 'cdn.tailwindcss.com', 'code.jquery.com']
const STYLE_HOSTS = [...SCRIPT_HOSTS, 'fonts.googleapis.com']
const FONT_HOSTS = [...SCRIPT_HOSTS, 'fonts.gstatic.com']
const MEDIA_TAGS = new Set(['img', 'video', 'audio', 'source', 'track', 'iframe', 'embed', 'object', 'input'])
const CODE_RULES = [
  [/\b(fetch\s*\(|new\s+XMLHttpRequest|new\s+WebSocket|new\s+EventSource|navigator\.sendBeacon)/, (m) => `${m} is blocked: pages have no network (connect-src 'none'); inline the data, publish it as a .js file, or use the db capability`],
  [/\b(localStorage|sessionStorage|document\.cookie|indexedDB)\b/, (m) => `${m} throws in the sandbox (opaque origin); keep state in the db capability`],
  [/\bwindow\.open\s*\(/, () => 'window.open does nothing: the sandbox has no popups; show the URL as text']
]

// An uploaded asset: <content>/a/<id>/assets/<assetId>/<name>. The CLI does not know the content host, so the path decides.
const ASSET_PATH = /^\/a\/([^/]+)\/assets\//

/** Heuristic, not a parser: names what the sandbox will block or what is missing, before anyone opens the page. */
function checkFiles(paths, id) {
  const warnings = []
  const warn = (file, message) => {
    const line = `${file}: ${message}`
    if (!warnings.includes(line)) warnings.push(line)
  }
  for (const [file, full] of Object.entries(paths)) {
    const isPage = /\.(html?|md)$/i.test(file)
    const isCss = /\.css$/i.test(file)
    const isJs = /\.m?js$/i.test(file) && !/\.min\.js$/i.test(file)
    if (!isPage && !isCss && !isJs) continue
    let text = readFileSync(full, 'utf8')
    let externalLinks = 0
    const checkRef = (tag, rel, ref) => {
      if (!ref || ref.startsWith('#') || /^(data|blob|mailto|tel|javascript|about):/i.test(ref) || /\$\{|\{\{/.test(ref)) return
      if (/^(https?:)?\/\//i.test(ref)) {
        let url
        try { url = new URL(ref, 'https://base.invalid') } catch { return }
        const host = url.host
        const assetOf = ASSET_PATH.exec(url.pathname)?.[1]
        if (tag === 'a' || tag === 'area') externalLinks++
        else if (assetOf && (MEDIA_TAGS.has(tag) || tag === 'css')) {
          if (id && assetOf !== id) warn(file, `${ref} is an asset of artifact ${assetOf}; a page loads only its own assets, so upload it to this one`)
        } else if (tag === 'script') {
          if (!SCRIPT_HOSTS.includes(host)) warn(file, `<script> from ${host} is blocked; use ${SCRIPT_HOSTS.join(', ')}, or publish the file`)
          else checkPinned(ref, url)
        } else if (tag === 'link' && /\bstylesheet\b/.test(rel)) {
          if (!STYLE_HOSTS.includes(host)) warn(file, `stylesheet from ${host} is blocked; use Google Fonts, the script CDNs, or publish the file`)
          else checkPinned(ref, url)
        } else if (MEDIA_TAGS.has(tag)) warn(file, `<${tag}> from ${host} is blocked; publish the file, use a data URI, or "artifact asset upload"`)
        else if (tag === 'css' && !FONT_HOSTS.includes(host)) warn(file, `url(${ref}) is blocked; only fonts from the CDNs load from other hosts`)
        return
      }
      if (ref.startsWith('/')) {
        warn(file, `${ref} is root-relative and resolves outside the artifact; make it relative to ${file}`)
        return
      }
      let target
      try { target = decodeURI(ref.replace(/[?#].*$/, '')) } catch { target = ref }
      const resolved = posix.normalize(posix.join(posix.dirname(file), target))
      if (resolved.startsWith('..')) warn(file, `${ref} points outside the published files`)
      else if (!(resolved in paths) && !(posix.join(resolved, 'index.html') in paths)) warn(file, `${ref} is not among the published files`)
    }
    const checkPinned = (ref, url) => {
      if (url.host !== 'unpkg.com' && !(url.host === 'cdn.jsdelivr.net' && url.pathname.startsWith('/npm/'))) return
      const spec = /^\/(?:npm\/)?((?:@[^/]+\/)?[^/@]+)(?:@([^/]+))?/.exec(url.pathname)
      if (spec && !/^\d+\.\d+\.\d+/.test(spec[2] ?? '')) warn(file, `${ref} has no exact version; pin one (${spec[1]}@1.2.3) so the page cannot change under you`)
    }
    let code = isJs ? text : ''
    if (isPage) {
      if (/\.md$/i.test(file)) {
        text = text.replace(/```[\s\S]*?```/g, '').replace(/`[^`\n]*`/g, '')
        for (const m of text.matchAll(/(!?)\[[^\]]*\]\(\s*<?([^)\s>]+)/g)) checkRef(m[1] ? 'img' : 'a', '', m[2])
      }
      text = text.replace(/<!--[\s\S]*?-->/g, '')
      for (const m of text.matchAll(/<([a-z][\w-]*)\b([^>]*)>/gi)) {
        const tag = m[1].toLowerCase()
        const attrs = m[2]
        const rel = (/\brel\s*=\s*["']?([^"'>]+)/i.exec(attrs)?.[1] ?? '').toLowerCase()
        for (const ref of attrs.matchAll(/(?:^|\s)(?:src|href)\s*=\s*["']([^"']*)["']/gi)) checkRef(tag, rel, ref[1].trim())
        if (/\btarget\s*=\s*["']?_blank/i.test(attrs)) warn(file, 'target="_blank" does nothing: the sandbox has no popups; show the URL as text')
        for (const handler of attrs.matchAll(/\son\w+\s*=\s*(?:"([^"]*)"|'([^']*)')/gi)) code += `\n${handler[1] ?? handler[2]}`
      }
      for (const m of text.matchAll(/<script\b([^>]*)>([\s\S]*?)<\/script>/gi)) if (!/\bsrc\s*=/i.test(m[1])) code += `\n${m[2]}`
    }
    if (isCss) for (const m of text.matchAll(/url\(\s*["']?([^"')]+)/gi)) checkRef('css', '', m[1].trim())
    for (const [pattern, message] of CODE_RULES) {
      const found = pattern.exec(code)?.[1] ?? pattern.exec(code)?.[0]
      if (found) warn(file, message(found.replace(/\s+/g, ' ')))
    }
    if (externalLinks > 0) warn(file, `${externalLinks} link(s) leave the artifact and will not open from the viewer; show those URLs as text`)
  }
  return warnings
}

function reportWarnings(warnings) {
  if (warnings.length === 0) return
  console.error(`artifact: ${warnings.length} problem(s) the page sandbox will hit (published anyway; "artifact check" rechecks):`)
  for (const line of warnings) console.error(`  ${line}`)
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
  const warnings = checkFiles(paths, knownId)
  reportWarnings(warnings)
  const hashes = hashFiles(paths)
  const capabilities = values.cap === undefined ? undefined : values.cap.split(',').map((cap) => cap.trim()).filter(Boolean)
  const common = {
    description: values.description,
    icon: values.icon,
    entry: values.entry,
    label: values.label,
    capabilities,
    source: publishSource(source)
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
    const title = values.title ?? guessTitle(paths, values.entry) ?? basename(source).replace(/\.(pptx|pptm|ppsx|ppsm|potx|potm|ppt|docx|docm|dotx|dotm|doc|xlsx|xlsm|xltx|xltm|xls)$/i, '')
    result = await api('/api/artifacts', { method: 'POST', auth: true, json: { ...common, title, files: encode(Object.keys(paths)) } })
  }
  const recordedNow = { id: result.id, url: result.url, version: result.currentVersion, files: hashes }
  updateState((latest) => {
    latest[source] = recordedNow
  })
  print(values.json ? { ...result, warnings } : result, values.json, (r) => `${r.url}  (v${r.currentVersion}${r.currentVersion > 1 ? ', same URL' : ''})`)
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

function formatComments(comments) {
  return comments.map((c) => `${c.parentId ? '    ↳ ' : ''}[${c.id}${c.resolved ? ' resolved' : ''}] v${c.version} ${c.author}: ${c.body}`).join('\n') || '(no comments)'
}

// A thread is open while its root is unresolved; replies follow their root.
function openThreads(comments) {
  const open = new Set(comments.filter((c) => !c.parentId && !c.resolved).map((c) => c.id))
  return comments.filter((c) => open.has(c.parentId ?? c.id))
}

// macOS has a readable computer name ("Audi's Mac mini"); elsewhere the short host name.
function machineName() {
  if (process.platform === 'darwin') {
    try {
      const name = execFileSync('scutil', ['--get', 'ComputerName'], { encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] }).trim()
      if (name) return name
    } catch {}
  }
  return hostname().split('.')[0]
}

// Shown in the gallery so the reader knows which machine and folder to continue from.
function publishSource(source) {
  const home = homedir()
  const path = source === home || source.startsWith(home + sep) ? `~${source.slice(home.length).split(sep).join('/')}` : source
  const host = process.env.ARTIFACTS_MACHINE || readJson(CONFIG_FILE, {}).machine || machineName()
  return { host: host.slice(0, 100), path: path.slice(0, 500) }
}

function commentAuthor(values) {
  return values.author || readJson(CONFIG_FILE, {}).author || `agent@${hostname().split('.')[0]}`
}

// Follows the viewer event stream; a (re)connect or a comments change re-reads the list, so nothing is missed.
async function waitForComments(id, timeoutSeconds) {
  const path = `/api/artifacts/${encodeURIComponent(id)}`
  const list = async () => (await api(`${path}/comments`)).comments
  const seen = new Set((await list()).map((c) => c.id))
  const deadline = Date.now() + timeoutSeconds * 1000
  let announced = false
  while (Date.now() < deadline) {
    const controller = new AbortController()
    const timer = setTimeout(() => controller.abort(), deadline - Date.now())
    let found = null
    try {
      const response = await fetch(`${settings().url}${path}/events`, { signal: controller.signal })
      if (!response.ok) fail(`GET ${path}/events → ${response.status}`)
      const decoder = new TextDecoder()
      let buffer = ''
      stream: for await (const chunk of response.body) {
        buffer += decoder.decode(chunk, { stream: true })
        let end
        while ((end = buffer.indexOf('\n\n')) !== -1) {
          const block = buffer.slice(0, end)
          buffer = buffer.slice(end + 2)
          if (/"kind":"deleted"/.test(block)) fail(`${id} was deleted`)
          if (!/^event: ready$/m.test(block) && !/"kind":"comments"/.test(block)) continue
          if (!announced) {
            announced = true
            console.error(`artifact: waiting up to ${timeoutSeconds}s for new comments on ${id}…`)
          }
          const fresh = (await list()).filter((c) => !seen.has(c.id))
          if (fresh.length > 0) {
            found = fresh
            break stream
          }
        }
      }
    } catch {
      if (controller.signal.aborted) break
      await new Promise((resolve) => setTimeout(resolve, 1000))
    } finally {
      clearTimeout(timer)
      controller.abort()
    }
    if (found) return found
  }
  console.error(`artifact: no new comments on ${id} within ${timeoutSeconds}s`)
  process.exit(2)
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
      open: { type: 'boolean' },
      wait: { type: 'boolean' },
      timeout: { type: 'string' },
      author: { type: 'string' },
      machine: { type: 'string' },
      all: { type: 'boolean' },
      keep: { type: 'string' },
      from: { type: 'string' },
      to: { type: 'string' },
      resolve: { type: 'boolean' },
      out: { type: 'string' },
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
    case 'check': {
      if (!args[0]) fail('check needs a file or directory')
      const source = resolve(args[0])
      if (!existsSync(source)) fail(`${args[0]} does not exist`)
      const warnings = checkFiles(collectFiles(source))
      print({ warnings }, json, (d) => d.warnings.join('\n') || 'no problems found')
      return process.exit(warnings.length ? 1 : 0)
    }
    case 'errors': {
      const id = args[0] ?? fail('errors needs an id')
      const query = values.v ? `?v=${enc(values.v)}` : ''
      const data = await api(`/api/artifacts/${enc(id)}/errors${query}`, { hints: { not_found: 'this server keeps no page errors yet; it needs a newer nas-artifacts image' } })
      return print(data, json, (d) =>
        d.errors.map((e) => `[${e.kind}] ${e.path} ×${e.count}  ${e.message}${e.source ? `  (${e.source})` : ''}`).join('\n') || '(no errors reported for this version)'
      )
    }
    case 'diff': {
      const id = args[0] ?? fail('diff needs an id')
      const params = new URLSearchParams()
      if (values.from) params.set('from', values.from)
      if (values.to) params.set('to', values.to)
      const data = await api(`/api/artifacts/${enc(id)}/diff?${params}`, { hints: { not_found: 'this server cannot diff versions yet; it needs a newer nas-artifacts image' } })
      const notes = { binary: '(not text)', too_large: '(too large to show)', rewritten: '(rewritten past showing)' }
      return print(data, json, (d) => {
        const head = `v${d.to}${d.from ? ` against v${d.from}` : ' (oldest kept)'}${d.label ? ` · ${d.label}` : ''}${d.source ? ` · ${d.source.host} ${d.source.path}` : ''}`
        const files = d.files.map((f) => {
          const lines = (f.lines ?? []).map((l) => (l.op === '…' ? `@@ ${l.skipped} unchanged lines` : `${l.op}${l.text}`))
          return [`${f.status} ${f.path} ${f.note ? notes[f.note] : ''}`.trimEnd(), ...lines].join('\n')
        })
        return [head, ...files].join('\n\n') + (d.files.length ? '' : '\n(no file changes)')
      })
    }
    case 'export': {
      const id = args[0] ?? fail('export needs an id')
      const query = values.v ? `?v=${enc(values.v)}` : ''
      const bytes = await api(`/api/artifacts/${enc(id)}/export.pptx${query}`, {
        raw: true,
        hints: { not_found: 'this server cannot export decks yet; it needs a newer nas-artifacts image', not_a_deck: 'only slide decks (*.slides.md, *.pptx) export to .pptx; save any other published file with: artifact read <id> --out FILE' }
      })
      const out = resolve(values.out ?? `${id}.pptx`)
      writeFileSync(out, bytes)
      return print({ path: out, bytes: bytes.length }, json, (d) => `saved ${d.path} (${kb(d.bytes)})`)
    }
    case 'checkout':
      if (!args[0]) fail('checkout needs an id or viewer URL')
      return checkout(args[0], args[1], values)
    case 'list': {
      const params = new URLSearchParams({ limit: '200' })
      if (values.q) params.set('q', values.q)
      const data = await api(`/api/artifacts?${params}`)
      return print(data, json, (d) =>
        d.artifacts.map((a) => `${a.id}  v${a.currentVersion}  ${a.updatedAt.slice(0, 16).replace('T', ' ')}  ${a.icon ?? ''} ${a.title}${a.source ? `  [${a.source.host} ${a.source.path}]` : ''}`).join('\n') || '(none)'
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
      if (values.out) {
        writeFileSync(resolve(values.out), bytes)
        return print({ path, out: resolve(values.out), bytes: bytes.length }, values.json, (r) => `saved ${r.path} to ${r.out} (${kb(r.bytes)})`)
      }
      // A zip or an image in a terminal or a model's context is noise: save it instead.
      if (bytes.subarray(0, 8192).includes(0)) fail(`${path} is a binary file; pass --out FILE to save it`)
      process.stdout.write(bytes)
      return
    }
    case 'comments': {
      const id = args[0] ?? fail('comments needs an id')
      if (values.wait) {
        const timeout = Number(values.timeout ?? 540)
        if (!(timeout > 0)) fail('--timeout is a number of seconds')
        const fresh = await waitForComments(id, timeout)
        return print({ comments: fresh }, json, (d) => formatComments(d.comments))
      }
      const data = await api(`/api/artifacts/${enc(id)}/comments`)
      const comments = values.open ? openThreads(data.comments) : data.comments
      return print({ comments }, json, (d) => formatComments(d.comments))
    }
    case 'comment': {
      const [action, id, commentId, text] = args
      if (!['reply', 'resolve', 'reopen'].includes(action) || !id || !commentId) fail('comment reply|resolve|reopen <id> <commentId> [text]')
      const base = `/api/artifacts/${enc(id)}/comments`
      const setResolved = (rootId, resolved) =>
        api(`${base}/${enc(rootId)}`, {
          method: 'PATCH',
          auth: true,
          json: { resolved },
          hints: { comment_not_found: `${rootId} is not a thread root on ${id}; pass the id of the thread's first comment (see "artifact comments ${id}")` }
        })
      if (action === 'reply') {
        if (!text) fail('comment reply needs the reply text')
        const created = await api(base, {
          method: 'POST',
          auth: true,
          json: { author: commentAuthor(values), body: text, parentId: commentId },
          hints: { comment_not_found: `no comment ${commentId} on ${id}` }
        })
        if (values.resolve) await setResolved(created.parentId, true)
        return print(created, json, (c) => `replied ${c.id} in thread ${c.parentId}${values.resolve ? ' and resolved it' : ''}`)
      }
      await setResolved(commentId, action === 'resolve')
      return console.log(`${action === 'resolve' ? 'resolved' : 'reopened'} ${commentId}`)
    }
    case 'prune': {
      const keep = Number(values.keep)
      if (!Number.isInteger(keep) || keep < 1) fail('prune needs --keep N (N ≥ 1 versions to keep)')
      if (!values.all && !args[0]) fail('prune <id> --keep N, or prune --all --keep N')
      let ids = [args[0]]
      if (values.all) {
        ids = []
        for (let offset = 0; ; offset += 200) {
          const page = await api(`/api/artifacts?limit=200&offset=${offset}`)
          ids.push(...page.artifacts.filter((a) => a.currentVersion > keep).map((a) => a.id))
          if (offset + 200 >= page.total) break
        }
      }
      const hints = { not_found: 'this server cannot prune yet; it needs a newer nas-artifacts image' }
      const results = []
      for (const id of ids) results.push({ id, ...(await api(`/api/artifacts/${enc(id)}/prune`, { method: 'POST', auth: true, json: { keep }, hints })) })
      return print({ results }, json, (d) =>
        d.results.map((r) => `${r.id}: ${r.removed.length ? `dropped v${r.removed.join(', v')}, freed ${kb(r.freedBytes)}` : 'nothing to drop'}`).join('\n') || 'nothing to drop'
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
      if (values.author) file.author = values.author
      if (values.machine) file.machine = values.machine
      writePrivateJson(CONFIG_FILE, file)
      return console.log(`url ${file.url ?? DEFAULT_URL} · token ${file.token ? 'saved' : 'not set'} · author ${file.author ?? '(agent@host)'} · machine ${file.machine ?? machineName()}`)
    }
    default:
      console.log(USAGE)
      process.exit(command && !values.help ? 1 : 0)
  }
}

await main()

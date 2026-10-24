const data = JSON.parse(document.getElementById('page-data').textContent)
let artifact = data.artifact
let version = data.version
const id = artifact.id

const frame = document.getElementById('frame')
const versionSelect = document.getElementById('version')
const titleEl = document.getElementById('title')
const fullscreenButton = document.getElementById('fullscreen')
const pinButton = document.getElementById('pin')
const toastEl = document.getElementById('toast')
const commentsPanel = document.getElementById('comments')
const toggleComments = document.getElementById('toggle-comments')
const commentCount = document.getElementById('comment-count')
const threadsEl = document.getElementById('threads')
const form = document.getElementById('comment-form')
const bodyInput = document.getElementById('comment-body')
const replyTarget = document.getElementById('reply-target')
const showResolved = document.getElementById('show-resolved')

const storage = {
  get(key) {
    try { return localStorage.getItem(key) } catch { return null }
  },
  set(key, value) {
    try { localStorage.setItem(key, value) } catch {}
  }
}

let toastTimer
function toast(message) {
  toastEl.textContent = message
  toastEl.hidden = false
  clearTimeout(toastTimer)
  toastTimer = setTimeout(() => { toastEl.hidden = true }, 2600)
}

async function api(path, options = {}) {
  const headers = { 'x-artifact-client': 'viewer', ...(options.headers || {}) }
  if (options.json !== undefined) headers['content-type'] = 'application/json'
  const response = await fetch(`/api/artifacts/${id}${path}`, {
    method: options.method || 'GET',
    headers,
    body: options.json !== undefined ? JSON.stringify(options.json) : options.body
  })
  if (response.status === 204) return null
  const payload = await response.json().catch(() => null)
  if (!response.ok) {
    const error = new Error(payload?.error?.message || `HTTP ${response.status}`)
    error.code = payload?.error?.code
    error.status = response.status
    throw error
  }
  return payload
}

// Versions

function contentUrl(v, path) {
  const base = `${data.contentOrigin}/a/${id}/v/${v}/`
  return path ? base + path.split('/').map(encodeURIComponent).join('/') : base
}

function renderVersions() {
  versionSelect.replaceChildren(
    ...artifact.versions.map((row) => {
      const option = document.createElement('option')
      option.value = String(row.version)
      const latest = row.version === artifact.currentVersion ? ' · 最新' : ''
      option.textContent = `v${row.version}${row.label ? ` · ${row.label}` : ''}${latest}`
      return option
    })
  )
  versionSelect.value = String(version)
}

function show(v, path) {
  version = v
  const row = artifact.versions.find((candidate) => candidate.version === v)
  titleEl.textContent = row?.title || artifact.title
  document.title = titleEl.textContent
  frame.src = contentUrl(v, path)
  versionSelect.value = String(v)
  const url = new URL(location.href)
  if (v === artifact.currentVersion) url.searchParams.delete('v')
  else url.searchParams.set('v', String(v))
  if (path) url.searchParams.set('path', path)
  else url.searchParams.delete('path')
  history.replaceState(null, '', url)
}

fullscreenButton.addEventListener('click', () => {
  frame.requestFullscreen?.().catch(() => toast('無法進入全螢幕'))
})

versionSelect.addEventListener('change', () => show(Number(versionSelect.value)))

async function refreshArtifact() {
  artifact = await api('')
  renderVersions()
  pinButton.textContent = artifact.pinned ? '★' : '☆'
  pinButton.setAttribute('aria-pressed', String(artifact.pinned))
}

// Page bridge: the sandboxed page has an opaque origin, so it is recognised by its window, not its origin.

const viewerName = () => storage.get('artifacts.viewerName') || ''

function capability(name) {
  if (!artifact.capabilities.includes(name)) {
    throw Object.assign(new Error(`This artifact did not declare the ${name} capability.`), { code: 'capability_not_declared' })
  }
}

const methods = {
  hello: () => ({
    id,
    version,
    latestVersion: artifact.currentVersion,
    title: artifact.title,
    capabilities: artifact.capabilities,
    viewer: { name: viewerName() || null }
  }),
  'db.get': async ({ collection, id: docId }) => {
    capability('db')
    try {
      return await api(`/db/${encodeURIComponent(collection)}/${encodeURIComponent(docId)}`)
    } catch (error) {
      if (error.status === 404 && error.code === 'doc_not_found') return null
      throw error
    }
  },
  'db.set': ({ collection, id: docId, data: doc }) => {
    capability('db')
    return api(`/db/${encodeURIComponent(collection)}/${encodeURIComponent(docId)}`, { method: 'PUT', json: { data: doc } })
  },
  'db.update': ({ collection, id: docId, data: doc }) => {
    capability('db')
    return api(`/db/${encodeURIComponent(collection)}/${encodeURIComponent(docId)}`, { method: 'PATCH', json: { data: doc } })
  },
  'db.delete': async ({ collection, id: docId }) => {
    capability('db')
    return (await api(`/db/${encodeURIComponent(collection)}/${encodeURIComponent(docId)}`, { method: 'DELETE' })).deleted
  },
  'db.list': ({ collection, limit, after }) => {
    capability('db')
    const params = new URLSearchParams()
    if (limit) params.set('limit', String(limit))
    if (after) params.set('after', after)
    return api(`/db/${encodeURIComponent(collection)}?${params}`)
  },
  'db.query': ({ collection, where, limit, after }) => {
    capability('db')
    return api(`/db/${encodeURIComponent(collection)}/_query`, { method: 'POST', json: { where, limit, after } })
  },
  'db.batch': async ({ ops }) => {
    capability('db')
    return (await api('/db/_batch', { method: 'POST', json: { ops } })).results
  },
  'uploads.put': ({ name, bytes }) => {
    capability('uploads')
    return api(`/assets?name=${encodeURIComponent(name)}`, { method: 'POST', body: bytes })
  },
  'uploads.list': async () => {
    capability('uploads')
    return (await api('/assets')).assets.filter((asset) => asset.source === 'page')
  }
}

window.addEventListener('message', async (event) => {
  if (event.source !== frame.contentWindow) return
  const message = event.data
  if (!message || message.__artifact !== 1 || typeof message.method !== 'string') return
  const reply = (payload) => frame.contentWindow?.postMessage({ __artifact: 1, id: message.id, ...payload }, '*')
  const handler = Object.hasOwn(methods, message.method) ? methods[message.method] : null
  if (!handler) {
    reply({ error: { code: 'unknown_method', message: `Unknown method ${message.method}` } })
    return
  }
  try {
    reply({ result: await handler(message.params || {}) })
  } catch (error) {
    reply({ error: { code: error.code || 'error', message: error.message } })
  }
})

function notifyPage(event) {
  frame.contentWindow?.postMessage({ __artifact: 1, event }, '*')
}

// Comments

let comments = []
let replyTo = null

function setReply(comment) {
  replyTo = comment
  replyTarget.textContent = comment ? `回覆 ${comment.author}` : ''
  if (comment) bodyInput.focus()
}

function commentNode(comment) {
  const node = document.createElement('div')
  node.className = 'comment'
  const meta = document.createElement('div')
  meta.className = 'comment-meta'
  const author = document.createElement('span')
  author.className = 'comment-author'
  author.textContent = comment.author
  const when = document.createElement('time')
  when.dateTime = comment.createdAt
  when.textContent = new Date(comment.createdAt).toLocaleString('zh-Hant', { dateStyle: 'short', timeStyle: 'short' })
  const ver = document.createElement('span')
  ver.textContent = `v${comment.version}`
  meta.append(author, when, ver)
  const body = document.createElement('p')
  body.className = 'comment-body'
  body.textContent = comment.body
  node.append(meta, body)
  return node
}

function actionButton(label, onClick) {
  const button = document.createElement('button')
  button.type = 'button'
  button.className = 'link-button'
  button.textContent = label
  button.addEventListener('click', onClick)
  return button
}

function renderComments() {
  const roots = comments.filter((comment) => !comment.parentId)
  const open = roots.filter((root) => !root.resolved).length
  commentCount.textContent = open ? String(open) : ''
  const visible = roots.filter((root) => showResolved.checked || !root.resolved)
  threadsEl.replaceChildren(
    ...visible.map((root) => {
      const thread = document.createElement('li')
      thread.className = `thread${root.resolved ? ' resolved' : ''}`
      thread.append(commentNode(root))
      for (const reply of comments.filter((comment) => comment.parentId === root.id)) {
        thread.append(commentNode(reply))
      }
      const actions = document.createElement('div')
      actions.className = 'thread-actions'
      actions.append(
        actionButton('回覆', () => setReply(root)),
        actionButton(root.resolved ? '重新開啟' : '解決', async () => {
          await api(`/comments/${root.id}`, { method: 'PATCH', json: { resolved: !root.resolved } })
        }),
        actionButton('刪除', async () => {
          if (confirm('刪除這串留言?')) await api(`/comments/${root.id}`, { method: 'DELETE' })
        })
      )
      thread.append(actions)
      return thread
    })
  )
  if (visible.length === 0) {
    const empty = document.createElement('li')
    empty.className = 'muted'
    empty.textContent = '還沒有留言。'
    threadsEl.append(empty)
  }
}

async function loadComments() {
  comments = (await api('/comments')).comments
  renderComments()
}

showResolved.addEventListener('change', renderComments)

function setCommentsOpen(open) {
  commentsPanel.hidden = !open
  toggleComments.setAttribute('aria-expanded', String(open))
  storage.set('artifacts.commentsOpen', open ? '1' : '0')
}

toggleComments.addEventListener('click', () => setCommentsOpen(commentsPanel.hidden))
document.getElementById('close-comments').addEventListener('click', () => setCommentsOpen(false))

form.addEventListener('submit', async (event) => {
  event.preventDefault()
  const body = bodyInput.value.trim()
  if (!body) return
  let author = viewerName()
  if (!author) {
    author = (prompt('留言要用的名字(只存在這個瀏覽器):') || '').trim()
    if (!author) return
    storage.set('artifacts.viewerName', author)
  }
  try {
    await api('/comments', { method: 'POST', json: { author, body, parentId: replyTo?.id, version } })
    bodyInput.value = ''
    setReply(null)
  } catch (error) {
    toast(`留言失敗:${error.message}`)
  }
})
bodyInput.addEventListener('keydown', (event) => {
  if (event.key === 'Enter' && (event.metaKey || event.ctrlKey)) form.requestSubmit()
  if (event.key === 'Escape') setReply(null)
})

// Toolbar

pinButton.addEventListener('click', async () => {
  try {
    const next = await api('', { method: 'PATCH', json: { pinned: !artifact.pinned } })
    artifact = { ...artifact, pinned: next.pinned }
    pinButton.textContent = next.pinned ? '★' : '☆'
    pinButton.setAttribute('aria-pressed', String(next.pinned))
  } catch (error) {
    toast(error.message)
  }
})

document.getElementById('copy').addEventListener('click', async () => {
  try {
    await navigator.clipboard.writeText(location.href)
    toast('已複製連結')
  } catch {
    toast('無法複製連結')
  }
})

document.getElementById('menu-delete').addEventListener('click', async () => {
  if (!confirm(`永久刪除「${artifact.title}」與它所有的版本、資料與留言?`)) return
  let token = storage.get('artifacts.publisherToken')
  if (!token) {
    token = (prompt('刪除需要發佈 token(只存在這個瀏覽器):') || '').trim()
    if (!token) return
  }
  try {
    await api('', { method: 'DELETE', headers: { authorization: `Bearer ${token}` } })
    storage.set('artifacts.publisherToken', token)
    location.href = '/'
  } catch (error) {
    if (error.status === 401) storage.set('artifacts.publisherToken', '')
    toast(`刪除失敗:${error.message}`)
  }
})

// Live changes

let lastSync = 0

// Events emitted while the stream was down are lost, so a reconnect or a wake-up re-reads everything.
async function resync() {
  if (Date.now() - lastSync < 2000) return
  lastSync = Date.now()
  const before = artifact.currentVersion
  const wasLatest = version === before
  try {
    await refreshArtifact()
    await loadComments()
  } catch {
    lastSync = 0
    return
  }
  if (artifact.currentVersion !== before && wasLatest) {
    show(artifact.currentVersion)
    toast(`已更新到 v${artifact.currentVersion}`)
  } else {
    // A retained frame keeps its stale db data unless told; a reloaded one re-reads on load.
    notifyPage({ type: 'db', resync: true })
  }
}

document.addEventListener('visibilitychange', () => {
  if (document.visibilityState === 'visible') resync()
})

function connectEvents() {
  const source = new EventSource(`/api/artifacts/${id}/events`)
  let connected = false
  source.addEventListener('ready', () => {
    if (connected) resync()
    else lastSync = Date.now()
    connected = true
  })
  source.addEventListener('change', async (event) => {
    const change = JSON.parse(event.data)
    if (change.kind === 'version') {
      const wasLatest = version === artifact.currentVersion
      await refreshArtifact()
      if (wasLatest) {
        show(change.version)
        toast(`已更新到 v${change.version}`)
      } else {
        toast(`有新版本 v${change.version}`)
      }
    } else if (change.kind === 'deleted') {
      toast('這個 artifact 已被刪除')
      setTimeout(() => { location.href = '/' }, 1500)
    } else if (change.kind === 'comments') {
      loadComments()
    } else if (change.kind === 'db') {
      notifyPage({ type: 'db', collection: change.collection, id: change.docId, deleted: change.deleted })
    } else if (change.kind === 'assets') {
      notifyPage({ type: 'uploads' })
    }
  })
}

renderVersions()
pinButton.textContent = artifact.pinned ? '★' : '☆'
pinButton.setAttribute('aria-pressed', String(artifact.pinned))
if (storage.get('artifacts.commentsOpen') === '1') {
  commentsPanel.hidden = false
  toggleComments.setAttribute('aria-expanded', 'true')
}
show(version, data.path)
loadComments().catch(() => {})
connectEvents()

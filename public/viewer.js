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
const toastText = document.getElementById('toast-text')
const toastAction = document.getElementById('toast-action')
const changesButton = document.getElementById('changes')
const diffDialog = document.getElementById('diff-dialog')
const commentsPanel = document.getElementById('comments')
const toggleComments = document.getElementById('toggle-comments')
const commentCount = document.getElementById('comment-count')
const threadsEl = document.getElementById('threads')
const form = document.getElementById('comment-form')
const bodyInput = document.getElementById('comment-body')
const replyTarget = document.getElementById('reply-target')
const showResolved = document.getElementById('show-resolved')
const errorsButton = document.getElementById('errors')
const errorCount = document.getElementById('error-count')
const errorDialog = document.getElementById('error-dialog')
const errorList = document.getElementById('error-list')

const storage = {
  get(key) {
    try { return localStorage.getItem(key) } catch { return null }
  },
  set(key, value) {
    try { localStorage.setItem(key, value) } catch {}
  }
}

let toastTimer
let toastRun = null
function toast(message, action) {
  toastText.textContent = message
  toastAction.hidden = !action
  toastAction.textContent = action?.label ?? ''
  toastRun = action?.run ?? null
  toastEl.hidden = false
  clearTimeout(toastTimer)
  toastTimer = setTimeout(() => { toastEl.hidden = true }, action ? 6000 : 2600)
}
toastAction.addEventListener('click', () => {
  toastEl.hidden = true
  toastRun?.()
})

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

// The file, hash and scroll position the page last reported; path null is the entry.
let place = { path: null, hash: '', scrollY: 0 }
let pendingScroll = null

// A hash the page or the address bar names: an anchor or a slide number, nothing that could be a URL.
function validHash(hash) {
  if (typeof hash !== 'string' || hash.length > 400) return ''
  try {
    const decoded = decodeURIComponent(hash)
    return /^#[\p{L}\p{N}_-]{1,100}$/u.test(decoded) ? decoded : ''
  } catch {
    return ''
  }
}

function syncUrl() {
  const url = new URL(location.href)
  if (version === artifact.currentVersion) url.searchParams.delete('v')
  else url.searchParams.set('v', String(version))
  if (place.path) url.searchParams.set('path', place.path)
  else url.searchParams.delete('path')
  url.hash = place.hash
  // Safari throws once a page replaces its URL too often; the address bar is a convenience.
  try { history.replaceState(null, '', url) } catch {}
}

function show(v, path, hash = '') {
  version = v
  const row = artifact.versions.find((candidate) => candidate.version === v)
  titleEl.textContent = row?.title || artifact.title
  document.title = titleEl.textContent
  frame.src = contentUrl(v, path) + hash
  versionSelect.value = String(v)
  place = { path: path ?? null, hash, scrollY: 0 }
  syncUrl()
  loadErrors()
  syncChangesButton()
  syncEditButton()
}

// Shown when a kept version older than the one on screen exists to compare against.
function syncChangesButton() {
  changesButton.hidden = !artifact.versions.some((candidate) => candidate.version < version)
}

// Opens version v on the file the reader is on, scrolled where they were, if that file exists there.
async function showKeepingPlace(v) {
  const { path, hash, scrollY } = place
  let files = v === artifact.currentVersion ? artifact.files : undefined
  if (path && !files) files = (await api(`?v=${v}`).catch(() => null))?.files
  const keep = !path || (files ? Object.hasOwn(files, path) : false)
  show(v, keep ? path ?? undefined : undefined, keep ? hash : '')
  pendingScroll = keep ? scrollY : null
}

frame.addEventListener('load', () => {
  if (pendingScroll === null) return
  notifyPage({ type: 'restore', scrollY: pendingScroll })
  pendingScroll = null
})

const currentEntry = () => artifact.versions.find((row) => row.version === version)?.entry ?? artifact.entry

// The file a page URL in the frame names, for the version on show; undefined for anything else.
function framePath(pathname) {
  if (typeof pathname !== 'string' || pathname.length > 4096) return undefined
  const match = /^\/a\/([^/]+)\/v\/(\d+)\/(.*)$/.exec(pathname)
  if (!match || match[1] !== id || Number(match[2]) !== version) return undefined
  try {
    const path = match[3].split('/').map(decodeURIComponent).join('/')
    return path === '' ? currentEntry() : path
  } catch {
    return undefined
  }
}

// The page is untrusted: its report only picks which of this version's files the address bar names.
function onPlaceReport(report) {
  const path = framePath(report.path)
  if (path === undefined || !Number.isFinite(report.scrollY)) return
  const next = path === currentEntry() ? null : path
  const hash = validHash(report.hash)
  const moved = next !== place.path || hash !== place.hash
  place = { path: next, hash, scrollY: Math.max(0, report.scrollY) }
  if (moved) syncUrl()
}

// Slide decks: arrow keys pressed outside the frame still turn the page.
const isDeck = () => /\.slides\.md$/i.test(place.path ?? currentEntry())
const DECK_KEYS = new Set(['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'PageUp', 'PageDown', 'Home', 'End'])
document.addEventListener('keydown', (event) => {
  if (!isDeck() || !DECK_KEYS.has(event.key) || event.metaKey || event.ctrlKey || event.altKey) return
  if (event.target.closest('input, textarea, select, [contenteditable], dialog') || document.querySelector('dialog[open]')) return
  event.preventDefault()
  notifyPage({ type: 'deck-key', key: event.key })
})

fullscreenButton.addEventListener('click', () => {
  frame.requestFullscreen?.().catch(() => toast('無法進入全螢幕'))
})

versionSelect.addEventListener('change', () => showKeepingPlace(Number(versionSelect.value)))

function versionLabel(v) {
  const label = artifact.versions.find((row) => row.version === v)?.label
  return `v${v}${label ? ` · ${label}` : ''}`
}

async function refreshArtifact() {
  artifact = await api('')
  renderVersions()
  syncChangesButton()
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
  if (!message || message.__artifact !== 1) return
  if (message.report?.type === 'place') {
    onPlaceReport(message.report)
    return
  }
  if (message.report?.type === 'errors') {
    onErrorReport(message.report)
    return
  }
  if (typeof message.method !== 'string') return
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

// What this version changed against the one before it.

const STATUS = { added: '新增', removed: '刪除', modified: '修改' }
const NOTES = { binary: '非文字檔,只列出有變更', too_large: '檔案太大,只列出有變更', rewritten: '改動太多,幾乎整份重寫' }

async function openDiff() {
  const shown = version
  let diff
  try {
    diff = await api(`/diff?to=${shown}`)
  } catch (error) {
    toast(`讀不到變更:${error.message}`)
    return
  }
  document.getElementById('diff-title').textContent = diff.from ? `v${diff.to} 相較 v${diff.from}` : `v${diff.to}(第一個保留的版本)`
  const meta = [diff.label, diff.source && `${diff.source.host} · ${diff.source.path}`, new Date(diff.createdAt).toLocaleString('zh-Hant', { dateStyle: 'short', timeStyle: 'short' })]
  document.getElementById('diff-meta').textContent = `${meta.filter(Boolean).join(' · ')} · ${diff.files.length} 個檔案有變更`
  document.getElementById('diff-files').replaceChildren(
    ...diff.files.map((file) => {
      const section = document.createElement('section')
      section.className = 'diff-file'
      const head = document.createElement('div')
      head.className = 'diff-file-head'
      const status = document.createElement('span')
      status.className = `diff-status diff-${file.status}`
      status.textContent = STATUS[file.status]
      const path = document.createElement('span')
      path.className = 'diff-path'
      path.textContent = file.path
      head.append(status, path)
      section.append(head)
      if (file.note) {
        const note = document.createElement('p')
        note.className = 'muted small'
        note.textContent = NOTES[file.note]
        section.append(note)
      } else if (file.lines) {
        const pre = document.createElement('pre')
        pre.className = 'diff-lines'
        for (const line of file.lines) {
          const row = document.createElement('span')
          row.className = line.op === '+' ? 'add' : line.op === '-' ? 'del' : line.op === '…' ? 'gap' : 'same'
          row.textContent = line.op === '…' ? `⋯ ${line.skipped} 行未變` : `${line.op} ${line.text}`
          pre.append(row)
        }
        section.append(pre)
      }
      return section
    })
  )
  diffDialog.showModal()
}

changesButton.addEventListener('click', openDiff)
document.getElementById('close-diff').addEventListener('click', () => diffDialog.close())

// Page errors: the page reports what broke, the server keeps a bounded log per version.

let pageErrors = []
const ERROR_KINDS = { error: '腳本錯誤', rejection: '未處理的 Promise', resource: '載入失敗', csp: '被沙箱擋下' }

// "<content>/a/<id>/v/<n>/x.js:3:9" reads as "x.js:3:9"; other sources stay as they are.
function shortSource(source) {
  if (!source.startsWith(`${data.contentOrigin}/`)) return source
  const match = /^([^:]*?)((?::\d+)*)$/.exec(source.slice(data.contentOrigin.length))
  const path = match && framePath(match[1].replace(/[?#].*$/, ''))
  return path === undefined || path === null ? source : `${path}${match[2]}`
}

function onErrorReport(report) {
  const path = framePath(report.path)
  if (path === undefined || !Array.isArray(report.errors)) return
  const errors = report.errors.slice(0, 20).map((error) => ({
    kind: String(error?.kind),
    message: String(error?.message ?? ''),
    source: typeof error?.source === 'string' ? shortSource(error.source) : null
  }))
  api('/errors', { method: 'POST', json: { version, path, errors } }).catch(() => {})
}

async function loadErrors() {
  const shown = version
  if (pageErrors.some((error) => error.version !== shown)) {
    pageErrors = []
    renderErrors()
  }
  let errors
  try {
    errors = (await api(`/errors?v=${shown}`)).errors
  } catch {
    return
  }
  if (shown !== version) return
  pageErrors = errors
  renderErrors()
}

function renderErrors() {
  errorsButton.hidden = pageErrors.length === 0
  errorCount.textContent = String(pageErrors.length)
  errorList.replaceChildren(
    ...pageErrors.map((error) => {
      const item = document.createElement('li')
      item.className = 'thread'
      const meta = document.createElement('div')
      meta.className = 'comment-meta'
      for (const text of [ERROR_KINDS[error.kind] ?? error.kind, error.path, `×${error.count}`]) {
        const span = document.createElement('span')
        span.textContent = text
        meta.append(span)
      }
      const message = document.createElement('p')
      message.className = 'comment-body'
      message.textContent = error.message
      item.append(meta, message)
      if (error.source) {
        const source = document.createElement('p')
        source.className = 'error-source'
        source.textContent = error.source
        item.append(source)
      }
      return item
    })
  )
}

errorsButton.addEventListener('click', () => errorDialog.showModal())
document.getElementById('close-errors').addEventListener('click', () => errorDialog.close())

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

toggleComments.addEventListener('click', () => {
  if (commentsPanel.hidden && !editorPanel.hidden) {
    if (dirty) return toast('先儲存或關閉編輯,再開留言')
    closeEditor()
  }
  setCommentsOpen(commentsPanel.hidden)
})
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

// Deck editor: the source on the left; every pause in typing renders the slides on the server and swaps
// them into the frame, so what is on screen is the deck as it would publish. Saving makes a new version.

const editorPanel = document.getElementById('editor')
const editButton = document.getElementById('edit')
const sourceInput = document.getElementById('editor-source')
const editorStatus = document.getElementById('editor-status')
const saveButton = document.getElementById('save-deck')
const viewerEl = document.querySelector('.viewer')
// The version the draft started from, the version this viewer saved last, and whether the draft differs.
let editBase = null
let savedVersion = null
let dirty = false
let previewSeq = 0
let previewTimer

const canEdit = () => isDeck() && artifact.capabilities.length === 0
const exportLink = document.getElementById('export-pptx')
function syncEditButton() {
  editButton.hidden = !canEdit()
  exportLink.hidden = !isDeck()
  exportLink.href = `/api/artifacts/${id}/export.pptx?v=${version}`
}

function setEditorStatus(text, isDirty = dirty) {
  editorStatus.textContent = text
  editorStatus.classList.toggle('dirty', isDirty)
  saveButton.disabled = !isDirty
}

// Which slide the cursor is on, splitting the way the server does: --- lines outside fenced code.
function slideIndexAt(text, offset) {
  const chunks = [{ start: 0, text: '' }]
  let fence = null
  let at = 0
  for (const line of text.split('\n')) {
    const marker = /^ {0,3}(`{3,}|~{3,})/.exec(line)?.[1]
    if (fence) {
      if (marker && marker[0] === fence[0] && marker.length >= fence.length && line.trim() === marker) fence = null
    } else if (marker) {
      fence = marker
    } else if (/^---\s*$/.test(line)) {
      chunks.push({ start: at + line.length + 1, text: '' })
      at += line.length + 1
      continue
    }
    chunks[chunks.length - 1].text += `${line}\n`
    at += line.length + 1
  }
  const kept = chunks.filter((chunk) => chunk.text.trim())
  let index = 0
  kept.forEach((chunk, i) => { if (chunk.start <= offset) index = i })
  return index
}

async function loadSource() {
  const response = await fetch(`/api/artifacts/${id}/files/${artifact.entry.split('/').map(encodeURIComponent).join('/')}?v=${artifact.currentVersion}`)
  if (!response.ok) throw new Error(`HTTP ${response.status}`)
  return response.text()
}

async function openEditor() {
  if (version !== artifact.currentVersion) {
    await showKeepingPlace(artifact.currentVersion)
  }
  let text
  try {
    text = await loadSource()
  } catch (error) {
    toast(`讀不到簡報原始碼:${error.message}`)
    return
  }
  setCommentsOpen(false)
  editBase = artifact.currentVersion
  dirty = false
  sourceInput.value = text
  editorPanel.hidden = false
  viewerEl.classList.add('editing')
  editButton.setAttribute('aria-expanded', 'true')
  setEditorStatus(`根據 v${editBase}`)
  sourceInput.focus()
  sourceInput.setSelectionRange(0, 0)
}

function closeEditor() {
  if (dirty && !confirm('放棄還沒儲存的修改?')) return
  editorPanel.hidden = true
  viewerEl.classList.remove('editing')
  editButton.setAttribute('aria-expanded', 'false')
  clearTimeout(previewTimer)
  previewSeq++
  // The frame may hold a preview of the discarded draft: show the saved deck again.
  if (dirty) show(version, place.path ?? undefined, place.hash)
  dirty = false
  editBase = null
}

async function preview() {
  const seq = ++previewSeq
  let result
  try {
    result = await api('/slides/preview', { method: 'POST', json: { markdown: sourceInput.value } })
  } catch (error) {
    if (seq === previewSeq) setEditorStatus(`預覽失敗:${error.message}`)
    return
  }
  if (seq !== previewSeq || editorPanel.hidden) return
  notifyPage({ type: 'deck-preview', html: result.html, index: slideIndexAt(sourceInput.value, sourceInput.selectionStart) })
}

sourceInput.addEventListener('input', () => {
  dirty = true
  setEditorStatus(`根據 v${editBase} · 未儲存`)
  clearTimeout(previewTimer)
  previewTimer = setTimeout(preview, 250)
})

// Moving the cursor to another slide's source shows that slide.
let cursorSlide = -1
function followCursor() {
  const index = slideIndexAt(sourceInput.value, sourceInput.selectionStart)
  if (index === cursorSlide) return
  cursorSlide = index
  notifyPage({ type: 'deck-go', index })
}
sourceInput.addEventListener('keyup', followCursor)
sourceInput.addEventListener('click', followCursor)

async function saveDeck() {
  if (!dirty || saveButton.disabled) return
  saveButton.disabled = true
  setEditorStatus('儲存中…', false)
  try {
    const saved = await api('/slides', { method: 'PUT', json: { markdown: sourceInput.value, baseVersion: editBase, author: viewerName() || undefined } })
    savedVersion = saved.currentVersion
    editBase = saved.currentVersion
    dirty = false
    setEditorStatus(`已存成 v${saved.currentVersion}`)
  } catch (error) {
    dirty = true
    if (error.code === 'version_conflict') {
      setEditorStatus('別人已存了新版本', true)
      if (confirm('在你編輯時,別人已存了新版本。\n\n確定:放棄你的修改,載入最新版。\n取消:保留你的草稿(可先複製起來)。')) {
        dirty = false
        closeEditor()
        await refreshArtifact()
        await showKeepingPlace(artifact.currentVersion)
        await openEditor()
      }
    } else {
      setEditorStatus(`儲存失敗:${error.message}`, true)
    }
  }
}

editButton.addEventListener('click', () => (editorPanel.hidden ? openEditor() : closeEditor()))
document.getElementById('close-editor').addEventListener('click', closeEditor)
saveButton.addEventListener('click', saveDeck)
sourceInput.addEventListener('keydown', (event) => {
  if (event.key === 's' && (event.metaKey || event.ctrlKey)) {
    event.preventDefault()
    saveDeck()
  }
})
addEventListener('beforeunload', (event) => {
  if (dirty) event.preventDefault()
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
    loadErrors()
  } catch {
    lastSync = 0
    return
  }
  if (artifact.currentVersion !== before && wasLatest && editorPanel.hidden) {
    await showKeepingPlace(artifact.currentVersion)
    toast(`已更新到 ${versionLabel(artifact.currentVersion)}`)
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
      if (change.version === savedVersion) {
        // This viewer's own save: show it quietly, on the same slide.
        await showKeepingPlace(change.version)
      } else if (!editorPanel.hidden) {
        if (dirty) {
          // Keep the draft and its preview; saving it will report the conflict.
          toast(`有人存了新版本 ${versionLabel(change.version)},你的修改還沒儲存`)
        } else {
          await showKeepingPlace(change.version)
          editBase = change.version
          sourceInput.value = await loadSource().catch(() => sourceInput.value)
          setEditorStatus(`根據 v${editBase}`)
          toast(`已載入別人存的 ${versionLabel(change.version)}`)
        }
      } else if (wasLatest) {
        await showKeepingPlace(change.version)
        toast(`已更新到 ${versionLabel(change.version)}`, { label: '看變更', run: openDiff })
      } else {
        toast(`有新版本 ${versionLabel(change.version)}`)
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
    } else if (change.kind === 'errors' && change.version === version) {
      loadErrors()
    } else if (change.kind === 'pruned') {
      await refreshArtifact()
      if (change.versions.includes(version)) {
        await showKeepingPlace(artifact.currentVersion)
        toast(`這個版本已被清理,改看最新版 v${artifact.currentVersion}`)
      }
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
show(version, data.path, validHash(location.hash))
loadComments().catch(() => {})
connectEvents()

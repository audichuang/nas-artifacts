const data = JSON.parse(document.getElementById('page-data').textContent)
const root = `/api/artifacts/${data.artifactId}/office-edits/${data.sessionId}`
const statusEl = document.getElementById('office-status')
const saveButton = document.getElementById('office-save')
const closeButton = document.getElementById('office-close')
const download = document.getElementById('office-result')
const back = document.getElementById('office-back')
let editor
let closing = false
let changed = false
let stopped = false
let lastVersion
let saveRequested = false
closeButton.disabled = true

async function api(path = '', method = 'GET') {
  const response = await fetch(root + path, { method, headers: { 'x-artifact-client': 'viewer' } })
  const value = await response.json()
  if (!response.ok) throw new Error(value.error?.message || `HTTP ${response.status}`)
  return value
}

async function poll() {
  if (stopped) return
  try {
    const state = await api()
    if (state.downloadUrl) { download.href = state.downloadUrl; download.hidden = false }
    if (state.status === 'conflict' || state.status === 'error') {
      statusEl.textContent = state.message
      back.hidden = false
      saveButton.disabled = state.status === 'conflict' || state.closed
      closeButton.hidden = state.closed
      closing = false
      closeButton.disabled = false
      closeButton.textContent = '結束編輯'
    } else if (state.closed && (closing || state.status === 'saved')) {
      stopped = true
      changed = false
      saveButton.disabled = true
      back.hidden = false
      statusEl.textContent = state.savedVersion ? `已儲存為 v${state.savedVersion}` : '編輯已結束'
      if (closing) location.assign(`/a/${data.artifactId}${state.savedVersion ? `?v=${state.savedVersion}` : ''}`)
    } else if (state.closed) {
      statusEl.textContent = 'Office 連線中斷，等待重新連線…'
      saveButton.disabled = true
    } else if (state.status === 'editing' && !closing && !saveRequested) {
      saveButton.disabled = false
      statusEl.textContent = changed ? '編輯中' : '可直接點選文字與拖動物件'
    } else if (state.savedVersion && state.savedVersion !== lastVersion) {
      lastVersion = state.savedVersion
      statusEl.textContent = `已儲存為 v${state.savedVersion}`
      saveButton.disabled = false
    }
  } catch (error) { statusEl.textContent = error.message }
  if (!stopped) setTimeout(poll, 2000)
}

function finish() {
  if (closing || !editor) return
  closing = true
  saveRequested = false
  closeButton.disabled = true
  saveButton.disabled = true
  statusEl.textContent = '正在結束編輯並保存，請稍候…'
  // Wait for local edits to reach Office before requesting its close handshake.
  if (!changed) editor.requestClose()
}
function closeReady() {
  closing = true
  closeButton.disabled = true
  saveButton.disabled = true
  statusEl.textContent = '正在保存並返回，請稍候…'
  editor.destroyEditor()
}
async function saveVersion() {
  saveButton.disabled = true
  if (changed) { saveRequested = true; statusEl.textContent = '正在同步修改…'; return }
  try {
    const { queued } = await api('/save', 'POST')
    statusEl.textContent = queued ? '正在儲存版本…' : '目前沒有待儲存的修改'
  } catch (error) { statusEl.textContent = error.message }
  finally { if (!closing) saveButton.disabled = false }
}
closeButton.addEventListener('click', finish)
saveButton.addEventListener('click', saveVersion)
addEventListener('beforeunload', (event) => { if (changed && !stopped) event.preventDefault() })

try {
  const config = await api('/config')
  const script = document.createElement('script')
  script.src = `${data.officeUrl}/web-apps/apps/api/documents/api.js`
  const loaded = new Promise((resolve, reject) => {
    script.onload = resolve
    script.onerror = () => reject(new Error('無法連到 Office 編輯服務，請稍後重新開啟。'))
  })
  document.head.append(script)
  await loaded
  config.events = {
    onDocumentReady() { statusEl.textContent = '可直接點選文字與拖動物件'; saveButton.disabled = false; closeButton.disabled = false },
    onDocumentStateChange(event) {
      changed = Boolean(event.data)
      if (closing) { if (!changed) editor.requestClose() }
      else if (saveRequested) { if (!changed) { saveRequested = false; saveVersion() } }
      else if (changed) { statusEl.textContent = '編輯中'; saveButton.disabled = false }
    },
    onRequestClose: closeReady,
    onError() { statusEl.textContent = 'Office 編輯器發生錯誤，請先取回編輯結果。'; back.hidden = false }
  }
  editor = new window.DocsAPI.DocEditor('office-editor', config)
  poll()
} catch (error) {
  statusEl.textContent = error.message
  back.hidden = false
  closeButton.hidden = true
}

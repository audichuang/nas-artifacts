const data = JSON.parse(document.getElementById('page-data').textContent)
const root = `/api/artifacts/${data.artifactId}/office-edits/${data.sessionId}`
const statusEl = document.getElementById('office-status')
const closeButton = document.getElementById('office-close')
const download = document.getElementById('office-result')
const back = document.getElementById('office-back')
let editor
let closing = false
let changed = false
let stopped = false
let lastVersion
let closingSince = 0
// Office normally confirms a close within seconds; past this its callback is likely not reaching the app.
const CLOSE_WAIT_MS = 60_000
closeButton.disabled = true

async function api(path = '') {
  const response = await fetch(root + path, { headers: { 'x-artifact-client': 'viewer' } })
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
      closeButton.hidden = state.closed
      closing = false
      closeButton.disabled = false
      closeButton.textContent = '結束編輯'
    } else if (state.closed && (closing || state.status === 'saved')) {
      stopped = true
      changed = false
      back.hidden = false
      statusEl.textContent = state.savedVersion ? `已儲存為 v${state.savedVersion}` : '編輯已結束'
      if (closing) location.assign(`/a/${data.artifactId}${state.savedVersion ? `?v=${state.savedVersion}` : ''}`)
    } else if (closing && Date.now() - closingSince > CLOSE_WAIT_MS) {
      statusEl.textContent = 'Office 尚未確認保存結果，可先返回檢視；保存完成後會出現新版本。'
      back.hidden = false
    } else if (state.closed) {
      statusEl.textContent = 'Office 連線中斷，等待重新連線…'
    } else if (state.status === 'editing' && !closing) {
      statusEl.textContent = changed ? '編輯中，請使用 Office 工具列儲存' : '可直接點選文字與拖動物件'
    } else if (state.savedVersion && state.savedVersion !== lastVersion) {
      lastVersion = state.savedVersion
      statusEl.textContent = `已儲存為 v${state.savedVersion}`
    }
  } catch (error) { statusEl.textContent = error.message }
  if (!stopped) setTimeout(poll, 2000)
}

function finish() {
  if (closing || !editor) return
  // Native Save and requestClose also work when the reader disables Office autosave.
  editor.requestClose()
}
function closeReady() {
  closing = true
  closingSince = Date.now()
  closeButton.disabled = true
  statusEl.textContent = '正在保存並返回，請稍候…'
  editor.destroyEditor()
}
closeButton.addEventListener('click', finish)
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
    onDocumentReady() { statusEl.textContent = '可直接點選文字與拖動物件'; closeButton.disabled = false },
    onDocumentStateChange(event) {
      changed = Boolean(event.data)
      if (changed) statusEl.textContent = '編輯中，請使用 Office 工具列儲存'
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

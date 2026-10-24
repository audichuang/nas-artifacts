const list = document.getElementById('list')
const status = document.getElementById('status')
const more = document.getElementById('more')
const search = document.getElementById('search')
const PAGE = 60
let offset = 0
let query = ''
let generation = 0

const timeFormat = new Intl.RelativeTimeFormat('zh-Hant', { numeric: 'auto' })
function relativeTime(iso) {
  const seconds = (new Date(iso).getTime() - Date.now()) / 1000
  const units = [[31536000, 'year'], [2592000, 'month'], [86400, 'day'], [3600, 'hour'], [60, 'minute']]
  const [size, unit] = units.find(([unitSeconds]) => Math.abs(seconds) >= unitSeconds) ?? [1, 'second']
  return timeFormat.format(Math.round(seconds / size), unit)
}

function card(artifact) {
  const item = document.createElement('li')
  const link = document.createElement('a')
  link.className = 'card'
  link.href = `/a/${artifact.id}`
  const head = document.createElement('div')
  head.className = 'card-head'
  const icon = document.createElement('span')
  icon.className = 'artifact-icon'
  icon.textContent = artifact.icon || '📄'
  const title = document.createElement('span')
  title.className = 'card-title'
  title.textContent = artifact.title
  head.append(icon, title)
  if (artifact.pinned) {
    const pin = document.createElement('span')
    pin.className = 'pinned-mark'
    pin.textContent = '★'
    pin.title = '已釘選'
    head.append(pin)
  }
  link.append(head)
  if (artifact.description) {
    const desc = document.createElement('p')
    desc.className = 'card-desc'
    desc.style.margin = '0'
    desc.textContent = artifact.description
    link.append(desc)
  }
  const meta = document.createElement('div')
  meta.className = 'card-meta'
  const parts = [`v${artifact.currentVersion}`, relativeTime(artifact.updatedAt)]
  if (artifact.capabilities.length) parts.push(artifact.capabilities.join(' · '))
  for (const part of parts) {
    const span = document.createElement('span')
    span.textContent = part
    meta.append(span)
  }
  link.append(meta)
  if (artifact.snippet) {
    const snippet = document.createElement('p')
    snippet.className = 'card-snippet'
    // The hit is marked with text nodes only: the snippet is page text, never markup.
    const at = artifact.snippet.toLowerCase().indexOf(query.toLowerCase())
    if (at >= 0) {
      const mark = document.createElement('mark')
      mark.textContent = artifact.snippet.slice(at, at + query.length)
      snippet.append(artifact.snippet.slice(0, at), mark, artifact.snippet.slice(at + query.length))
    } else {
      snippet.textContent = artifact.snippet
    }
    link.append(snippet)
  }
  if (artifact.source) {
    const source = document.createElement('div')
    source.className = 'card-source'
    source.textContent = `${artifact.source.host} · ${artifact.source.path}`
    source.title = '最近一次從這台機器、這個資料夾發佈'
    link.append(source)
  }
  item.append(link)
  return item
}

async function load(reset) {
  const mine = ++generation
  if (reset) offset = 0
  const params = new URLSearchParams({ limit: String(PAGE), offset: String(offset) })
  if (query) params.set('q', query)
  status.textContent = reset ? '載入中…' : ''
  try {
    const response = await fetch(`/api/artifacts?${params}`)
    if (!response.ok) throw new Error(`HTTP ${response.status}`)
    const data = await response.json()
    if (mine !== generation) return
    if (reset) list.replaceChildren()
    list.append(...data.artifacts.map(card))
    offset += data.artifacts.length
    more.hidden = offset >= data.total
    status.textContent = data.total === 0 ? (query ? '沒有符合的 artifact。' : '還沒有 artifact。用 CLI 或請 agent 發佈第一個。') : `共 ${data.total} 個`
  } catch (error) {
    if (mine === generation) status.textContent = `載入失敗:${error.message}`
  }
}

let debounce
search.addEventListener('input', () => {
  clearTimeout(debounce)
  debounce = setTimeout(() => {
    query = search.value.trim()
    load(true)
  }, 200)
})
more.addEventListener('click', () => load(false))
document.addEventListener('keydown', (event) => {
  if (event.key === '/' && document.activeElement !== search) {
    event.preventDefault()
    search.focus()
  }
})
load(true)

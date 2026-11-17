import { createPasskey, describe, post } from './webauthn.js'

const list = document.getElementById('passkeys')
const sessions = document.getElementById('sessions')
const status = document.getElementById('status')

async function load() {
  const response = await fetch('/api/account')
  if (!response.ok) {
    location.href = '/login?next=/account'
    return
  }
  const account = await response.json()
  sessions.textContent = String(account.sessions)
  list.replaceChildren(...account.passkeys.map((key) => {
    const item = document.createElement('li')
    const label = document.createElement('span')
    label.textContent = `${key.name} · 建立於 ${key.createdAt.slice(0, 10)}${key.lastUsedAt ? ` · 上次使用 ${key.lastUsedAt.slice(0, 10)}` : ''}`
    const remove = document.createElement('button')
    remove.className = 'link-button danger'
    remove.type = 'button'
    remove.textContent = '刪除'
    remove.disabled = account.passkeys.length <= 1
    remove.addEventListener('click', async () => {
      if (!confirm(`刪除「${key.name}」?那台裝置就不能再登入,所有裝置(包括這一台)也會登出,要用其他 passkey 重新登入。`)) return
      const result = await fetch(`/api/account/passkeys/${encodeURIComponent(key.id)}`, { method: 'DELETE', headers: { 'x-artifact-client': 'viewer' } })
      if (result.ok) location.href = '/login?next=/account'
      else status.textContent = '刪除失敗。'
    })
    item.append(label, ' ', remove)
    return item
  }))
}

// Fetched ahead so the passkey prompt opens straight from the tap (Safari refuses one after a round trip).
let pending = null
function prepare() {
  pending = { startedAt: Date.now(), options: post('/auth/register/options') }
  pending.options.catch(() => {})
}
prepare()

document.getElementById('add-passkey').addEventListener('click', async () => {
  status.textContent = '等待裝置建立 passkey…'
  try {
    if (!pending || Date.now() - pending.startedAt > 4 * 60_000) prepare()
    const { flow, options } = await pending.options
    pending = null
    const response = await createPasskey(options)
    const name = prompt('這把 passkey 的名稱(例如 MacBook)') ?? ''
    await post('/auth/register/verify', { flow, response, name })
    status.textContent = '已新增。'
    load()
  } catch (error) {
    status.textContent = describe(error)
  }
  prepare()
})

document.getElementById('logout').addEventListener('click', async () => {
  await post('/auth/logout').catch(() => {})
  location.href = '/login'
})

document.getElementById('logout-all').addEventListener('click', async () => {
  if (!confirm('登出所有裝置(包括這一台)?')) return
  await post('/auth/logout-all').catch(() => {})
  location.href = '/login'
})

load()

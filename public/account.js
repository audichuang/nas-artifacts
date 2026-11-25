import { createPasskey, describe, post, signalAccepted, signalUnknown, unsupported } from './webauthn.js'

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
  const clients = document.getElementById('clients')
  clients.replaceChildren(...(account.clients.length ? account.clients.map((client) => {
    const item = document.createElement('li')
    const label = document.createElement('span')
    label.textContent = `${client.name}${client.lastUsedAt ? ` · 上次使用 ${client.lastUsedAt.slice(0, 10)}` : ''}`
    const revoke = document.createElement('button')
    revoke.className = 'link-button danger'
    revoke.type = 'button'
    revoke.textContent = '撤銷'
    revoke.addEventListener('click', async () => {
      if (!confirm(`撤銷「${client.name}」?它要重新授權才能再使用。`)) return
      const result = await fetch(`/api/account/clients/${encodeURIComponent(client.id)}`, { method: 'DELETE', headers: { 'x-artifact-client': 'viewer' } })
      status.textContent = result.ok ? '已撤銷。' : '撤銷失敗。'
      load()
    })
    item.append(label, ' ', revoke)
    return item
  }) : [Object.assign(document.createElement('li'), { className: 'muted', textContent: '(沒有)' })]))
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
      if (result.ok) {
        // The server's list after the removal, not this page's: a passkey enrolled elsewhere since must stay offered.
        const remaining = await result.json().catch(() => ({ passkeys: [] }))
        await signalAccepted(remaining.rpId, remaining.userId, remaining.passkeys)
        location.href = '/login?next=/account'
      }
      else status.textContent = '刪除失敗。'
    })
    item.append(label, ' ', remove)
    return item
  }))
  signalAccepted(account.rpId, account.userId, account.passkeys.map((key) => key.id))
}

// Fetched ahead so the passkey prompt opens straight from the tap (Safari refuses one after a round trip).
let pending = null
function prepare() {
  pending = { startedAt: Date.now(), options: post('/auth/register/options') }
  pending.options.catch(() => {})
}
const add = document.getElementById('add-passkey')
if (!unsupported(add)) prepare()

add.addEventListener('click', async () => {
  status.textContent = '等待裝置建立 passkey…'
  try {
    if (!pending || Date.now() - pending.startedAt > 4 * 60_000) prepare()
    const { flow, options } = await pending.options
    pending = null
    const response = await createPasskey(options)
    const name = prompt('這把 passkey 的名稱(例如 MacBook)') ?? ''
    await post('/auth/register/verify', { flow, response, name }).catch(async (error) => {
      // Same rule as the enrolment page: only a refusal (4xx) means the server stored nothing.
      if (error.code && error.status >= 400 && error.status < 500) await signalUnknown(options.rp.id, response.id)
      throw error
    })
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

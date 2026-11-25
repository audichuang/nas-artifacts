import { createPasskey, describe, getPasskey, post, signalUnknown, unsupported } from './webauthn.js'

const card = document.querySelector('.auth-card')
const button = document.getElementById('go')
const status = document.getElementById('status')
const mode = card.dataset.mode
// The enrolment code rides in the fragment, which never reaches a server log; drop it from the address bar.
const code = mode === 'enroll' ? location.hash.slice(1) : ''
if (code) history.replaceState(null, '', location.pathname)

// Options are fetched before the tap, so the passkey prompt opens straight from the tap: Safari refuses one
// that follows a network round trip.
let ready = null
function prepare() {
  const startedAt = Date.now()
  const options = mode === 'enroll' ? post('/auth/register/options', { code }) : post('/auth/login/options')
  ready = { startedAt, options }
  options.catch(() => {})
}

async function run() {
  button.disabled = true
  status.textContent = mode === 'enroll' ? '等待裝置建立 passkey…' : '等待 passkey…'
  try {
    if (mode === 'enroll' && !code) throw new Error('連結不完整，請重新產生註冊連結。')
    // A ceremony lives five minutes on the server.
    if (!ready || Date.now() - ready.startedAt > 4 * 60_000) prepare()
    const { flow, options } = await ready.options
    ready = null
    if (mode === 'enroll') {
      const response = await createPasskey(options)
      // Refused (4xx), so stored nowhere: the device should not keep offering it. A 5xx or a lost response may
      // come after the passkey was stored, so those leave it alone.
      await post('/auth/register/verify', { flow, response, name: document.getElementById('name').value }).catch(async (error) => {
        if (error.status >= 400 && error.status < 500) await signalUnknown(options.rp.id, response.id)
        throw error
      })
    } else {
      const response = await getPasskey(options)
      await post('/auth/login/verify', { flow, response }).catch(async (error) => {
        if (error.status === 404) await signalUnknown(options.rpId, response.id)
        throw error
      })
    }
    status.textContent = '完成，正在前往…'
    location.replace(card.dataset.next || '/')
  } catch (error) {
    status.textContent = describe(error)
    button.disabled = false
    if (mode !== 'enroll' || code) prepare()
  }
}

if (!unsupported(button)) {
  if (mode !== 'enroll' || code) prepare()
  button.addEventListener('click', run)
}

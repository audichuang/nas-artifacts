import { createPasskey, describe, getPasskey, post } from './webauthn.js'

const card = document.querySelector('.auth-card')
const button = document.getElementById('go')
const status = document.getElementById('status')
const mode = card.dataset.mode
// The enrolment code rides in the fragment, which never reaches a server log; drop it from the address bar.
const code = mode === 'enroll' ? location.hash.slice(1) : ''
if (code) history.replaceState(null, '', location.pathname)

async function run() {
  button.disabled = true
  status.textContent = mode === 'enroll' ? '等待裝置建立 passkey…' : '等待 passkey…'
  try {
    if (mode === 'enroll') {
      if (!code) throw new Error('連結不完整，請重新產生註冊連結。')
      const { flow, options } = await post('/auth/register/options', { code })
      const response = await createPasskey(options)
      await post('/auth/register/verify', { flow, response, name: document.getElementById('name').value })
    } else {
      const { flow, options } = await post('/auth/login/options')
      const response = await getPasskey(options)
      await post('/auth/login/verify', { flow, response })
    }
    status.textContent = '完成，正在前往…'
    location.replace(card.dataset.next || '/')
  } catch (error) {
    status.textContent = describe(error)
    button.disabled = false
  }
}

button.addEventListener('click', run)

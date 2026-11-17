import { post } from './webauthn.js'

const card = document.querySelector('.auth-card')
const status = document.getElementById('status')
const buttons = [document.getElementById('allow'), document.getElementById('deny')]

async function decide(allow) {
  for (const button of buttons) button.disabled = true
  status.textContent = allow ? '授權中…' : '已拒絕，正在返回…'
  try {
    const { redirect } = await post('/oauth/authorize/decision', { request: card.dataset.request, allow })
    status.textContent = allow ? '已授權，正在返回應用程式…(可以關掉這頁)' : '已拒絕。'
    location.href = redirect
  } catch (error) {
    status.textContent = error.message
  }
}

buttons[0].addEventListener('click', () => decide(true))
buttons[1].addEventListener('click', () => decide(false))

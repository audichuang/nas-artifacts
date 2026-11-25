// WebAuthn JSON <-> browser objects, and the calls the sign-in pages share. The browser's own JSON helpers
// (parse*OptionsFromJSON, toJSON) where it has them; a few hand-rolled lines each way otherwise.

const toBytes = (value) => Uint8Array.from(atob(value.replace(/-/g, '+').replace(/_/g, '/')), (c) => c.charCodeAt(0))
const toBase64url = (buffer) => btoa(String.fromCharCode(...new Uint8Array(buffer))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')

export async function post(path, body) {
  const response = await fetch(path, {
    method: 'POST',
    headers: { 'content-type': 'application/json', 'x-artifact-client': 'viewer' },
    body: JSON.stringify(body ?? {})
  })
  const payload = await response.json().catch(() => null)
  if (!response.ok) throw Object.assign(new Error(payload?.error?.message || `HTTP ${response.status}`), { status: response.status, code: payload?.error?.code })
  return payload
}

/** Without WebAuthn the page says so and keeps its passkey buttons off. */
export function unsupported(...buttons) {
  if (window.PublicKeyCredential) return false
  for (const button of buttons) button.disabled = true
  document.getElementById('status').textContent = '這個瀏覽器不支援 passkey，請換一個瀏覽器。'
  return true
}

// Per call and per object: a password manager's extension may hand back a credential without toJSON, or one
// whose toJSON throws, so the hand-rolled conversions stay as the fallback.
function parsed(parse, options, fallback) {
  try {
    if (typeof parse === 'function') return parse.call(PublicKeyCredential, options)
  } catch {}
  return fallback()
}

function encoded(credential, fallback) {
  try {
    if (typeof credential.toJSON === 'function') return credential.toJSON()
  } catch {}
  return fallback(credential)
}

export async function createPasskey(options) {
  const publicKey = parsed(PublicKeyCredential.parseCreationOptionsFromJSON, options, () => ({
    ...options,
    challenge: toBytes(options.challenge),
    user: { ...options.user, id: toBytes(options.user.id) },
    excludeCredentials: (options.excludeCredentials ?? []).map((item) => ({ ...item, id: toBytes(item.id) }))
  }))
  return encoded(await navigator.credentials.create({ publicKey }), (credential) => ({
    id: credential.id,
    rawId: toBase64url(credential.rawId),
    type: credential.type,
    clientExtensionResults: credential.getClientExtensionResults(),
    authenticatorAttachment: credential.authenticatorAttachment ?? undefined,
    response: {
      clientDataJSON: toBase64url(credential.response.clientDataJSON),
      attestationObject: toBase64url(credential.response.attestationObject),
      transports: credential.response.getTransports?.() ?? []
    }
  }))
}

export async function getPasskey(options) {
  const publicKey = parsed(PublicKeyCredential.parseRequestOptionsFromJSON, options, () => ({
    ...options,
    challenge: toBytes(options.challenge),
    allowCredentials: (options.allowCredentials ?? []).map((item) => ({ ...item, id: toBytes(item.id) }))
  }))
  return encoded(await navigator.credentials.get({ publicKey }), (credential) => ({
    id: credential.id,
    rawId: toBase64url(credential.rawId),
    type: credential.type,
    clientExtensionResults: credential.getClientExtensionResults(),
    authenticatorAttachment: credential.authenticatorAttachment ?? undefined,
    response: {
      clientDataJSON: toBase64url(credential.response.clientDataJSON),
      authenticatorData: toBase64url(credential.response.authenticatorData),
      signature: toBase64url(credential.response.signature),
      userHandle: credential.response.userHandle ? toBase64url(credential.response.userHandle) : undefined
    }
  }))
}

export function describe(error) {
  if (error?.name === 'NotAllowedError') return '已取消，或這台裝置沒有可用的 passkey。'
  if (error?.name === 'AbortError') return '已取消。'
  if (error?.name === 'InvalidStateError') return '這台裝置已經有這裡的 passkey 了。'
  if (error?.name === 'SecurityError') return '這個網址不能使用 passkey（網域設定不符）。'
  return error?.message || String(error)
}

// Signal API (Chrome 132+, Safari 26+) keeps the passkey provider in step with the server. Feature-detected and
// never fatal; the RP ID always comes from the server, not from location.hostname.
const isBase64url = (value) => typeof value === 'string' && value.length > 0 && value.length % 4 !== 1 && /^[A-Za-z0-9_-]+$/.test(value)

/** Only after the server answered that it does not hold this credential; never on a network failure, when it may. */
export async function signalUnknown(rpId, credentialId) {
  if (!window.PublicKeyCredential?.signalUnknownCredential || !rpId || !isBase64url(credentialId)) return
  try {
    await PublicKeyCredential.signalUnknownCredential({ rpId, credentialId })
  } catch (error) {
    console.warn('signalUnknownCredential:', error)
  }
}

/** The owner's complete passkey list: the provider hides any other passkey it holds for this account. */
export async function signalAccepted(rpId, userId, credentialIds) {
  // Never an empty or malformed list: a short list would hide passkeys that still work.
  if (!window.PublicKeyCredential?.signalAllAcceptedCredentials || !rpId || !isBase64url(userId) || !credentialIds.length || !credentialIds.every(isBase64url)) return
  try {
    await PublicKeyCredential.signalAllAcceptedCredentials({ rpId, userId, allAcceptedCredentialIds: credentialIds })
  } catch (error) {
    console.warn('signalAllAcceptedCredentials:', error)
  }
}

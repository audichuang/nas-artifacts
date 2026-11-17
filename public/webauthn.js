// WebAuthn JSON <-> browser objects, and the calls the sign-in pages share. No library: a few lines each way.

const toBytes = (value) => Uint8Array.from(atob(value.replace(/-/g, '+').replace(/_/g, '/')), (c) => c.charCodeAt(0))
const toBase64url = (buffer) => btoa(String.fromCharCode(...new Uint8Array(buffer))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')

export async function post(path, body) {
  const response = await fetch(path, {
    method: 'POST',
    headers: { 'content-type': 'application/json', 'x-artifact-client': 'viewer' },
    body: JSON.stringify(body ?? {})
  })
  const payload = await response.json().catch(() => null)
  if (!response.ok) throw new Error(payload?.error?.message || `HTTP ${response.status}`)
  return payload
}

export async function createPasskey(options) {
  const credential = await navigator.credentials.create({
    publicKey: {
      ...options,
      challenge: toBytes(options.challenge),
      user: { ...options.user, id: toBytes(options.user.id) },
      excludeCredentials: (options.excludeCredentials ?? []).map((item) => ({ ...item, id: toBytes(item.id) }))
    }
  })
  const response = credential.response
  return {
    id: credential.id,
    rawId: toBase64url(credential.rawId),
    type: credential.type,
    clientExtensionResults: credential.getClientExtensionResults(),
    authenticatorAttachment: credential.authenticatorAttachment ?? undefined,
    response: {
      clientDataJSON: toBase64url(response.clientDataJSON),
      attestationObject: toBase64url(response.attestationObject),
      transports: response.getTransports?.() ?? []
    }
  }
}

export async function getPasskey(options) {
  const credential = await navigator.credentials.get({
    publicKey: {
      ...options,
      challenge: toBytes(options.challenge),
      allowCredentials: (options.allowCredentials ?? []).map((item) => ({ ...item, id: toBytes(item.id) }))
    }
  })
  const response = credential.response
  return {
    id: credential.id,
    rawId: toBase64url(credential.rawId),
    type: credential.type,
    clientExtensionResults: credential.getClientExtensionResults(),
    authenticatorAttachment: credential.authenticatorAttachment ?? undefined,
    response: {
      clientDataJSON: toBase64url(response.clientDataJSON),
      authenticatorData: toBase64url(response.authenticatorData),
      signature: toBase64url(response.signature),
      userHandle: response.userHandle ? toBase64url(response.userHandle) : undefined
    }
  }
}

export function describe(error) {
  if (error?.name === 'NotAllowedError') return '已取消，或這台裝置沒有可用的 passkey。'
  if (error?.name === 'InvalidStateError') return '這台裝置已經註冊過了。'
  return error?.message || String(error)
}

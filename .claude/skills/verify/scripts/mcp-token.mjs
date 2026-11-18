#!/usr/bin/env node
// Gets an OAuth access token for the remote MCP the way an MCP client does: passkey sign-in (simulated
// authenticator), dynamic registration, PKCE S256 authorize + consent click, loopback callback, token exchange.
// Usage: node .claude/skills/verify/scripts/mcp-token.mjs
//   writes .verify/run/mcp-token.json (mode 0600) and prints only the client id and expiry, never the token.
// As a module: authorizeClient(page, feature, clientName) for a page already signed in as the owner.
import { createHash, randomBytes } from 'node:crypto'
import { writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { V, enrolOwner, launch, log, secret } from './browser.mjs'

export const TOKEN_FILE = join(V.run, 'mcp-token.json')
export const MCP_URL = `${V.publicApp}/mcp`

export async function authorizeClient(page, feature, clientName = 'verify MCP client') {
  const reg = await fetch(`${V.publicApp}/oauth/register`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ client_name: clientName, redirect_uris: [V.callback] })
  })
  const client = await reg.json()
  if (reg.status !== 201) throw new Error(`register -> ${reg.status} ${JSON.stringify(client)}`)
  log(feature, `POST /oauth/register -> ${reg.status} client_id ${client.client_id}`)
  const verifier = randomBytes(32).toString('base64url')
  const challenge = createHash('sha256').update(verifier).digest('base64url')
  const state = randomBytes(8).toString('hex')
  // The client's loopback listener, played by Playwright: nothing binds the port.
  await page.route(`${V.callback}**`, (route) => route.fulfill({ status: 200, contentType: 'text/plain', body: 'callback received' }))
  const query = new URLSearchParams({ response_type: 'code', client_id: client.client_id, redirect_uri: V.callback, code_challenge: challenge, code_challenge_method: 'S256', state, resource: MCP_URL, scope: 'artifacts' })
  await page.goto(`${V.publicApp}/oauth/authorize?${query}`)
  await page.locator('#allow').waitFor({ timeout: 5000 })
  const consent = await page.locator('body').innerText()
  await Promise.all([page.waitForURL(`${V.callback}**`, { timeout: 10_000 }), page.locator('#allow').click()])
  const back = new URL(page.url())
  const code = secret(back.searchParams.get('code'))
  if (back.searchParams.get('state') !== state) throw new Error('callback state does not match')
  log(feature, `consent allowed; callback has code, state matches, iss=${back.searchParams.get('iss')}`)
  const tokenResponse = await fetch(`${V.publicApp}/oauth/token`, {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ grant_type: 'authorization_code', code, redirect_uri: V.callback, client_id: client.client_id, code_verifier: verifier, resource: MCP_URL })
  })
  const tokens = await tokenResponse.json()
  if (tokenResponse.status !== 200) throw new Error(`token -> ${tokenResponse.status} ${JSON.stringify(tokens)}`)
  secret(tokens.access_token)
  secret(tokens.refresh_token)
  const saved = { client_id: client.client_id, client_name: clientName, obtained_at: new Date().toISOString(), ...tokens }
  writeFileSync(TOKEN_FILE, JSON.stringify(saved, null, 2), { mode: 0o600 })
  log(feature, `POST /oauth/token -> 200, token_type ${tokens.token_type}, expires_in ${tokens.expires_in}s, saved to .verify/run/mcp-token.json (0600)`)
  return { clientId: client.client_id, iss: back.searchParams.get('iss'), consent, tokens }
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const browser = await launch()
  try {
    const page = await (await browser.newContext()).newPage()
    await enrolOwner(page, 'mcp-token')
    const { clientId, tokens } = await authorizeClient(page, 'mcp-token')
    console.log(`client ${clientId}; access token valid ${tokens.expires_in}s; in ${TOKEN_FILE}`)
  } finally {
    await browser.close()
  }
}

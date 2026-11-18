#!/usr/bin/env node
// Calls the remote MCP (Streamable HTTP, JSON-RPC) with the token mcp-token.mjs saved.
// Usage: node .claude/skills/verify/scripts/mcp-call.mjs <method> [json-params]
//        node .claude/skills/verify/scripts/mcp-call.mjs tool <name> <json-args>
//   prints the HTTP status and the JSON reply (for a tool: the parsed result.content[0].text).
// As a module: rpc(method, params, token?) and tool(name, args, token?).
import { existsSync, readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { MCP_URL, TOKEN_FILE } from './mcp-token.mjs'

let nextId = 0
export function savedToken() {
  if (!existsSync(TOKEN_FILE)) throw new Error('no .verify/run/mcp-token.json: run mcp-token.mjs first')
  return JSON.parse(readFileSync(TOKEN_FILE, 'utf8')).access_token
}

/** One JSON-RPC call; token null sends none. Returns { status, headers, body }. */
export async function rpc(method, params = {}, token = savedToken()) {
  const response = await fetch(MCP_URL, {
    method: 'POST',
    headers: {
      'content-type': 'application/json',
      accept: 'application/json, text/event-stream',
      'mcp-protocol-version': '2025-11-25',
      ...(token ? { authorization: `Bearer ${token}` } : {})
    },
    body: JSON.stringify({ jsonrpc: '2.0', id: ++nextId, method, params })
  })
  const text = await response.text()
  let body = null
  try {
    body = JSON.parse(text)
  } catch {
    body = text
  }
  return { status: response.status, headers: Object.fromEntries(response.headers), body }
}

/** tools/call; returns { status, isError, value } with value parsed from result.content[0].text. */
export async function tool(name, args, token = savedToken()) {
  const { status, body } = await rpc('tools/call', { name, arguments: args }, token)
  const result = body?.result
  if (!result) return { status, isError: true, value: body }
  const text = result.content?.[0]?.text ?? ''
  let value = text
  try {
    value = JSON.parse(text)
  } catch {}
  return { status, isError: result.isError === true, value }
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const [first, ...rest] = process.argv.slice(2)
  if (!first) {
    console.error('usage: mcp-call.mjs <method> [json-params] | mcp-call.mjs tool <name> <json-args>')
    process.exit(64)
  }
  if (first === 'tool') {
    const out = await tool(rest[0], JSON.parse(rest[1] ?? '{}'))
    console.log(`HTTP ${out.status}${out.isError ? ' (tool error)' : ''}`)
    console.log(JSON.stringify(out.value, null, 2))
    process.exitCode = out.status === 200 && !out.isError ? 0 : 1
  } else {
    const out = await rpc(first, JSON.parse(rest[0] ?? '{}'))
    console.log(`HTTP ${out.status}`)
    console.log(typeof out.body === 'string' ? out.body : JSON.stringify(out.body, null, 2))
    process.exitCode = out.status === 200 && !out.body?.error ? 0 : 1
  }
}

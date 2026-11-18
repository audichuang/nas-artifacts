// Shared helpers for the verify journeys: instance settings, the CLI wrapper, Chromium, evidence.
// Nothing here imports src/ or touches the store: every action goes through the CLI, a browser or HTTP.
import { spawn, spawnSync } from 'node:child_process'
import { appendFileSync, existsSync, mkdirSync, readFileSync, readdirSync, renameSync, writeFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { homedir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
export const ROOT = join(here, '..', '..', '..', '..')
const RUN = join(ROOT, '.verify', 'run')
const readRun = (name) => (existsSync(join(RUN, name)) ? readFileSync(join(RUN, name), 'utf8').trim() : '')
// Same rule as env.sh: the block launch.sh recorded wins over an exported VERIFY_PORT_BASE, so a stale export
// never sends this worktree's journeys to another worktree's instance. Written back so the CLI wrapper agrees.
const recordedBase = readRun('port-base')
if (recordedBase && process.env.VERIFY_PORT_BASE && process.env.VERIFY_PORT_BASE !== recordedBase) {
  console.error(`verify: ignoring exported VERIFY_PORT_BASE=${process.env.VERIFY_PORT_BASE}; this worktree's instance was launched on ${recordedBase} (.verify/run/port-base)`)
}
const base = Number(recordedBase || process.env.VERIFY_PORT_BASE || 28790)
process.env.VERIFY_PORT_BASE = String(base)

/** Where everything lives for the running instance (same values as env.sh). */
export const V = {
  root: ROOT,
  run: RUN,
  wrapper: join(here, 'artifact'),
  app: `http://localhost:${base}`,
  content: `http://127.0.0.1:${base + 1}`,
  publicApp: `http://localhost:${base + 2}`,
  publicContent: `http://127.0.0.1:${base + 3}`,
  callback: `http://127.0.0.1:${base + 9}/callback`,
  stamp: readRun('stamp')
}

const claimed = new Set()
/**
 * The evidence folder for one feature of this run; survives stop.sh. The first use in a process claims it: a
 * folder an earlier run of the same feature left under this stamp moves to <feature>.run-<n>, so steps.log,
 * cli.log and result.json in <feature>/ always describe one run.
 */
export function evidence(feature) {
  if (!V.stamp) throw new Error('no .verify/run/stamp: run scripts/launch.sh first')
  const dir = join(ROOT, '.verify', 'evidence', V.stamp, feature)
  if (!claimed.has(feature)) {
    claimed.add(feature)
    if (existsSync(dir) && readdirSync(dir).length > 0) {
      let n = 1
      while (existsSync(`${dir}.run-${n}`)) n++
      renameSync(dir, `${dir}.run-${n}`)
      console.error(`verify: an earlier ${feature} run in this stamp moved to ${feature}.run-${n}/`)
    }
  }
  mkdirSync(dir, { recursive: true })
  return dir
}

const secrets = new Set()
/** Register a value that must never reach evidence (tokens, codes). */
export function secret(value) {
  if (value) secrets.add(String(value))
  return value
}
/** Masks registered secrets, login-link codes and content tokens in text bound for evidence. */
export function redact(text) {
  let out = String(text)
  for (const value of secrets) out = out.split(value).join('<redacted>')
  return out
    .replace(/(\/enroll#)[A-Za-z0-9_-]+/g, '$1<redacted>')
    .replace(/(\/t\/)[A-Za-z0-9._~-]+(\/a\/)/g, '$1<token>$2')
    .replace(/(Bearer )(?!resource_metadata)[A-Za-z0-9._~-]{12,}/g, '$1<redacted>')
}

/** Appends a line to <feature>/steps.log and echoes it. */
export function log(feature, line) {
  const text = redact(line)
  appendFileSync(join(evidence(feature), 'steps.log'), `${text}\n`)
  console.log(text)
}

/** Writes a JSON or text file into the feature's evidence (redacted). */
export function save(feature, name, value) {
  const text = typeof value === 'string' ? value : JSON.stringify(value, null, 2)
  writeFileSync(join(evidence(feature), name), redact(text))
}

/**
 * Runs the CLI wrapper (never the bare CLI) and records command, exit code, stdout and stderr.
 * opts.env adds variables (e.g. VERIFY_CLI_DIR for a second machine); opts.quiet skips the evidence record.
 */
export function cli(feature, args, opts = {}) {
  const result = spawnSync(V.wrapper, args, { cwd: opts.cwd ?? ROOT, env: { ...process.env, ...opts.env }, encoding: 'utf8' })
  const record = { code: result.status, stdout: result.stdout ?? '', stderr: result.stderr ?? '' }
  if (!opts.quiet) {
    const prefix = opts.env?.VERIFY_CLI_DIR ? 'VERIFY_CLI_DIR=<second> ' : ''
    appendFileSync(join(evidence(feature), 'cli.log'), redact(`$ ${prefix}artifact ${args.join(' ')}\n[exit ${record.code}]\n${record.stdout}${record.stderr ? `[stderr]\n${record.stderr}` : ''}\n`))
  }
  return record
}

/** Starts the CLI wrapper in the background (e.g. `comments --wait`); resolves its result when it exits. */
export function cliBackground(feature, args) {
  const child = spawn(V.wrapper, args, { cwd: ROOT, env: process.env })
  let stdout = ''
  let stderr = ''
  child.stdout.on('data', (d) => (stdout += d))
  child.stderr.on('data', (d) => (stderr += d))
  const done = new Promise((resolve) =>
    child.on('exit', (code) => {
      appendFileSync(join(evidence(feature), 'cli.log'), redact(`$ artifact ${args.join(' ')}   (background, pid ${child.pid})\n[exit ${code}]\n${stdout}${stderr ? `[stderr]\n${stderr}` : ''}\n`))
      resolve({ code, stdout, stderr })
    })
  )
  return { child, done, stderr: () => stderr }
}

/** CHROMIUM_PATH, else the newest Playwright headless shell (macOS or Linux cache), else a system Chromium. */
export function findChromium() {
  if (process.env.CHROMIUM_PATH) return existsSync(process.env.CHROMIUM_PATH) ? process.env.CHROMIUM_PATH : undefined
  const caches = [join(homedir(), 'Library/Caches/ms-playwright'), join(homedir(), '.cache/ms-playwright')]
  for (const cache of caches) {
    if (!existsSync(cache)) continue
    const shells = readdirSync(cache)
      .filter((name) => name.startsWith('chromium_headless_shell-'))
      .sort((a, b) => Number(b.split('-')[1]) - Number(a.split('-')[1]))
    for (const shell of shells) {
      for (const sub of readdirSync(join(cache, shell))) {
        const binary = join(cache, shell, sub, 'chrome-headless-shell')
        if (sub.startsWith('chrome-headless-shell') && existsSync(binary)) return binary
      }
    }
  }
  return ['/usr/bin/chromium', '/usr/bin/chromium-browser'].find((path) => existsSync(path))
}

/** Launches Chromium through the repo's playwright-core. Close it in a finally block. */
export async function launch() {
  const executablePath = findChromium()
  if (!executablePath) throw new Error('no Chromium: set CHROMIUM_PATH or run scripts/doctor.sh')
  const { chromium } = createRequire(join(ROOT, 'package.json'))('playwright-core')
  return chromium.launch({ executablePath })
}

/** Screenshot into the feature's evidence folder. */
export async function shot(page, feature, name) {
  const path = join(evidence(feature), `${name}.png`)
  await page.screenshot({ path, fullPage: false })
  log(feature, `screenshot ${name}.png  (${redact(page.url())})`)
  return path
}

/** The sandboxed page frame inside a viewer page, found by its origin. */
export function contentFrame(page, origin) {
  return page.frames().find((frame) => frame.url().startsWith(`${origin}/`))
}

/** A simulated passkey device (CDP virtual authenticator) for this page. */
export async function withAuthenticator(page) {
  const cdp = await page.context().newCDPSession(page)
  await cdp.send('WebAuthn.enable')
  await cdp.send('WebAuthn.addVirtualAuthenticator', {
    options: { protocol: 'ctap2', transport: 'internal', hasResidentKey: true, hasUserVerification: true, isUserVerified: true, automaticPresenceSimulation: true }
  })
}

/** Registers a passkey from a fresh `artifact login-link` and lands signed in on the public gallery. */
export async function enrolOwner(page, feature, deviceName = 'verify device') {
  await withAuthenticator(page)
  const link = cli(feature, ['login-link', '--json'], { quiet: true })
  if (link.code !== 0) throw new Error(`login-link failed: ${link.stderr}`)
  const url = JSON.parse(link.stdout).url
  secret(url.split('#')[1])
  log(feature, `artifact login-link -> ${url}  (exit 0)`)
  await page.goto(url)
  await page.locator('#name').fill(deviceName)
  await page.locator('#go').click()
  await page.waitForURL(`${V.publicApp}/`, { timeout: 10_000 })
  log(feature, `enrolled passkey "${deviceName}", now on ${page.url()}`)
}

/** Records pass/fail checks and turns them into an exit code. */
export function checks(feature) {
  const results = []
  return {
    ok(condition, label, detail = '') {
      results.push({ pass: Boolean(condition), label, detail: redact(String(detail)) })
      log(feature, `${condition ? 'PASS' : 'FAIL'}  ${label}${detail ? `  [${String(detail).slice(0, 300)}]` : ''}`)
      return Boolean(condition)
    },
    finish(error) {
      if (error) results.push({ pass: false, label: 'journey threw', detail: redact(error.stack ?? String(error)) })
      const failed = results.filter((r) => !r.pass)
      save(feature, 'result.json', { feature, pass: failed.length === 0, checks: results })
      log(feature, `${feature}: ${results.length - failed.length}/${results.length} checks passed${error ? ` (threw: ${redact(error.message)})` : ''}`)
      process.exitCode = failed.length === 0 ? 0 : 1
    }
  }
}

/** A scratch folder for the files a journey publishes (inside .verify/run, removed by stop.sh). */
export function scratch(name) {
  const dir = join(RUN, 'work', name)
  mkdirSync(dir, { recursive: true })
  return dir
}

#!/usr/bin/env bash
# Read-only health check: starts nothing, writes nothing, kills nothing, prints no token.
# Usage: scripts/doctor.sh      exit 0 = all ok, 1 = at least one FAIL
set -uo pipefail
. "$(dirname "$0")/env.sh" || exit 1

fails=0
ok() { echo "ok    $1"; }
fail() { echo "FAIL  $1"; echo "      fix: $2"; fails=1; }
warn() { echo "warn  $1"; }

major="$(node -p 'process.versions.node.split(".")[0]' 2>/dev/null || echo 0)"
if [ "${major}" -ge 26 ] 2>/dev/null; then ok "node $(node -v)"; else fail "node >= 26 needed (the image runs Node 26; .node-version pins it), found ${major}" "install Node 26 (mise picks it up from .node-version)"; fi

for mod in hono playwright-core pptxgenjs; do
  if [ -d "${VERIFY_ROOT}/node_modules/${mod}" ]; then ok "node_modules/${mod}"; else fail "node_modules/${mod} missing" "pnpm install --frozen-lockfile (in ${VERIFY_ROOT})"; fi
done

chromium="$(cd "${VERIFY_ROOT}" && node -e 'import(process.argv[1]).then((m) => console.log(m.findChromium() ?? ""))' "${VERIFY_ROOT}/.claude/skills/verify/scripts/browser.mjs" 2>/dev/null)"
if [ -n "${chromium}" ]; then ok "chromium ${chromium}"; else fail "no Chromium found" "set CHROMIUM_PATH, or: pnpm exec playwright-core install chromium-headless-shell"; fi

ours="$(verify_server_pid)"
for port in ${VERIFY_PORTS}; do
  holders="$(verify_port_pids "${port}")"
  if [ -z "${holders}" ]; then ok "port ${port} free"
  elif [ -n "${ours}" ] && [ "${holders}" = "${ours}" ]; then ok "port ${port} held by our server (pid ${ours}, already running)"
  else fail "port ${port} held by pid ${holders}, not this run's server" "leave it alone; launch with another VERIFY_PORT_BASE (never 18790-18799, the e2e suite's)"; fi
done
echo "note  ports 18790-18793 and 18799 belong to test/bridge.e2e.test.ts; this skill uses ${VERIFY_PORT_BASE}-$((VERIFY_PORT_BASE + 3)) and callback ${VERIFY_CALLBACK_PORT}"

case "${ARTIFACTS_URL:-}" in
  ''|http://localhost:*|http://127.0.0.1:*) ok "caller ARTIFACTS_URL is unset or loopback" ;;
  *) warn "your shell exports a non-loopback ARTIFACTS_URL; the bare CLI would publish there. Always use scripts/artifact" ;;
esac
if [ -e "${HOME}/.config/nas-artifacts/config.json" ]; then
  warn "a real CLI config exists in ~/.config/nas-artifacts (not read). The bare CLI would use its token: always use scripts/artifact"
else
  ok "no real CLI config in ~/.config/nas-artifacts"
fi

if grep -qx '\.verify/' "${VERIFY_ROOT}/.gitignore" 2>/dev/null; then ok ".gitignore has .verify/"; else fail ".gitignore lacks .verify/" "add a line '.verify/' to .gitignore"; fi

if [ -n "${ours}" ]; then
  if out="$(verify_ready)"; then ok "running instance (pid ${ours}) is ready: 4 x /healthz, log lines, CLI list, /mcp 401"
  else fail "running instance (pid ${ours}) is not ready: ${out}" "read ${VERIFY_RUN}/server.log; stop.sh then launch.sh"; fi
else
  echo "note  no verify instance running (launch.sh starts one)"
fi
exit "${fails}"

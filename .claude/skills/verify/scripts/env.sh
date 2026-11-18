# Sourced by the other verify scripts: paths, ports, origins and the throwaway token of the local instance.
# bash 3.2 compatible (macOS): no associative arrays, no ${!var}.

VERIFY_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/../../../.." && pwd)"
VERIFY_DIR="${VERIFY_ROOT}/.verify"
VERIFY_RUN="${VERIFY_DIR}/run"

# The port block launch.sh recorded wins over an exported VERIFY_PORT_BASE, so every later call (and stop.sh)
# talks to this worktree's instance, never to another worktree's block a stale export points at.
# browser.mjs applies the same rule. To move to another block: stop.sh, then launch.sh with the new value.
if [ -f "${VERIFY_RUN}/port-base" ]; then
  verify_recorded_base="$(cat "${VERIFY_RUN}/port-base")"
  if [ -n "${VERIFY_PORT_BASE:-}" ] && [ "${VERIFY_PORT_BASE}" != "${verify_recorded_base}" ] && [ "${verify_base_warn:-1}" = 1 ]; then
    echo "verify: ignoring exported VERIFY_PORT_BASE=${VERIFY_PORT_BASE}; this worktree's instance was launched on ${verify_recorded_base} (.verify/run/port-base). Run stop.sh first to move to another block." >&2
  fi
  VERIFY_PORT_BASE="${verify_recorded_base}"
fi
VERIFY_PORT_BASE="${VERIFY_PORT_BASE:-28790}"
case "${VERIFY_PORT_BASE}" in
  ''|*[!0-9]*) echo "verify: VERIFY_PORT_BASE must be a number, got '${VERIFY_PORT_BASE}'" >&2; return 1 2>/dev/null || exit 1 ;;
esac
# The block uses base..base+9 (base+9 is the OAuth callback); 18790-18799 belongs to test/bridge.e2e.test.ts.
if [ $((VERIFY_PORT_BASE + 9)) -ge 18790 ] && [ "${VERIFY_PORT_BASE}" -le 18799 ]; then
  echo "verify: ports ${VERIFY_PORT_BASE}-$((VERIFY_PORT_BASE + 9)) overlap the e2e suite's 18790-18799; pick another VERIFY_PORT_BASE" >&2
  return 1 2>/dev/null || exit 1
fi

VERIFY_APP_PORT="${VERIFY_PORT_BASE}"
VERIFY_CONTENT_PORT=$((VERIFY_PORT_BASE + 1))
VERIFY_PUBLIC_APP_PORT=$((VERIFY_PORT_BASE + 2))
VERIFY_PUBLIC_CONTENT_PORT=$((VERIFY_PORT_BASE + 3))
VERIFY_CALLBACK_PORT=$((VERIFY_PORT_BASE + 9))
# App origins on localhost (passkey rpID, __Host- cookie), content origins on 127.0.0.1 (a different origin).
VERIFY_APP_ORIGIN="http://localhost:${VERIFY_APP_PORT}"
VERIFY_CONTENT_ORIGIN="http://127.0.0.1:${VERIFY_CONTENT_PORT}"
VERIFY_PUBLIC_APP_ORIGIN="http://localhost:${VERIFY_PUBLIC_APP_PORT}"
VERIFY_PUBLIC_CONTENT_ORIGIN="http://127.0.0.1:${VERIFY_PUBLIC_CONTENT_PORT}"
VERIFY_TOKEN="verify-throwaway-token-0001"
VERIFY_PORTS="${VERIFY_APP_PORT} ${VERIFY_CONTENT_PORT} ${VERIFY_PUBLIC_APP_PORT} ${VERIFY_PUBLIC_CONTENT_PORT}"

export VERIFY_ROOT VERIFY_DIR VERIFY_RUN VERIFY_PORT_BASE VERIFY_APP_PORT VERIFY_CONTENT_PORT VERIFY_PUBLIC_APP_PORT \
  VERIFY_PUBLIC_CONTENT_PORT VERIFY_CALLBACK_PORT VERIFY_APP_ORIGIN VERIFY_CONTENT_ORIGIN VERIFY_PUBLIC_APP_ORIGIN \
  VERIFY_PUBLIC_CONTENT_ORIGIN VERIFY_TOKEN VERIFY_PORTS

# PID of our server if the pidfile names a live `node src/server.ts`; empty otherwise.
verify_server_pid() {
  local pid
  [ -f "${VERIFY_RUN}/server.pid" ] || return 0
  pid="$(cat "${VERIFY_RUN}/server.pid")"
  case "${pid}" in ''|*[!0-9]*) return 0 ;; esac
  if ps -p "${pid}" -o command= 2>/dev/null | grep -q 'src/server.ts'; then
    echo "${pid}"
  fi
}

# PIDs listening on a TCP port (empty when free).
verify_port_pids() {
  { lsof -nP -t -iTCP:"$1" -sTCP:LISTEN 2>/dev/null || true; } | tr '\n' ' ' | sed 's/ $//'
}

# The four "ready means usable" checks; prints one line per failure, returns 1 if any failed.
verify_ready() {
  local failed=0 port code headers
  for port in ${VERIFY_PORTS}; do
    code="$(curl -s -o /dev/null -w '%{http_code}' --max-time 2 "http://127.0.0.1:${port}/healthz" || true)"
    if [ "${code}" != "200" ]; then echo "not ready: /healthz on :${port} -> ${code}"; failed=1; fi
  done
  if ! grep -q '^artifacts: app ' "${VERIFY_RUN}/server.log" 2>/dev/null || ! grep -q '^artifacts: public app ' "${VERIFY_RUN}/server.log" 2>/dev/null; then
    echo "not ready: server.log lacks the 'artifacts: app' and 'artifacts: public app' lines"; failed=1
  fi
  if [ "${failed}" = 0 ]; then
    if ! list_err="$("${VERIFY_ROOT}/.claude/skills/verify/scripts/artifact" list 2>&1 >/dev/null)"; then
      echo "not ready: 'artifact list' failed: ${list_err}"; failed=1
    fi
    headers="$(curl -s -D - -o /dev/null --max-time 2 -X POST -H 'content-type: application/json' --data '{}' "${VERIFY_PUBLIC_APP_ORIGIN}/mcp" || true)"
    if ! printf '%s' "${headers}" | head -1 | grep -q ' 401' || ! printf '%s' "${headers}" | grep -qi '^www-authenticate: Bearer'; then
      echo "not ready: POST /mcp without a token did not answer 401 + WWW-Authenticate"; failed=1
    fi
  fi
  return "${failed}"
}

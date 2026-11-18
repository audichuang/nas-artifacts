#!/usr/bin/env bash
# Start the local verify instance (all four listeners) and wait until it is usable.
# Usage: scripts/launch.sh [--fresh]     --fresh wipes .verify/run/data first. VERIFY_PORT_BASE shifts the ports.
set -euo pipefail
requested_base="${VERIFY_PORT_BASE:-}"
verify_base_warn=0
. "$(dirname "$0")/env.sh"

fresh=0
if [ "${1:-}" = "--fresh" ]; then fresh=1; elif [ -n "${1:-}" ]; then echo "usage: launch.sh [--fresh]" >&2; exit 64; fi

pid="$(verify_server_pid)"
if [ -n "${pid}" ]; then
  if [ -n "${requested_base}" ] && [ "${requested_base}" != "${VERIFY_PORT_BASE}" ]; then
    echo "launch: this worktree's instance (pid ${pid}) runs on ${VERIFY_PORT_BASE}, not the requested ${requested_base}; run stop.sh first to move it" >&2
  fi
  if [ "${fresh}" = 1 ]; then echo "launch: already running (pid ${pid}); run stop.sh before --fresh" >&2; exit 1; fi
  if verify_ready; then echo "launch: already running and ready (pid ${pid})"; exit 0; fi
  echo "launch: pid ${pid} is running but not ready; see ${VERIFY_RUN}/server.log or run stop.sh" >&2; exit 1
fi
rm -f "${VERIFY_RUN}/server.pid"
# No live server of ours: a recorded block from a dead launch must not override the one asked for now.
if [ -n "${requested_base}" ] && [ "${requested_base}" != "${VERIFY_PORT_BASE}" ]; then
  rm -f "${VERIFY_RUN}/port-base"
  VERIFY_PORT_BASE="${requested_base}"
  . "$(dirname "$0")/env.sh"
fi

busy=0
for port in ${VERIFY_PORTS}; do
  holders="$(verify_port_pids "${port}")"
  if [ -n "${holders}" ]; then echo "launch: port ${port} is held by pid ${holders} (not ours; not killing it). Use another VERIFY_PORT_BASE." >&2; busy=1; fi
done
[ "${busy}" = 0 ] || exit 1
mkdir -p "${VERIFY_RUN}"

if [ "${fresh}" = 1 ]; then rm -rf "${VERIFY_RUN}/data"; fi
mkdir -p "${VERIFY_RUN}/data" "${VERIFY_RUN}/cli"
echo "${VERIFY_PORT_BASE}" > "${VERIFY_RUN}/port-base"
# One stamp per server start. A previous launch that died without stop.sh keeps its own stamp folder, and its
# server.log goes there before the new server overwrites it.
if [ -f "${VERIFY_RUN}/stamp" ]; then
  old_evidence="${VERIFY_DIR}/evidence/$(cat "${VERIFY_RUN}/stamp")"
  mkdir -p "${old_evidence}"
  [ -f "${VERIFY_RUN}/server.log" ] && cp "${VERIFY_RUN}/server.log" "${old_evidence}/server.log"
fi
stamp="$(date -u +%Y%m%d-%H%M%S)"
while [ -e "${VERIFY_DIR}/evidence/${stamp}" ]; do sleep 1; stamp="$(date -u +%Y%m%d-%H%M%S)"; done
echo "${stamp}" > "${VERIFY_RUN}/stamp"
evidence="${VERIFY_DIR}/evidence/${stamp}"
mkdir -p "${evidence}"

cd "${VERIFY_ROOT}"
# node directly, not `pnpm dev`: `node --watch` would make the recorded PID a watcher, not the server.
env -u NODE_ENV -u ONLYOFFICE_URL -u APP_ORIGIN -u CONTENT_ORIGIN -u ARTIFACTS_TOKEN_FILE \
  LISTEN_HOST=127.0.0.1 \
  DATA_DIR="${VERIFY_RUN}/data" \
  ARTIFACTS_TOKENS="${VERIFY_TOKEN}" \
  APP_PORT="${VERIFY_APP_PORT}" CONTENT_PORT="${VERIFY_CONTENT_PORT}" \
  PUBLIC_APP_PORT="${VERIFY_PUBLIC_APP_PORT}" PUBLIC_CONTENT_PORT="${VERIFY_PUBLIC_CONTENT_PORT}" \
  PUBLIC_APP_ORIGIN="${VERIFY_PUBLIC_APP_ORIGIN}" PUBLIC_CONTENT_ORIGIN="${VERIFY_PUBLIC_CONTENT_ORIGIN}" \
  nohup node src/server.ts > "${VERIFY_RUN}/server.log" 2>&1 < /dev/null &
echo $! > "${VERIFY_RUN}/server.pid"

for _ in $(seq 1 40); do
  sleep 0.25
  if ! kill -0 "$(cat "${VERIFY_RUN}/server.pid")" 2>/dev/null; then
    echo "launch: the server exited; log tail:" >&2; tail -n 30 "${VERIFY_RUN}/server.log" >&2; exit 1
  fi
  if verify_ready >/dev/null 2>&1; then
    echo "ready (pid $(cat "${VERIFY_RUN}/server.pid"))"
    echo "  tailnet app      ${VERIFY_APP_ORIGIN}"
    echo "  tailnet content  ${VERIFY_CONTENT_ORIGIN}"
    echo "  public app       ${VERIFY_PUBLIC_APP_ORIGIN}"
    echo "  public content   ${VERIFY_PUBLIC_CONTENT_ORIGIN}"
    echo "  evidence         ${evidence}"
    exit 0
  fi
done
echo "launch: not ready after 10 s:" >&2
verify_ready >&2 || true
tail -n 30 "${VERIFY_RUN}/server.log" >&2
exit 1

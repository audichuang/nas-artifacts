#!/usr/bin/env bash
# Stop the instance launch.sh started (pidfile only, never by name), keep the evidence, remove .verify/run.
# Usage: scripts/stop.sh
set -euo pipefail
. "$(dirname "$0")/env.sh"

[ -d "${VERIFY_RUN}" ] || { echo "stop: nothing to stop (${VERIFY_RUN} absent)"; exit 0; }
pid="$(verify_server_pid)"
if [ -n "${pid}" ]; then
  kill -TERM "${pid}"
  for _ in $(seq 1 40); do kill -0 "${pid}" 2>/dev/null || break; sleep 0.25; done
  if kill -0 "${pid}" 2>/dev/null; then echo "stop: pid ${pid} ignored SIGTERM for 10 s; sending SIGKILL" >&2; kill -KILL "${pid}"; sleep 0.5; fi
  echo "stopped pid ${pid}"
elif [ -f "${VERIFY_RUN}/server.pid" ]; then
  echo "stop: pidfile names pid $(cat "${VERIFY_RUN}/server.pid"), which is not a running src/server.ts; not killing anything"
fi

held=0
for port in ${VERIFY_PORTS}; do
  holders="$(verify_port_pids "${port}")"
  if [ -n "${holders}" ]; then echo "stop: port ${port} still held by pid ${holders} (not ours; left alone)" >&2; held=1; fi
done

if [ -f "${VERIFY_RUN}/stamp" ]; then
  evidence="${VERIFY_DIR}/evidence/$(cat "${VERIFY_RUN}/stamp")"
  mkdir -p "${evidence}"
  [ -f "${VERIFY_RUN}/server.log" ] && cp "${VERIFY_RUN}/server.log" "${evidence}/server.log"
  echo "evidence kept in ${evidence}"
fi
rm -rf "${VERIFY_RUN}"
[ "${held}" = 0 ] && echo "ports ${VERIFY_PORTS} free"
exit "${held}"

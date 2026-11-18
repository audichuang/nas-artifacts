#!/usr/bin/env bash
# Runs every feature journey in the map's order against the running instance and prints a summary.
# Usage: .claude/skills/verify/scripts/journeys/run-all.sh [feature …]     (default: all seven)
# Exit 0 only if every journey passed, 1 if any failed. The last line is the verdict, so it survives a `tail`,
# but a pipe replaces this exit code with the last command's: redirect to a file instead of piping.
set -uo pipefail
here="$(cd "$(dirname "$0")" && pwd)"
. "${here}/../env.sh" || exit 1
features="$*"
[ -n "${features}" ] || features="publish-and-view versions-and-conflict comments capabilities-live-page public-sharing owner-signin-and-mcp office-files"
status=0
summary=""
passed=0
total=0
failed=""
cd "${VERIFY_ROOT}"
for feature in ${features}; do
  echo "== ${feature}"
  total=$((total + 1))
  if node "${here}/${feature}.mjs"; then
    summary="${summary}pass  ${feature}"$'\n'; passed=$((passed + 1))
  else
    summary="${summary}FAIL  ${feature}"$'\n'; failed="${failed} ${feature}"; status=1
  fi
done
echo "== summary (evidence: ${VERIFY_DIR}/evidence/$(cat "${VERIFY_RUN}/stamp" 2>/dev/null))"
printf '%s' "${summary}"
if [ "${status}" = 0 ]; then
  echo "run-all: PASS ${passed}/${total} journeys (exit 0)"
else
  echo "run-all: FAIL ${passed}/${total} journeys passed; failed:${failed} (exit 1)"
fi
exit "${status}"

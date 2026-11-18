#!/usr/bin/env bash
# Runs every feature journey in the map's order against the running instance and prints a summary.
# Usage: .claude/skills/verify/scripts/journeys/run-all.sh [feature …]     (default: all seven)
set -uo pipefail
here="$(cd "$(dirname "$0")" && pwd)"
. "${here}/../env.sh" || exit 1
features="$*"
[ -n "${features}" ] || features="publish-and-view versions-and-conflict comments capabilities-live-page public-sharing owner-signin-and-mcp office-files"
status=0
summary=""
cd "${VERIFY_ROOT}"
for feature in ${features}; do
  echo "== ${feature}"
  if node "${here}/${feature}.mjs"; then summary="${summary}pass  ${feature}"$'\n'; else summary="${summary}FAIL  ${feature}"$'\n'; status=1; fi
done
echo "== summary (evidence: ${VERIFY_DIR}/evidence/$(cat "${VERIFY_RUN}/stamp" 2>/dev/null))"
printf '%s' "${summary}"
exit "${status}"

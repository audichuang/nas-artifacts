# nas-artifacts

- Node 24 runs the TypeScript in `src/` directly (type stripping): keep syntax erasable, import with `.ts` extensions. `pnpm tc` typechecks, `pnpm test` runs node:test (browser e2e needs Chromium, see README).
- The two-origin split is the security boundary: anything served from the content port must keep the `sandbox` CSP without `allow-same-origin`, and the app origin must never serve artifact files. Page capabilities go only through the viewer's `postMessage` bridge (`public/viewer.js` ↔ `content/bridge.js`), and the server re-checks the declared capability on every viewer call.
- `skill/nas-artifacts/scripts/artifact.mjs` is the only CLI; it must stay dependency-free so agents can run it from the skill folder.
- Deploy config lives in the nas-ops repo (`hosts/ttnas/artifacts/`), not here.

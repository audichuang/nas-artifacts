# Comments

Readers comment on a page from the viewer's panel. Agents read the comments with the CLI, reply, and resolve or reopen threads. The viewer updates live. An agent can block on `comments --wait` until someone writes.

## Sub-features

- `post`: `#toggle-comments` opens `#comments`. `#comment-body` plus `#comment-form button[type="submit"]` posts. The first post asks for an author name in a JS `prompt`, which is kept in localStorage after that.
- `read`: `artifact comments <id> [--json] [--open]`.
- `reply/resolve`: `artifact comment reply <id> <rootId> "<text>" --resolve --author NAME`. `comment resolve|reopen <id> <rootId>`.
- `live`: resolved threads leave the default list. `#show-resolved` shows them as `.thread.resolved` with their replies.
- `wait`: `artifact comments <id> --wait --timeout S` prints only new comments and exits 0, or exits 2 on timeout.

## How to get to it (user POV)

- Reader: on `/a/<id>`, clicks 留言, types, and sends. The first time, they give a name.
- Agent: `artifact comments <id> --open`, then `artifact comment reply …`. Or it waits with `--wait` after asking for feedback.

## Driving it with Playwright Chromium and the artifact CLI

Preconditions: instance ready.

```bash
node .claude/skills/verify/scripts/journeys/comments.mjs
```

- Publishes `page.md --new` and opens the viewer, setting a no-reload marker. It clicks `#toggle-comments`, registers `page.once('dialog', d => d.accept('Verifier'))` before clicking submit, and posts "first comment <nonce>". `#threads` shows it.
- `comments <id> --json` has the comment with author `Verifier` and `resolved: false`.
- `comment reply <id> <rootId> "agent reply <nonce>" --resolve --author Agent` prints `replied … and resolved it`. The thread leaves `#threads` live. After `#show-resolved` is checked, `.thread.resolved` shows both texts and the marker is intact.
- `comments <id> --open` no longer lists it. `comment reopen` brings it back.
- Starts `comments <id> --wait --timeout 20` as a child process and waits for its stderr `waiting up to`. Posts "second comment <nonce>" from the viewer. The child exits 0 printing only the new comment.

Evidence: `cli.log` (including the background command with its pid and exit), `01-comment-posted.png` … `04-second-comment.png`, `result.json`.

## Gotchas

- Always use `scripts/artifact`. The bare CLI defaults to production and reads the real token.
- Register the dialog handler before the click. Otherwise Playwright auto-dismisses the prompt and nothing is posted.
- Without `--author`, a CLI reply is signed `agent@<hostname>`. Pass `--author` so evidence carries no machine name.
- `reply --resolve` resolves the thread root, not the reply. Resolve and reopen take the root id.
- Post only after `--wait` prints `waiting up to`. A comment that lands before the CLI's first snapshot of the list counts as already seen, and the wait then times out with exit 2.

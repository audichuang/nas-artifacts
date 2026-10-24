export type DiffLine = { op: ' ' | '-' | '+'; text: string }
export type ShownLine = DiffLine | { op: '…'; skipped: number }

/**
 * Myers' line diff. The common head and tail are trimmed first, since most revisions touch a few places;
 * returns null when the middle needs more than `maxEdits` insertions and deletions (a rewrite, not an edit).
 */
export function diffLines(a: string[], b: string[], maxEdits = 2000): DiffLine[] | null {
  let head = 0
  while (head < a.length && head < b.length && a[head] === b[head]) head++
  let tail = 0
  while (tail < a.length - head && tail < b.length - head && a[a.length - 1 - tail] === b[b.length - 1 - tail]) tail++
  const x0 = a.slice(head, a.length - tail)
  const y0 = b.slice(head, b.length - tail)
  const middle = myers(x0, y0, maxEdits)
  if (!middle) {
    return null
  }
  const kept = (lines: string[]) => lines.map((text): DiffLine => ({ op: ' ', text }))
  return [...kept(a.slice(0, head)), ...middle, ...kept(a.slice(a.length - tail))]
}

function myers(a: string[], b: string[], maxEdits: number): DiffLine[] | null {
  const n = a.length
  const m = b.length
  const max = n + m
  const offset = max + 1
  const v = new Int32Array(2 * max + 3)
  // trace[d] is v before step d, kept only for k in -d..d (all step d reads), so memory grows with d², not d·(n+m).
  const trace: Int32Array[] = []
  let found = -1
  for (let d = 0; d <= Math.min(max, maxEdits) && found < 0; d++) {
    trace.push(v.slice(offset - d, offset + d + 1))
    for (let k = -d; k <= d; k += 2) {
      let x = k === -d || (k !== d && v[offset + k - 1] < v[offset + k + 1]) ? v[offset + k + 1] : v[offset + k - 1] + 1
      let y = x - k
      while (x < n && y < m && a[x] === b[y]) {
        x++
        y++
      }
      v[offset + k] = x
      if (x >= n && y >= m) {
        found = d
        break
      }
    }
  }
  if (found < 0) {
    return null
  }
  const ops: DiffLine[] = []
  let x = n
  let y = m
  for (let d = found; d > 0; d--) {
    const before = trace[d]
    const at = (k: number) => before[k + d]
    const k = x - y
    const down = k === -d || (k !== d && at(k - 1) < at(k + 1))
    const prevK = down ? k + 1 : k - 1
    const prevX = at(prevK)
    const prevY = prevX - prevK
    while (x > (down ? prevX : prevX + 1) && y > (down ? prevY + 1 : prevY)) {
      ops.push({ op: ' ', text: a[--x] })
      y--
    }
    if (down) {
      ops.push({ op: '+', text: b[--y] })
    } else {
      ops.push({ op: '-', text: a[--x] })
    }
  }
  while (x > 0 && y > 0) {
    ops.push({ op: ' ', text: a[--x] })
    y--
  }
  return ops.reverse()
}

/** Unchanged runs collapse to a count, keeping `context` lines around each change. */
export function hunks(lines: DiffLine[], context = 3): ShownLine[] {
  const near = new Uint8Array(lines.length)
  lines.forEach((line, index) => {
    if (line.op === ' ') return
    for (let i = Math.max(0, index - context); i <= Math.min(lines.length - 1, index + context); i++) near[i] = 1
  })
  const shown: ShownLine[] = []
  let skipped = 0
  lines.forEach((line, index) => {
    if (near[index]) {
      if (skipped) shown.push({ op: '…', skipped })
      skipped = 0
      shown.push(line)
    } else {
      skipped++
    }
  })
  if (skipped) shown.push({ op: '…', skipped })
  return shown
}

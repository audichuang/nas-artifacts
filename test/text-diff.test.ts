import assert from 'node:assert/strict'
import { describe, test } from 'node:test'
import { diffLines, hunks } from '../src/text-diff.ts'

const rebuild = (ops: ReturnType<typeof diffLines> & object, side: '-' | '+') =>
  ops.filter((op) => op.op === ' ' || op.op === side).map((op) => op.text)

describe('line diff', () => {
  test('marks inserted, removed and kept lines', () => {
    const ops = diffLines(['a', 'b', 'c', 'd'], ['a', 'c', 'x', 'd'])
    assert.deepEqual(ops, [
      { op: ' ', text: 'a' },
      { op: '-', text: 'b' },
      { op: ' ', text: 'c' },
      { op: '+', text: 'x' },
      { op: ' ', text: 'd' }
    ])
  })

  test('always rebuilds both sides, and is minimal enough on random edits', () => {
    let seed = 7
    const random = () => (seed = (seed * 1103515245 + 12345) % 2 ** 31) / 2 ** 31
    for (let round = 0; round < 300; round++) {
      const a = Array.from({ length: Math.floor(random() * 40) }, () => String.fromCharCode(97 + Math.floor(random() * 5)))
      const b = a.filter(() => random() > 0.2).flatMap((line) => (random() > 0.85 ? [line, 'n'] : [line]))
      const ops = diffLines(a, b)
      assert.ok(ops)
      assert.deepEqual(rebuild(ops, '-'), a)
      assert.deepEqual(rebuild(ops, '+'), b)
    }
  })

  test('gives up on rewrites too large to show', () => {
    const a = Array.from({ length: 3000 }, (_, i) => `a${i}`)
    const b = Array.from({ length: 3000 }, (_, i) => `b${i}`)
    assert.equal(diffLines(a, b, 500), null)
  })

  test('hunks keep three lines of context and say how many were skipped', () => {
    const a = Array.from({ length: 30 }, (_, i) => `line ${i}`)
    const b = [...a]
    b[15] = 'changed'
    const shown = hunks(diffLines(a, b)!)
    assert.deepEqual(shown[0], { op: '…', skipped: 12 })
    assert.deepEqual(shown.slice(1, 9).map((line) => line.op), [' ', ' ', ' ', '-', '+', ' ', ' ', ' '])
    assert.deepEqual(shown.at(-1), { op: '…', skipped: 11 })
  })
})

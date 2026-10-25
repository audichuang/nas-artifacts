/**
 * A small XML reader for Office parts: elements, attributes and text, in document order. Office XML
 * never needs a DTD, so a document carrying one is refused rather than expanded (no entity tricks).
 */
export type XmlNode = { name: string; attrs: Record<string, string>; children: XmlNode[]; text: string }

export class XmlError extends Error {}

/** How deep elements may nest. Office files stay far below it; a crafted one would overflow the stack. */
export const MAX_DEPTH = 1000

const ENTITIES: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" }

function decode(text: string): string {
  if (!text.includes('&')) return text
  return text.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|[a-z]+);/g, (whole, ref: string) => {
    if (ref[0] === '#') {
      const code = ref[1] === 'x' ? parseInt(ref.slice(2), 16) : parseInt(ref.slice(1), 10)
      return code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : ''
    }
    return ENTITIES[ref] ?? whole
  })
}

/** Just past the next `token`, or -1 when the source ends first. */
function skipPast(source: string, token: string, from: number): number {
  const found = source.indexOf(token, from)
  return found < 0 ? -1 : found + token.length
}

const isSpace = (char: string) => char === ' ' || char === '\t' || char === '\n' || char === '\r'

/**
 * The attributes inside a tag, scanned once from left to right: `name = "value"` or `'value'`. A token
 * with no `=` is skipped. (A regex retried from every position took quadratic time on a long run of them.)
 */
function readAttrs(inner: string, from: number, into: Record<string, string>): void {
  let at = from
  const end = inner.length
  while (at < end) {
    const char = inner[at]
    if (isSpace(char) || char === '=' || char === '/' || char === '>' || char === '"' || char === "'") {
      if (char === '"' || char === "'") {
        const close = inner.indexOf(char, at + 1)
        at = close < 0 ? end : close + 1
      } else at++
      continue
    }
    const nameStart = at
    while (at < end && !isSpace(inner[at]) && inner[at] !== '=' && inner[at] !== '/' && inner[at] !== '>' && inner[at] !== '"' && inner[at] !== "'") at++
    const name = inner.slice(nameStart, at)
    let next = at
    while (next < end && isSpace(inner[next])) next++
    if (inner[next] !== '=') continue
    next++
    while (next < end && isSpace(inner[next])) next++
    const quote = inner[next]
    if (quote !== '"' && quote !== "'") {
      at = next
      continue
    }
    const close = inner.indexOf(quote, next + 1)
    if (close < 0) return
    into[name] = decode(inner.slice(next + 1, close))
    at = close + 1
  }
}

/** The attributes of one start tag given as raw text (`<row r='3' ht = "20">`), read by the same rules as parseXml. */
export function tagAttrs(tag: string): Record<string, string> {
  const attrs: Record<string, string> = {}
  const nameEnd = tag.search(/[\s/>]/)
  if (nameEnd > 0) readAttrs(tag, nameEnd, attrs)
  return attrs
}

/**
 * `cut` reads the first part of a document whose end was not read (a part too large to inflate whole):
 * elements still open at the end are closed there instead of failing.
 */
export function parseXml(source: string, maxNodes = 2_000_000, cut = false): XmlNode {
  const root: XmlNode = { name: '#root', attrs: {}, children: [], text: '' }
  const stack: XmlNode[] = [root]
  let at = 0
  let nodes = 0
  while (at < source.length) {
    const open = source.indexOf('<', at)
    if (open < 0) break
    if (open > at) stack[stack.length - 1].text += decode(source.slice(at, open))
    // Markup the source ends inside: an error, or in a document read only in part, where it stops.
    if (source.startsWith('<?', open) || source.startsWith('<!--', open)) {
      at = skipPast(source, source[open + 1] === '?' ? '?>' : '-->', open)
      if (at < 0 && cut) break
      if (at < 0) throw new XmlError('XML 標記沒有結束')
    } else if (source.startsWith('<![CDATA[', open)) {
      const close = source.indexOf(']]>', open)
      if (close < 0 && cut) break
      if (close < 0) throw new XmlError('XML 的 CDATA 區段沒有結束')
      stack[stack.length - 1].text += source.slice(open + 9, close)
      at = close + 3
    } else if (source.startsWith('<!', open)) {
      throw new XmlError('不接受含 DTD 的 XML')
    } else if (source[open + 1] === '/') {
      const past = skipPast(source, '>', open)
      if (past < 0 && cut) break
      if (past < 0) throw new XmlError('XML 標記沒有結束')
      const close = past - 1
      const name = source.slice(open + 2, close).trim()
      if (stack.length < 2 || stack[stack.length - 1].name !== name) throw new XmlError(`XML 多出了 </${name}>`)
      stack.pop()
      at = close + 1
    } else {
      // A '>' inside a quoted attribute value is legal, so find the tag's end outside quotes.
      let close = open + 1
      let quote = ''
      for (; close < source.length; close++) {
        const char = source[close]
        if (quote) {
          if (char === quote) quote = ''
        } else if (char === '"' || char === "'") quote = char
        else if (char === '>') break
      }
      if (close >= source.length && cut) break
      if (close >= source.length) throw new XmlError('XML 標籤沒有結束')
      const selfClosing = source[close - 1] === '/'
      const inner = source.slice(open + 1, selfClosing ? close - 1 : close)
      const nameEnd = inner.search(/[\s/]|$/)
      const node: XmlNode = { name: inner.slice(0, nameEnd), attrs: {}, children: [], text: '' }
      readAttrs(inner, nameEnd, node.attrs)
      if (++nodes > maxNodes) throw new XmlError('XML 內容太多')
      stack[stack.length - 1].children.push(node)
      if (!selfClosing) {
        // Renderers walk the tree recursively; past this depth they would run out of stack.
        if (stack.length > MAX_DEPTH) throw new XmlError(`XML 巢狀超過 ${MAX_DEPTH} 層`)
        stack.push(node)
      }
      at = close + 1
    }
  }
  if (stack.length > 1 && !cut) throw new XmlError(`XML 的 <${stack[stack.length - 1].name}> 沒有結束`)
  const element = root.children[0]
  if (!element) throw new XmlError('XML 是空的')
  return element
}

/** Local name without its namespace prefix: `a:solidFill` → `solidFill`. */
export const local = (name: string): string => name.slice(name.indexOf(':') + 1)

export const child = (node: XmlNode | null | undefined, name: string): XmlNode | null =>
  node?.children.find((candidate) => local(candidate.name) === name) ?? null

export const children = (node: XmlNode | null | undefined, name: string): XmlNode[] =>
  node ? node.children.filter((candidate) => local(candidate.name) === name) : []

/** Follows a path of local names: `path(sp, 'spPr', 'xfrm', 'off')`. */
export function path(node: XmlNode | null | undefined, ...names: string[]): XmlNode | null {
  let current: XmlNode | null = node ?? null
  for (const name of names) current = child(current, name)
  return current
}

/** Attribute by local name, so `r:embed` and `embed` both match. */
export function attr(node: XmlNode | null | undefined, name: string): string | undefined {
  if (!node) return undefined
  if (name in node.attrs) return node.attrs[name]
  for (const [key, value] of Object.entries(node.attrs)) if (local(key) === name) return value
  return undefined
}

/** A relationship attribute (`r:embed`, `r:id`), whatever prefix the part bound to it. */
export function relAttr(node: XmlNode | null | undefined, name: string): string | undefined {
  if (!node) return undefined
  for (const [key, value] of Object.entries(node.attrs)) if (key.includes(':') && local(key) === name) return value
  return undefined
}

export function num(node: XmlNode | null | undefined, name: string): number | undefined {
  const value = attr(node, name)
  if (value === undefined) return undefined
  const parsed = Number(value)
  return Number.isFinite(parsed) ? parsed : undefined
}

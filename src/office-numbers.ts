// Numbers written the ways Office lists, notes and number formats spell them: Roman, letters, Chinese counting,
// kana and English words. Shared by the Word, PowerPoint and Excel readers so each spells a number the same way.

export const CN_DIGITS = '〇一二三四五六七八九'
export const CN_LEGAL = '零壹貳參肆伍陸柒捌玖'

export function roman(n: number): string {
  if (n <= 0 || n >= 4000) return String(n)
  const table: [number, string][] = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']]
  let out = ''
  for (const [value, glyphs] of table) while (n >= value) {
    out += glyphs
    n -= value
  }
  return out
}

/** One symbol per number from `set`; past the end Word repeats it (A … Z, AA … ZZ); past a few dozen repeats the label is a number. */
export function repeating(n: number, set: readonly string[]): string {
  if (n <= 0 || n > set.length * 40) return String(n)
  return set[(n - 1) % set.length].repeat(Math.floor((n - 1) / set.length) + 1)
}

export const letters = (n: number) => repeating(n, LATIN)
const LATIN = [...'ABCDEFGHIJKLMNOPQRSTUVWXYZ']

/** 1234 → 一千二百三十四 (digits and units from the caller: everyday or legal characters). */
export function chineseCount(n: number, digits: string, units: string[], big: string, lowTen: boolean): string {
  if (n === 0) return digits[0]
  if (n < 0 || n >= 100_000_000) return String(n)
  if (n >= 10000) {
    const high = Math.floor(n / 10000)
    const low = n % 10000
    return chineseCount(high, digits, units, big, lowTen) + big + (low ? (low < 1000 ? digits[0] : '') + chineseCount(low, digits, units, big, false) : '')
  }
  let out = ''
  let zero = false
  for (let place = 3; place >= 0; place--) {
    const digit = Math.floor(n / 10 ** place) % 10
    if (digit === 0) {
      if (out) zero = true
      continue
    }
    if (zero) out += digits[0]
    zero = false
    out += place === 1 && digit === 1 && !out && lowTen ? units[1] : digits[digit] + units[place]
  }
  return out
}

export const chineseCounting = (n: number) => chineseCount(n, CN_DIGITS, ['', '十', '百', '千'], '萬', true)
export const chineseLegal = (n: number) => chineseCount(n, CN_LEGAL, ['', '拾', '佰', '仟'], '萬', false)

// Japanese kana in aiueo (gojūon) and iroha order; the half-width forms are what Word's plain `aiueo`/`iroha` draw.
export const AIUEO = [...'アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン']
export const AIUEO_HALF = [...'ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜｦﾝ']
export const IROHA = [...'イロハニホヘトチリヌルヲワカヨタレソツネナラムウヰノオクヤマケフコエテアサキユメミシヱヒモセス']
// ヰ and ヱ have no half-width form; Word keeps them full width.
export const IROHA_HALF = [...'ｲﾛﾊﾆﾎﾍﾄﾁﾘﾇﾙｦﾜｶﾖﾀﾚｿﾂﾈﾅﾗﾑｳヰﾉｵｸﾔﾏｹﾌｺｴﾃｱｻｷﾕﾒﾐｼヱﾋﾓｾｽ']
export const GANADA = [...'가나다라마바사아자차카타파하']
export const CHOSUNG = [...'ㄱㄴㄷㄹㅁㅂㅅㅇㅈㅊㅋㅌㅍㅎ']
export const RUSSIAN = [...'абвгдежзиклмнопрстуфхцчшщыэюя']

const ONES = ['zero', 'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine', 'ten', 'eleven', 'twelve', 'thirteen', 'fourteen', 'fifteen', 'sixteen', 'seventeen', 'eighteen', 'nineteen']
const TENS = ['', '', 'twenty', 'thirty', 'forty', 'fifty', 'sixty', 'seventy', 'eighty', 'ninety']

function words(n: number): string {
  if (n < 20) return ONES[n]
  if (n < 100) return TENS[Math.floor(n / 10)] + (n % 10 ? `-${ONES[n % 10]}` : '')
  if (n < 1000) return `${ONES[Math.floor(n / 100)]} hundred${n % 100 ? ` ${words(n % 100)}` : ''}`
  return `${words(Math.floor(n / 1000))} thousand${n % 1000 ? ` ${words(n % 1000)}` : ''}`
}

const capital = (text: string) => text.charAt(0).toUpperCase() + text.slice(1)

/** 21 → "Twenty-one", as Word's cardinalText list format writes it; past 999,999 the number itself. */
export function cardinalText(n: number): string {
  return n < 0 || n > 999_999 || !Number.isInteger(n) ? String(n) : capital(words(n))
}

const ORDINAL_WORDS: Record<string, string> = { one: 'first', two: 'second', three: 'third', five: 'fifth', eight: 'eighth', nine: 'ninth', twelve: 'twelfth' }

/** 23 → "Twenty-third", as Word's ordinalText list format writes it. */
export function ordinalText(n: number): string {
  if (n < 0 || n > 999_999 || !Number.isInteger(n)) return String(n)
  const text = words(n).replace(/[a-z]+$/, (last) => ORDINAL_WORDS[last] ?? (last.endsWith('y') ? `${last.slice(0, -1)}ieth` : `${last}th`))
  return capital(text)
}

/** 2 → "2nd". */
export function ordinal(n: number): string {
  const tens = Math.abs(n) % 100
  const suffix = tens >= 11 && tens <= 13 ? 'th' : ['th', 'st', 'nd', 'rd'][Math.abs(n) % 10] ?? 'th'
  return `${n}${suffix}`
}

/** 1 → ①, up to ㊿; past fifty the number itself. */
export function circled(n: number): string {
  if (n >= 1 && n <= 20) return String.fromCodePoint(0x2460 + n - 1)
  if (n >= 21 && n <= 35) return String.fromCodePoint(0x3251 + n - 21)
  if (n >= 36 && n <= 50) return String.fromCodePoint(0x32b1 + n - 36)
  return String(n)
}

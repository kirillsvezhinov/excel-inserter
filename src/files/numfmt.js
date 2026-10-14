// Excel number-format engine: renders a cell value the way Excel displays it,
// with the viewer's locale for decimal/group separators, month and day names.

/** Built-in formats (ECMA-376 §18.8.30). 14 and 22 follow the system locale in Excel. */
const BUILTIN = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
  5: '#,##0 _₽;-#,##0 _₽', 6: '#,##0 _₽;[Red]-#,##0 _₽', 7: '#,##0.00 _₽;-#,##0.00 _₽', 8: '#,##0.00 _₽;[Red]-#,##0.00 _₽',
  9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
  14: 'SHORT_DATE', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM',
  20: 'h:mm', 21: 'h:mm:ss', 22: 'SHORT_DATE h:mm',
  37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)',
  41: '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)', 42: '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)', 44: '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
  45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mm:ss.0', 48: '##0.0E+0', 49: '@'
}
// East-Asian locale date ids map to the short date.
for (const id of [27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 50, 51, 52, 53, 54, 55, 56, 57, 58]) BUILTIN[id] = 'SHORT_DATE'

export function builtinFormat(id) {
  return BUILTIN[id] ?? 'General'
}

const COLOR_NAMES = { black: '#000000', blue: '#0000ff', cyan: '#00ffff', green: '#00ff00', magenta: '#ff00ff', red: '#ff0000', white: '#ffffff', yellow: '#ffff00' }

const BOOL_NAMES = { ru: ['ИСТИНА', 'ЛОЖЬ'], uk: ['ІСТИНА', 'ХИБНІСТЬ'], de: ['WAHR', 'FALSCH'], fr: ['VRAI', 'FAUX'], es: ['VERDADERO', 'FALSO'], it: ['VERO', 'FALSO'], pt: ['VERDADEIRO', 'FALSO'] }

const localeCache = new Map()
/** Separators and names for a locale. */
export function localeInfo(locale) {
  const key = locale || ''
  let info = localeCache.get(key)
  if (info) return info
  let decimal = '.'
  let group = ','
  try {
    for (const part of new Intl.NumberFormat(locale || undefined).formatToParts(1234567.5)) {
      if (part.type === 'decimal') decimal = part.value
      if (part.type === 'group') group = part.value
    }
  } catch { /* unknown locale */ }
  const fmt = (opts) => {
    try { return new Intl.DateTimeFormat(locale || undefined, { timeZone: 'UTC', ...opts }) } catch { return new Intl.DateTimeFormat('en-US', { timeZone: 'UTC', ...opts }) }
  }
  const months = (style) => Array.from({ length: 12 }, (_, m) => fmt({ month: style }).format(Date.UTC(2021, m, 15)))
  const days = (style) => Array.from({ length: 7 }, (_, d) => fmt({ weekday: style }).format(Date.UTC(2021, 1, 7 + d))) // 2021-02-07 is a Sunday
  const lang = (locale || (typeof navigator !== 'undefined' ? navigator.language : 'en') || 'en').toLowerCase().split('-')[0]
  info = {
    decimal,
    group,
    monthsLong: months('long'),
    monthsShort: months('short').map((s) => s.replace(/\.$/, '')),
    daysLong: days('long'),
    daysShort: days('short').map((s) => s.replace(/\.$/, '')),
    shortDate: fmt({ year: 'numeric', month: '2-digit', day: '2-digit' }),
    bools: BOOL_NAMES[lang] || ['TRUE', 'FALSE']
  }
  localeCache.set(key, info)
  return info
}

/** Excel serial date → UTC Date (1900 system incl. the fake 1900-02-29, or 1904 system). */
export function serialToDate(serial, date1904 = false) {
  let days = Math.floor(serial)
  const ms = Math.round((serial - days) * 86400000)
  if (date1904) return new Date(Date.UTC(1904, 0, 1) + days * 86400000 + ms)
  if (days < 60) days += 1 // before the phantom leap day
  return new Date(Date.UTC(1899, 11, 30) + days * 86400000 + ms)
}

/** Split a format code into sections on `;`, respecting quotes, escapes and brackets. */
function splitSections(code) {
  const out = []
  let cur = ''
  for (let i = 0; i < code.length; i++) {
    const ch = code[i]
    if (ch === '"') {
      const end = code.indexOf('"', i + 1)
      const stop = end < 0 ? code.length : end
      cur += code.slice(i, stop + 1)
      i = stop
    } else if (ch === '\\' || ch === '_' || ch === '*') {
      cur += ch + (code[i + 1] ?? '')
      i++
    } else if (ch === '[') {
      const end = code.indexOf(']', i)
      const stop = end < 0 ? code.length : end
      cur += code.slice(i, stop + 1)
      i = stop
    } else if (ch === ';') {
      out.push(cur)
      cur = ''
    } else cur += ch
  }
  out.push(cur)
  return out
}

/** Tokenize one section. Literal tokens keep their text; placeholders keep their char. */
function tokenize(section) {
  const tokens = []
  let color = null
  let condition = null
  let elapsed = false
  const lit = (text) => {
    const last = tokens[tokens.length - 1]
    if (last && last.t === 'lit') last.v += text
    else tokens.push({ t: 'lit', v: text })
  }
  for (let i = 0; i < section.length; i++) {
    const ch = section[i]
    if (ch === '"') {
      const end = section.indexOf('"', i + 1)
      const stop = end < 0 ? section.length : end
      lit(section.slice(i + 1, stop))
      i = stop
    } else if (ch === '\\') { lit(section[i + 1] ?? ''); i++ }
    else if (ch === '_') { lit(' '); i++ }
    else if (ch === '*') { i++ } // fill character: no fixed width here
    else if (ch === '[') {
      const end = section.indexOf(']', i)
      const inner = section.slice(i + 1, end < 0 ? section.length : end)
      i = end < 0 ? section.length : end
      let m
      if ((m = /^\$([^-\]]*)(?:-[0-9a-f]+)?$/i.exec(inner))) lit(m[1])
      else if (/^(h+|m+|s+)$/i.test(inner)) { tokens.push({ t: 'date', v: inner.toLowerCase(), elapsed: true }); elapsed = true }
      else if ((m = /^(<=|>=|<>|<|>|=)\s*(-?[\d.]+)$/.exec(inner))) condition = { op: m[1], n: parseFloat(m[2]) }
      else if (COLOR_NAMES[inner.toLowerCase()]) color = COLOR_NAMES[inner.toLowerCase()]
      // [ColorN], [DBNum], locale ids etc. are ignored
    } else if (/[yYmMdDhHsS]/.test(ch)) {
      let j = i
      while (j < section.length && section[j].toLowerCase() === ch.toLowerCase()) j++
      tokens.push({ t: 'date', v: section.slice(i, j).toLowerCase() })
      i = j - 1
    } else if (/^(AM\/PM|A\/P)/i.test(section.slice(i))) {
      const v = /^AM\/PM/i.test(section.slice(i)) ? 'AM/PM' : 'A/P'
      tokens.push({ t: 'ampm', v })
      i += v.length - 1
    } else if (ch === 'E' || ch === 'e') {
      if (section[i + 1] === '+' || section[i + 1] === '-') { tokens.push({ t: 'exp', v: section[i + 1] }); i++ }
      else lit(ch)
    } else if ('0#?'.includes(ch)) tokens.push({ t: 'ph', v: ch })
    else if (ch === '.') tokens.push({ t: 'dot' })
    else if (ch === ',') tokens.push({ t: 'comma' })
    else if (ch === '%') tokens.push({ t: 'pct' })
    else if (ch === '/') tokens.push({ t: 'slash' })
    else if (ch === '@') tokens.push({ t: 'text' })
    else lit(ch)
  }
  // A '.' followed by zeros right after seconds is fractional seconds, not a decimal point.
  const isDate = tokens.some((t) => t.t === 'date' || t.t === 'ampm')
  return { tokens, color, condition, isDate, elapsed }
}

const sectionCache = new Map()
function parseFormat(code) {
  let parsed = sectionCache.get(code)
  if (!parsed) {
    parsed = splitSections(code).map(tokenize)
    if (sectionCache.size > 500) sectionCache.clear()
    sectionCache.set(code, parsed)
  }
  return parsed
}

function roundTo(n, d) {
  if (!isFinite(n)) return n
  return Number(Math.round(Number(n + 'e' + d)) + 'e-' + d)
}

function groupDigits(int, sep) {
  return int.replace(/\B(?=(\d{3})+(?!\d))/g, sep)
}

/** General format: up to 11 characters, switching to scientific for very large/small numbers. */
function formatGeneral(n, loc) {
  if (n === 0) return '0'
  const abs = Math.abs(n)
  let s
  if (abs >= 1e11 || abs < 1e-9) {
    s = n.toExponential(5).replace(/\.?0+e/, 'e').replace(/e([+-])(\d)$/, 'e$10$2').toUpperCase()
  } else if (Number.isInteger(n)) s = String(n)
  else {
    const intDigits = Math.max(1, Math.floor(Math.log10(abs)) + 1)
    s = String(roundTo(n, Math.max(0, 10 - intDigits - (n < 0 ? 1 : 0))))
    if (s.includes('e')) s = n.toPrecision(10).replace(/\.?0+$/, '')
  }
  return s.replace('.', loc.decimal)
}

function formatFraction(n, tokens, loc) {
  // "# ?/?" style: integer part + best fraction with as many denominator digits as placeholders.
  const slash = tokens.findIndex((t) => t.t === 'slash')
  const denomPh = tokens.slice(slash + 1).filter((t) => t.t === 'ph').length || 1
  const fixedDenom = tokens.slice(slash + 1).map((t) => (t.t === 'lit' ? t.v : '')).join('').trim()
  const hasInt = tokens.slice(0, slash).some((t) => t.t === 'lit' && /\s/.test(t.v))
  const abs = Math.abs(n)
  const whole = hasInt ? Math.floor(abs) : 0
  const frac = abs - whole
  let best = [0, 1]
  if (/^\d+$/.test(fixedDenom)) best = [Math.round(frac * +fixedDenom), +fixedDenom]
  else {
    const maxD = 10 ** denomPh - 1
    let err = Infinity
    for (let d = 1; d <= maxD; d++) {
      const num = Math.round(frac * d)
      const e = Math.abs(frac - num / d)
      if (e < err - 1e-12) { err = e; best = [num, d] }
    }
  }
  const sign = n < 0 ? '-' : ''
  if (best[0] === 0) return sign + String(whole)
  if (best[0] === best[1]) return sign + String(whole + 1)
  return sign + (hasInt && whole ? `${whole} ` : '') + `${best[0]}/${best[1]}`.replace('.', loc.decimal)
}

function formatNumberSection(n, sec, loc, withSign) {
  const { tokens } = sec
  if (tokens.some((t) => t.t === 'slash') && tokens.some((t) => t.t === 'ph')) return formatFraction(n, tokens, loc)
  const expIdx = tokens.findIndex((t) => t.t === 'exp')
  const numTokens = expIdx >= 0 ? tokens.slice(0, expIdx) : tokens
  const firstPh = numTokens.findIndex((t) => t.t === 'ph')
  let lastPh = -1
  numTokens.forEach((t, i) => { if (t.t === 'ph') lastPh = i })
  if (firstPh < 0) {
    // No digits at all (e.g. `"-"` zero section): just literals.
    return tokens.map((t) => (t.t === 'lit' ? t.v : t.t === 'pct' ? '%' : '')).join('')
  }

  let value = Math.abs(n)
  const pctCount = tokens.filter((t) => t.t === 'pct').length
  value *= 100 ** pctCount
  // Trailing commas after the last digit placeholder scale by 1000 each.
  let scale = 0
  for (let i = lastPh + 1; i < numTokens.length && numTokens[i].t === 'comma'; i++) scale++
  value /= 1000 ** scale

  const dotIdx = numTokens.findIndex((t, i) => t.t === 'dot' && i > firstPh - 1 && i <= lastPh + 1)
  const intTokens = numTokens.slice(0, dotIdx >= 0 ? dotIdx : lastPh + 1)
  const fracTokens = dotIdx >= 0 ? numTokens.slice(dotIdx + 1, lastPh + 1) : []
  const grouping = intTokens.some((t, i) => t.t === 'comma' && i > firstPh && intTokens.slice(i + 1).some((x) => x.t === 'ph'))
  const fracPh = fracTokens.filter((t) => t.t === 'ph')

  let exponent = 0
  if (expIdx >= 0) {
    const intPh = intTokens.filter((t) => t.t === 'ph').length || 1
    if (value !== 0) {
      exponent = Math.floor(Math.log10(value))
      // Engineering-style formats (##0.0E+0) step the exponent by the integer placeholder count.
      exponent = intPh > 1 && intTokens.some((t) => t.v === '#') ? Math.floor(exponent / intPh) * intPh : exponent - (intPh - 1)
      value /= 10 ** exponent
    }
  }

  const rounded = roundTo(value, fracPh.length)
  let [intStr, fracStr = ''] = (fracPh.length ? rounded.toFixed(fracPh.length) : String(Math.round(rounded))).split('.')
  if (intStr === '0') intStr = ''

  // Integer part: fill placeholders right to left; the leftmost one takes any extra digits.
  const intOut = []
  let digits = intStr
  const intPhIdx = intTokens.map((t, i) => (t.t === 'ph' ? i : -1)).filter((i) => i >= 0)
  for (let k = intTokens.length - 1; k >= 0; k--) {
    const t = intTokens[k]
    if (t.t === 'ph') {
      const isLeftmost = k === intPhIdx[0]
      let piece
      if (isLeftmost) { piece = digits || (t.v === '0' ? '0' : t.v === '?' ? ' ' : ''); digits = '' }
      else if (digits) { piece = digits.slice(-1); digits = digits.slice(0, -1) }
      else piece = t.v === '0' ? '0' : t.v === '?' ? ' ' : ''
      intOut.unshift(piece)
    } else if (t.t === 'lit') intOut.unshift(t.v)
    else if (t.t === 'pct') intOut.unshift('%')
    else if (t.t === 'dot' || t.t === 'comma' || t.t === 'slash') intOut.unshift('')
    else if (t.t === 'text') intOut.unshift('')
  }
  let intText = intOut.join('')
  if (grouping) {
    // Group the digit run (literals inside a grouped integer part are rare).
    intText = intText.replace(/\d+/, (d) => groupDigits(d, loc.group))
  }

  // Fraction part: fill left to right; '#' drops trailing zeros, '?' pads with spaces.
  let fracOut = ''
  if (dotIdx >= 0) {
    // Trim trailing zeros where the placeholder is optional.
    const chars = fracPh.map((t, i) => ({ t, d: fracStr[i] ?? '0' }))
    for (let i = chars.length - 1; i >= 0 && chars[i].d === '0' && chars[i].t.v !== '0'; i--) chars[i].d = chars[i].t.v === '?' ? ' ' : ''
    let pi = 0
    for (const t of fracTokens) {
      if (t.t === 'ph') fracOut += chars[pi++].d
      else if (t.t === 'lit') fracOut += t.v
      else if (t.t === 'pct') fracOut += '%'
    }
  }
  // Excel keeps the decimal point even when every fraction digit is optional ("#.##" → "5.").
  let text = intText + (dotIdx >= 0 ? loc.decimal + fracOut : '')

  if (expIdx >= 0) {
    const expTokens = tokens.slice(expIdx + 1)
    const expPh = expTokens.filter((t) => t.t === 'ph').length
    const sign = exponent < 0 ? '-' : tokens[expIdx].v === '+' ? '+' : ''
    text += 'E' + sign + String(Math.abs(exponent)).padStart(expPh, '0')
    text += expTokens.filter((t) => t.t === 'lit').map((t) => t.v).join('')
  }
  // Literals/percent after the last placeholder.
  const tail = numTokens.slice(lastPh + 1).map((t) => (t.t === 'lit' ? t.v : t.t === 'pct' ? '%' : '')).join('')
  if (expIdx < 0) text += tail
  return (withSign && n < 0 && Number(rounded) !== 0 ? '-' : '') + text
}

function formatDateSection(serial, sec, loc, date1904) {
  const { tokens } = sec
  const d = serialToDate(serial, date1904)
  const hasAmPm = tokens.some((t) => t.t === 'ampm')
  // Fractional seconds: a `.0`/`.00` right after seconds.
  const out = []
  let ms = d.getUTCMilliseconds()
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i]
    if (t.t === 'date') {
      const v = t.v
      const prevTime = (() => { for (let j = i - 1; j >= 0; j--) if (tokens[j].t === 'date') return tokens[j].v[0] === 'h'; return false })()
      const nextSec = (() => { for (let j = i + 1; j < tokens.length; j++) if (tokens[j].t === 'date') return tokens[j].v[0] === 's'; return false })()
      if (t.elapsed) {
        const total = serial * 24
        out.push(String(Math.floor(v[0] === 'h' ? total : v[0] === 'm' ? total * 60 : total * 3600)).padStart(v.length, '0'))
      } else if (v[0] === 'y') out.push(v.length <= 2 ? String(d.getUTCFullYear() % 100).padStart(2, '0') : String(d.getUTCFullYear()))
      else if (v[0] === 'm' && v.length <= 2 && (prevTime || nextSec)) out.push(String(d.getUTCMinutes()).padStart(v.length, '0'))
      else if (v[0] === 'm') {
        const m = d.getUTCMonth()
        out.push(v.length === 1 ? String(m + 1) : v.length === 2 ? String(m + 1).padStart(2, '0') : v.length === 3 ? loc.monthsShort[m] : v.length === 5 ? loc.monthsLong[m][0] : loc.monthsLong[m])
      } else if (v[0] === 'd') {
        const wd = d.getUTCDay()
        out.push(v.length === 1 ? String(d.getUTCDate()) : v.length === 2 ? String(d.getUTCDate()).padStart(2, '0') : v.length === 3 ? loc.daysShort[wd] : loc.daysLong[wd])
      } else if (v[0] === 'h') {
        let h = d.getUTCHours()
        if (hasAmPm) h = h % 12 || 12
        out.push(String(h).padStart(v.length >= 2 ? 2 : 1, '0'))
      } else if (v[0] === 's') {
        let s = d.getUTCSeconds()
        const frac = tokens[i + 1]?.t === 'dot' ? tokens.slice(i + 2).findIndex((x) => !(x.t === 'ph' && x.v === '0')) : -1
        const places = tokens[i + 1]?.t === 'dot' ? (frac < 0 ? tokens.length - i - 2 : frac) : 0
        if (!places && ms >= 500) s = (s + 1) % 60
        out.push(String(s).padStart(v.length >= 2 ? 2 : 1, '0'))
        if (places) {
          out.push(loc.decimal + String(Math.round(ms / 10 ** (3 - places))).padStart(places, '0'))
          i += 1 + places
          ms = 0
        }
      }
    } else if (t.t === 'ampm') {
      const pm = d.getUTCHours() >= 12
      out.push(t.v === 'AM/PM' ? (pm ? 'PM' : 'AM') : (pm ? 'P' : 'A'))
    } else if (t.t === 'lit') out.push(t.v)
    else if (t.t === 'dot') out.push('.')
    else if (t.t === 'comma') out.push(',')
    else if (t.t === 'slash') out.push('/')
    else if (t.t === 'ph') out.push(t.v === '0' ? '0' : '')
  }
  return out.join('')
}

function matches(cond, n) {
  switch (cond.op) {
    case '<': return n < cond.n
    case '<=': return n <= cond.n
    case '>': return n > cond.n
    case '>=': return n >= cond.n
    case '=': return n === cond.n
    case '<>': return n !== cond.n
  }
  return false
}

/**
 * Format a cell value with an Excel number format code.
 * @param {number|string|boolean|null} value
 * @param {string} code  format code, e.g. `#,##0.00 "₽";[Red]-#,##0.00 "₽"`
 * @returns {{ text: string, color: string|null, isDate: boolean }}
 */
export function formatValue(value, code, { locale, date1904 = false } = {}) {
  const loc = localeInfo(locale)
  if (value == null || value === '') return { text: '', color: null, isDate: false }
  if (typeof value === 'boolean') return { text: loc.bools[value ? 0 : 1], color: null, isDate: false }
  code = code || 'General'

  if (code.includes('SHORT_DATE') && typeof value === 'number') {
    const d = serialToDate(value, date1904)
    let text = loc.shortDate.format(d)
    if (/h:mm/.test(code)) text += ' ' + String(d.getUTCHours()) + ':' + String(d.getUTCMinutes()).padStart(2, '0')
    return { text, color: null, isDate: true }
  }

  const sections = parseFormat(code)
  if (typeof value === 'string') {
    const sec = sections.length >= 4 ? sections[3] : sections.find((s) => s.tokens.some((t) => t.t === 'text'))
    if (!sec) return { text: value, color: null, isDate: false }
    return {
      text: sec.tokens.map((t) => (t.t === 'text' ? value : t.t === 'lit' ? t.v : '')).join(''),
      color: sec.color,
      isDate: false
    }
  }
  if (typeof value !== 'number' || !isFinite(value)) return { text: String(value), color: null, isDate: false }

  // Pick the section: explicit conditions first, else positive;negative;zero.
  let sec
  let withSign = true
  const numeric = sections.filter((s) => !s.tokens.some((t) => t.t === 'text') || s.tokens.some((t) => t.t === 'ph'))
  if (sections.some((s) => s.condition)) {
    sec = sections.find((s) => s.condition && matches(s.condition, value))
    if (!sec) sec = sections.find((s) => !s.condition) || sections[sections.length - 1]
    withSign = !sec.condition || sec.condition.n >= 0 ? value < 0 && !sec.condition : false
    if (sec.condition && value < 0 && sec.condition.op.includes('<')) withSign = false
  } else if (sections.length === 1 || numeric.length === 1) sec = sections[0]
  else if (value > 0 || (value === 0 && sections.length < 3)) sec = sections[0]
  else if (value < 0) { sec = sections[1]; withSign = false }
  else sec = sections[2] || sections[0]

  const isGeneral = sec.tokens.length === 1 && sec.tokens[0].t === 'lit' && /^general$/i.test(sec.tokens[0].v.trim())
  const generalWithText = !isGeneral && sec.tokens.some((t) => t.t === 'lit' && /general/i.test(t.v))
  if (isGeneral || generalWithText || !sec.tokens.length) {
    const g = formatGeneral(withSign ? value : Math.abs(value), loc)
    const text = generalWithText ? sec.tokens.map((t) => (t.t === 'lit' ? t.v.replace(/general/i, g) : '')).join('') : g
    return { text, color: sec.color, isDate: false }
  }
  if (sec.isDate) {
    if (value < 0) return { text: '#'.repeat(8), color: null, isDate: true }
    return { text: formatDateSection(value, sec, loc, date1904), color: sec.color, isDate: true }
  }
  return { text: formatNumberSection(value, sec, loc, withSign), color: sec.color, isDate: false }
}

/** True when the format code renders dates/times. */
export function isDateFormat(code) {
  if (!code) return false
  if (code.includes('SHORT_DATE')) return true
  return parseFormat(code).some((s) => s.isDate)
}

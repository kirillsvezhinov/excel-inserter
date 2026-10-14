// Minimal CSS handling tuned for what spreadsheet apps put on the clipboard.
// We do not rely on the browser's CSSOM: DOMParser documents don't reliably
// expose `styleSheets`, browsers drop `mso-*` properties, and they rewrite
// colors to rgb() — so declarations are parsed by hand.

const SIDES = ['top', 'right', 'bottom', 'left']

// Colors Excel/LibreOffice actually emit by name, plus CSS2 system colors.
const NAMED_COLORS = {
  black: '#000000', white: '#ffffff', red: '#ff0000', lime: '#00ff00',
  blue: '#0000ff', yellow: '#ffff00', aqua: '#00ffff', cyan: '#00ffff',
  fuchsia: '#ff00ff', magenta: '#ff00ff', silver: '#c0c0c0', gray: '#808080',
  grey: '#808080', maroon: '#800000', olive: '#808000', green: '#008000',
  purple: '#800080', teal: '#008080', navy: '#000080', orange: '#ffa500',
  windowtext: '#000000', window: '#ffffff', buttontext: '#000000',
  buttonface: '#f0f0f0', captiontext: '#000000', infotext: '#000000',
  infobackground: '#ffffe1', graytext: '#6d6d6d', highlight: '#0078d7',
  highlighttext: '#ffffff', menu: '#f0f0f0', menutext: '#000000'
}

const NO_COLOR = new Set(['auto', 'none', 'transparent', 'inherit', 'initial', 'unset', 'currentcolor', 'invert'])

/**
 * Normalize a CSS color to lowercase `#rrggbb` where possible.
 * `loose` also accepts unknown alphabetic names (valid only in `color`/`background-color`,
 * not inside shorthands where words like `no-repeat` would be misread).
 */
export function normalizeColor(value, loose = true) {
  if (value == null) return null
  let v = String(value).trim().toLowerCase().replace(/\s*!important$/, '')
  if (!v || NO_COLOR.has(v)) return null
  let m
  if ((m = /^#([0-9a-f]{3})$/.exec(v))) return '#' + m[1].split('').map((c) => c + c).join('')
  if (/^#[0-9a-f]{6}$/.test(v)) return v
  if ((m = /^#([0-9a-f]{6})([0-9a-f]{2})$/.exec(v))) return m[2] === '00' ? null : '#' + m[1]
  if ((m = /^rgba?\(\s*([\d.]+)%?\s*[,\s]\s*([\d.]+)%?\s*[,\s]\s*([\d.]+)%?\s*(?:[,/]\s*([\d.]+%?)\s*)?\)$/.exec(v))) {
    if (m[4] != null && parseFloat(m[4]) === 0) return null
    return '#' + [m[1], m[2], m[3]].map((n) => Math.max(0, Math.min(255, Math.round(+n))).toString(16).padStart(2, '0')).join('')
  }
  if (NAMED_COLORS[v]) return NAMED_COLORS[v]
  if (loose && /^[a-z]{3,30}$/.test(v)) return v
  return null
}

const FONT_SIZE_KEYWORDS = {
  'xx-small': 7, 'x-small': 7.5, small: 10, medium: 12, large: 13.5, 'x-large': 18, 'xx-large': 24
}

/** Convert a CSS length to points. `basePt` resolves em/%; returns null when unknown. */
export function toPt(value, basePt = 12) {
  if (value == null) return null
  const v = String(value).trim().toLowerCase()
  if (FONT_SIZE_KEYWORDS[v]) return FONT_SIZE_KEYWORDS[v]
  const m = /^(-?[\d.]+)(pt|px|em|rem|%|in|cm|mm)?$/.exec(v)
  if (!m) return null
  const n = parseFloat(m[1])
  if (!isFinite(n)) return null
  switch (m[2]) {
    case 'px': return n * 0.75
    case 'em': case 'rem': return n * basePt
    case '%': return (n / 100) * basePt
    case 'in': return n * 72
    case 'cm': return (n * 72) / 2.54
    case 'mm': return (n * 72) / 25.4
    default: return n // 'pt' or unitless (Excel writes `width=64` in px, but that is an attribute, not CSS)
  }
}

/** Split a declaration list on `;`, honoring quotes and backslash escapes (Excel number formats contain `\;`). */
function splitTopLevel(text, sep) {
  const parts = []
  let cur = ''
  let quote = null
  let depth = 0
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]
    if (ch === '\\' && i + 1 < text.length) { cur += ch + text[++i]; continue }
    if (quote) { if (ch === quote) quote = null; cur += ch; continue }
    if (ch === '"' || ch === "'") { quote = ch; cur += ch; continue }
    if (ch === '(') depth++
    else if (ch === ')') depth = Math.max(0, depth - 1)
    if (ch === sep && depth === 0) { parts.push(cur); cur = '' } else cur += ch
  }
  parts.push(cur)
  return parts
}

function splitTokens(value) {
  return splitTopLevel(value.trim().replace(/\s+/g, ' '), ' ').filter(Boolean)
}

const BORDER_STYLES = /^(none|hidden|solid|dashed|dotted|double|groove|ridge|inset|outset|hairline|dot-dash|dot-dot-dash|slant-dash-dot|dash-dot|dash-dot-dot)$/
const BORDER_WIDTHS = /^(thin|medium|thick|[\d.]+(pt|px|em|in|cm|mm)?)$/

function parseBorderValue(value) {
  const out = {}
  for (const tok of splitTokens(value.toLowerCase())) {
    if (BORDER_STYLES.test(tok)) out.style = tok
    else if (BORDER_WIDTHS.test(tok)) out.width = tok
    else out.color = tok
  }
  return out
}

function boxValues(value) {
  const t = splitTokens(value)
  const [a, b = a, c = a, d = b] = t
  return [a, b, c, d]
}

const FONT_SHORTHAND = /^((?:(?:italic|oblique|normal|bold|bolder|lighter|small-caps|\d{3})\s+)*)([\d.]+(?:pt|px|em|rem|%)|xx-small|x-small|small|medium|large|x-large|xx-large)(?:\s*\/\s*\S+)?\s+(.+)$/i

/** Expand one declaration into the longhands the rest of the code reads. */
function expandInto(out, prop, value) {
  switch (prop) {
    case 'background': {
      const tok = splitTokens(value).find((t) => normalizeColor(t, false))
      out['background-color'] = tok || 'none'
      return
    }
    case 'border': {
      const b = parseBorderValue(value)
      for (const s of SIDES) setBorderSide(out, s, b)
      return
    }
    case 'border-width': case 'border-style': case 'border-color': {
      const kind = prop.slice(7)
      boxValues(value).forEach((v, i) => { out[`border-${SIDES[i]}-${kind}`] = v })
      return
    }
    case 'font': {
      const m = FONT_SHORTHAND.exec(value.trim())
      if (!m) return
      const pre = m[1].toLowerCase()
      out['font-style'] = /italic|oblique/.test(pre) ? 'italic' : 'normal'
      const w = /(bold|bolder|lighter|\d{3})/.exec(pre)
      out['font-weight'] = w ? w[1] : 'normal'
      out['font-size'] = m[2]
      out['font-family'] = m[3]
      return
    }
    case 'text-decoration-line':
      out['text-decoration'] = value
      return
    case 'text-align':
      // Excel's `general` is invalid CSS; browsers ignore it so `align=right` on numbers wins.
      if (/^general$/i.test(value.trim())) return
      break
  }
  const side = /^border-(top|right|bottom|left)$/.exec(prop)
  if (side) { setBorderSide(out, side[1], parseBorderValue(value)); return }
  out[prop] = value
}

function setBorderSide(out, side, b) {
  out[`border-${side}-style`] = b.style || 'none'
  out[`border-${side}-width`] = b.width || 'medium'
  out[`border-${side}-color`] = b.color || 'currentcolor'
}

/** Parse a `style` attribute / rule body into an object of lowercase longhand properties. */
export function parseDeclarations(text) {
  const out = {}
  if (!text) return out
  for (const part of splitTopLevel(String(text), ';')) {
    const idx = part.indexOf(':')
    if (idx < 0) continue
    const prop = part.slice(0, idx).trim().toLowerCase()
    const value = part.slice(idx + 1).trim().replace(/\s*!important$/i, '')
    if (!prop || !value || !/^[a-z-]+$/.test(prop)) continue
    expandInto(out, prop, value)
  }
  return out
}

/** Read one CSS block starting after `{`; returns [body, indexAfterClosingBrace]. */
function readBlock(text, start) {
  let depth = 1
  let quote = null
  for (let i = start; i < text.length; i++) {
    const ch = text[i]
    if (ch === '\\') { i++; continue }
    if (quote) { if (ch === quote) quote = null; continue }
    if (ch === '"' || ch === "'") quote = ch
    else if (ch === '{') depth++
    else if (ch === '}' && --depth === 0) return [text.slice(start, i), i + 1]
  }
  return [text.slice(start), text.length]
}

const SIMPLE_SELECTOR = /^([a-z][a-z0-9]*)?((?:\.[\w-]+)*)$/i

/**
 * A stylesheet restricted to simple selectors (`td`, `.xl65`, `td.td1`, comma lists) —
 * that's all Excel, LibreOffice, Google Sheets and Cocoa HTML Writer emit for cells.
 */
export class StyleSheet {
  constructor(texts = [], { ignoreBorders = false } = {}) {
    this.rules = []
    this.cache = new Map()
    let order = 0
    for (const raw of texts) {
      const text = String(raw).replace(/\/\*[\s\S]*?\*\//g, '').replace(/<!--|-->/g, ' ')
      let i = 0
      while (i < text.length) {
        const open = text.indexOf('{', i)
        if (open < 0) break
        const selectorText = text.slice(i, open).trim()
        const [body, next] = readBlock(text, open + 1)
        i = next
        if (selectorText.startsWith('@')) continue // @page, @media, @font-face …
        const decls = parseDeclarations(body)
        if (ignoreBorders) for (const k of Object.keys(decls)) if (k.startsWith('border-')) delete decls[k]
        for (const sel of selectorText.split(',')) {
          const m = SIMPLE_SELECTOR.exec(sel.trim())
          if (!m || (!m[1] && !m[2])) continue
          const classes = m[2] ? m[2].slice(1).split('.').map((c) => c.toLowerCase()) : []
          this.rules.push({
            tag: m[1] ? m[1].toLowerCase() : null,
            classes,
            specificity: classes.length * 10 + (m[1] ? 1 : 0),
            order: order++,
            decls
          })
        }
      }
    }
  }

  /** Merged declarations of every rule matching an element with this tag and classes. */
  match(tag, classes = []) {
    const key = tag + '|' + classes.join(' ')
    let hit = this.cache.get(key)
    if (hit) return hit
    hit = {}
    this.rules
      .filter((r) => (!r.tag || r.tag === tag) && r.classes.every((c) => classes.includes(c)))
      .sort((a, b) => a.specificity - b.specificity || a.order - b.order)
      .forEach((r) => Object.assign(hit, r.decls))
    this.cache.set(key, hit)
    return hit
  }
}

/** `bold`, `700`, … → true; `normal`, `400`, … → false; unknown → fallback. */
export function isBold(weight, fallback = false) {
  const v = String(weight).trim().toLowerCase()
  if (v === 'bold' || v === 'bolder') return true
  if (v === 'normal' || v === 'lighter') return false
  const n = parseInt(v, 10)
  return isNaN(n) ? fallback : n >= 600
}

/** Keep only characters that can appear in font names (incl. CJK/Cyrillic names). */
export function sanitizeFontFamily(value) {
  if (!value) return null
  const v = String(value).replace(/[^\p{L}\p{N}\s,'"\-._]/gu, '').replace(/\s+/g, ' ').trim()
  return v && v.length <= 200 ? v : null
}

export function primaryFontFamily(value) {
  if (!value) return null
  return value.split(',')[0].trim().replace(/^['"]|['"]$/g, '').toLowerCase()
}

function borderWidthPx(width) {
  const v = String(width).trim().toLowerCase()
  if (v === 'thin') return 1
  if (v === 'medium') return 2
  if (v === 'thick') return 3
  const m = /^([\d.]+)(px)$/.exec(v)
  if (m) return Math.max(1, Math.min(6, Math.round(+m[1])))
  const pt = toPt(v)
  if (pt == null) return 1
  if (pt === 0) return 0
  // Excel: thin = .5pt, medium = 1pt, thick = 1.5pt
  return pt <= 0.75 ? 1 : pt <= 1.25 ? 2 : 3
}

function borderStyle(style) {
  const v = String(style).toLowerCase()
  if (v === 'none' || v === 'hidden') return null
  if (v === 'hairline') return 'solid'
  if (/^(solid|dashed|dotted|double|groove|ridge|inset|outset)$/.test(v)) return v
  if (v.includes('dash')) return 'dashed'
  if (v.includes('dot')) return 'dotted'
  return 'solid'
}

/** Build a normalized `1px solid #000000` value for one side from longhand declarations. */
export function resolveBorder(decls, side, textColor) {
  const style = decls[`border-${side}-style`]
  if (!style) return null
  const s = borderStyle(style)
  if (!s) return null
  let w = borderWidthPx(decls[`border-${side}-width`] ?? 'medium')
  if (!w) return null
  if (s === 'double') w = Math.max(w, 3)
  const c = normalizeColor(decls[`border-${side}-color`], true) || textColor || '#000000'
  return `${w}px ${s} ${c}`
}

/** Normalize a border value read back from a cell's DOM (used by schema helpers). */
export function normalizeBorder(value) {
  if (!value) return null
  const decls = {}
  setBorderSide(decls, 'top', parseBorderValue(value))
  return resolveBorder(decls, 'top', null)
}

export { SIDES }

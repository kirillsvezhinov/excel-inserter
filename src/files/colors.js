// Resolving OOXML colors ({argb} | {theme, tint} | {indexed}) to #rrggbb.

// Default legacy palette (ECMA-376 §18.8.27). 64/65 = system foreground/background.
const INDEXED = (
  '000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF ' +
  '000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF ' +
  '800000 008000 000080 808000 800080 008080 C0C0C0 808080 ' +
  '9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF ' +
  '000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF ' +
  '00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 ' +
  '3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 ' +
  '003366 339966 003300 333300 993300 993366 333399 333333 ' +
  '000000 FFFFFF'
).split(' ')

// Office 2013+ default theme, used when the workbook has none.
const DEFAULT_THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72']
const SCHEME_ORDER = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']

/** Parse `<a:clrScheme>` from theme1.xml into the 12 theme colors (index order as cells use them). */
export function parseThemeColors(xml) {
  if (!xml) return DEFAULT_THEME
  try {
    const doc = new DOMParser().parseFromString(xml, 'application/xml')
    const scheme = doc.getElementsByTagName('a:clrScheme')[0]
    if (!scheme) return DEFAULT_THEME
    const byName = {}
    for (const el of scheme.children) {
      const c = el.firstElementChild
      if (c) byName[el.localName] = c.getAttribute('val') && c.localName === 'srgbClr' ? c.getAttribute('val') : c.getAttribute('lastClr')
    }
    return SCHEME_ORDER.map((n, i) => byName[n] || DEFAULT_THEME[i])
  } catch {
    return DEFAULT_THEME
  }
}

function hexToHsl(hex) {
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255)
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  const l = (max + min) / 2
  if (max === min) return [0, 0, l]
  const d = max - min
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
  const h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4
  return [h / 6, s, l]
}

function hslToHex([h, s, l]) {
  const hue = (p, q, t) => {
    if (t < 0) t += 1
    if (t > 1) t -= 1
    if (t < 1 / 6) return p + (q - p) * 6 * t
    if (t < 1 / 2) return q
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6
    return p
  }
  let rgb
  if (s === 0) rgb = [l, l, l]
  else {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s
    const p = 2 * l - q
    rgb = [hue(p, q, h + 1 / 3), hue(p, q, h), hue(p, q, h - 1 / 3)]
  }
  return rgb.map((v) => Math.round(v * 255).toString(16).padStart(2, '0')).join('')
}

/** ECMA-376 tint: darken (tint < 0) or lighten (tint > 0) in HSL luminance. */
export function applyTint(hex, tint) {
  if (!tint) return hex
  const [h, s, l] = hexToHsl(hex)
  const nl = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint
  return hslToHex([h, s, Math.max(0, Math.min(1, nl))])
}

/** Resolve an ExcelJS/SheetJS color object to `#rrggbb`, or null for "automatic". */
export function resolveColor(c, theme = DEFAULT_THEME) {
  if (!c) return null
  let hex = null
  const argb = c.argb || c.rgb
  if (typeof argb === 'string' && /^[0-9a-f]{6,8}$/i.test(argb)) hex = argb.slice(-6)
  else if (c.theme != null && theme[c.theme]) hex = theme[c.theme]
  else if (c.indexed != null && INDEXED[c.indexed]) hex = INDEXED[c.indexed]
  if (!hex) return null
  return '#' + applyTint(hex, c.tint).toLowerCase()
}

/** Linear interpolation between two #rrggbb colors (for color scales). */
export function mixColors(a, b, t) {
  const pa = [1, 3, 5].map((i) => parseInt(a.slice(i, i + 2), 16))
  const pb = [1, 3, 5].map((i) => parseInt(b.slice(i, i + 2), 16))
  return '#' + pa.map((v, i) => Math.round(v + (pb[i] - v) * t).toString(16).padStart(2, '0')).join('')
}

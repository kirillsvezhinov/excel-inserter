import { parseDelimited, tsvToModel } from '../tsv.js'

// Locales whose Excel writes "ANSI" CSV in windows-1251.
const CYRILLIC_LANGS = /^(ru|uk|be|bg|sr|mk|kk|ky|tg|mn)\b/i

/**
 * Decode CSV bytes: BOM → UTF-8/UTF-16; valid UTF-8 → UTF-8; otherwise the legacy
 * code page Excel uses for "CSV (comma delimited)" in the viewer's locale.
 */
export function decodeText(bytes, { encoding, locale } = {}) {
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) return new TextDecoder('utf-8').decode(bytes.subarray(3))
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes.subarray(2))
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes.subarray(2))
  if (encoding) return new TextDecoder(encoding).decode(bytes)
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes)
  } catch {
    const lang = locale || (typeof navigator !== 'undefined' ? navigator.language : '') || ''
    const cyrillic = CYRILLIC_LANGS.test(lang) || looksLikeCp1251(bytes)
    return new TextDecoder(cyrillic ? 'windows-1251' : 'windows-1252').decode(bytes)
  }
}

/**
 * Russian text in windows-1251 is words made of runs of bytes 0xC0–0xFF; Western text in
 * windows-1252 has isolated accented letters (é, ü) inside ASCII words.
 */
function looksLikeCp1251(bytes) {
  let high = 0
  let inRuns = 0
  let run = 0
  const n = Math.min(bytes.length, 65536)
  for (let i = 0; i <= n; i++) {
    const b = i < n ? bytes[i] : 0
    if (b >= 0xc0 || b === 0xa8 || b === 0xb8) { high++; run++ } else {
      if (run >= 3) inRuns += run
      run = 0
    }
  }
  return high >= 3 && inRuns / high > 0.6
}

/** Pick the delimiter that splits the first lines into the most consistent number of columns. */
export function detectDelimiter(text) {
  const sep = /^sep=(.)\r?\n/i.exec(text) // Excel's explicit "sep=;" hint
  if (sep) return sep[1]
  const sample = text.slice(0, 20000)
  let best = ','
  let bestScore = 0
  for (const d of ['\t', ';', ',', '|']) {
    const rows = parseDelimited(sample, d).slice(0, 20)
    if (rows.length > 1 && !/\n$/.test(sample)) rows.pop() // last sampled line may be cut
    const widths = rows.map((r) => r.length)
    const first = widths[0] || 0
    if (first < 2) continue
    const consistent = widths.filter((w) => w === first).length / widths.length
    const score = consistent * first
    if (score > bestScore) { bestScore = score; best = d }
  }
  return best
}

export function readDelimited(bytes, opts, delimiter) {
  let text = decodeText(bytes, opts.files)
  const d = delimiter || detectDelimiter(text)
  text = text.replace(/^sep=.\r?\n/i, '')
  let rows = parseDelimited(text, d)
  let truncated = false
  const width = rows.reduce((w, r) => Math.max(w, r.length), 0)
  if (width && rows.length * width > opts.files.maxCells) {
    rows = rows.slice(0, Math.max(1, Math.floor(opts.files.maxCells / width)))
    truncated = true
  }
  if (!rows.length || !width) return null
  const model = tsvToModel(rows, 'csv')
  if (truncated) model.truncated = true
  return model
}

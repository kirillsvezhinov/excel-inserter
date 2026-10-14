/**
 * Parse tab-separated text as written by Excel/Sheets: fields containing tabs,
 * newlines or quotes are wrapped in double quotes, with `""` as an escaped quote.
 * @returns {string[][]}
 */
export function parseTSV(text) {
  return parseDelimited(text, '\t')
}

/** Same as parseTSV with any single-character delimiter (`,` `;` `|` …). */
export function parseDelimited(text, delim) {
  const s = String(text).replace(/\r\n?/g, '\n')
  const rows = []
  let row = []
  let field = ''
  let i = 0
  let fieldStart = true
  while (i < s.length) {
    const ch = s[i]
    if (fieldStart && ch === '"') {
      // Quoted field: only treat as quoted if the closing quote is followed by a delimiter.
      let j = i + 1
      let value = ''
      for (;;) {
        if (j >= s.length) { j = -1; break }
        if (s[j] === '"') {
          if (s[j + 1] === '"') { value += '"'; j += 2; continue }
          j++
          break
        }
        value += s[j++]
      }
      if (j !== -1 && (j === s.length || s[j] === delim || s[j] === '\n')) {
        field = value
        i = j
        fieldStart = false
        continue
      }
    }
    fieldStart = false
    if (ch === delim) { row.push(field); field = ''; fieldStart = true }
    else if (ch === '\n') { row.push(field); rows.push(row); row = []; field = ''; fieldStart = true }
    else field += ch
    i++
  }
  if (!fieldStart || field || row.length) { row.push(field); rows.push(row) }
  return rows
}

/** True when text has a consistent tabular shape (≥2 columns in every row). */
export function looksLikeTSV(rows) {
  if (!rows.length) return false
  const width = rows[0].length
  return width >= 2 && rows.every((r) => r.length === width)
}

const NUMERIC = /^[-+(]?[$€£¥₽]?\s?\d[\d\s\u00a0\u202f.,']*(?:[eE][-+]?\d+)?\s?[%$€£¥₽)]?$/

/** Build a table model (see parseSpreadsheetHTML) from TSV rows; numeric-looking cells are right-aligned. */
export function tsvToModel(rows, source = 'tsv') {
  const width = rows.reduce((w, r) => Math.max(w, r.length), 0)
  return {
    source,
    width,
    height: rows.length,
    colWidths: null,
    rows: rows.map((r) =>
      Array.from({ length: width }, (_, c) => {
        const lines = (r[c] ?? '').split('\n')
        const runs = []
        lines.forEach((line, k) => {
          if (k) runs.push({ br: true })
          if (line) runs.push({ text: line, marks: {} })
        })
        const style = NUMERIC.test((r[c] ?? '').trim()) ? { textAlign: 'right' } : {}
        return { colspan: 1, rowspan: 1, style, content: [runs] }
      })
    )
  }
}

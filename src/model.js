import { primaryFontFamily } from './css.js'

/**
 * Text formatting state → marks, keeping only what differs from the source's default
 * cell look (`base`), filtered by the `styles` option.
 */
export function marksFor(s, base, styles) {
  const m = {}
  if (styles.bold && s.bold && !base.bold) m.bold = true
  if (styles.italic && s.italic && !base.italic) m.italic = true
  if (styles.underline && s.underline && !base.underline && !s.link) m.underline = true
  if (styles.strikethrough && s.strike && !base.strike) m.strike = true
  if (styles.script && s.script) m[s.script] = true
  if (styles.color && s.color && s.color !== base.color && !s.link) m.color = s.color
  if (styles.fontSize && s.fontSize && (!base.fontSize || Math.abs(s.fontSize - base.fontSize) > 0.05)) {
    m.fontSize = `${s.fontSize}pt`
  }
  if (styles.fontFamily && s.fontFamily && primaryFontFamily(s.fontFamily) !== primaryFontFamily(base.fontFamily)) {
    m.fontFamily = s.fontFamily
  }
  if (styles.links && s.link) m.link = s.link
  return m
}

export const sameMarks = (a, b) => {
  const ka = Object.keys(a)
  return ka.length === Object.keys(b).length && ka.every((k) => a[k] === b[k])
}

/** Append a run, merging it into the previous one when the marks match. */
export function pushRun(runs, text, marks) {
  if (!text) return
  const prev = runs[runs.length - 1]
  if (prev && !prev.br && sameMarks(prev.marks, marks)) prev.text += text
  else runs.push({ text, marks })
}

export function emptyCell() {
  return { colspan: 1, rowspan: 1, style: {}, content: [[]] }
}

/**
 * Turn raw grid placements into a rectangular TableModel: drops hidden rows/columns
 * (shrinking spans that cross them), pads holes with empty cells, maps column widths.
 *
 * @param {object} g
 * @param {Array<{row:number,col:number,rowspan:number,colspan:number,hidden?:boolean}>} g.placements
 * @param {(p: object) => {style: object, content: any[][]}} g.build  called only for kept cells
 */
export function assembleModel({ source, placements, nRows, width, hiddenRows = [], hiddenCols = [], colWidths = [], opts, build }) {
  const keepRow = (r) => !(opts.skipHidden && hiddenRows[r])
  const keepCol = (c) => !(opts.skipHidden && hiddenCols[c])
  const rowMap = []
  let nr = 0
  for (let r = 0; r < nRows; r++) rowMap[r] = keepRow(r) ? nr++ : -1
  const colMap = []
  let nc = 0
  for (let c = 0; c < width; c++) colMap[c] = keepCol(c) ? nc++ : -1
  if (!nr || !nc) return null

  const grid = Array.from({ length: nr }, () => [])
  const filled = Array.from({ length: nr }, () => [])
  for (const p of placements) {
    if (opts.skipHidden && p.hidden) continue
    const rows = []
    for (let r = p.row; r < p.row + p.rowspan; r++) if (rowMap[r] >= 0) rows.push(rowMap[r])
    const cols = []
    for (let c = p.col; c < p.col + p.colspan; c++) if (colMap[c] >= 0) cols.push(colMap[c])
    if (!rows.length || !cols.length) continue
    const cell = { ...build(p), colspan: cols.length, rowspan: rows.length }
    grid[rows[0]].push({ col: cols[0], cell })
    for (const r of rows) for (const c of cols) filled[r][c] = true
  }
  // Partial selections and ragged rows: pad to a rectangle with empty cells.
  for (let r = 0; r < nr; r++) {
    for (let c = 0; c < nc; c++) {
      if (!filled[r][c]) grid[r].push({ col: c, cell: emptyCell() })
    }
    grid[r].sort((a, b) => a.col - b.col)
  }

  const widths = opts.columnWidths
    ? colMap.map((m, c) => (m >= 0 ? colWidths[c] ?? null : undefined)).filter((w) => w !== undefined)
    : null
  return {
    source,
    width: nc,
    height: nr,
    colWidths: widths && widths.some((w) => w) ? widths : null,
    rows: grid.map((row) => row.map((x) => x.cell))
  }
}

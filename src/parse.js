import {
  StyleSheet, SIDES, isBold, normalizeColor, parseDeclarations, primaryFontFamily,
  resolveBorder, sanitizeFontFamily, toPt
} from './css.js'
import { detectSource } from './detect.js'
import { resolveOptions } from './options.js'

// Elements whose content never becomes cell text.
const SKIP_TAGS = new Set([
  'style', 'script', 'template', 'head', 'title', 'meta', 'link', 'noscript',
  'img', 'svg', 'object', 'iframe', 'input', 'button', 'select', 'textarea', 'comment', 'xml'
])
const BLOCK_TAGS = new Set(['p', 'div', 'li', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'blockquote', 'pre'])

// Browser default styles for inline formatting tags (lowest cascade priority).
const UA_STYLES = {
  b: { 'font-weight': 'bold' }, strong: { 'font-weight': 'bold' }, th: { 'font-weight': 'bold' },
  i: { 'font-style': 'italic' }, em: { 'font-style': 'italic' }, cite: { 'font-style': 'italic' },
  u: { 'text-decoration': 'underline' }, ins: { 'text-decoration': 'underline' },
  s: { 'text-decoration': 'line-through' }, strike: { 'text-decoration': 'line-through' },
  del: { 'text-decoration': 'line-through' },
  sub: { 'vertical-align': 'sub' }, sup: { 'vertical-align': 'super' }
}

// <font size=N> → pt
const FONT_TAG_SIZES = [null, 8, 10, 12, 14, 18, 24, 36]

const INITIAL_TEXT = {
  bold: false, italic: false, underline: false, strike: false, script: null,
  color: null, fontSize: null, fontFamily: null, link: null,
  textAlign: null, verticalAlign: null, background: null
}

/** Strip a raw Windows CF_HTML header ("Version:0.9 StartHTML:…") if a browser left it in. */
function stripClipboardHeader(html) {
  return /^\s*Version:\d/.test(html) ? html.replace(/^[\s\S]*?(?=<)/, '') : html
}

function presentational(el, tag) {
  const out = {}
  const align = el.getAttribute('align')
  if (align && tag !== 'table') out['text-align'] = align
  const valign = el.getAttribute('valign')
  if (valign) out['vertical-align'] = valign
  const bg = el.getAttribute('bgcolor')
  if (bg) out['background-color'] = bg
  if (tag === 'font') {
    const color = el.getAttribute('color')
    if (color) out.color = color
    const face = el.getAttribute('face')
    if (face) out['font-family'] = face
    const size = el.getAttribute('size')
    if (size) {
      const n = /^[+-]/.test(size) ? 3 + parseInt(size, 10) : parseInt(size, 10)
      const pt = FONT_TAG_SIZES[Math.max(1, Math.min(7, n))]
      if (pt) out['font-size'] = pt + 'pt'
    }
  }
  return out
}

function isHidden(decls) {
  return /^none$/i.test((decls.display || '').trim()) || /^(hidden|collapse)$/i.test((decls.visibility || '').trim())
}

function textAlignValue(v) {
  if (!v) return undefined
  v = v.trim().toLowerCase()
  if (v === 'general' || v === 'inherit' || v === 'initial') return undefined // Excel's "General": let content decide
  if (v === 'left' || v === 'start' || v === 'fill') return 'left'
  if (v === 'right' || v === 'end') return 'right'
  if (v === 'center' || v === 'centeracross' || v === 'center-across' || v === 'middle') return 'center'
  if (v === 'justify' || v === 'distributed') return 'justify'
  return undefined
}

function verticalAlignValue(v) {
  if (!v) return undefined
  v = v.trim().toLowerCase()
  if (v === 'top' || v === 'text-top') return 'top'
  if (v === 'middle' || v === 'center' || v === 'justify' || v === 'distributed') return 'middle'
  if (v === 'bottom' || v === 'text-bottom' || v === 'baseline') return 'bottom'
  return undefined
}

function safeHref(href) {
  if (!href) return null
  const h = href.trim()
  return /^(https?:|mailto:|tel:)/i.test(h) ? h : null
}

/** Apply declarations to an inherited text/box state. `inline` = inside a cell (vertical-align means sub/sup). */
function applyDecls(state, d, inline) {
  const s = { ...state }
  if (d['font-weight'] != null) s.bold = isBold(d['font-weight'], s.bold)
  if (d['font-style'] != null) s.italic = /italic|oblique/i.test(d['font-style'])
  if (d['text-decoration'] != null) {
    // Excel font classes are a full spec, so `none` resets inherited decoration (unlike browsers).
    const v = d['text-decoration'].toLowerCase()
    s.underline = v.includes('underline')
    s.strike = v.includes('line-through')
  }
  if (d['text-underline-style'] != null) s.underline = !/^none$/i.test(d['text-underline-style'].trim())
  if (d.color != null) {
    const c = normalizeColor(d.color)
    if (c) s.color = c
  }
  if (d['font-size'] != null) {
    const pt = toPt(d['font-size'], s.fontSize || 12)
    if (pt && pt > 0 && pt < 500) s.fontSize = Math.round(pt * 100) / 100
  }
  if (d['font-family'] != null) {
    const f = sanitizeFontFamily(d['font-family'])
    if (f) s.fontFamily = f
  }
  if (inline) {
    const va = (d['vertical-align'] || '').trim().toLowerCase()
    if (va === 'super') s.script = 'sup'
    else if (va === 'sub') s.script = 'sub'
    else if (va === 'baseline') s.script = null
  } else {
    const ta = textAlignValue(d['text-align'])
    if (ta) s.textAlign = ta
    const va = verticalAlignValue(d['vertical-align'])
    if (va) s.verticalAlign = va
  }
  return s
}

class Context {
  constructor(doc, source, opts) {
    this.source = source
    this.opts = opts
    this.sheet = new StyleSheet(
      [...doc.querySelectorAll('style')].map((s) => s.textContent || ''),
      { ignoreBorders: source === 'google' } // Google's `td {border:1px solid #ccc}` is just gridlines
    )
    this.declCache = new WeakMap()
    this.stateCache = new WeakMap()
  }

  /** Cascade: UA tag style < presentational attrs < mso-style-parent < matching rules < inline style. */
  declared(el) {
    let d = this.declCache.get(el)
    if (d) return d
    const tag = el.localName.toLowerCase()
    const classes = (el.getAttribute('class') || '').toLowerCase().split(/\s+/).filter(Boolean)
    const matched = this.sheet.match(tag, classes)
    const parent = matched['mso-style-parent']
    d = Object.assign(
      {},
      UA_STYLES[tag],
      presentational(el, tag),
      parent ? this.sheet.match(tag, [parent.trim().toLowerCase()]) : null,
      matched,
      parseDeclarations(el.getAttribute('style'))
    )
    this.declCache.set(el, d)
    return d
  }

  /** State inherited by `el`'s children: walks from <body> down through `el`. */
  stateFor(el) {
    const hit = this.stateCache.get(el)
    if (hit) return hit
    const chain = []
    for (let n = el; n && n.nodeType === 1 && n.localName !== 'html'; n = n.parentElement) chain.unshift(n)
    let s = INITIAL_TEXT
    for (const n of chain) {
      const d = this.declared(n)
      s = applyDecls(s, d, false)
      // Background is not inherited in CSS, but a row's fill visually fills its cells.
      s.background = n.localName === 'tr' || n.localName === 'td' || n.localName === 'th'
        ? normalizeColor(d['background-color']) || (n.localName === 'tr' ? null : s.background)
        : null
    }
    this.stateCache.set(el, s)
    return s
  }
}

function marksFor(s, base, styles) {
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

const sameMarks = (a, b) => {
  const ka = Object.keys(a)
  return ka.length === Object.keys(b).length && ka.every((k) => a[k] === b[k])
}

/** Turn a cell's DOM into paragraphs of styled runs. */
function extractContent(td, cellState, base, ctx) {
  const { opts } = ctx
  const paras = [[]]
  const cur = () => paras[paras.length - 1]
  const breakPara = () => { if (cur().length) paras.push([]) }

  const walk = (node, st) => {
    for (const child of node.childNodes) {
      if (child.nodeType === 3) {
        const text = child.data.replace(/[ \t\n\r\f]+/g, ' ')
        if (text) cur().push({ text, state: st })
      } else if (child.nodeType === 1) {
        const tag = child.localName.toLowerCase()
        if (SKIP_TAGS.has(tag) || tag.includes(':')) continue // <o:p>, <v:shape> …
        const d = ctx.declared(child)
        if (opts.skipHidden && isHidden(d)) continue
        if (tag === 'br') {
          if (opts.lineBreaks === 'br') cur().push({ br: true })
          else paras.push([])
          continue
        }
        const block = BLOCK_TAGS.has(tag)
        if (block) breakPara()
        let s = applyDecls(st, d, true)
        if (tag === 'a') s = { ...s, link: safeHref(child.getAttribute('href')) || s.link }
        walk(child, s)
        if (block) breakPara()
      }
    }
  }
  walk(td, cellState)

  const result = paras.map((runs) => {
    // Collapse whitespace across run boundaries and trim at paragraph/line edges.
    const out = []
    let atLineStart = true
    for (const run of runs) {
      if (run.br) { trimEnd(out); out.push(run); atLineStart = true; continue }
      let text = run.text
      if (atLineStart || (out.length && !out[out.length - 1].br && out[out.length - 1].text.endsWith(' '))) {
        text = text.replace(/^ /, '')
      }
      if (!text) continue
      atLineStart = false
      const marks = marksFor(run.state, base, opts.styles)
      const prev = out[out.length - 1]
      if (prev && !prev.br && sameMarks(prev.marks, marks)) prev.text += text
      else out.push({ text, marks })
    }
    trimEnd(out)
    return out
  })
  while (result.length > 1 && !result[result.length - 1].length) result.pop()
  // A cell holding only &nbsp; is empty in spreadsheet terms.
  if (result.every((p) => p.every((r) => r.br || /^[\s ]*$/.test(r.text)))) return [[]]
  return result
}

function trimEnd(runs) {
  const last = runs[runs.length - 1]
  if (last && !last.br) {
    last.text = last.text.replace(/ $/, '')
    if (!last.text) runs.pop()
  }
}

function cellBox(s, d, base, ctx) {
  const { styles } = ctx.opts
  const box = {}
  if (styles.background && s.background) box.background = s.background
  if (styles.textAlign && s.textAlign && s.textAlign !== base.textAlign) box.textAlign = s.textAlign
  if (styles.verticalAlign && s.verticalAlign && s.verticalAlign !== base.verticalAlign) box.verticalAlign = s.verticalAlign
  if (styles.borders) {
    for (const side of SIDES) {
      const b = resolveBorder(d, side, s.color)
      if (b) box['border' + side[0].toUpperCase() + side.slice(1)] = b
    }
  }
  return box
}

function widthPx(el, d) {
  const attr = el.getAttribute('width')
  if (attr && /^\d+(\.\d+)?$/.test(attr.trim())) return Math.round(+attr)
  if (d.width) {
    const v = d.width.trim().toLowerCase()
    if (v.endsWith('px')) return Math.round(parseFloat(v)) || null
    const pt = toPt(v)
    if (pt && !v.endsWith('%') && !v.endsWith('em')) return Math.round((pt * 4) / 3)
  }
  return null
}

const intAttr = (el, name, fallback) => {
  const n = parseInt(el.getAttribute(name), 10)
  return isNaN(n) ? fallback : n
}

/**
 * Parse spreadsheet clipboard HTML into a table model.
 * Returns null when the HTML is not from an enabled source or contains no table.
 */
export function parseSpreadsheetHTML(html, options) {
  const opts = resolveOptions(options)
  html = stripClipboardHeader(String(html || ''))
  const doc = new DOMParser().parseFromString(html, 'text/html')
  const source = detectSource(doc, html)
  if (!source || !opts.sources.includes(source)) return null
  const table = doc.querySelector('table')
  if (!table) return null
  const ctx = new Context(doc, source, opts)

  // "Base" = what an unstyled cell looks like in the source app (Calibri 11pt black in Excel,
  // the table font in Google Sheets). Only differences from it become marks.
  const tableState = ctx.stateFor(table)
  const base = applyDecls(tableState, ctx.sheet.match('td', []), false)
  base.verticalAlign ||= 'bottom' // every spreadsheet app's default

  // Pass 1: lay every cell (hidden ones too) onto the grid so column indices match <col>s.
  const rowEls = [...table.rows].filter((tr) => tr.closest('table') === table)
  const placements = []
  const occupied = []
  const occupy = (r, c) => { (occupied[r] ||= [])[c] = true }
  const hiddenRows = rowEls.map((tr) => isHidden(ctx.declared(tr)))
  const nRows = rowEls.length

  rowEls.forEach((tr, r) => {
    let c = 0
    for (const td of tr.cells) {
      const d = ctx.declared(td)
      const colspan = Math.max(1, Math.min(1000, intAttr(td, 'colspan', 1)))
      const rawRowspan = intAttr(td, 'rowspan', 1)
      const rowspan = Math.max(1, Math.min(nRows - r, rawRowspan === 0 ? nRows - r : rawRowspan))
      while (occupied[r]?.[c]) c++
      for (let i = r; i < r + rowspan; i++) for (let j = c; j < c + colspan; j++) occupy(i, j)
      const hidden = isHidden(d)
      // Excel writes text overflowing into empty neighbours as `colspan=N style='mso-ignore:colspan'`.
      // That's not a merge — split it back into separate cells.
      if (colspan > 1 && /colspan/i.test(d['mso-ignore'] || '')) {
        placements.push({ td, d, row: r, col: c, colspan: 1, rowspan, hidden, overflow: 'first', span: colspan })
        for (let j = 1; j < colspan; j++) {
          placements.push({ td, d, row: r, col: c + j, colspan: 1, rowspan, hidden, overflow: j === colspan - 1 ? 'last' : 'middle', filler: true })
        }
      } else {
        placements.push({ td, d, row: r, col: c, colspan, rowspan, hidden })
      }
      c += colspan
    }
  })
  let width = occupied.reduce((w, row) => Math.max(w, row ? row.length : 0), 0)
  if (!width || !nRows) return null

  // Column metadata from <col>/<colgroup>.
  const colWidths = []
  const hiddenCols = []
  for (const child of table.children) {
    const tag = child.localName
    if (tag !== 'colgroup' && tag !== 'col') continue
    const items = tag === 'colgroup' && child.querySelector('col') ? [...child.querySelectorAll('col')] : [child]
    for (const el of items) {
      const d = ctx.declared(el)
      const span = Math.max(1, Math.min(1000, intAttr(el, 'span', 1)))
      const hidden = isHidden(d)
      const w = widthPx(el, d)
      for (let i = 0; i < span; i++) { colWidths.push(hidden ? null : w); hiddenCols.push(hidden) }
    }
  }
  for (const p of placements) {
    if (p.colspan === 1 && !p.filler && colWidths[p.col] == null && !hiddenCols[p.col]) {
      const w = widthPx(p.td, p.d)
      if (w) colWidths[p.col] = w
    }
  }
  // A column is hidden when its <col> says so, or when every cell starting in it is display:none.
  for (let c = 0; c < width; c++) {
    const own = placements.filter((p) => p.col === c && p.colspan === 1)
    if (own.length && own.every((p) => p.hidden)) hiddenCols[c] = true
  }

  // Pass 2: drop hidden rows/columns and remap coordinates.
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
    const cell = buildCell(p, ctx, base, opts)
    cell.colspan = cols.length
    cell.rowspan = rows.length
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

  const widths = opts.columnWidths ? colMap.map((m, c) => (m >= 0 ? colWidths[c] ?? null : undefined)).filter((w) => w !== undefined) : null
  return {
    source,
    width: nc,
    height: nr,
    colWidths: widths && widths.some((w) => w) ? widths : null,
    rows: grid.map((row) => row.map((x) => x.cell))
  }
}

function emptyCell() {
  return { colspan: 1, rowspan: 1, style: {}, content: [[]] }
}

function buildCell(p, ctx, base, opts) {
  const rowState = ctx.stateFor(p.td.parentElement)
  const cellState = applyDecls(rowState, p.d, false)
  cellState.background = normalizeColor(p.d['background-color']) || rowState.background
  const style = cellBox(cellState, p.d, base, ctx)
  if (p.overflow) {
    // Split overflow cells share one class: keep outer borders on the outer cells only.
    if (p.overflow !== 'first') delete style.borderLeft
    if (p.overflow !== 'last') delete style.borderRight
  }
  return {
    colspan: 1,
    rowspan: 1,
    style,
    content: p.filler ? [[]] : extractContent(p.td, cellState, base, ctx)
  }
}

export { stripClipboardHeader }

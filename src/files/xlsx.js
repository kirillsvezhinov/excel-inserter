// .xlsx / .xlsm / .xltx reader: ZIP + SpreadsheetML → TableModel, with styles,
// rich text, merges, hidden rows/columns, hyperlinks and conditional formatting.

import { readZip } from './zip.js'
import { parseThemeColors, resolveColor, mixColors } from './colors.js'
import { builtinFormat, formatValue } from './numfmt.js'
import { assembleModel, marksFor, pushRun } from '../model.js'
import { safeHref } from '../parse.js'

const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

// ---------- XML helpers ----------

function parseXML(str, name) {
  const doc = new DOMParser().parseFromString(str, 'application/xml')
  if (doc.getElementsByTagName('parsererror').length) throw new Error(`Invalid XML in ${name}`)
  return doc.documentElement
}
function kids(el, name) {
  const out = []
  if (el) for (let c = el.firstElementChild; c; c = c.nextElementSibling) if (c.localName === name) out.push(c)
  return out
}
function kid(el, name) {
  if (el) for (let c = el.firstElementChild; c; c = c.nextElementSibling) if (c.localName === name) return c
  return null
}
const num = (v, d = null) => (v == null || v === '' || isNaN(+v) ? d : +v)
const flag = (el) => !!el && !/^(0|false|none)$/i.test(el.getAttribute('val') ?? '1')
const relId = (el) => el.getAttribute('r:id') || el.getAttributeNS(REL_NS, 'id')
// OOXML escapes control chars as _xHHHH_.
const unescapeOOXML = (s) => s.replace(/_x([0-9a-fA-F]{4})_/g, (_, h) => String.fromCharCode(parseInt(h, 16)))

function resolvePath(base, target) {
  if (target.startsWith('/')) return target.slice(1)
  const parts = base.split('/').slice(0, -1)
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop()
    else if (seg && seg !== '.') parts.push(seg)
  }
  return parts.join('/')
}

/** "AB12" → { row: 11, col: 27 } (zero-based) */
export function decodeRef(ref) {
  const m = /^\$?([A-Z]{1,3})\$?(\d+)$/i.exec(ref.trim())
  if (!m) return null
  let col = 0
  for (const ch of m[1].toUpperCase()) col = col * 26 + ch.charCodeAt(0) - 64
  return { row: +m[2] - 1, col: col - 1 }
}
function decodeRange(ref) {
  const [a, b = a] = ref.split(':')
  const s = decodeRef(a)
  const e = decodeRef(b)
  if (!s || !e) return null
  return { r0: Math.min(s.row, e.row), c0: Math.min(s.col, e.col), r1: Math.max(s.row, e.row), c1: Math.max(s.col, e.col) }
}
const inRange = (rg, r, c) => r >= rg.r0 && r <= rg.r1 && c >= rg.c0 && c <= rg.c1

// ---------- styles ----------

const FONT_DEFAULTS = { bold: false, italic: false, underline: false, strike: false, script: null }

function colorOf(el) {
  if (!el || /^(1|true)$/.test(el.getAttribute('auto') || '')) return null
  return { rgb: el.getAttribute('rgb'), theme: num(el.getAttribute('theme')), tint: num(el.getAttribute('tint'), 0), indexed: num(el.getAttribute('indexed')) }
}

/** Only the properties present in the element (dxf fonts are partial). */
function parseFont(el, theme) {
  const f = {}
  if (!el) return f
  for (let c = el.firstElementChild; c; c = c.nextElementSibling) {
    switch (c.localName) {
      case 'b': f.bold = flag(c); break
      case 'i': f.italic = flag(c); break
      case 'u': f.underline = flag(c); break
      case 'strike': f.strike = flag(c); break
      case 'vertAlign': {
        const v = c.getAttribute('val')
        f.script = v === 'superscript' ? 'sup' : v === 'subscript' ? 'sub' : null
        break
      }
      case 'sz': f.size = num(c.getAttribute('val')); break
      case 'name': case 'rFont': f.name = c.getAttribute('val'); break
      case 'color': {
        const col = colorOf(c)
        f.color = col ? resolveColor(col, theme) : null
        break
      }
    }
  }
  return f
}

function parseFill(el, theme, dxf) {
  if (!el) return undefined
  const pattern = kid(el, 'patternFill')
  if (pattern) {
    const type = pattern.getAttribute('patternType')
    if (type === 'none') return null
    const fg = resolveColor(colorOf(kid(pattern, 'fgColor')), theme)
    const bg = resolveColor(colorOf(kid(pattern, 'bgColor')), theme)
    // In differential formats Excel stores a solid fill's color in bgColor.
    if (dxf) return bg || fg || undefined
    if (!type) return null
    return fg || bg || null
  }
  const gradient = kid(el, 'gradientFill')
  if (gradient) {
    const stop = kid(gradient, 'stop')
    return stop ? resolveColor(colorOf(kid(stop, 'color')), theme) : null
  }
  return undefined
}

const BORDER_STYLES = {
  hair: [1, 'solid'], thin: [1, 'solid'], medium: [2, 'solid'], thick: [3, 'solid'],
  dashed: [1, 'dashed'], mediumDashed: [2, 'dashed'], dotted: [1, 'dotted'], double: [3, 'double'],
  dashDot: [1, 'dashed'], mediumDashDot: [2, 'dashed'], dashDotDot: [1, 'dashed'],
  mediumDashDotDot: [2, 'dashed'], slantDashDot: [2, 'dashed']
}

function parseBorder(el, theme) {
  const out = {}
  if (!el) return out
  const sides = { top: 'top', bottom: 'bottom', left: 'left', right: 'right', start: 'left', end: 'right' }
  for (const [tag, side] of Object.entries(sides)) {
    const s = kid(el, tag)
    if (!s) continue
    const spec = BORDER_STYLES[s.getAttribute('style')]
    if (!spec) { if (!out[side]) out[side] = null; continue }
    const color = resolveColor(colorOf(kid(s, 'color')), theme) || '#000000'
    out[side] = `${spec[0]}px ${spec[1]} ${color}`
  }
  return out
}

function parseStyles(root, theme) {
  const numFmts = {}
  for (const n of kids(kid(root, 'numFmts'), 'numFmt')) numFmts[n.getAttribute('numFmtId')] = n.getAttribute('formatCode')
  const fmtCode = (id) => numFmts[id] ?? builtinFormat(num(id, 0))
  const fonts = kids(kid(root, 'fonts'), 'font').map((f) => ({ ...FONT_DEFAULTS, ...parseFont(f, theme) }))
  const fills = kids(kid(root, 'fills'), 'fill').map((f) => parseFill(f, theme, false) ?? null)
  const borders = kids(kid(root, 'borders'), 'border').map((b) => parseBorder(b, theme))
  const xfs = kids(kid(root, 'cellXfs'), 'xf').map((xf) => {
    const align = kid(xf, 'alignment')
    return {
      numFmt: fmtCode(xf.getAttribute('numFmtId') ?? '0'),
      font: fonts[num(xf.getAttribute('fontId'), 0)] || fonts[0] || FONT_DEFAULTS,
      fill: fills[num(xf.getAttribute('fillId'), 0)] ?? null,
      border: borders[num(xf.getAttribute('borderId'), 0)] || {},
      h: align?.getAttribute('horizontal') || null,
      v: align?.getAttribute('vertical') || null
    }
  })
  const dxfs = kids(kid(root, 'dxfs'), 'dxf').map((d) => ({
    font: parseFont(kid(d, 'font'), theme),
    fill: parseFill(kid(d, 'fill'), theme, true),
    border: parseBorder(kid(d, 'border'), theme),
    numFmt: kid(d, 'numFmt')?.getAttribute('formatCode') || null
  }))
  const defaultFont = fonts[0] || { ...FONT_DEFAULTS, name: 'Calibri', size: 11 }
  return { xfs, dxfs, defaultFont, defaultXf: xfs[0] || { numFmt: 'General', font: defaultFont, fill: null, border: {} } }
}

/** <si>/<is> → plain string or rich runs [{text, font}] */
function parseStringItem(si, theme) {
  const runs = kids(si, 'r')
  if (!runs.length) return unescapeOOXML(kids(si, 't').map((t) => t.textContent).join(''))
  return runs.map((r) => {
    const rPr = kid(r, 'rPr')
    return { text: unescapeOOXML(kids(r, 't').map((t) => t.textContent).join('')), font: rPr ? parseFont(rPr, theme) : null }
  })
}

// ---------- conditional formatting ----------

function parseCF(root) {
  const blocks = []
  for (const cf of kids(root, 'conditionalFormatting')) {
    const ranges = (cf.getAttribute('sqref') || '').split(/\s+/).map(decodeRange).filter(Boolean)
    const rules = kids(cf, 'cfRule').map((r) => {
      const scale = kid(r, 'colorScale')
      return {
        type: r.getAttribute('type'),
        priority: num(r.getAttribute('priority'), 1e9),
        dxfId: num(r.getAttribute('dxfId')),
        operator: r.getAttribute('operator'),
        text: r.getAttribute('text'),
        stopIfTrue: /^(1|true)$/.test(r.getAttribute('stopIfTrue') || ''),
        rank: num(r.getAttribute('rank'), 10),
        percent: /^(1|true)$/.test(r.getAttribute('percent') || ''),
        bottom: /^(1|true)$/.test(r.getAttribute('bottom') || ''),
        aboveAverage: !/^(0|false)$/.test(r.getAttribute('aboveAverage') || '1'),
        equalAverage: /^(1|true)$/.test(r.getAttribute('equalAverage') || ''),
        stdDev: r.getAttribute('stdDev'),
        formulas: kids(r, 'formula').map((f) => f.textContent),
        scale: scale && {
          cfvo: kids(scale, 'cfvo').map((v) => ({ type: v.getAttribute('type'), val: v.getAttribute('val') })),
          colors: kids(scale, 'color').map(colorOf)
        }
      }
    })
    if (ranges.length && rules.length) blocks.push({ ranges, rules })
  }
  return blocks
}

function percentile(sorted, p) {
  if (!sorted.length) return 0
  const idx = (sorted.length - 1) * p
  const lo = Math.floor(idx)
  return sorted[lo] + (sorted[Math.min(lo + 1, sorted.length - 1)] - sorted[lo]) * (idx - lo)
}

/**
 * Evaluate the supported rule types (cellIs, text, blanks/errors, top/bottom N, above/below
 * average, duplicate/unique, 2-/3-color scales). `expression`, data bars and icon sets are skipped.
 * @returns {Map<string, {background?, color?, bold?, italic?, underline?, strike?, border?}>} by "r,c"
 */
function evaluateCF(blocks, cells, dxfs, theme) {
  const result = new Map()
  const valueAt = (r, c) => cells.get(r)?.get(c)?.raw
  const literal = (f) => {
    const s = f.trim()
    if (/^-?[\d.]+(e[+-]?\d+)?$/i.test(s)) return +s
    if (/^".*"$/.test(s)) return s.slice(1, -1).replace(/""/g, '"')
    if (/^(TRUE|FALSE)$/i.test(s)) return /^TRUE$/i.test(s)
    const ref = decodeRef(s.replace(/^'?[^!]*'?!/, ''))
    if (ref && !s.includes('!')) return valueAt(ref.row, ref.col)
    return undefined // formulas are not evaluated
  }
  const eq = (a, b) => (typeof a === 'string' || typeof b === 'string'
    ? String(a ?? '').toLowerCase() === String(b ?? '').toLowerCase()
    : a === b)

  // Collect the cells of each block, sorted rules by priority.
  const applied = []
  for (const block of blocks) {
    const members = []
    for (const rg of block.ranges) {
      for (const [r, row] of cells) {
        if (r < rg.r0 || r > rg.r1) continue
        for (const [c, cell] of row) if (c >= rg.c0 && c <= rg.c1) members.push({ r, c, cell })
      }
    }
    const nums = members.map((m) => m.cell.raw).filter((v) => typeof v === 'number')
    const sorted = [...nums].sort((a, b) => a - b)
    const avg = nums.length ? nums.reduce((s, v) => s + v, 0) / nums.length : 0
    const counts = new Map()
    for (const m of members) {
      if (m.cell.raw == null || m.cell.raw === '') continue
      const k = String(m.cell.raw).toLowerCase()
      counts.set(k, (counts.get(k) || 0) + 1)
    }
    for (const rule of block.rules) applied.push({ rule, members, sorted, avg, counts })
  }
  applied.sort((a, b) => a.rule.priority - b.rule.priority)

  const stopped = new Set()
  for (const { rule, members, sorted, avg, counts } of applied) {
    const cfvoValue = (v) => {
      const min = sorted[0] ?? 0
      const max = sorted[sorted.length - 1] ?? 0
      switch (v.type) {
        case 'min': return min
        case 'max': return max
        case 'num': return num(literal(v.val ?? ''), min)
        case 'percent': return min + ((max - min) * num(v.val, 0)) / 100
        case 'percentile': return percentile(sorted, num(v.val, 50) / 100)
        default: return num(literal(v.val ?? ''), null)
      }
    }
    let threshold = null
    if (rule.type === 'top10' && sorted.length) {
      const n = rule.percent ? Math.max(1, Math.floor((sorted.length * rule.rank) / 100)) : rule.rank
      threshold = rule.bottom ? sorted[Math.min(n, sorted.length) - 1] : sorted[Math.max(0, sorted.length - n)]
    }
    for (const { r, c, cell } of members) {
      const key = `${r},${c}`
      if (stopped.has(key)) continue
      const v = cell.raw
      let style = null
      const dxf = rule.dxfId != null ? dxfs[rule.dxfId] : null
      const fromDxf = () => (dxf ? {
        background: dxf.fill || undefined,
        color: dxf.font.color || undefined,
        bold: dxf.font.bold,
        italic: dxf.font.italic,
        underline: dxf.font.underline,
        strike: dxf.font.strike,
        border: dxf.border
      } : null)
      let hit = false
      switch (rule.type) {
        case 'cellIs': {
          const [a, b] = rule.formulas.map(literal)
          if (a === undefined || v == null || v === '') break
          const lo = typeof a === 'number' && typeof b === 'number' ? Math.min(a, b) : a
          const hi = typeof a === 'number' && typeof b === 'number' ? Math.max(a, b) : b
          switch (rule.operator) {
            case 'equal': hit = eq(v, a); break
            case 'notEqual': hit = !eq(v, a); break
            case 'greaterThan': hit = v > a; break
            case 'greaterThanOrEqual': hit = v >= a; break
            case 'lessThan': hit = v < a; break
            case 'lessThanOrEqual': hit = v <= a; break
            case 'between': hit = b !== undefined && v >= lo && v <= hi; break
            case 'notBetween': hit = b !== undefined && (v < lo || v > hi); break
          }
          break
        }
        case 'containsText': case 'notContainsText': case 'beginsWith': case 'endsWith': {
          const s = String(v ?? '').toLowerCase()
          const t = String(rule.text ?? '').toLowerCase()
          hit = rule.type === 'containsText' ? s.includes(t)
            : rule.type === 'notContainsText' ? !s.includes(t)
              : rule.type === 'beginsWith' ? s.startsWith(t) : s.endsWith(t)
          break
        }
        case 'containsBlanks': hit = v == null || String(v).trim() === ''; break
        case 'notContainsBlanks': hit = !(v == null || String(v).trim() === ''); break
        case 'containsErrors': hit = !!cell.error; break
        case 'notContainsErrors': hit = !cell.error; break
        case 'top10': hit = typeof v === 'number' && threshold != null && (rule.bottom ? v <= threshold : v >= threshold); break
        case 'aboveAverage':
          if (typeof v === 'number' && !rule.stdDev) {
            hit = rule.aboveAverage ? (rule.equalAverage ? v >= avg : v > avg) : (rule.equalAverage ? v <= avg : v < avg)
          }
          break
        case 'duplicateValues': case 'uniqueValues': {
          if (v == null || v === '') break
          const n = counts.get(String(v).toLowerCase()) || 0
          hit = rule.type === 'duplicateValues' ? n > 1 : n === 1
          break
        }
        case 'colorScale': {
          if (typeof v !== 'number' || !rule.scale) break
          const stops = rule.scale.cfvo.map(cfvoValue)
          const colors = rule.scale.colors.map((col) => resolveColor(col, theme))
          if (stops.length < 2 || colors.length !== stops.length || colors.some((x) => !x) || stops.some((x) => x == null)) break
          let bg
          if (v <= stops[0]) bg = colors[0]
          else if (v >= stops[stops.length - 1]) bg = colors[colors.length - 1]
          else {
            const i = stops.findIndex((s, k) => k > 0 && v <= s)
            const span = stops[i] - stops[i - 1]
            bg = mixColors(colors[i - 1], colors[i], span ? (v - stops[i - 1]) / span : 0)
          }
          style = { background: bg }
          hit = true
          break
        }
      }
      if (!hit) continue
      style ||= fromDxf()
      if (style) {
        // Lower priority number wins: only fill properties not set by an earlier rule.
        const cur = result.get(key) || {}
        for (const [k, val] of Object.entries(style)) {
          if (val === undefined || cur[k] !== undefined) continue
          if (k === 'border') {
            cur.border = { ...val, ...(cur.border || {}) }
            continue
          }
          cur[k] = val
        }
        result.set(key, cur)
      }
      if (rule.stopIfTrue) stopped.add(key)
    }
  }
  return result
}

// ---------- sheet ----------

function cellValue(c, ctx) {
  const t = c.getAttribute('t') || 'n'
  const v = kid(c, 'v')?.textContent
  switch (t) {
    case 's': {
      const item = ctx.sst[num(v, -1)]
      return item === undefined ? { raw: '' } : typeof item === 'string' ? { raw: item } : { raw: item.map((r) => r.text).join(''), rich: item }
    }
    case 'inlineStr': {
      const item = parseStringItem(kid(c, 'is'), ctx.theme)
      return typeof item === 'string' ? { raw: item } : { raw: item.map((r) => r.text).join(''), rich: item }
    }
    case 'str': return { raw: unescapeOOXML(v ?? '') }
    case 'b': return { raw: v === '1' || v === 'true' }
    case 'e': return { raw: v ?? '#N/A', error: true }
    case 'd': {
      // Excel's ISO dates have no zone; parse them as UTC wall-clock, not local time.
      const iso = v && !/(Z|[+-]\d\d:?\d\d)$/.test(v) ? (v.includes('T') ? v + 'Z' : v + 'T00:00:00Z') : v
      const ms = Date.parse(iso)
      return isNaN(ms) ? { raw: v ?? '' } : { raw: (ms - Date.UTC(1899, 11, 30)) / 86400000, isoDate: true }
    }
    default: return v == null || v === '' ? { raw: null } : { raw: +v }
  }
}

const H_ALIGN = { left: 'left', center: 'center', right: 'right', justify: 'justify', distributed: 'justify', fill: 'left', centerContinuous: 'center' }
const V_ALIGN = { top: 'top', center: 'middle', justify: 'middle', distributed: 'middle' }

function textToRuns(text, marks, lineBreaks) {
  const paras = [[]]
  // Trailing line breaks (often a stray _x000D_) are invisible in Excel.
  String(text).replace(/\r\n?/g, '\n').replace(/\n+$/, '').split('\n').forEach((line, i) => {
    if (i) {
      if (lineBreaks === 'br') paras[paras.length - 1].push({ br: true })
      else paras.push([])
    }
    pushRun(paras[paras.length - 1], line, marks)
  })
  return paras
}

function buildContent(cell, xf, ctx, cf) {
  const { opts, base } = ctx
  const font = xf.font
  const cellState = (f) => ({
    bold: cf?.bold ?? f.bold,
    italic: cf?.italic ?? f.italic,
    underline: cf?.underline ?? f.underline,
    strike: cf?.strike ?? f.strike,
    script: f.script || null,
    color: cf?.color || f.color || base.color,
    fontSize: f.size ?? base.fontSize,
    fontFamily: f.name ?? base.fontFamily,
    link: cell.link || null
  })
  if (cell.rich) {
    const paras = [[]]
    for (const run of cell.rich) {
      // A run's rPr is a full spec; fall back to the cell font for missing name/size/color.
      const f = run.font ? { ...FONT_DEFAULTS, name: font.name, size: font.size, color: font.color, ...run.font } : font
      const marks = marksFor(cellState(f), base, opts.styles)
      const pieces = textToRuns(run.text, marks, opts.lineBreaks)
      pieces.forEach((p, i) => {
        if (i && opts.lineBreaks !== 'br') paras.push([])
        for (const r of p) r.br ? paras[paras.length - 1].push(r) : pushRun(paras[paras.length - 1], r.text, r.marks)
      })
    }
    return paras
  }
  let text = ''
  let fmtColor = null
  const v = cell.raw
  if (v == null || v === '') text = ''
  else if (cell.error) text = String(v)
  else {
    const out = formatValue(v, xf.numFmt, { locale: opts.files.locale, date1904: ctx.date1904 })
    text = out.text
    fmtColor = out.color
  }
  const state = cellState(font)
  if (fmtColor && !cf?.color) state.color = fmtColor
  return textToRuns(text, marksFor(state, base, opts.styles), opts.lineBreaks)
}

function cellStyle(cell, xf, ctx, cf) {
  const { styles } = ctx.opts
  const style = {}
  const bg = cf?.background ?? xf.fill
  if (styles.background && bg) style.background = bg
  let h = H_ALIGN[xf.h]
  if (!h && cell) {
    // "General" alignment: numbers and dates right, booleans and errors centered.
    if (typeof cell.raw === 'number' && !cell.error) h = 'right'
    else if (typeof cell.raw === 'boolean' || cell.error) h = 'center'
  }
  if (styles.textAlign && h) style.textAlign = h
  if (styles.verticalAlign && V_ALIGN[xf.v]) style.verticalAlign = V_ALIGN[xf.v]
  if (styles.borders) {
    const border = { ...xf.border, ...(cf?.border || {}) }
    for (const side of ['top', 'right', 'bottom', 'left']) {
      if (border[side]) style['border' + side[0].toUpperCase() + side.slice(1)] = border[side]
    }
  }
  return style
}

async function readSheet(path, name, ctx) {
  const { opts } = ctx
  const src = await ctx.read(path)
  if (src == null) throw new Error(`Missing sheet part ${path}`)
  const root = parseXML(src, path)
  const rels = await ctx.rels(path)

  // Columns
  const fmtPr = kid(root, 'sheetFormatPr')
  const defaultWidth = num(fmtPr?.getAttribute('defaultColWidth'))
  const defaultPx = defaultWidth ? Math.round(defaultWidth * 7) : 64
  const colInfo = []
  for (const col of kids(kid(root, 'cols'), 'col')) {
    const min = num(col.getAttribute('min'), 1) - 1
    const max = Math.min(num(col.getAttribute('max'), min + 1) - 1, 16383)
    const width = num(col.getAttribute('width'))
    const hidden = /^(1|true)$/.test(col.getAttribute('hidden') || '') || width === 0
    for (let c = min; c <= max && c < min + 1024; c++) colInfo[c] = { px: width ? Math.round(width * 7) : defaultPx, hidden, style: num(col.getAttribute('style')) }
  }

  // Cells
  const cells = new Map() // row → Map(col → cell)
  const hiddenRowSet = new Set()
  let rMin = Infinity, rMax = -1, cMin = Infinity, cMax = -1
  const sheetData = kid(root, 'sheetData')
  let nextRow = 0
  for (const rowEl of kids(sheetData, 'row')) {
    const r = num(rowEl.getAttribute('r'), nextRow + 1) - 1
    nextRow = r + 1
    if (/^(1|true)$/.test(rowEl.getAttribute('hidden') || '')) hiddenRowSet.add(r)
    let nextCol = 0
    const rowMap = new Map()
    for (const c of kids(rowEl, 'c')) {
      const pos = c.getAttribute('r') ? decodeRef(c.getAttribute('r')) : { row: r, col: nextCol }
      if (!pos) continue
      nextCol = pos.col + 1
      const cell = { ...cellValue(c, ctx), xf: num(c.getAttribute('s'), 0) }
      rowMap.set(pos.col, cell)
      if (cell.raw != null && cell.raw !== '') {
        rMin = Math.min(rMin, r); rMax = Math.max(rMax, r)
        cMin = Math.min(cMin, pos.col); cMax = Math.max(cMax, pos.col)
      }
    }
    if (rowMap.size) cells.set(r, rowMap)
  }

  // Merges
  const merges = kids(kid(root, 'mergeCells'), 'mergeCell').map((m) => decodeRange(m.getAttribute('ref') || '')).filter(Boolean)
  for (const m of merges) {
    rMin = Math.min(rMin, m.r0); rMax = Math.max(rMax, m.r1)
    cMin = Math.min(cMin, m.c0); cMax = Math.max(cMax, m.c1)
  }
  if (rMax < 0) return null

  // Hyperlinks (external only; in-workbook "location" links are dropped)
  for (const h of kids(kid(root, 'hyperlinks'), 'hyperlink')) {
    const rg = decodeRange(h.getAttribute('ref') || '')
    const rel = rels.get(relId(h) || '')
    const href = rel ? safeHref(rel.target) : null
    if (!rg || !href) continue
    for (let r = rg.r0; r <= rg.r1; r++) for (let c = rg.c0; c <= rg.c1; c++) {
      const cell = cells.get(r)?.get(c)
      if (cell) cell.link = href
    }
  }

  // Size guard
  let truncated = false
  const width = cMax - cMin + 1
  if ((rMax - rMin + 1) * width > opts.files.maxCells) {
    rMax = rMin + Math.max(1, Math.floor(opts.files.maxCells / width)) - 1
    truncated = true
  }

  const cfStyles = opts.files.conditionalFormatting ? evaluateCF(parseCF(root), cells, ctx.styles.dxfs, ctx.theme) : new Map()

  // Placements
  const nRows = rMax - rMin + 1
  const masters = new Map()
  const covered = new Set()
  for (const m of merges) {
    masters.set(`${m.r0},${m.c0}`, m)
    for (let r = m.r0; r <= m.r1; r++) for (let c = m.c0; c <= m.c1; c++) if (r !== m.r0 || c !== m.c0) covered.add(`${r},${c}`)
  }
  const placements = []
  for (let r = rMin; r <= rMax; r++) {
    const row = cells.get(r)
    for (let c = cMin; c <= cMax; c++) {
      const key = `${r},${c}`
      if (covered.has(key)) continue
      const cell = row?.get(c)
      const m = masters.get(key)
      if (!cell && !m) continue
      placements.push({
        row: r - rMin,
        col: c - cMin,
        rowspan: m ? Math.min(m.r1, rMax) - r + 1 : 1,
        colspan: m ? Math.min(m.c1, cMax) - c + 1 : 1,
        cell: cell || { raw: null, xf: colInfo[c]?.style ?? 0 },
        cf: cfStyles.get(key)
      })
    }
  }

  const model = assembleModel({
    source: 'xlsx',
    placements,
    nRows,
    width,
    hiddenRows: Array.from({ length: nRows }, (_, i) => hiddenRowSet.has(rMin + i)),
    hiddenCols: Array.from({ length: width }, (_, i) => !!colInfo[cMin + i]?.hidden),
    colWidths: Array.from({ length: width }, (_, i) => colInfo[cMin + i]?.px ?? defaultPx),
    opts,
    build: (p) => {
      const xf = ctx.styles.xfs[p.cell.xf] || ctx.styles.defaultXf
      return { style: cellStyle(p.cell, xf, ctx, p.cf), content: buildContent(p.cell, xf, ctx, p.cf) }
    }
  })
  if (model) {
    model.sheet = name
    if (truncated) model.truncated = true
  }
  return model
}

// ---------- workbook ----------

/**
 * Read the first sheet of an .xlsx/.xlsm/.xltx workbook (the first visible one, in tab
 * order). Other sheets are never decompressed or parsed.
 * @param {ArrayBuffer|Uint8Array} buffer
 * @param {object} opts resolved options
 * @returns {Promise<{ name: string, model: object } | null>} null when the sheet is empty
 */
export async function readXlsx(buffer, opts) {
  const zip = readZip(buffer)
  const decoder = new TextDecoder()
  const read = async (path) => {
    const entry = zip.get(path)
    return entry ? decoder.decode(await entry()) : null
  }
  const rels = async (partPath) => {
    const dir = partPath.includes('/') ? partPath.slice(0, partPath.lastIndexOf('/') + 1) : ''
    const file = partPath.slice(dir.length)
    const src = await read(`${dir}_rels/${file}.rels`)
    const map = new Map()
    if (!src) return map
    for (const rel of kids(parseXML(src, 'rels'), 'Relationship')) {
      const target = rel.getAttribute('Target') || ''
      const external = rel.getAttribute('TargetMode') === 'External'
      map.set(rel.getAttribute('Id'), { type: rel.getAttribute('Type') || '', target: external ? target : resolvePath(partPath, target), external })
    }
    return map
  }

  const rootRels = await rels('')
  const wbPath = [...rootRels.values()].find((r) => r.type.endsWith('/officeDocument'))?.target || 'xl/workbook.xml'
  const wbSrc = await read(wbPath)
  if (!wbSrc) throw new Error('Not an Excel workbook (no workbook part)')
  const wb = parseXML(wbSrc, wbPath)
  const wbRels = await rels(wbPath)
  const partOf = (suffix) => [...wbRels.values()].find((r) => r.type.endsWith(suffix))?.target

  const themeSrc = partOf('/theme') ? await read(partOf('/theme')) : null
  const theme = parseThemeColors(themeSrc)
  const stylesSrc = partOf('/styles') ? await read(partOf('/styles')) : null
  const styles = stylesSrc ? parseStyles(parseXML(stylesSrc, 'styles'), theme) : parseStyles(null, theme)
  const sstSrc = partOf('/sharedStrings') ? await read(partOf('/sharedStrings')) : null
  const sst = sstSrc ? kids(parseXML(sstSrc, 'sharedStrings'), 'si').map((si) => parseStringItem(si, theme)) : []

  const date1904 = /^(1|true)$/.test(kid(wb, 'workbookPr')?.getAttribute('date1904') || '')
  // <sheets> lists sheets in tab order. Hidden sheets are skipped: the user can't see them in Excel.
  const first = kids(kid(wb, 'sheets'), 'sheet').find((s) => !/hidden/i.test(s.getAttribute('state') || '') && wbRels.get(relId(s))?.target)
  if (!first) throw new Error('The workbook has no visible sheets')
  const name = first.getAttribute('name') || 'Sheet'

  const df = styles.defaultFont
  const base = {
    bold: false, italic: false, underline: false, strike: false, script: null,
    color: df.color || '#000000', fontSize: df.size ?? 11, fontFamily: df.name || 'Calibri', link: null
  }
  const ctx = { opts, read, rels, theme, styles, sst, date1904, base }

  const model = await readSheet(wbRels.get(relId(first)).target, name, ctx)
  return model ? { name, model } : null
}


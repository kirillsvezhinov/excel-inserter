// Test-only .xlsx writer: describe a workbook as plain objects, get the bytes Excel would write.
// Covers what the reader handles: styles (deduplicated into styles.xml), shared/inline/rich
// strings, formulas with cached values, merges, hidden rows/cols, hyperlinks, raw CF XML.
import { zip } from './xlsx-builder.js'

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'

export const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

export const OFFICE_THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72']

export function themeXML(colors = OFFICE_THEME) {
  const [lt1, dk1, lt2, dk2, ...rest] = colors
  const names = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Test"><a:themeElements><a:clrScheme name="Test">` +
    `<a:dk1><a:sysClr val="windowText" lastClr="${dk1}"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="${lt1}"/></a:lt1>` +
    `<a:dk2><a:srgbClr val="${dk2}"/></a:dk2><a:lt2><a:srgbClr val="${lt2}"/></a:lt2>` +
    names.map((n, i) => `<a:${n}><a:srgbClr val="${rest[i]}"/></a:${n}>`).join('') +
    `</a:clrScheme></a:themeElements></a:theme>`
}

/** 'FF0000' | { theme, tint } | { indexed } | { auto: true } → <tag .../> */
function colorXML(c, tag = 'color') {
  if (c == null) return ''
  if (typeof c === 'string') return `<${tag} rgb="${c.length === 6 ? 'FF' + c : c}"/>`
  if (c.auto) return `<${tag} auto="1"/>`
  const attrs = Object.entries(c).map(([k, v]) => `${k}="${v}"`).join(' ')
  return `<${tag} ${attrs}/>`
}

/** font: { b, i, u: true|'double', strike, vertAlign, sz, color, name } */
export function fontXML(f = {}, tag = 'font', nameTag = 'name') {
  return `<${tag}>` +
    (f.b ? '<b/>' : '') + (f.i ? '<i/>' : '') + (f.strike ? '<strike/>' : '') +
    (f.u ? (f.u === true ? '<u/>' : `<u val="${f.u}"/>`) : '') +
    (f.vertAlign ? `<vertAlign val="${f.vertAlign}"/>` : '') +
    (f.sz ? `<sz val="${f.sz}"/>` : '') + colorXML(f.color) +
    (f.name ? `<${nameTag} val="${esc(f.name)}"/>` : '') + `</${tag}>`
}

const BORDER_SIDES = ['left', 'right', 'top', 'bottom']

class StyleRegistry {
  constructor(defaultFont) {
    this.fonts = [fontXML(defaultFont)]
    this.fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>']
    this.borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>']
    this.numFmts = new Map()
    this.xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>']
    this.dxfs = []
    this.cache = new Map()
  }
  idx(list, xml) {
    let i = list.indexOf(xml)
    if (i < 0) { list.push(xml); i = list.length - 1 }
    return i
  }
  numFmtId(code) {
    if (typeof code === 'number') return code // built-in id
    if (!this.numFmts.has(code)) this.numFmts.set(code, 164 + this.numFmts.size)
    return this.numFmts.get(code)
  }
  /** style: { font, fill: color, border: { top: 'thin' | ['medium', color] }, numFmt: code|id, align: { h, v, wrap } } */
  xf(style) {
    if (!style) return 0
    const key = JSON.stringify(style)
    if (this.cache.has(key)) return this.cache.get(key)
    const fontId = style.font ? this.idx(this.fonts, fontXML({ sz: 11, color: { theme: 1 }, name: 'Calibri', ...style.font })) : 0
    const fillId = style.fill ? this.idx(this.fills, `<fill><patternFill patternType="solid">${colorXML(style.fill, 'fgColor')}<bgColor indexed="64"/></patternFill></fill>`) : 0
    const borderId = style.border
      ? this.idx(this.borders, '<border>' + BORDER_SIDES.map((s) => {
        const b = style.border[s]
        if (!b) return `<${s}/>`
        const [kind, color = { indexed: 64 }] = Array.isArray(b) ? b : [b]
        return `<${s} style="${kind}">${colorXML(color)}</${s}>`
      }).join('') + '<diagonal/></border>')
      : 0
    const numFmtId = style.numFmt != null ? this.numFmtId(style.numFmt) : 0
    const a = style.align
    const align = a ? `<alignment${a.h ? ` horizontal="${a.h}"` : ''}${a.v ? ` vertical="${a.v}"` : ''}${a.wrap ? ' wrapText="1"' : ''}/>` : ''
    const xml = `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"${align ? ' applyAlignment="1">' + align + '</xf>' : '/>'}`
    const id = this.idx(this.xfs, xml)
    this.cache.set(key, id)
    return id
  }
  /** dxf: { font, fill } — note: Excel stores a dxf solid fill in bgColor */
  dxf(d) {
    const xml = '<dxf>' + (d.font ? fontXML(d.font) : '') +
      (d.fill ? `<fill><patternFill>${colorXML(d.fill, 'bgColor')}</patternFill></fill>` : '') + '</dxf>'
    return this.idx(this.dxfs, xml)
  }
  toXML() {
    const fmts = [...this.numFmts].map(([code, id]) => `<numFmt numFmtId="${id}" formatCode="${esc(code)}"/>`).join('')
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="${MAIN}">` +
      (fmts ? `<numFmts count="${this.numFmts.size}">${fmts}</numFmts>` : '') +
      `<fonts count="${this.fonts.length}">${this.fonts.join('')}</fonts>` +
      `<fills count="${this.fills.length}">${this.fills.join('')}</fills>` +
      `<borders count="${this.borders.length}">${this.borders.join('')}</borders>` +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      `<cellXfs count="${this.xfs.length}">${this.xfs.join('')}</cellXfs>` +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      (this.dxfs.length ? `<dxfs count="${this.dxfs.length}">${this.dxfs.join('')}</dxfs>` : '') +
      '</styleSheet>'
  }
}

export function colName(c) {
  let s = ''
  for (c += 1; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s
  return s
}

/** JS Date → Excel serial (1900 or 1904 system). */
export function toSerial(date, date1904 = false) {
  const days = (date.getTime() - Date.UTC(1899, 11, 30)) / 86400000
  return date1904 ? days - 1462 : days
}

/**
 * Cell value forms:
 *   number | string | boolean | Date | null
 *   { v, s }                     value + style object
 *   { rich: [{ text, font }] }   rich text (shared string)
 *   { inline: 'text' }           inline string
 *   { f: 'A1*2', v, t }          formula with cached value (t: 'str' | 'b' | 'e')
 *   { error: '#N/A' }            error
 *   { iso: '2024-01-31T00:00:00' } t="d" cell
 *   { empty: true, s }           styled cell without value
 */
function cellXML(ref, cell, ctx) {
  // Without r="…" attributes position is implied by order, so gaps must be written out.
  if (cell == null) return ctx.noRefs ? '<c/>' : ''
  const isObj = typeof cell === 'object' && !(cell instanceof Date)
  const value = isObj && 'v' in cell && !cell.f ? cell.v : isObj ? undefined : cell
  const s = ctx.styles.xf(isObj ? cell.s : null)
  const r = ctx.noRefs ? '' : ` r="${ref}"`
  const sAttr = s ? ` s="${s}"` : ''
  if (isObj && cell.empty) return `<c${r}${sAttr}/>`
  if (isObj && cell.rich) return `<c${r}${sAttr} t="s"><v>${ctx.shared(richXML(cell.rich))}</v></c>`
  if (isObj && cell.inline != null) return `<c${r}${sAttr} t="inlineStr"><is><t xml:space="preserve">${esc(cell.inline)}</t></is></c>`
  if (isObj && cell.error) return `<c${r}${sAttr} t="e"><v>${esc(cell.error)}</v></c>`
  if (isObj && cell.iso) return `<c${r}${sAttr} t="d"><v>${cell.iso}</v></c>`
  if (isObj && cell.f) {
    const f = `<f${cell.shared ? ` t="shared" si="${cell.shared.si}"${cell.shared.ref ? ` ref="${cell.shared.ref}"` : ''}` : ''}>${esc(cell.f)}</f>`
    const t = cell.t ? ` t="${cell.t}"` : ''
    return `<c${r}${sAttr}${t}>${f}${cell.v === undefined ? '' : `<v>${esc(cell.v === true ? 1 : cell.v === false ? 0 : cell.v)}</v>`}</c>`
  }
  if (value == null) return ''
  if (typeof value === 'number') return `<c${r}${sAttr}><v>${value}</v></c>`
  if (typeof value === 'boolean') return `<c${r}${sAttr} t="b"><v>${value ? 1 : 0}</v></c>`
  if (value instanceof Date) return `<c${r}${sAttr}><v>${toSerial(value, ctx.date1904)}</v></c>`
  return `<c${r}${sAttr} t="s"><v>${ctx.shared(`<t xml:space="preserve">${esc(value)}</t>`)}</v></c>`
}

function richXML(runs) {
  return runs.map((run) => (run.font ? `<r>${fontXML(run.font, 'rPr', 'rFont')}` : '<r>') + `<t xml:space="preserve">${esc(run.text)}</t></r>`).join('')
}

/**
 * sheet: {
 *   name, state?: 'hidden',
 *   rows: Array<Array<cell>>   (row-major, starting at `origin`, default A1)
 *   origin?: 'J20',
 *   cols?: [{ min, max, width, hidden }],
 *   hiddenRows?: [rowNumber1Based], merges?: ['A1:B2'],
 *   hyperlinks?: [{ ref, url } | { ref, location }],
 *   cf?: (dxf) => xml   (build <conditionalFormatting> XML; dxf(style) returns an id)
 *   extra?: xml appended before </worksheet>
 * }
 */
function sheetXML(sheet, ctx) {
  const [oc, or] = (() => {
    const m = /^([A-Z]+)(\d+)$/.exec(sheet.origin || 'A1')
    let col = 0
    for (const ch of m[1]) col = col * 26 + ch.charCodeAt(0) - 64
    return [col - 1, +m[2] - 1]
  })()
  const hidden = new Set(sheet.hiddenRows || [])
  const rows = (sheet.rows || []).map((cells, i) => {
    const rn = or + i + 1
    const body = cells.map((cell, j) => cellXML(colName(oc + j) + rn, cell, ctx)).join('')
    const attrs = (ctx.noRefs ? '' : ` r="${rn}"`) + (hidden.has(rn) ? ' hidden="1"' : '')
    return body || hidden.has(rn) ? `<row${attrs}>${body}</row>` : ''
  }).join('')
  const cols = sheet.cols?.length
    ? '<cols>' + sheet.cols.map((c) => `<col min="${c.min}" max="${c.max ?? c.min}"${c.width != null ? ` width="${c.width}" customWidth="1"` : ''}${c.hidden ? ' hidden="1"' : ''}/>`).join('') + '</cols>'
    : ''
  const merges = sheet.merges?.length ? `<mergeCells count="${sheet.merges.length}">${sheet.merges.map((m) => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>` : ''
  const cf = sheet.cf ? sheet.cf((d) => ctx.styles.dxf(d)) : ''
  const rels = []
  const links = sheet.hyperlinks?.length
    ? '<hyperlinks>' + sheet.hyperlinks.map((h) => {
      if (h.location) return `<hyperlink ref="${h.ref}" location="${esc(h.location)}"/>`
      rels.push(h.url)
      return `<hyperlink ref="${h.ref}" r:id="rId${rels.length}"/>`
    }).join('') + '</hyperlinks>'
    : ''
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${MAIN}" xmlns:r="${REL}">${cols}<sheetData>${rows}</sheetData>${merges}${cf}${links}${sheet.extra || ''}</worksheet>`
  const relsXML = rels.length
    ? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PKG_REL}">${rels.map((u, i) => `<Relationship Id="rId${i + 1}" Type="${REL}/hyperlink" Target="${esc(u)}" TargetMode="External"/>`).join('')}</Relationships>`
    : null
  return { xml, relsXML }
}

/** Prefix SpreadsheetML elements (`<x:worksheet xmlns:x=…>`), as the Open XML SDK writes them. */
function prefixMain(xml) {
  return xml
    .replace(`xmlns="${MAIN}"`, `xmlns:x="${MAIN}"`)
    .replace(/<(\/?)(?![?!])([A-Za-z]+)(?=[\s/>])/g, '<$1x:$2')
}

/**
 * workbook: { sheets, activeTab?, date1904?, theme?: colors | false, styles?: false,
 *             sharedStrings?: false (use inline strings), noRefs?, prefixed?, store?, defaultFont? }
 */
export function writeXlsx(wb) {
  const styles = new StyleRegistry(wb.defaultFont || { sz: 11, color: { theme: 1 }, name: 'Calibri' })
  const sst = []
  const ctx = {
    styles,
    date1904: !!wb.date1904,
    noRefs: !!wb.noRefs,
    shared: (xml) => {
      let i = sst.indexOf(xml)
      if (i < 0) { sst.push(xml); i = sst.length - 1 }
      return i
    }
  }
  const files = {}
  const pre = wb.prefixed ? prefixMain : (x) => x
  const sheetRels = []
  wb.sheets.forEach((sheet, i) => {
    const { xml, relsXML } = sheetXML(sheet, ctx)
    files[`xl/worksheets/sheet${i + 1}.xml`] = pre(xml)
    if (relsXML) files[`xl/worksheets/_rels/sheet${i + 1}.xml.rels`] = relsXML
    sheetRels.push(`<Relationship Id="rId${i + 1}" Type="${REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`)
  })
  let n = wb.sheets.length
  const wbRels = [...sheetRels]
  if (wb.theme !== false) {
    files['xl/theme/theme1.xml'] = themeXML(wb.theme || OFFICE_THEME)
    wbRels.push(`<Relationship Id="rId${++n}" Type="${REL}/theme" Target="theme/theme1.xml"/>`)
  }
  if (wb.styles !== false) {
    files['xl/styles.xml'] = pre(styles.toXML())
    wbRels.push(`<Relationship Id="rId${++n}" Type="${REL}/styles" Target="styles.xml"/>`)
  }
  if (sst.length) {
    files['xl/sharedStrings.xml'] = pre(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="${MAIN}" count="${sst.length}" uniqueCount="${sst.length}">${sst.map((x) => `<si>${x}</si>`).join('')}</sst>`)
    wbRels.push(`<Relationship Id="rId${++n}" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/>`)
  }
  files['xl/workbook.xml'] = pre(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="${MAIN}" xmlns:r="${REL}">${wb.date1904 ? '<workbookPr date1904="1"/>' : ''}` +
    `<bookViews><workbookView activeTab="${wb.activeTab || 0}"/></bookViews><sheets>` +
    wb.sheets.map((s, i) => `<sheet name="${esc(s.name || `Лист${i + 1}`)}" sheetId="${i + 1}"${s.state ? ` state="${s.state}"` : ''} r:id="rId${i + 1}"/>`).join('') +
    '</sheets></workbook>')
  files['xl/_rels/workbook.xml.rels'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PKG_REL}">${wbRels.join('')}</Relationships>`
  files['_rels/.rels'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PKG_REL}"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`
  files['[Content_Types].xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>`
  return zip(files, { store: !!wb.store })
}

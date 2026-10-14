import { resolveOptions } from './options.js'

const CELL_STYLE_PROPS = [
  ['background', 'background-color'],
  ['textAlign', 'text-align'],
  ['verticalAlign', 'vertical-align'],
  ['borderTop', 'border-top'],
  ['borderRight', 'border-right'],
  ['borderBottom', 'border-bottom'],
  ['borderLeft', 'border-left']
]

function renderRuns(doc, parent, runs) {
  for (const run of runs) {
    if (run.br) { parent.appendChild(doc.createElement('br')); continue }
    const m = run.marks || {}
    let node = doc.createTextNode(run.text)
    const wrap = (tag, attrs) => {
      const el = doc.createElement(tag)
      if (attrs) for (const [k, v] of Object.entries(attrs)) el.setAttribute(k, v)
      el.appendChild(node)
      node = el
    }
    const css = []
    if (m.color) css.push(`color: ${m.color}`)
    if (m.fontSize) css.push(`font-size: ${m.fontSize}`)
    if (m.fontFamily) css.push(`font-family: ${m.fontFamily}`)
    if (css.length) wrap('span', { style: css.join('; ') })
    if (m.sub) wrap('sub')
    if (m.sup) wrap('sup')
    if (m.strike) wrap('s')
    if (m.underline) wrap('u')
    if (m.italic) wrap('em')
    if (m.bold) wrap('strong')
    if (m.link) wrap('a', { href: m.link })
    parent.appendChild(node)
  }
}

function renderParagraphs(doc, parent, content) {
  for (const runs of content) {
    const p = doc.createElement('p')
    renderRuns(doc, p, runs)
    parent.appendChild(p)
  }
}

/**
 * Render a table model as clean HTML: a <table> whose cells carry inline styles and
 * `data-colwidth`, or — for a single cell with `singleCell: 'inline'` — just its content.
 */
export function renderTableHTML(model, options) {
  const opts = resolveOptions(options)
  const doc = document.implementation.createHTMLDocument('')
  const root = doc.createElement('div')

  if (opts.singleCell === 'inline' && model.width === 1 && model.height === 1) {
    const content = model.rows[0][0].content
    if (content.length === 1) renderRuns(doc, root, content[0])
    else renderParagraphs(doc, root, content)
    return root.innerHTML
  }

  const table = doc.createElement('table')
  const tbody = doc.createElement('tbody')
  table.appendChild(tbody)
  root.appendChild(table)
  const occupied = []

  model.rows.forEach((row, r) => {
    const tr = doc.createElement('tr')
    let c = 0
    for (const cell of row) {
      while (occupied[r]?.[c]) c++
      for (let i = r; i < r + cell.rowspan; i++) for (let j = c; j < c + cell.colspan; j++) (occupied[i] ||= [])[j] = true
      const td = doc.createElement('td')
      if (cell.colspan > 1) td.setAttribute('colspan', String(cell.colspan))
      if (cell.rowspan > 1) td.setAttribute('rowspan', String(cell.rowspan))
      if (model.colWidths) {
        const widths = model.colWidths.slice(c, c + cell.colspan)
        if (widths.length === cell.colspan && widths.every((w) => w > 0)) td.setAttribute('data-colwidth', widths.join(','))
      }
      const css = CELL_STYLE_PROPS.filter(([k]) => cell.style[k]).map(([k, prop]) => `${prop}: ${cell.style[k]}`)
      if (css.length) td.setAttribute('style', css.join('; '))
      renderParagraphs(doc, td, cell.content.length ? cell.content : [[]])
      tr.appendChild(td)
      c += cell.colspan
    }
    tbody.appendChild(tr)
  })
  return root.innerHTML
}

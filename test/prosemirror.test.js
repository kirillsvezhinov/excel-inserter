import { describe, expect, it } from 'vitest'
import { Schema } from 'prosemirror-model'
import { EditorState, TextSelection } from 'prosemirror-state'
import { EditorView } from 'prosemirror-view'
import { schema as basic } from 'prosemirror-schema-basic'
import { tableEditing, tableNodes, CellSelection } from 'prosemirror-tables'
import { excelPaste, excelCellAttributes, excelMarks } from '../src/index.js'
import { fixture } from './helpers.js'

const schema = new Schema({
  nodes: basic.spec.nodes.append(tableNodes({
    tableGroup: 'block',
    cellContent: 'block+',
    cellAttributes: excelCellAttributes
  })),
  marks: basic.spec.marks.append(excelMarks)
})

const { table, table_row: row, table_cell: cell, paragraph: p, doc } = schema.nodes
const emptyTable = (rows, cols) =>
  table.create(null, Array.from({ length: rows }, () =>
    row.create(null, Array.from({ length: cols }, () => cell.create(null, p.create())))))

function makeView(docNode, options) {
  const state = EditorState.create({ doc: docNode, plugins: [excelPaste(options), tableEditing()] })
  return new EditorView(document.createElement('div'), { state })
}

// A real Ctrl+V: goes through the DOM paste handler (pasteText() would act like Ctrl+Shift+V).
function firePaste(view, { text = '', html = '' }) {
  const clipboardData = { getData: (type) => (type === 'text/html' ? html : type === 'text/plain' ? text : '') }
  view.dom.dispatchEvent(new ClipboardEvent('paste', { clipboardData, bubbles: true, cancelable: true }))
}

const cellTexts = (tableNode) => {
  const out = []
  tableNode.forEach((r) => { const line = []; r.forEach((c) => line.push(c.textContent)); out.push(line) })
  return out
}

describe('pasting into a ProseMirror editor', () => {
  it('pastes an Excel range as a styled table', () => {
    const view = makeView(doc.create(null, p.create()))
    view.pasteHTML(fixture('excel-windows-range'))
    const t = view.state.doc.firstChild
    expect(t.type.name).toBe('table')
    const header = t.firstChild.firstChild
    expect(header.attrs).toMatchObject({
      rowspan: 2, colwidth: [120], background: '#4472c4', textAlign: 'center',
      verticalAlign: 'middle', borderBottom: '2px solid #000000'
    })
    const text = header.firstChild.firstChild
    expect(text.text).toBe('Регион')
    expect(text.marks.map((m) => m.type.name).sort()).toEqual(['strong', 'textColor'])

    const rich = t.child(2).child(3).firstChild
    expect(rich.childCount).toBe(3)
    expect(rich.child(1).marks.map((m) => [m.type.name, m.attrs.color ?? null])).toEqual([['strong', null], ['textColor', '#c00000']])

    const link = t.child(3).child(3).firstChild.firstChild
    expect(link.marks[0].type.name).toBe('link')
    expect(t.child(3).child(0).firstChild.child(1).type.name).toBe('hard_break')

    const struck = t.child(4).child(2).firstChild.firstChild
    expect(struck.marks.map((m) => m.type.name).sort()).toEqual(['fontFamily', 'fontSize', 'strikethrough'])
  })

  it('round-trips cell attributes through toDOM and back', () => {
    const view = makeView(doc.create(null, p.create()))
    view.pasteHTML(fixture('excel-windows-range'))
    const header = view.state.doc.firstChild.firstChild.firstChild
    // The browser re-serializes colors as rgb(); parsing must still give back the same attrs.
    const dom = view.dom.querySelector('td')
    const parsed = schema.nodes.table_cell.spec.parseDOM[0].getAttrs(dom)
    for (const k of ['background', 'textAlign', 'verticalAlign', 'borderTop', 'borderRight', 'borderBottom', 'borderLeft']) {
      expect(parsed[k]).toBe(header.attrs[k])
    }
  })

  it('pastes a partial range into an existing table at the cursor cell', () => {
    const d = doc.create(null, emptyTable(3, 3))
    const view = makeView(d)
    const html = `<meta name=ProgId content=Excel.Sheet><style>td{font-family:Calibri}</style>
      <table><tr><td>a</td><td>b</td></tr><tr><td>c</td><td style="background:yellow">d</td></tr></table>`
    const cells = []
    d.descendants((node, pos) => { if (node.type === cell) cells.push(pos) })
    // Cursor in the middle cell (row 1, col 1).
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, cells[4] + 2)))
    view.pasteHTML(html)
    const t = view.state.doc.firstChild
    expect(cellTexts(t)).toEqual([['', '', ''], ['', 'a', 'b'], ['', 'c', 'd']])
    expect(t.child(2).child(2).attrs.background).toBe('#ffff00')
  })

  it('grows the table when the pasted range does not fit', () => {
    const d = doc.create(null, emptyTable(2, 2))
    const view = makeView(d)
    const cells = []
    d.descendants((node, pos) => { if (node.type === cell) cells.push(pos) })
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, cells[3] + 2)))
    view.pasteHTML(`<meta name=ProgId content=Excel.Sheet><table><tr><td>1</td><td>2</td></tr><tr><td>3</td><td>4</td></tr></table>`)
    expect(cellTexts(view.state.doc.firstChild)).toEqual([['', '', ''], ['', '1', '2'], ['', '3', '4']])
  })

  it('fills a CellSelection with the pasted range', () => {
    const d = doc.create(null, emptyTable(2, 2))
    const view = makeView(d)
    const cells = []
    d.descendants((node, pos) => { if (node.type === cell) cells.push(pos) })
    view.dispatch(view.state.tr.setSelection(CellSelection.create(view.state.doc, cells[0], cells[3])))
    view.pasteHTML(`<meta name=ProgId content=Excel.Sheet><table><tr><td>x</td></tr></table>`)
    expect(cellTexts(view.state.doc.firstChild)).toEqual([['x', 'x'], ['x', 'x']])
  })

  it('pastes a single cell inline into a paragraph', () => {
    const view = makeView(doc.create(null, p.create(null, schema.text('Say: '))))
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 6)))
    view.pasteHTML(fixture('excel-single-cell'))
    const para = view.state.doc.firstChild
    expect(view.state.doc.childCount).toBe(1)
    expect(para.textContent).toBe('Say: Hello')
    expect(para.lastChild.marks[0].type.name).toBe('strong')
  })

  it('turns TSV text into a table when enabled', () => {
    const view = makeView(doc.create(null, p.create()), { tsv: true })
    firePaste(view, { text: 'a\tb\n"multi\nline"\td\n' })
    const t = view.state.doc.firstChild
    expect(t.type.name).toBe('table')
    expect(cellTexts(t)).toEqual([['a', 'b'], ['multiline', 'd']])
    expect(t.child(1).child(0).firstChild.child(1).type.name).toBe('hard_break')
  })

  it('leaves TSV text alone by default', () => {
    const view = makeView(doc.create(null, p.create()))
    firePaste(view, { text: 'a\tb\nc\td' })
    expect(view.state.doc.firstChild.type.name).toBe('paragraph')
  })

  it('leaves TSV text alone on paste-as-plain-text (Ctrl+Shift+V)', () => {
    const view = makeView(doc.create(null, p.create()), { tsv: true })
    view.pasteText('a\tb\nc\td')
    expect(view.state.doc.firstChild.type.name).toBe('paragraph')
  })

  it('handles a real paste event carrying both HTML and TSV from Excel', () => {
    const view = makeView(doc.create(null, p.create()), { tsv: true })
    firePaste(view, { html: fixture('excel-windows-range'), text: 'ignored\tbecause html wins' })
    expect(view.state.doc.firstChild.firstChild.firstChild.attrs.background).toBe('#4472c4')
  })
})

import 'prosemirror-view/style/prosemirror.css'
import 'prosemirror-tables/style/tables.css'
import 'prosemirror-example-setup/style/style.css'
import { Schema } from 'prosemirror-model'
import { EditorState } from 'prosemirror-state'
import { EditorView } from 'prosemirror-view'
import { schema as basic } from 'prosemirror-schema-basic'
import { exampleSetup } from 'prosemirror-example-setup'
import { tableEditing, tableNodes, columnResizing, fixTables } from 'prosemirror-tables'
import { excelPaste, excelCellAttributes, excelMarks, transformSpreadsheetHTML } from '../src/index.js'

const schema = new Schema({
  nodes: basic.spec.nodes.append(tableNodes({
    tableGroup: 'block',
    cellContent: 'block+',
    cellAttributes: excelCellAttributes
  })),
  marks: basic.spec.marks.append(excelMarks)
})

let view
function createView(options) {
  const doc = view ? view.state.doc : schema.nodes.doc.create(null, schema.nodes.paragraph.create())
  view?.destroy()
  const plugins = [
    columnResizing(),
    tableEditing(),
    ...exampleSetup({ schema, menuBar: false })
  ]
  if (options) plugins.unshift(excelPaste(options))
  let state = EditorState.create({ doc, plugins })
  const fix = fixTables(state)
  if (fix) state = state.apply(fix.setMeta('addToHistory', false))
  view = new EditorView(document.querySelector('#editor'), { state })
  window.view = view
}

const enabled = document.querySelector('#enabled')
const tsv = document.querySelector('#tsv')
const rebuild = () => createView(enabled.checked ? { tsv: tsv.checked } : null)
enabled.addEventListener('change', rebuild)
tsv.addEventListener('change', rebuild)
rebuild()

// Capture the raw clipboard before ProseMirror sees it, to build real fixtures.
document.querySelector('#editor').addEventListener('paste', (e) => {
  const html = e.clipboardData.getData('text/html')
  document.querySelector('#raw').value = html || '(нет text/html)\n\n' + e.clipboardData.getData('text/plain')
  document.querySelector('#out').value = html ? transformSpreadsheetHTML(html, { tsv: tsv.checked }) : ''
}, true)

document.querySelector('#save-raw').addEventListener('click', () => {
  const blob = new Blob([document.querySelector('#raw').value], { type: 'text/html' })
  const a = Object.assign(document.createElement('a'), { href: URL.createObjectURL(blob), download: 'clipboard.html' })
  a.click()
  URL.revokeObjectURL(a.href)
})

document.querySelector('#insert-table').addEventListener('click', () => {
  const { table, table_row, table_cell, paragraph } = schema.nodes
  const t = table.create(null, Array.from({ length: 4 }, () =>
    table_row.create(null, Array.from({ length: 4 }, () => table_cell.create(null, paragraph.create())))))
  view.dispatch(view.state.tr.replaceSelectionWith(t))
  view.focus()
})

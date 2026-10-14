import { Plugin, PluginKey } from 'prosemirror-state'
import { DOMParser as PMDOMParser } from 'prosemirror-model'
import { resolveOptions } from './options.js'
import { parseSpreadsheetHTML } from './parse.js'
import { renderTableHTML } from './render.js'
import { looksLikeTSV, parseTSV, tsvToModel } from './tsv.js'

export const excelPasteKey = new PluginKey('excelPaste')

/**
 * Rewrite spreadsheet clipboard HTML into clean, inline-styled table HTML.
 * Returns the input unchanged for anything that isn't a spreadsheet table, and on any error.
 */
export function transformSpreadsheetHTML(html, options) {
  try {
    const model = parseSpreadsheetHTML(html, options)
    return model ? renderTableHTML(model, options) : html
  } catch (err) {
    if (typeof console !== 'undefined') console.warn('[prosemirror-excel-paste] falling back to default paste:', err)
    return html
  }
}

/**
 * ProseMirror plugin. Put it alongside `tableEditing()`: this plugin only normalizes the
 * pasted HTML; prosemirror-tables then inserts a partial range into an existing table
 * (at the cursor cell or over a CellSelection) or pastes a new table elsewhere.
 */
export function excelPaste(options = {}) {
  const opts = resolveOptions(options)
  return new Plugin({
    key: excelPasteKey,
    props: {
      transformPastedHTML(html) {
        return transformSpreadsheetHTML(html, opts)
      },
      clipboardTextParser(text, $context, plain, view) {
        if (!opts.tsv || plain) return null
        const rows = parseTSV(text)
        if (!looksLikeTSV(rows)) return null
        const html = renderTableHTML(tsvToModel(rows), { ...opts, singleCell: 'table' })
        const dom = document.createElement('div')
        dom.innerHTML = html
        return PMDOMParser.fromSchema(view.state.schema).parseSlice(dom, { context: $context })
      }
    }
  })
}

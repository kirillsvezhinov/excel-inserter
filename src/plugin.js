import { Plugin, PluginKey } from 'prosemirror-state'
import { DOMParser as PMDOMParser } from 'prosemirror-model'
import { resolveOptions } from './options.js'
import { parseSpreadsheetHTML } from './parse.js'
import { renderTableHTML } from './render.js'
import { looksLikeTSV, parseTSV, tsvToModel } from './tsv.js'
import { readSpreadsheetFile, spreadsheetFileKind } from './files/index.js'

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
 * ProseMirror plugin. List it before `tableEditing()`: it normalizes pasted spreadsheet
 * HTML and reads pasted .xlsx/.csv files; prosemirror-tables then inserts the result into
 * an existing table (at the cursor cell or over a CellSelection) or pastes a new table.
 */
export function excelPaste(options = {}) {
  const opts = resolveOptions(options)
  // Set while re-pasting HTML we rendered from a file, so it isn't transformed twice.
  let pastingRendered = false

  return new Plugin({
    key: excelPasteKey,
    props: {
      transformPastedHTML(html) {
        return pastingRendered ? html : transformSpreadsheetHTML(html, opts)
      },

      // Files on the clipboard (an .xlsx/.csv copied in Finder/Explorer). Must run before
      // tableEditing's handlePaste, so list excelPaste() first.
      handlePaste(view, event) {
        if (!opts.files.enabled) return false
        const file = [...(event.clipboardData?.files || [])].find((f) => spreadsheetFileKind(f))
        if (!file) return false
        pasteSpreadsheetFile(view, file, opts, (html) => {
          pastingRendered = true
          try { view.pasteHTML(html) } finally { pastingRendered = false }
        })
        return true
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

/**
 * Read a spreadsheet file and insert it at the current selection, through the normal
 * paste path (so prosemirror-tables can merge it into a table under the cursor).
 * The callbacks in `options.files` report progress; errors never throw.
 */
export async function pasteSpreadsheetFile(view, file, options, insert) {
  const opts = resolveOptions(options)
  const { onLoadStart, onLoad, onError } = opts.files
  onLoadStart?.(file)
  try {
    const sheet = await readSpreadsheetFile(file, opts)
    if (view.isDestroyed) return
    if (!sheet) throw new Error('The file contains no data')
    const html = renderTableHTML(sheet.model, opts)
    if (insert) insert(html)
    else view.pasteHTML(html)
    onLoad?.(file, sheet)
  } catch (err) {
    if (onError) onError(err, file)
    else if (typeof console !== 'undefined') console.warn('[prosemirror-excel-paste] could not read', file?.name, err)
  }
}

import { useEffect, useRef } from 'react'
import { Schema } from 'prosemirror-model'
import { EditorState } from 'prosemirror-state'
import { EditorView } from 'prosemirror-view'
import { keymap } from 'prosemirror-keymap'
import { baseKeymap } from 'prosemirror-commands'
import { history, undo, redo } from 'prosemirror-history'
import { schema as basic } from 'prosemirror-schema-basic'
import { tableEditing, tableNodes, columnResizing, goToNextCell } from 'prosemirror-tables'
import { excelPaste, excelCellAttributes, excelMarks } from '../../src/index.js'

export const schema = new Schema({
  nodes: basic.spec.nodes.append(tableNodes({
    tableGroup: 'block',
    cellContent: 'block+',
    cellAttributes: excelCellAttributes
  })),
  marks: basic.spec.marks.append(excelMarks)
})

function buildPlugins(options) {
  return [
    ...(options ? [excelPaste(options)] : []),
    columnResizing(),
    tableEditing(),
    history(),
    keymap({
      'Mod-z': undo,
      'Mod-y': redo,
      'Shift-Mod-z': redo,
      Tab: goToNextCell(1),
      'Shift-Tab': goToNextCell(-1)
    }),
    keymap(baseKeymap)
  ]
}

/**
 * ProseMirror owns its DOM; React only mounts it once and swaps plugins when options change.
 * `onPaste` gets the raw clipboard (captured before ProseMirror), `onChange` the new doc.
 */
export function Editor({ options, viewRef, onPaste, onChange }) {
  const mountRef = useRef(null)
  const callbacks = useRef({ onPaste, onChange })
  callbacks.current = { onPaste, onChange }

  useEffect(() => {
    const view = new EditorView(mountRef.current, {
      state: EditorState.create({
        doc: schema.nodes.doc.create(null, schema.nodes.paragraph.create()),
        plugins: buildPlugins(options)
      }),
      dispatchTransaction(tr) {
        view.updateState(view.state.apply(tr))
        if (tr.docChanged) callbacks.current.onChange?.(view.state.doc)
      }
    })
    const capture = (e) => callbacks.current.onPaste?.({
      html: e.clipboardData?.getData('text/html') || '',
      text: e.clipboardData?.getData('text/plain') || ''
    })
    view.dom.addEventListener('paste', capture, true)
    viewRef.current = view
    return () => {
      view.dom.removeEventListener('paste', capture, true)
      view.destroy()
      viewRef.current = null
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  useEffect(() => {
    const view = viewRef.current
    if (view) view.updateState(view.state.reconfigure({ plugins: buildPlugins(options) }))
  }, [options, viewRef])

  return <div className="editor" ref={mountRef} />
}

import { useMemo, useRef, useState } from 'react'
import { Editor, schema } from './Editor.jsx'
import { transformSpreadsheetHTML, parseSpreadsheetHTML, renderTableHTML } from '../../src/index.js'

const STYLE_KEYS = ['background', 'borders', 'textAlign', 'verticalAlign', 'color', 'fontSize', 'fontFamily', 'links']

export function App() {
  const viewRef = useRef(null)
  const [enabled, setEnabled] = useState(true)
  const [singleCell, setSingleCell] = useState('inline')
  const [lineBreaks, setLineBreaks] = useState('br')
  const [tsv, setTsv] = useState(false)
  const [styles, setStyles] = useState(() => Object.fromEntries(STYLE_KEYS.map((k) => [k, true])))
  const [clip, setClip] = useState(null)
  const [docJSON, setDocJSON] = useState('')
  const [tab, setTab] = useState('raw')
  const [cf, setCf] = useState(true)
  // { state: 'loading' | 'done' | 'error', name, message?, sheet? }
  const [fileStatus, setFileStatus] = useState(null)

  const options = useMemo(
    () => (enabled ? {
      singleCell, lineBreaks, tsv, styles,
      files: {
        conditionalFormatting: cf,
        onLoadStart: (file) => setFileStatus({ state: 'loading', name: file.name }),
        onLoad: (file, sheet) => {
          setFileStatus({ state: 'done', name: file.name, sheet })
          setClip({ file: file.name, sheet })
        },
        onError: (err, file) => setFileStatus({ state: 'error', name: file.name, message: err.message })
      }
    } : null),
    [enabled, singleCell, lineBreaks, tsv, styles, cf]
  )

  const result = useMemo(() => {
    if (clip?.sheet && options) {
      return { model: clip.sheet.model, html: renderTableHTML(clip.sheet.model, options) }
    }
    if (!clip?.html || !options) return null
    return {
      model: parseSpreadsheetHTML(clip.html, options),
      html: transformSpreadsheetHTML(clip.html, options)
    }
  }, [clip, options])

  const insertTable = () => {
    const view = viewRef.current
    const { table, table_row, table_cell, paragraph } = schema.nodes
    const t = table.create(null, Array.from({ length: 4 }, () =>
      table_row.create(null, Array.from({ length: 4 }, () => table_cell.create(null, paragraph.create())))))
    view.dispatch(view.state.tr.replaceSelectionWith(t).scrollIntoView())
    view.focus()
  }

  const clear = () => {
    const view = viewRef.current
    view.dispatch(view.state.tr.delete(0, view.state.doc.content.size))
    view.focus()
  }

  const saveRaw = () => {
    if (!clip?.html) return
    const url = URL.createObjectURL(new Blob([clip.html], { type: 'text/html' }))
    Object.assign(document.createElement('a'), { href: url, download: `clipboard-${result?.model?.source ?? 'raw'}.html` }).click()
    URL.revokeObjectURL(url)
  }

  const panel = {
    raw: clip?.sheet ? `(файл ${clip.file} — читается напрямую, без text/html)` : clip ? clip.html || `(нет text/html)\n\n${clip.text}` : '',
    out: result?.html ?? '',
    model: result?.model ? JSON.stringify(result.model, null, 2) : '',
    doc: docJSON
  }[tab]

  return (
    <div className="app">
      <header>
        <h1>Excel → ProseMirror</h1>
        <p>
          Скопируйте диапазон в Excel / Google Sheets / LibreOffice и нажмите <kbd>Ctrl</kbd>+<kbd>V</kbd>.
          Курсор внутри таблицы — диапазон встанет в неё, начиная с этой ячейки; выделите ячейки
          мышью — заполнит выделение. Можно скопировать сам файл <code>.xlsx</code> / <code>.csv</code> в
          Finder или Проводнике и вставить его — из книги берётся только первый лист.
        </p>
      </header>

      <div className="toolbar">
        <label><input type="checkbox" checked={enabled} onChange={(e) => setEnabled(e.target.checked)} /> плагин</label>
        <label>
          одна ячейка:{' '}
          <select value={singleCell} onChange={(e) => setSingleCell(e.target.value)} disabled={!enabled}>
            <option value="inline">как текст</option>
            <option value="table">как таблица 1×1</option>
          </select>
        </label>
        <label>
          Alt+Enter:{' '}
          <select value={lineBreaks} onChange={(e) => setLineBreaks(e.target.value)} disabled={!enabled}>
            <option value="br">перенос строки</option>
            <option value="paragraph">абзацы</option>
          </select>
        </label>
        <label><input type="checkbox" checked={tsv} disabled={!enabled} onChange={(e) => setTsv(e.target.checked)} /> TSV → таблица</label>
        <label><input type="checkbox" checked={cf} disabled={!enabled} onChange={(e) => setCf(e.target.checked)} /> условное форматирование</label>
        <span className="spacer" />
        <button onClick={insertTable}>Пустая таблица 4×4</button>
        <button onClick={clear}>Очистить</button>
      </div>

      <div className="toolbar styles">
        {STYLE_KEYS.map((k) => (
          <label key={k}>
            <input
              type="checkbox"
              checked={styles[k]}
              disabled={!enabled}
              onChange={(e) => setStyles((s) => ({ ...s, [k]: e.target.checked }))}
            />{' '}
            {k}
          </label>
        ))}
      </div>

      {fileStatus && (
        <div className={`file-status ${fileStatus.state}`}>
          {fileStatus.state === 'loading' && <>Читаю <b>{fileStatus.name}</b>…</>}
          {fileStatus.state === 'done' && (
            <>
              Вставлен <b>{fileStatus.name}</b>: лист «{fileStatus.sheet.name}», {fileStatus.sheet.model.height}×{fileStatus.sheet.model.width}
              {fileStatus.sheet.model.truncated && ' (обрезано)'}
            </>
          )}
          {fileStatus.state === 'error' && <>Не удалось прочитать <b>{fileStatus.name}</b>: {fileStatus.message}</>}
          <button onClick={() => setFileStatus(null)} aria-label="Скрыть">×</button>
        </div>
      )}

      <Editor
        options={options}
        viewRef={viewRef}
        onPaste={(c) => { if (c.html || c.text) setClip(c) }}
        onChange={(doc) => setDocJSON(JSON.stringify(doc.toJSON(), null, 2))}
      />

      <section className="inspector">
        <div className="tabs">
          {[['raw', 'Буфер: text/html'], ['out', 'После преобразования'], ['model', 'Модель'], ['doc', 'Документ PM']].map(([id, label]) => (
            <button key={id} className={tab === id ? 'active' : ''} onClick={() => setTab(id)}>{label}</button>
          ))}
          <span className="spacer" />
          {result?.model && <span className="badge">{result.model.sheet ?? result.model.source} · {result.model.height}×{result.model.width}</span>}
          <button onClick={saveRaw} disabled={!clip?.html}>Скачать как фикстуру</button>
        </div>
        <pre>{panel || 'Вставьте что-нибудь в редактор.'}</pre>
      </section>
    </div>
  )
}

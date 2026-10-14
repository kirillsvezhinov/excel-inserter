import { describe, expect, it } from 'vitest'
import zlib from 'node:zlib'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { Schema } from 'prosemirror-model'
import { EditorState, TextSelection } from 'prosemirror-state'
import { EditorView } from 'prosemirror-view'
import { schema as basic } from 'prosemirror-schema-basic'
import { tableEditing, tableNodes } from 'prosemirror-tables'
import { excelPaste, excelCellAttributes, excelMarks, readSpreadsheetFile, spreadsheetFileKind } from '../src/index.js'
import { readZip } from '../src/files/zip.js'
import { inflateRaw } from '../src/files/inflate.js'
import { decodeText, detectDelimiter } from '../src/files/csv.js'
import { localeInfo } from '../src/files/numfmt.js'
import { buildWorkbook, zip } from './xlsx-builder.js'

const G = localeInfo('ru-RU').group // U+00A0 in ru-RU, like Excel

const texts = (model) => model.rows.map((r) => r.map((c) => c.content.map((p) => p.map((x) => (x.br ? '\n' : x.text)).join('')).join('¶')))
const ru = { kind: 'xlsx', files: { locale: 'ru-RU' } }

describe('zip + inflate', () => {
  it('reads stored and deflated entries', async () => {
    for (const store of [false, true]) {
      const entries = readZip(zip({ 'a.txt': 'hello', 'dir/b.xml': '<x>Привет</x>'.repeat(100) }, { store }))
      expect(new TextDecoder().decode(await entries.get('a.txt')())).toBe('hello')
      expect(new TextDecoder().decode(await entries.get('dir/b.xml')())).toBe('<x>Привет</x>'.repeat(100))
    }
  })

  it('pure-JS inflate matches zlib (stored, fixed and dynamic blocks)', () => {
    const data = Buffer.from('Excel '.repeat(5000) + Math.random().toString(36).repeat(300))
    for (const opts of [{ level: 0 }, { strategy: zlib.constants.Z_FIXED }, { level: 9 }]) {
      expect(Buffer.from(inflateRaw(new Uint8Array(zlib.deflateRawSync(data, opts)))).equals(data)).toBe(true)
    }
  })

  it('rejects non-zip input', () => {
    expect(() => readZip(new Uint8Array([1, 2, 3]))).toThrow(/Not a ZIP/)
  })
})

describe('xlsx', () => {
  it('reads the first sheet with layout: merges, hidden rows/cols, widths', async () => {
    const { name, model } = await readSpreadsheetFile(buildWorkbook(), ru)
    expect(name).toBe('Отчёт')
    expect(model.width).toBe(5) // B..G minus hidden E
    expect(model.height).toBe(6) // rows 2..8 minus hidden row 4
    expect(model.colWidths).toEqual([131, 96, 96, 64, 64])
    expect(model.rows[0][1].colspan).toBe(2)
    expect(model.rows[4][3]).toMatchObject({ colspan: 2, rowspan: 2 })
    expect(texts(model).flat().join()).not.toMatch(/Скрытая|ID|^17$/)
  })

  it('formats values like Excel in the viewer locale', async () => {
    const { model } = await readSpreadsheetFile(buildWorkbook(), ru)
    const t = texts(model)
    expect(t[1]).toEqual(['Москва', `1${G}234,50 ₽`, `2${G}469,00 ₽`, '-12,00%', '15.03.2023'])
    expect(t[3].slice(1, 4)).toEqual(['-500', 'ИСТИНА', '#DIV/0!'])
    expect(t[4][2]).toBe('3,14159')

    const { model: en } = await readSpreadsheetFile(buildWorkbook(), { kind: 'xlsx', files: { locale: 'en-US' } })
    expect(texts(en)[1]).toEqual(['Москва', '1,234.50 ₽', '2,469.00 ₽', '-12.00%', '15.03.2023'])
    expect(texts(en)[2][4]).toBe('03/15/2023') // built-in short date follows the locale
  })

  it('resolves theme colors, tints, fonts, borders and alignment', async () => {
    const { model } = await readSpreadsheetFile(buildWorkbook(), ru)
    const header = model.rows[0][0]
    expect(header.style).toEqual({
      background: '#4472c4', textAlign: 'center', verticalAlign: 'middle',
      borderTop: '1px solid #000000', borderRight: '1px solid #000000',
      borderBottom: '2px solid #000000', borderLeft: '1px solid #000000'
    })
    expect(header.content[0][0].marks).toEqual({ bold: true, color: '#ffffff' })
    expect(model.rows[4][2].style.background).toBe('#dae3f3') // accent1 + tint 0.8
    expect(model.rows[4][1].content[0][0].marks).toEqual({ italic: true, color: '#ff0000', fontSize: '14pt', fontFamily: 'Times New Roman' })
    expect(model.rows[1][0].style.verticalAlign).toBe('top')
    expect(model.rows[3][2].style.textAlign).toBe('center') // booleans centered under General
  })

  it('keeps rich text, line breaks, [Red] formats and links', async () => {
    const { model } = await readSpreadsheetFile(buildWorkbook(), ru)
    expect(model.rows[3][0].content[0]).toEqual([
      { text: 'Итого: ', marks: {} },
      { text: 'важно', marks: { bold: true, color: '#c00000' } },
      { text: ' проверить', marks: {} }
    ])
    expect(model.rows[4][0].content[0]).toEqual([{ text: 'Строка 1', marks: {} }, { br: true }, { text: 'Строка 2', marks: {} }])
    expect(model.rows[3][1].content[0][0].marks).toEqual({ color: '#ff0000' })
    expect(model.rows[3][4].content[0]).toEqual([{ text: 'отчёт', marks: { link: 'https://example.com/report?id=1' } }])
    expect(texts(model)[2][0]).toBe('Санкт-Петербург') // in-workbook link dropped
    expect(texts(model)[4][3]).toBe('H2O') // stray _x000D_ trimmed
  })

  it('evaluates conditional formatting', async () => {
    const { model } = await readSpreadsheetFile(buildWorkbook(), ru)
    // cellIs >= $C$5 (green, bold) on C5; cellIs < 100 (red) on C8; formula rules are skipped.
    expect(model.rows[2][1].style.background).toBe('#c6efce')
    expect(model.rows[2][1].content[0][0].marks).toEqual({ bold: true, color: '#006100' })
    expect(model.rows[5][1].style.background).toBe('#ffc7ce')
    expect(model.rows[1][1].style.background).toBeUndefined()
    // 2-color scale min → red, max → green
    expect(model.rows[1][3].style.background).toBe('#f8696b')
    expect(model.rows[2][3].style.background).toBe('#63be7b')
    // containsText (case-insensitive)
    expect(model.rows[5][0].style.background).toBe('#c6efce')
  })

  it('can skip conditional formatting', async () => {
    const { model } = await readSpreadsheetFile(buildWorkbook(), { ...ru, files: { ...ru.files, conditionalFormatting: false } })
    expect(model.rows[2][1].style.background).toBeUndefined()
  })

  it('reads only the first sheet, even when another tab is active', async () => {
    const { name, model } = await readSpreadsheetFile(buildWorkbook({ activeTab: 1 }), { kind: 'xlsx' })
    expect(name).toBe('Отчёт')
    expect(texts(model)[0][0]).toBe('Регион')
  })

  it('never parses the other sheets', async () => {
    const broken = buildWorkbook({ parts: { 'xl/worksheets/sheet2.xml': '<<< not xml', 'xl/worksheets/sheet3.xml': '' } })
    const { name } = await readSpreadsheetFile(broken, { kind: 'xlsx' })
    expect(name).toBe('Отчёт')
  })

  it('skips a hidden first sheet', async () => {
    const { name, model } = await readSpreadsheetFile(buildWorkbook({ hideFirst: true }), { kind: 'xlsx' })
    expect(name).toBe('Второй')
    expect(texts(model)).toEqual([['Второй лист', '42']])
  })

  it('truncates huge sheets', async () => {
    const { model } = await readSpreadsheetFile(buildWorkbook(), { kind: 'xlsx', files: { maxCells: 12 } })
    expect(model.truncated).toBe(true)
    expect(model.height).toBeLessThan(6)
  })

  it('explains legacy .xls', async () => {
    await expect(readSpreadsheetFile(new Uint8Array([0xd0, 0xcf, 0x11, 0xe0]), { kind: 'xlsx' })).rejects.toThrow(/\.xls/)
  })
})

describe('csv', () => {
  const enc = (s, encoding) => new Uint8Array(Buffer.from(s, encoding))

  it('detects encodings', () => {
    expect(decodeText(new Uint8Array([0xef, 0xbb, 0xbf, ...Buffer.from('Привет')]))).toBe('Привет')
    expect(decodeText(enc('Привет', 'utf8'))).toBe('Привет')
    const cp1251 = new Uint8Array([0xcf, 0xf0, 0xe8, 0xe2, 0xe5, 0xf2]) // "Привет" in windows-1251
    expect(decodeText(cp1251, { locale: 'ru-RU' })).toBe('Привет')
    expect(decodeText(new Uint8Array([0xe9]), { locale: 'fr-FR' })).toBe('é')
  })

  it('detects delimiters, including Excel sep= hint', () => {
    expect(detectDelimiter('a;b;c\n1;2,5;3\n')).toBe(';')
    expect(detectDelimiter('a,b,c\n1,2,3\n')).toBe(',')
    expect(detectDelimiter('a\tb\n1\t2\n')).toBe('\t')
    expect(detectDelimiter('sep=|\na|b\n')).toBe('|')
  })

  it('reads a Russian Excel CSV and right-aligns numbers', async () => {
    const file = new File(['Город;Сумма\nМосква;"1 234,50"\n"Санкт-\nПетербург";7\n'], 'data.csv', { type: 'text/csv' })
    const { model } = await readSpreadsheetFile(file)
    expect(texts(model)).toEqual([['Город', 'Сумма'], ['Москва', '1 234,50'], ['Санкт-\nПетербург', '7']])
    expect(model.rows[1][1].style).toEqual({ textAlign: 'right' })
  })

  it('reads a windows-1251 CSV saved by Russian Excel', async () => {
    const bytes = new Uint8Array(readFileSync(join(process.cwd(), 'test/fixtures/sample-cp1251.csv')))
    const { model } = await readSpreadsheetFile(new File([bytes], 'sample-cp1251.csv'), { files: { locale: 'ru-RU' } })
    expect(texts(model)).toEqual([['Город', 'Сумма', 'Дата'], ['Москва', '1 234,50', '14.10.2026'], ['Санкт-\nПетербург', '7', '01.01.2026']])
  })

  it('recognises file kinds', () => {
    expect(spreadsheetFileKind({ name: 'Report.XLSX' })).toBe('xlsx')
    expect(spreadsheetFileKind({ name: 'x', type: 'text/csv' })).toBe('csv')
    expect(spreadsheetFileKind({ name: 'image.png', type: 'image/png' })).toBe(null)
  })
})

describe('pasting a file from the clipboard', () => {
  const schema = new Schema({
    nodes: basic.spec.nodes.append(tableNodes({ tableGroup: 'block', cellContent: 'block+', cellAttributes: excelCellAttributes })),
    marks: basic.spec.marks.append(excelMarks)
  })
  const { doc, paragraph: p, table, table_row: row, table_cell: cell } = schema.nodes

  function setup(docNode, files = {}) {
    let done
    const loaded = new Promise((resolve, reject) => { done = { resolve, reject } })
    const plugins = [
      excelPaste({ files: { locale: 'ru-RU', onLoad: (f, sheet) => done.resolve(sheet), onError: (e) => done.reject(e), ...files } }),
      tableEditing()
    ]
    const view = new EditorView(document.createElement('div'), { state: EditorState.create({ doc: docNode, plugins }) })
    return { view, loaded }
  }
  function pasteFiles(view, files, extra = {}) {
    const clipboardData = {
      files,
      getData: (type) => extra[type] || ''
    }
    const event = new ClipboardEvent('paste', { clipboardData, bubbles: true, cancelable: true })
    view.dom.dispatchEvent(event)
    return event
  }

  it('inserts an .xlsx copied in the file manager as a styled table', async () => {
    const { view, loaded } = setup(doc.create(null, p.create()))
    const file = new File([buildWorkbook()], 'Отчёт.xlsx')
    const event = pasteFiles(view, [file], { 'text/plain': 'Отчёт.xlsx' })
    expect(event.defaultPrevented).toBe(true)
    await loaded
    const t = view.state.doc.firstChild
    expect(t.type.name).toBe('table')
    expect(t.firstChild.firstChild.attrs).toMatchObject({ background: '#4472c4', colwidth: [131] })
    expect(t.child(1).child(1).textContent).toBe(`1${G}234,50 ₽`)
  })

  it('merges a CSV into the table under the cursor', async () => {
    const empty = () => cell.create(null, p.create())
    const d = doc.create(null, table.create(null, [row.create(null, [empty(), empty()]), row.create(null, [empty(), empty()])]))
    const { view, loaded } = setup(d)
    const cells = []
    d.descendants((n, pos) => { if (n.type === cell) cells.push(pos) })
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, cells[3] + 2)))
    pasteFiles(view, [new File(['a,b\n1,2\n'], 'x.csv', { type: 'text/csv' })])
    await loaded
    const out = []
    view.state.doc.firstChild.forEach((r) => { const line = []; r.forEach((c) => line.push(c.textContent)); out.push(line) })
    expect(out).toEqual([['', '', ''], ['', 'a', 'b'], ['', '1', '2']])
  })

  it('ignores non-spreadsheet files (Excel also puts a PNG of the range on the clipboard)', () => {
    const { view } = setup(doc.create(null, p.create()))
    const event = pasteFiles(view, [new File(['x'], 'image.png', { type: 'image/png' })], { 'text/plain': 'hi' })
    expect(event.defaultPrevented).toBe(true) // handled by the normal text paste
    expect(view.state.doc.textContent).toBe('hi')
  })

  it('reports unreadable files through onError', async () => {
    const { view, loaded } = setup(doc.create(null, p.create()))
    pasteFiles(view, [new File(['not a zip'], 'broken.xlsx')])
    await expect(loaded).rejects.toThrow(/Not an \.xlsx/)
    expect(view.state.doc.textContent).toBe('')
  })

  it('intercepts .xls and reports it instead of pasting the file name', async () => {
    const { view, loaded } = setup(doc.create(null, p.create()))
    pasteFiles(view, [new File([new Uint8Array([0xd0, 0xcf, 0x11, 0xe0])], 'old.xls')], { 'text/plain': 'old.xls' })
    await expect(loaded).rejects.toThrow(/save it as \.xlsx/)
    expect(view.state.doc.textContent).toBe('')
  })

  it('can be disabled', () => {
    const { view } = setup(doc.create(null, p.create()), { enabled: false })
    pasteFiles(view, [new File([buildWorkbook()], 'a.xlsx')], { 'text/plain': 'a.xlsx' })
    expect(view.state.doc.textContent).toBe('a.xlsx')
  })
})

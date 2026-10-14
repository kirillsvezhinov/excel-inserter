import { describe, expect, it } from 'vitest'
import { parseSpreadsheetHTML, renderTableHTML, transformSpreadsheetHTML, detectSource } from '../src/index.js'
import { fixture } from './helpers.js'

const texts = (cell) => cell.content.map((p) => p.map((r) => (r.br ? '\n' : r.text)).join(''))

describe('Excel (Windows) range', () => {
  const model = parseSpreadsheetHTML(fixture('excel-windows-range'))

  it('detects Excel and drops hidden rows and columns', () => {
    expect(model.source).toBe('excel')
    expect(model.width).toBe(4)
    expect(model.height).toBe(5)
    expect(model.rows.flat().some((c) => texts(c).join().includes('Скрытая'))).toBe(false)
    expect(model.rows.flat().some((c) => texts(c).join() === '17')).toBe(false)
  })

  it('keeps column widths in px, minus the hidden column', () => {
    expect(model.colWidths).toEqual([120, 96, 96, 128])
  })

  it('keeps merges', () => {
    const [region, revenue] = model.rows[0]
    expect(region.rowspan).toBe(2)
    expect(revenue.colspan).toBe(2)
    expect(model.rows[1].map((c) => texts(c)[0])).toEqual(['Q1', 'Q2', ''])
  })

  it('resolves class styles into cell styles and marks', () => {
    const header = model.rows[0][0]
    expect(header.style).toMatchObject({
      background: '#4472c4',
      textAlign: 'center',
      verticalAlign: 'middle',
      borderTop: '1px solid #000000',
      borderBottom: '2px solid #000000'
    })
    expect(header.content[0][0].marks).toEqual({ bold: true, color: '#ffffff' })
  })

  it('right-aligns numbers (align=right beats text-align:general) and keeps the displayed locale text', () => {
    const cell = model.rows[2][1]
    expect(texts(cell)).toEqual(['1 234,50 ₽'])
    expect(cell.style.textAlign).toBe('right')
  })

  it('keeps conditional-format colors that Excel bakes into the copy', () => {
    const cell = model.rows[2][2]
    expect(cell.style.background).toBe('#ffc7ce')
    expect(cell.content[0][0].marks).toEqual({ color: '#9c0006' })
  })

  it('keeps rich text runs from <font class=…>', () => {
    expect(model.rows[2][3].content[0]).toEqual([
      { text: 'Итого: ', marks: {} },
      { text: 'важно', marks: { bold: true, color: '#c00000' } },
      { text: ' проверить', marks: {} }
    ])
  })

  it('does not emit the workbook default font (Calibri 11pt black)', () => {
    expect(model.rows[2][0].content[0]).toEqual([{ text: 'Москва', marks: {} }])
  })

  it('turns in-cell line breaks into <br>', () => {
    expect(model.rows[3][0].content[0]).toEqual([
      { text: 'Санкт-', marks: {} }, { br: true }, { text: 'Петербург', marks: {} }
    ])
  })

  it('keeps links but drops Excel hyperlink styling', () => {
    expect(model.rows[3][3].content[0]).toEqual([{ text: 'отчёт', marks: { link: 'https://example.com/report?id=1' } }])
  })

  it('splits mso-ignore:colspan overflow back into separate cells', () => {
    const row = model.rows[4]
    expect(row.map((c) => c.colspan)).toEqual([1, 1, 1, 1])
    expect(texts(row[0])[0]).toBe('Длинный текст, который вылезает в соседние ячейки')
    expect(row[0].content[0][0].marks).toEqual({ italic: true, underline: true })
    expect(row[2].content[0][0].marks).toEqual({ strike: true, fontSize: '14pt', fontFamily: '"Times New Roman", serif' })
  })
})

describe('partial and irregular selections', () => {
  const wrap = (body, style = '') => `<html xmlns:x="urn:schemas-microsoft-com:office:excel"><head>
    <meta name=ProgId content=Excel.Sheet><style><!--td {color:black;font-size:11.0pt;font-family:Calibri;}${style}--></style></head>
    <body><table>${body}</table></body></html>`

  it('pads ragged rows to a rectangle', () => {
    const m = parseSpreadsheetHTML(wrap('<tr><td>a</td><td>b</td><td>c</td></tr><tr><td>d</td></tr>'))
    expect(m.rows.map((r) => r.length)).toEqual([3, 3])
  })

  it('handles a merged cell cut by the selection edge', () => {
    // Selecting B1:C2 when A1:B1 is merged: Excel sends the visible part of the merge.
    const m = parseSpreadsheetHTML(wrap('<tr><td rowspan=2>m</td><td>x</td></tr><tr><td>y</td></tr>'))
    expect(m.width).toBe(2)
    expect(m.rows[0][0].rowspan).toBe(2)
    expect(m.rows[1]).toHaveLength(1)
  })

  it('handles rowspan=0 and absurd spans', () => {
    const m = parseSpreadsheetHTML(wrap('<tr><td rowspan=0>a</td><td>b</td></tr><tr><td>c</td></tr>'))
    expect(m.rows[0][0].rowspan).toBe(2)
  })

  it('renders a single cell inline by default', () => {
    const html = transformSpreadsheetHTML(fixture('excel-single-cell'))
    expect(html).toBe('<strong>Hello</strong>')
  })

  it('renders a single cell as a table when asked', () => {
    const html = transformSpreadsheetHTML(fixture('excel-single-cell'), { singleCell: 'table' })
    expect(html).toMatch(/^<table><tbody><tr><td data-colwidth="64"><p><strong>Hello<\/strong><\/p><\/td><\/tr><\/tbody><\/table>$/)
  })
})

describe('other sources', () => {
  it('Google Sheets: ignores gridline borders, keeps inline styles and runs', () => {
    const m = parseSpreadsheetHTML(fixture('google-sheets'))
    expect(m.source).toBe('google')
    expect(m.colWidths).toEqual([100, 140])
    const [name, num] = m.rows[0]
    expect(name.style).toEqual({ background: '#ffff00', borderBottom: '1px solid #000000' })
    expect(num.style).toEqual({ textAlign: 'right' })
    expect(m.rows[1][0].content[0]).toEqual([
      { text: 'mixed ', marks: {} },
      { text: 'text', marks: { italic: true, color: '#ff0000' } }
    ])
    expect(m.rows[1][1].style).toEqual({ verticalAlign: 'top' })
    expect(m.rows[1][1].content[0][0].marks).toEqual({ fontSize: '12pt', fontFamily: 'Roboto' })
  })

  it('LibreOffice: legacy attributes, <font>, colgroup widths, hidden comments, non-Latin text', () => {
    const m = parseSpreadsheetHTML(fixture('libreoffice'))
    expect(m.source).toBe('libreoffice')
    expect(m.colWidths).toEqual([85, 120, 120])
    expect(m.rows[0][0].style.background).toBe('#ffff00')
    expect(m.rows[0][0].content[0][0].marks).toEqual({ bold: true, color: '#ff0000' })
    expect(m.rows[0][2].content[0]).toEqual([{ text: 'note', marks: {} }])
    expect(m.rows[1][0].content[0][0]).toEqual({ text: 'Ελληνικά', marks: { italic: true, fontFamily: 'DejaVu Serif' } })
    expect(m.rows[1][1]).toMatchObject({ colspan: 2, style: { textAlign: 'center', verticalAlign: 'middle' } })
  })

  it('leaves non-spreadsheet HTML untouched', () => {
    const html = '<p>Hello <b>world</b></p><table><tr><td>x</td></tr></table>'
    expect(transformSpreadsheetHTML(html)).toBe(html)
  })

  it('handles generic tables only when enabled', () => {
    const html = '<table><tr><td style="background:#eee">x</td><td>y</td></tr></table>'
    expect(transformSpreadsheetHTML(html, { sources: ['generic'] })).toContain('background-color: #eeeeee')
  })

  it('does not mistake Word for Excel', () => {
    const doc = new DOMParser().parseFromString(
      '<html xmlns:x="urn:schemas-microsoft-com:office:excel"><head><meta name=ProgId content=Word.Document></head><body><table><tr><td>x</td></tr></table></body></html>',
      'text/html'
    )
    expect(detectSource(doc, doc.documentElement.outerHTML)).toBe('generic')
  })

  it('strips a leftover CF_HTML header', () => {
    const raw = 'Version:0.9\r\nStartHTML:0000000105\r\nEndHTML:0000000400\r\n' + fixture('excel-single-cell')
    expect(transformSpreadsheetHTML(raw)).toBe('<strong>Hello</strong>')
  })
})

describe('options and safety', () => {
  it('lets you switch off individual styles', () => {
    const m = parseSpreadsheetHTML(fixture('excel-windows-range'), {
      styles: { fontFamily: false, fontSize: false, borders: false, background: false }
    })
    expect(m.rows[4][2].content[0][0].marks).toEqual({ strike: true })
    expect(m.rows[0][0].style).toEqual({ textAlign: 'center', verticalAlign: 'middle' })
  })

  it('can split in-cell line breaks into paragraphs', () => {
    const m = parseSpreadsheetHTML(fixture('excel-windows-range'), { lineBreaks: 'paragraph' })
    expect(texts(m.rows[3][0])).toEqual(['Санкт-', 'Петербург'])
  })

  it('drops unsafe links and sanitizes style values', () => {
    const html = `<meta name=ProgId content=Excel.Sheet><table><tr>
      <td><a href="javascript:alert(1)">bad</a></td>
      <td style="color:red;background:url(javascript:x) #00f;font-family:Arial&quot;;}&lt;/style&gt;">ok</td>
      <td style="border-top:1px solid expression(alert(1))">b</td>
    </tr></table>`
    const out = renderTableHTML(parseSpreadsheetHTML(html))
    expect(out).not.toMatch(/javascript|expression|<\/style/i)
    expect(out).toContain('<p>bad</p>')
    expect(out).toContain('background-color: #0000ff')
  })

  it('returns the original HTML if parsing throws', () => {
    expect(transformSpreadsheetHTML(null)).toBe(null)
  })
})

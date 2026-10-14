// Runs the reader over every generated file in test/fixtures (see test/fixtures.js).
import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { readSpreadsheetFile } from '../src/index.js'
import { localeInfo } from '../src/files/numfmt.js'
import { fixtures, FORMAT_CASES } from './fixtures.js'

const bytesOf = (name) => new Uint8Array(readFileSync(join(process.cwd(), 'test/fixtures', name)))
const read = (name, files = {}) => readSpreadsheetFile(new File([bytesOf(name)], name), { files: { locale: 'en-US', ...files } })
const text = (cell) => cell.content.map((p) => p.map((r) => (r.br ? '\n' : r.text)).join('')).join('\n')
const grid = (model) => model.rows.map((r) => r.map(text))

describe('fixture files', () => {
  it('are up to date with test/fixtures.js (run `npm run fixtures`)', () => {
    for (const [name, build] of Object.entries(fixtures)) {
      expect(Buffer.from(bytesOf(name)).equals(Buffer.from(build())), name).toBe(true)
    }
  })
})

describe('formats.xlsx', () => {
  it('displays every value the way Excel does (en-US)', async () => {
    const { model } = await read('formats.xlsx')
    expect(model.height).toBe(FORMAT_CASES.length + 1)
    model.rows.slice(1).forEach((row, i) => {
      expect(text(row[1]), `format ${FORMAT_CASES[i][0]}`).toBe(text(row[2]))
    })
  })

  it('applies [Red] and keeps formats locale-aware', async () => {
    const { model } = await read('formats.xlsx', { locale: 'de-DE' })
    const red = model.rows.find((r) => text(r[0]) === '#,##0;[Red]-#,##0')
    expect(red[1].content[0][0].marks).toEqual({ color: '#ff0000' })
    const money = model.rows.find((r) => text(r[0]) === '#,##0.00\\ "₽"')
    const { group, decimal } = localeInfo('de-DE')
    expect(text(money[1])).toBe(`1${group}234${decimal}50 ₽`)
  })
})

describe('rich-text.xlsx', () => {
  it('keeps runs, scripts, decorations and theme colors', async () => {
    const { model } = await read('rich-text.xlsx')
    const runs = (r) => model.rows[r][0].content[0]
    expect(runs(0).map((x) => [x.text, x.marks])).toEqual([
      ['H', {}], ['2', { sub: true }], ['O, E=mc', {}], ['2', { sup: true }]
    ])
    expect(runs(1).filter((x) => Object.keys(x.marks).length).map((x) => x.marks)).toEqual([
      { bold: true }, { italic: true }, { underline: true }, { strike: true }
    ])
    expect(runs(2)).toEqual([
      { text: 'красный', marks: { color: '#ff0000' } },
      { text: ' ', marks: {} },
      { text: 'тема', marks: { color: '#ed7d31' } },
      { text: ' ', marks: {} },
      { text: 'Arial 16', marks: { fontSize: '16pt', fontFamily: 'Arial' } }
    ])
  })

  it('keeps multi-line text, emoji, RTL and CJK text and full-width font names', async () => {
    const { model } = await read('rich-text.xlsx')
    const g = grid(model).map((r) => r[0])
    expect(g[3]).toBe('Первая\nВторая\nТретья')
    expect(model.rows[3][0].style.verticalAlign).toBe('top')
    expect(g.slice(4, 9)).toEqual(['emoji 🎉👍', 'مرحبا بالعالم', 'שלום עולם', '你好，世界', 'こんにちは'])
    expect(model.rows[8][0].content[0][0].marks).toEqual({ fontFamily: 'ＭＳ Ｐゴシック' })
    expect(model.rows[9][0].content[0][0].marks).toEqual({ bold: true, color: '#0000ff' })
    expect(g[10]).toBe('  отступы  ')
    expect(g[11]).toBe('<b>не тег</b> & "кавычки"')
  })
})

describe('merges.xlsx', () => {
  it('shrinks merges that cross a hidden column or row', async () => {
    const { model } = await read('merges.xlsx')
    expect([model.height, model.width]).toEqual([5, 5])
    expect(model.colWidths).toEqual([98, 70, 70, 70, 84])
    const spans = model.rows.map((r) => r.map((c) => `${text(c)}:${c.colspan}x${c.rowspan}`))
    expect(spans).toEqual([
      ['Квартальный отчёт:5x1'],
      ['Регион:1x2', 'Q1:2x1', 'Q2:1x1', 'Итого:1x2'],
      ['план:1x1', 'факт:1x1', 'план:1x1'],
      ['Север:1x2', '10:1x1', '12:1x1', '20:1x1', '64:1x1'],
      ['11:1x1', '13:1x1', '21:1x1', '68:1x1']
    ])
  })
})

describe('conditional.xlsx', () => {
  const highlighted = async () => {
    const { model } = await read('conditional.xlsx')
    const heads = model.rows[0].map(text)
    const out = {}
    heads.forEach((h, c) => {
      out[h] = model.rows.slice(1).map((r) => r[c]).filter((cell) => cell.style.background).map(text)
    })
    return { out, model }
  }

  it('evaluates each supported rule type', async () => {
    const { out } = await highlighted()
    expect(out['between 3..6']).toEqual(['3', '4', '5', '6'])
    expect(out['top 2']).toEqual(['9', '10'])
    expect(out['bottom 20%']).toEqual(['1', '2'])
    expect(out['above avg']).toEqual(['6', '7', '8', '9', '10'])
    expect(out['<= avg']).toEqual(['1', '2', '3', '4', '5'])
    expect(out['duplicates']).toEqual(['2', '2', '3', '3', '3'])
    expect(out['begins A']).toEqual(['Apple', 'Avocado', 'apricot'])
    expect(out['blanks']).toEqual(['', '', ''])
    expect(out['expression']).toEqual([]) // formula rules are not evaluated
  })

  it('interpolates a 3-color scale through the percentile midpoint', async () => {
    const { model } = await highlighted()
    const col = model.rows.slice(1).map((r) => r[5].style.background)
    expect(col[0]).toBe('#f8696b')
    expect(col[9]).toBe('#63be7b')
    expect(new Set(col).size).toBe(10)
  })

  it('honours priority and stopIfTrue', async () => {
    const { model } = await highlighted()
    const cells = model.rows.slice(1).map((r) => r[9])
    expect(cells[6].style.background).toBe('#c6efce') // 7 > 5: first rule, then stop
    expect(cells[6].content[0][0].marks).toEqual({ color: '#006100' })
    expect(cells[3].style.background).toBeUndefined() // 4: only the second rule (> 2)
    expect(cells[3].content[0][0].marks).toEqual({ bold: true, color: '#ff0000' })
    expect(cells[0].content[0][0].marks).toEqual({})
  })
})

describe('colors', () => {
  it('matches Excel for theme tints', async () => {
    const { model } = await read('theme-colors.xlsx')
    const bg = Object.fromEntries(model.rows.map((r) => [text(r[0]), r[1].style.background]))
    expect(bg).toMatchObject({
      'lighter 80%': '#dae3f3',
      'lighter 60%': '#b4c7e7',
      'lighter 40%': '#8faadc',
      'darker 25%': '#2f5597',
      'darker 50%': '#203864',
      'white darker 15%': '#d9d9d9',
      'black lighter 50%': '#7f7f7f',
      'indexed 10': '#ff0000',
      argb: '#123456'
    })
  })

  it('uses the colors of a custom theme', async () => {
    const { model } = await read('custom-theme.xlsx')
    expect(model.rows[0][0].style.background).toBe('#6b2c91')
    expect(model.rows[0][1].content[0][0].marks.color).toBe('#e36c09')
  })
})

describe('structure variants', () => {
  it('date1904.xlsx', async () => {
    expect(grid((await read('date1904.xlsx')).model)).toEqual([['01.01.1904'], ['15.03.2023'], ['2023-03-15 06:00'], ['12:00:00']])
  })

  it('offset.xlsx: data far from A1, whole-sheet column width', async () => {
    const { model } = await read('offset.xlsx')
    expect(grid(model)).toEqual([['a', 'b', ''], ['1', '', '3']])
    expect(model.colWidths).toEqual([140, 140, 140])
    expect(model.rows[0][2].style.background).toBe('#ffff00')
  })

  it('no-refs.xlsx: rows and cells without r="…"', async () => {
    expect(grid((await read('no-refs.xlsx')).model)).toEqual([['a', '', 'c'], ['1', '2', '3']])
  })

  it('minimal.xlsx: no styles, theme or shared strings', async () => {
    expect(grid((await read('minimal.xlsx')).model)).toEqual([['только', 'inline'], ['1', '2.5']])
  })

  it('prefixed.xlsx: x:-prefixed SpreadsheetML', async () => {
    const { model } = await read('prefixed.xlsx')
    expect(grid(model)).toEqual([['префикс', '50%']])
    expect(model.rows[0][0].content[0][0].marks).toEqual({ bold: true })
  })

  it('stored.xlsx: uncompressed ZIP entries', async () => {
    expect(grid((await read('stored.xlsx')).model)).toEqual([['stored', '1']])
  })

  it('formulas.xlsx: cached results of every type', async () => {
    expect(grid((await read('formulas.xlsx')).model)).toEqual([
      ['1', '2', '3'],
      ['10', '20', 'ab'],
      ['TRUE', '#DIV/0!', ''],
      ['31.01.2024', '#N/A', '=not a formula'],
      ['2024-01-31 00:00', '2024-01-31 23:30', '']
    ])
  })

  it('hyperlinks.xlsx: keeps safe external links only', async () => {
    const { model } = await read('hyperlinks.xlsx')
    const links = model.rows.map((r) => r.map((c) => c.content[0][0]?.marks.link ?? null))
    expect(links).toEqual([
      ['https://example.com/?q=1&r=2', null],
      ['mailto:team@example.com', null],
      [null, null],
      [null, null],
      ['https://example.org/', 'https://example.org/']
    ])
  })
})

describe('sheets', () => {
  it('multi-sheet.xlsx: reads the first sheet, not the active one', async () => {
    const { name, model } = await read('multi-sheet.xlsx')
    expect(name).toBe('Первый')
    expect(grid(model)).toEqual([['первый лист']])
  })

  it('hidden-first.xlsx: skips a hidden first sheet', async () => {
    expect((await read('hidden-first.xlsx')).name).toBe('Данные')
  })

  it('empty-first.xlsx: an empty first sheet means no data', async () => {
    expect(await read('empty-first.xlsx')).toBe(null)
  })

  it('big.xlsx: truncates to maxCells', async () => {
    const { model } = await read('big.xlsx')
    expect(model.truncated).toBe(true)
    expect([model.height, model.width]).toEqual([2500, 20])
    expect(text(model.rows[2499][0])).toBe('row 2500')
    const all = await read('big.xlsx', { maxCells: 100000 })
    expect(all.model.height).toBe(3000)
    expect(all.model.truncated).toBeUndefined()
  })
})

describe('broken and unsupported files', () => {
  it('rejects them with a readable message', async () => {
    await expect(read('broken-truncated.xlsx')).rejects.toThrow(/Not a ZIP/)
    await expect(read('not-a-workbook.xlsx')).rejects.toThrow(/no workbook part/)
    await expect(read('legacy.xls')).rejects.toThrow(/save it as \.xlsx/)
  })
})

describe('csv/tsv files', () => {
  it('windows-1251 is recognised even with a non-Russian locale', async () => {
    const { model } = await read('sample-cp1251.csv')
    expect(grid(model)[0]).toEqual(['Город', 'Сумма', 'Дата'])
    expect(grid(model)[2][0]).toBe('Санкт-\nПетербург')
  })

  it('windows-1252 Western text is not mistaken for Cyrillic', async () => {
    expect(grid((await read('cp1252-western.csv')).model)).toEqual([['Ville', 'Größe'], ['Café Zürich', 'Ça va']])
  })

  it('UTF-8 with BOM and quoted commas', async () => {
    expect(grid((await read('utf8-bom.csv')).model)).toEqual([['name', 'city'], ['Doe, John', 'Zürich'], ['say "hi"', 'Köln']])
  })

  it('UTF-16LE tab-separated (Excel "Unicode text")', async () => {
    expect(grid((await read('utf16le.tsv')).model)).toEqual([['Имя', 'Город'], ['Анна', 'Казань']])
  })

  it('Excel sep= hint', async () => {
    const { model } = await read('sep-hint.csv')
    expect(grid(model)).toEqual([['a', 'b'], ['1,5', '2,5']])
    expect(model.rows[1][0].style.textAlign).toBe('right')
  })
})

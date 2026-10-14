// Fixture definitions. `npm run fixtures` writes them to test/fixtures/; tests read the files
// from disk and also check they are up to date with these definitions.
import { writeXlsx, colName, toSerial } from './xlsx-writer.js'
import { buildWorkbook, zip } from './xlsx-builder.js'

const bold = { b: true }
const header = { font: { b: true, color: 'FFFFFF' }, fill: { theme: 4 }, align: { h: 'center', v: 'center' }, border: { top: 'thin', bottom: 'medium', left: 'thin', right: 'thin' } }
const thin = { top: 'thin', bottom: 'thin', left: 'thin', right: 'thin' }

// [format code (string) or built-in id (number), value, what Excel shows in en-US]
export const FORMAT_CASES = [
  ['General', 1234.5, '1234.5'],
  ['General', 1 / 3, '0.333333333'],
  [1, 2.5, '3'],
  [2, 3.14159, '3.14'],
  [3, 1234567, '1,234,567'],
  [4, -1234.567, '-1,234.57'],
  [9, 0.256, '26%'],
  [10, 0.1234, '12.34%'],
  [11, 12345, '1.23E+04'],
  [12, 1.5, '1 1/2'],
  ['#,##0.00\\ "₽"', 1234.5, '1,234.50 ₽'],
  ['[$€-407] #,##0.00', 99.9, '€ 99.90'],
  ['"$"#,##0.00', 5, '$5.00'],
  ['#,##0;(#,##0)', -42, '(42)'],
  ['#,##0;[Red]-#,##0', -5, '-5'],
  ['0;-0;"ноль"', 0, 'ноль'],
  ['[>=1000]#,##0,"K";0', 15300, '15K'],
  ['00000', 42, '00042'],
  ['0.0,,"M"', 2500000, '2.5M'],
  ['#.##', 5, '5.'],
  ['dd.mm.yyyy', 45000, '15.03.2023'],
  ['yyyy-mm-dd hh:mm:ss', 45000.5, '2023-03-15 12:00:00'],
  ['mmmm d, yyyy', 45000, 'March 15, 2023'],
  ['ddd, d mmm yy', 45000, 'Wed, 15 Mar 23'],
  [14, 45000, '03/15/2023'],
  [20, 0.75, '18:00'],
  [18, 0.75, '6:00 PM'],
  ['[h]:mm', 1.5, '36:00'],
  ['mm:ss.0', 0.0000063657, '00:00.6'],
  [49, '007', '007'],
  ['"Итого: "@', 'x', 'Итого: x'],
  ['General', true, 'TRUE']
]

export const fixtures = {
  // Hand-written XML parts, the way Excel writes them (see xlsx-builder.js).
  'sample.xlsx': () => buildWorkbook(),

  // A gallery of number formats: column C says what Excel displays (en-US).
  'formats.xlsx': () => writeXlsx({
    sheets: [{
      name: 'Форматы',
      cols: [{ min: 1, width: 24 }, { min: 2, width: 18 }, { min: 3, width: 24 }],
      rows: [
        [{ v: 'Формат', s: { font: bold } }, { v: 'Значение', s: { font: bold } }, { v: 'Excel (en-US)', s: { font: bold } }],
        ...FORMAT_CASES.map(([code, value, shown]) => [
          { inline: typeof code === 'number' ? `builtin ${code}` : code },
          { v: value, s: { numFmt: code } },
          { inline: shown }
        ])
      ]
    }]
  }),

  // Rich text runs, scripts, decorations, colors, multi-line, non-Latin scripts, escaping.
  'rich-text.xlsx': () => writeXlsx({
    sheets: [{
      name: 'Rich',
      cols: [{ min: 1, width: 40 }],
      rows: [
        [{ rich: [{ text: 'H' }, { text: '2', font: { vertAlign: 'subscript' } }, { text: 'O, E=mc' }, { text: '2', font: { vertAlign: 'superscript' } }] }],
        [{ rich: [{ text: 'обычный ' }, { text: 'жирный', font: { b: true } }, { text: ' ' }, { text: 'курсив', font: { i: true } }, { text: ' ' }, { text: 'двойное', font: { u: 'double' } }, { text: ' ' }, { text: 'зачёркнутый', font: { strike: true } }] }],
        [{ rich: [{ text: 'красный', font: { color: 'FF0000' } }, { text: ' ' }, { text: 'тема', font: { color: { theme: 5 } } }, { text: ' ' }, { text: 'Arial 16', font: { name: 'Arial', sz: 16 } }] }],
        [{ v: 'Первая\nВторая\nТретья', s: { align: { wrap: true, v: 'top' } } }],
        ['emoji 🎉👍'],
        ['مرحبا بالعالم'],
        ['שלום עולם'],
        [{ v: '你好，世界', s: { font: { name: 'Microsoft YaHei' } } }],
        [{ v: 'こんにちは', s: { font: { name: 'ＭＳ Ｐゴシック' } } }],
        [{ v: 'весь жирный и синий', s: { font: { b: true, color: '0000FF' } } }],
        ['  отступы  '],
        ['<b>не тег</b> & "кавычки"']
      ]
    }]
  }),

  // Nested merges, a merge crossing a hidden column, a vertical merge crossing a hidden row.
  'merges.xlsx': () => writeXlsx({
    sheets: [{
      name: 'Отчёт',
      cols: [{ min: 1, width: 14 }, { min: 2, max: 4, width: 10 }, { min: 5, hidden: true }, { min: 6, width: 12 }],
      hiddenRows: [5],
      merges: ['A1:F1', 'A2:A3', 'B2:C2', 'D2:E2', 'F2:F3', 'A4:A6'],
      rows: [
        [{ v: 'Квартальный отчёт', s: { font: { b: true, sz: 14 }, align: { h: 'center' } } }],
        [{ v: 'Регион', s: header }, { v: 'Q1', s: header }, { empty: true, s: header }, { v: 'Q2', s: header }, { empty: true, s: header }, { v: 'Итого', s: header }],
        [{ empty: true, s: header }, 'план', 'факт', 'план', 'факт', { empty: true, s: header }],
        [{ v: 'Север', s: { align: { v: 'center' } } }, 10, 12, 20, 22, 64],
        [null, 1, 1, 1, 1, 4],
        [null, 11, 13, 21, 23, 68]
      ]
    }]
  }),

  // One column per conditional-formatting rule type over the same data.
  'conditional.xlsx': () => {
    const nums = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
    const dups = [1, 2, 2, 3, 3, 3, 4, 5, 6, 7]
    const words = ['Apple', 'Banana', 'Avocado', 'Cherry', 'apricot', 'Grape', 'Kiwi', 'Mango', 'Lemon', 'Lime']
    const blanks = [1, null, 3, null, 5, 6, null, 8, 9, 10]
    const heads = ['between 3..6', 'top 2', 'bottom 20%', 'above avg', '<= avg', '3-color', 'duplicates', 'begins A', 'blanks', 'stopIfTrue', 'expression']
    const rows = [heads.map((h) => ({ v: h, s: { font: bold } }))]
    for (let i = 0; i < 10; i++) {
      rows.push([nums[i], nums[i], nums[i], nums[i], nums[i], nums[i], dups[i], words[i], blanks[i] ?? { empty: true, s: { border: thin } }, nums[i], nums[i]])
    }
    const col = (c) => `${colName(c)}2:${colName(c)}11`
    return writeXlsx({
      sheets: [{
        name: 'CF',
        rows,
        cf: (dxf) => {
          const red = dxf({ fill: 'FFC7CE', font: { color: '9C0006' } })
          const green = dxf({ fill: 'C6EFCE', font: { color: '006100' } })
          const boldRed = dxf({ font: { b: true, color: 'FF0000' } })
          const rule = (c, body) => `<conditionalFormatting sqref="${col(c)}">${body}</conditionalFormatting>`
          return [
            rule(0, `<cfRule type="cellIs" dxfId="${red}" priority="1" operator="between"><formula>3</formula><formula>6</formula></cfRule>`),
            rule(1, `<cfRule type="top10" dxfId="${green}" priority="2" rank="2"/>`),
            rule(2, `<cfRule type="top10" dxfId="${red}" priority="3" percent="1" bottom="1" rank="20"/>`),
            rule(3, `<cfRule type="aboveAverage" dxfId="${green}" priority="4"/>`),
            rule(4, `<cfRule type="aboveAverage" dxfId="${red}" priority="5" aboveAverage="0" equalAverage="1"/>`),
            rule(5, `<cfRule type="colorScale" priority="6"><colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale></cfRule>`),
            rule(6, `<cfRule type="duplicateValues" dxfId="${red}" priority="7"/>`),
            rule(7, `<cfRule type="beginsWith" dxfId="${green}" priority="8" operator="beginsWith" text="A"><formula>LEFT(H2,1)="A"</formula></cfRule>`),
            rule(8, `<cfRule type="containsBlanks" dxfId="${red}" priority="9"><formula>LEN(TRIM(I2))=0</formula></cfRule>`),
            rule(9, `<cfRule type="cellIs" dxfId="${green}" priority="10" operator="greaterThan" stopIfTrue="1"><formula>5</formula></cfRule><cfRule type="cellIs" dxfId="${boldRed}" priority="11" operator="greaterThan"><formula>2</formula></cfRule>`),
            rule(10, `<cfRule type="expression" dxfId="${red}" priority="12"><formula>MOD(K2,2)=0</formula></cfRule>`)
          ].join('')
        }
      }]
    })
  },

  // Theme colors with Excel's standard tints, a custom theme, indexed and ARGB colors.
  'theme-colors.xlsx': () => {
    const tints = [
      ['lighter 80%', 0.7999816888943144], ['lighter 60%', 0.5999938962981048], ['lighter 40%', 0.3999755851924192],
      ['darker 25%', -0.249977111117893], ['darker 50%', -0.499984740745262]
    ]
    return writeXlsx({
      sheets: [{
        name: 'Цвета',
        rows: [
          ...tints.map(([label, tint]) => [label, { empty: true, s: { fill: { theme: 4, tint } } }]),
          ['white darker 15%', { empty: true, s: { fill: { theme: 0, tint: -0.1499984740745262 } } }],
          ['black lighter 50%', { empty: true, s: { fill: { theme: 1, tint: 0.499984740745262 } } }],
          ['indexed 10', { empty: true, s: { fill: { indexed: 10 } } }],
          ['argb', { empty: true, s: { fill: '80123456' } }],
          ['font theme 5', { v: 'text', s: { font: { color: { theme: 5 } } } }]
        ]
      }]
    })
  },

  // A non-Office theme: colors must come from the file's theme1.xml.
  'custom-theme.xlsx': () => writeXlsx({
    theme: ['FFFFFF', '000000', 'EEECE1', '1F497D', '6B2C91', 'E36C09', '9BBB59', '8064A2', '4BACC6', 'F79646', '0000FF', '800080'],
    sheets: [{ rows: [[{ v: 'fill', s: { fill: { theme: 4 } } }, { v: 'text', s: { font: { color: { theme: 5 } } } }]] }]
  }),

  // The 1904 date system (old Mac Excel).
  'date1904.xlsx': () => writeXlsx({
    date1904: true,
    sheets: [{
      rows: [
        [{ v: 0, s: { numFmt: 'dd.mm.yyyy' } }],
        [{ v: 43538, s: { numFmt: 'dd.mm.yyyy' } }],
        [{ v: 43538.25, s: { numFmt: 'yyyy-mm-dd hh:mm' } }],
        [{ v: 0.5, s: { numFmt: 'hh:mm:ss' } }]
      ]
    }]
  }),

  // Data far from A1, a whole-sheet <col> range, styled empty cells around the data.
  'offset.xlsx': () => writeXlsx({
    sheets: [{
      origin: 'J20',
      cols: [{ min: 1, max: 16384, width: 20 }],
      rows: [
        [{ v: 'a', s: { border: thin } }, { v: 'b', s: { border: thin } }, { empty: true, s: { fill: 'FFFF00' } }],
        [1, { empty: true, s: { border: thin } }, 3]
      ]
    }]
  }),

  // Streaming writers omit r="…" on rows and cells.
  'no-refs.xlsx': () => writeXlsx({ noRefs: true, sheets: [{ rows: [['a', null, 'c'], [1, 2, 3]] }] }),

  // No styles.xml, no theme, no sharedStrings — inline strings only.
  'minimal.xlsx': () => writeXlsx({
    styles: false,
    theme: false,
    sheets: [{ rows: [[{ inline: 'только' }, { inline: 'inline' }], [1, 2.5]] }]
  }),

  // Open XML SDK style: every SpreadsheetML element carries an `x:` prefix.
  'prefixed.xlsx': () => writeXlsx({
    prefixed: true,
    sheets: [{ rows: [[{ v: 'префикс', s: { font: bold } }, { v: 0.5, s: { numFmt: 9 } }]] }]
  }),

  // Uncompressed ZIP entries.
  'stored.xlsx': () => writeXlsx({ store: true, sheets: [{ rows: [['stored', 1]] }] }),

  // Formulas with cached results of every type, shared formulas, a missing result, t="d".
  'formulas.xlsx': () => writeXlsx({
    sheets: [{
      rows: [
        [1, 2, { f: 'A1+B1', v: 3 }],
        [{ f: 'A1*10', v: 10, shared: { si: 0, ref: 'A2:B2' } }, { f: '', v: 20, shared: { si: 0 } }, { f: 'CONCAT("a","b")', v: 'ab', t: 'str' }],
        [{ f: 'A1>0', v: true, t: 'b' }, { f: '1/0', v: '#DIV/0!', t: 'e' }, { f: 'NOW()' }],
        [{ iso: '2024-01-31T00:00:00', s: { numFmt: 'dd.mm.yyyy' } }, { error: '#N/A' }, { v: '=not a formula' }],
        [{ iso: '2024-01-31', s: { numFmt: 'yyyy-mm-dd hh:mm' } }, { iso: '2024-01-31T23:30:00Z', s: { numFmt: 'yyyy-mm-dd hh:mm' } }]
      ]
    }]
  }),

  // External, mailto, unsafe and in-workbook links; a link over a range.
  'hyperlinks.xlsx': () => writeXlsx({
    sheets: [
      {
        rows: [['сайт'], ['почта'], ['опасная'], ['внутренняя'], ['диапазон', 'тоже']],
        hyperlinks: [
          { ref: 'A1', url: 'https://example.com/?q=1&r=2' },
          { ref: 'A2', url: 'mailto:team@example.com' },
          { ref: 'A3', url: 'javascript:alert(1)' },
          { ref: 'A4', location: "'Второй'!A1" },
          { ref: 'A5:B5', url: 'https://example.org/' }
        ]
      },
      { name: 'Второй', rows: [['x']] }
    ]
  }),

  // Several sheets: the first one is read even though the third is active.
  'multi-sheet.xlsx': () => writeXlsx({
    activeTab: 2,
    sheets: [
      { name: 'Первый', rows: [['первый лист']] },
      { name: 'Второй', rows: [['второй лист']] },
      { name: 'Третий (активный)', rows: [['третий лист']] }
    ]
  }),

  // First sheet hidden → the first visible one is read.
  'hidden-first.xlsx': () => writeXlsx({
    sheets: [
      { name: 'Служебный', state: 'hidden', rows: [['секрет']] },
      { name: 'Данные', rows: [['видимый']] }
    ]
  }),

  // First sheet empty → "no data" (the second sheet is not used).
  'empty-first.xlsx': () => writeXlsx({ sheets: [{ name: 'Пустой', rows: [] }, { name: 'Данные', rows: [['x']] }] }),

  // 3000×20 = 60 000 cells: over the default 50 000 limit.
  'big.xlsx': () => writeXlsx({
    sheets: [{
      rows: Array.from({ length: 3000 }, (_, r) => Array.from({ length: 20 }, (_, c) => (c === 0 ? `row ${r + 1}` : r * 20 + c)))
    }]
  }),

  // Broken inputs.
  'broken-truncated.xlsx': () => buildWorkbook().subarray(0, 2000),
  'not-a-workbook.xlsx': () => zip({ 'readme.txt': 'just a zip' }),
  'legacy.xls': () => new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, ...new Array(504).fill(0)]),

  // CSV/TSV variants.
  'sample-cp1251.csv': () => encode1251('Город;Сумма;Дата\r\nМосква;1 234,50;14.10.2026\r\n"Санкт-\nПетербург";7;01.01.2026\r\n'),
  'utf8-bom.csv': () => new Uint8Array([0xef, 0xbb, 0xbf, ...new TextEncoder().encode('name,city\n"Doe, John",Zürich\n"say ""hi""",Köln\n')]),
  'utf16le.tsv': () => utf16le('Имя\tГород\r\nАнна\tКазань\r\n'),
  'sep-hint.csv': () => new TextEncoder().encode('sep=;\r\na;b\r\n1,5;2,5\r\n'),
  'cp1252-western.csv': () => new Uint8Array([...'Ville;Größe\r\nCafé Zürich;Ça va\r\n'].map((ch) => ch.charCodeAt(0)))
}

function encode1251(text) {
  const out = []
  for (const ch of text) {
    const code = ch.codePointAt(0)
    if (code < 0x80) out.push(code)
    else if (code >= 0x410 && code <= 0x44f) out.push(code - 0x410 + 0xc0)
    else if (ch === 'Ё') out.push(0xa8)
    else if (ch === 'ё') out.push(0xb8)
    else throw new Error(`not in windows-1251: ${ch}`)
  }
  return new Uint8Array(out)
}

function utf16le(text) {
  const out = [0xff, 0xfe]
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i)
    out.push(c & 0xff, c >> 8)
  }
  return new Uint8Array(out)
}

export { toSerial }

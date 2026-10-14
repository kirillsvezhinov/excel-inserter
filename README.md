# prosemirror-excel-paste

Плагин ProseMirror, который превращает таблицы, скопированные из Excel (Windows/Mac/Online), Google Sheets, LibreOffice Calc и Numbers, в нормальные таблицы `prosemirror-tables`. Стили сохраняются. Без зависимостей, кроме самого ProseMirror.

Плагин только **нормализует HTML** из буфера обмена (`transformPastedHTML`). Остальное делают парсер вашей схемы и `prosemirror-tables`. Поэтому:

- курсор вне таблицы → вставляется новая таблица;
- курсор в ячейке → скопированный диапазон (в том числе **часть таблицы Excel**) вставляется в существующую таблицу, начиная с этой ячейки; при нехватке места таблица растёт;
- выделены ячейки (`CellSelection`) → диапазон заполняется скопированным, при необходимости с повторением;
- одна скопированная ячейка → вставляется как обычный текст со стилями (меняется опцией `singleCell`).

## Подключение

```js
import { Schema } from 'prosemirror-model'
import { schema as basic } from 'prosemirror-schema-basic'
import { tableNodes, tableEditing } from 'prosemirror-tables'
import { excelPaste, excelCellAttributes, excelMarks } from 'prosemirror-excel-paste'

const schema = new Schema({
  nodes: basic.spec.nodes.append(tableNodes({
    tableGroup: 'block',
    cellContent: 'block+',
    cellAttributes: excelCellAttributes   // background, textAlign, verticalAlign, border*
  })),
  marks: basic.spec.marks.append(excelMarks) // underline, strikethrough, sub/superscript, textColor, fontSize, fontFamily
})

const plugins = [
  excelPaste(),   // порядок относительно tableEditing не важен
  tableEditing(),
  // ...
]
```

`strong`, `em`, `link` и `hard_break` берутся из `prosemirror-schema-basic`. Если в вашей схеме свои марки или атрибуты ячеек, `excelCellAttributes` и `excelMarks` можно не подключать. Плагин выдаёт обычный HTML, и ваши правила `parseDOM` его разберут:

| Что | HTML на выходе |
|---|---|
| ячейка | `<td colspan rowspan data-colwidth="120,96" style="background-color; text-align; vertical-align; border-top/right/bottom/left">` |
| текст | `<strong> <em> <u> <s> <sub> <sup> <a href> <br>` и `<span style="color; font-size; font-family">` |

## Опции

```js
excelPaste({
  sources: ['excel', 'google', 'libreoffice', 'numbers'], // + 'generic' — любые HTML-таблицы
  singleCell: 'inline',     // 'table' — вставлять 1×1 как таблицу
  lineBreaks: 'br',         // 'paragraph' — Alt+Enter в ячейке → отдельные абзацы
  columnWidths: true,       // data-colwidth из ширин столбцов (px)
  skipHidden: true,         // выкидывать скрытые строки/столбцы
  tsv: false,               // true — plain text с табами → таблица (Ctrl+Shift+V не затрагивается)
  styles: { fontFamily: false, fontSize: false } // отключить отдельные стили; по умолчанию всё true
})
```

Без плагина тоже можно: `transformSpreadsheetHTML(html, options)` возвращает строку, а `parseSpreadsheetHTML(html)` — модель таблицы (`rows → cells → paragraphs → runs`), если нужно собрать узлы самостоятельно.

## Что поддерживается

**Стили ячеек, столбцов и таблицы.** Классы Excel (`.xl65`), `mso-style-parent` (именованные стили вроде «Плохой»), `td`-правила и стили, заданные прямо в ячейках (inline), сводятся в итоговые стили каждой ячейки: заливка, выравнивание по горизонтали и вертикали, границы по сторонам. Толщины Excel переводятся в пиксели: `.5pt` → 1px, `1pt` → 2px, `1.5pt` → 3px. Стили `hairline` и `dot-dash` заменяются ближайшими CSS-стилями. Ширины столбцов берутся из `<col>`/`<colgroup>` с учётом `span` и сохраняются как `colwidth` из prosemirror-tables. Объединения сохраняются (`colspan`/`rowspan`).

**Rich text.** Если в ячейке Excel текст оформлен по-разному, он приходит частями `<font class="fontN">`, и каждая часть получает свои марки. Шрифт книги по умолчанию (Calibri 11pt чёрный в Excel, шрифт таблицы в Google Sheets) **не** переносится: в марки попадает только то, что от него отличается. Иначе каждая вставка тянула бы за собой «Calibri 11pt».

**Условное форматирование.** Применённые цвета переносятся, если Excel положил их в буфер как обычные стили ячейки (это стоит проверить в демо на ваших файлах). Сами *правила* в буфер обмена не попадают.

**Проверка данных (Data Validation).** Через буфер обмена не передаётся ни одним приложением, поэтому при вставке её не получить. Для этого нужен импорт `.xlsx`-файла, например через ExcelJS.

**Исправление особенностей разных приложений:**
- Excel: `colspan` + `mso-ignore:colspan` — это не объединение, а текст, вылезающий в соседние ячейки; такая ячейка разбивается обратно;
- Excel: `text-align:general` + `align=right` у чисел → числа выравниваются вправо, как в Excel;
- Excel: системные цвета (`windowtext`), переносы строк внутри исходного HTML, `<br style='mso-data-placement:same-cell'>`, ячейки только с `&nbsp;`, `<o:p>`;
- Excel: скрытые строки и столбцы, которые Excel копирует вместе с диапазоном;
- Google Sheets: CSS-правило `td {border:1px solid #ccc}` — это сетка листа, а не границы, оно игнорируется;
- LibreOffice: `bgcolor`, `<font color face size>`, `<colgroup width span>`, скрытые примечания `<comment>`;
- рваные строки и частично попавшие в выделение объединения дополняются до прямоугольной таблицы;
- заголовок `Version:0.9 StartHTML:…` формата Windows CF_HTML, если браузер его не убрал;
- безопасность: пропускаются только ссылки `http(s)`/`mailto`/`tel`; цвета, границы и шрифты проходят через белые списки; результат строится через DOM, а не склейкой строк.

**Локали.** Excel кладёт в буфер уже отформатированный текст (`1 234,50 ₽`, `14.10.2026`), и он вставляется как есть, без переформатирования. Кириллица, CJK и другие алфавиты в тексте и названиях шрифтов сохраняются. TSV разбирается с учётом кавычек и переносов внутри полей, как его пишет Excel.

## Чего нет (по сравнению с SheetJS Pro)

Плагин работает только со вставкой из буфера обмена. В нём нет чтения и записи `.xlsx`/`.xls`/`.ods`, формул, правил условного форматирования, проверки данных и экспорта обратно в Excel. Если понадобится импорт файлов, лучше всего подойдёт ExcelJS (MIT): он читает стили, rich text, проверку данных и правила условного форматирования, а его модель приводится к той же `TableModel`.

## Разработка

```sh
npm test       # vitest + jsdom, включая вставку через EditorView
npm run demo        # демо на чистом JS
npm run demo:react  # демо на React (http://localhost:5174): переключатели опций, буфер, модель, документ PM
```

Тестовые фрагменты в `test/fixtures` воспроизведены по формату приложений, а не сняты с реального буфера обмена. Чтобы добавить настоящий фрагмент, вставьте диапазон в демо, нажмите «Скачать .html», положите файл в `test/fixtures` и напишите для него тест.

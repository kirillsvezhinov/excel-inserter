import type { Plugin, PluginKey } from 'prosemirror-state'
import type { EditorView } from 'prosemirror-view'
import type { MarkSpec } from 'prosemirror-model'

export type SpreadsheetSource = 'excel' | 'google' | 'libreoffice' | 'numbers' | 'generic'

export interface StyleOptions {
  background?: boolean
  borders?: boolean
  textAlign?: boolean
  verticalAlign?: boolean
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strikethrough?: boolean
  /** Subscript / superscript. */
  script?: boolean
  color?: boolean
  fontSize?: boolean
  fontFamily?: boolean
  links?: boolean
}

export interface FileOptions {
  /** Handle spreadsheet files on the clipboard (.xlsx/.xlsm/.xltx/.csv/.tsv). Default true. */
  enabled?: boolean
  /** Truncate sheets larger than this many cells (rows are cut). Default 50 000. */
  maxCells?: number
  /** Evaluate conditional formatting rules (cellIs, text, blanks, top/bottom, average, duplicates, color scales). Default true. */
  conditionalFormatting?: boolean
  /** Locale for number/date display and CSV code page. Default: the browser's. */
  locale?: string
  /** Force a CSV encoding (e.g. 'windows-1251'). Default: BOM → UTF-8 → locale code page. */
  encoding?: string
  onLoadStart?(file: File): void
  onLoad?(file: File, sheet: SheetResult): void
  /** Without it, errors are logged with console.warn. */
  onError?(error: Error, file: File): void
}

export interface SheetResult {
  name: string
  model: TableModel
}

export interface ExcelPasteOptions {
  /** Which producers to handle. Default: excel, google, libreoffice, numbers. Add 'generic' to restyle any HTML table. */
  sources?: SpreadsheetSource[]
  /** A 1×1 copy: insert its content inline (default) or as a 1×1 table. */
  singleCell?: 'inline' | 'table'
  /** In-cell line breaks (Alt+Enter): `<br>` / hard_break (default) or separate paragraphs. */
  lineBreaks?: 'br' | 'paragraph'
  /** Emit `data-colwidth` from the source column widths (px). Default true. */
  columnWidths?: boolean
  /** Drop hidden rows/columns (Excel copies them along with the range). Default true. */
  skipHidden?: boolean
  /** Turn tab-separated plain text into a table when no HTML is on the clipboard. Default false. */
  tsv?: boolean
  /** Toggle individual styles; all default to true. */
  styles?: StyleOptions
  /** Pasting files; `false` turns it off. */
  files?: FileOptions | false
}

export interface CellStyle {
  background?: string
  textAlign?: 'left' | 'center' | 'right' | 'justify'
  verticalAlign?: 'top' | 'middle' | 'bottom'
  borderTop?: string
  borderRight?: string
  borderBottom?: string
  borderLeft?: string
}

export interface RunMarks {
  bold?: true
  italic?: true
  underline?: true
  strike?: true
  sub?: true
  sup?: true
  /** `#rrggbb` (or a CSS color name). */
  color?: string
  /** e.g. `14pt`. */
  fontSize?: string
  fontFamily?: string
  link?: string
}

export type Inline = { text: string; marks: RunMarks } | { br: true }

export interface CellModel {
  colspan: number
  rowspan: number
  style: CellStyle
  /** Paragraphs of inline runs. */
  content: Inline[][]
}

export interface TableModel {
  source: SpreadsheetSource | 'tsv' | 'csv' | 'xlsx'
  /** Sheet name, for models read from a workbook. */
  sheet?: string
  /** Set when the sheet exceeded `files.maxCells`. */
  truncated?: boolean
  width: number
  height: number
  /** Column widths in px (null for unknown columns), or null if none are known. */
  colWidths: (number | null)[] | null
  rows: CellModel[][]
}

export const excelPasteKey: PluginKey
export const DEFAULT_OPTIONS: Required<ExcelPasteOptions>
export const DEFAULT_FILE_OPTIONS: FileOptions

/** ProseMirror plugin; list it before `tableEditing()` from prosemirror-tables. */
export function excelPaste(options?: ExcelPasteOptions): Plugin

/** Rewrite spreadsheet clipboard HTML into clean table HTML; returns the input unchanged otherwise. */
export function transformSpreadsheetHTML(html: string, options?: ExcelPasteOptions): string

/** Parse spreadsheet clipboard HTML into a model, or null if it isn't a spreadsheet table. */
export function parseSpreadsheetHTML(html: string, options?: ExcelPasteOptions): TableModel | null

export function renderTableHTML(model: TableModel, options?: ExcelPasteOptions): string

/** Read a spreadsheet file (or raw bytes with `kind`). Workbooks: only the first visible sheet. Null when empty. */
export function readSpreadsheetFile(
  file: File | Blob | ArrayBuffer | Uint8Array,
  options?: ExcelPasteOptions & { kind?: 'xlsx' | 'csv' | 'tsv' }
): Promise<SheetResult | null>

/** 'unsupported' = a spreadsheet format that is recognised (.xls, .ods, …) but can't be read. */
export function spreadsheetFileKind(file: { name?: string; type?: string }): 'xlsx' | 'csv' | 'tsv' | 'unsupported' | null

/** Read a file and paste it at the current selection (e.g. from a file picker or drop handler). */
export function pasteSpreadsheetFile(view: EditorView, file: File, options?: ExcelPasteOptions): Promise<void>

/** Format a value with an Excel number format code. */
export function formatValue(
  value: number | string | boolean | null,
  code: string,
  options?: { locale?: string; date1904?: boolean }
): { text: string; color: string | null; isDate: boolean }

/** Excel serial date → UTC Date. */
export function serialToDate(serial: number, date1904?: boolean): Date

export function parseTSV(text: string): string[][]
export function parseDelimited(text: string, delimiter: string): string[][]
export function looksLikeTSV(rows: string[][]): boolean
export function tsvToModel(rows: string[][], source?: 'tsv' | 'csv'): TableModel

export function detectSource(doc: Document, html?: string): SpreadsheetSource | null

export interface CellAttributeSpec {
  default: null
  getFromDOM(dom: HTMLElement): string | null
  setDOMAttr(value: unknown, attrs: Record<string, unknown>): void
}

/** Pass as `tableNodes({ cellAttributes: excelCellAttributes })`. */
export const excelCellAttributes: {
  background: CellAttributeSpec
  textAlign: CellAttributeSpec
  verticalAlign: CellAttributeSpec
  borderTop: CellAttributeSpec
  borderRight: CellAttributeSpec
  borderBottom: CellAttributeSpec
  borderLeft: CellAttributeSpec
}

/** Mark specs for underline, strikethrough, sub/superscript, text color, font size and family. */
export const excelMarks: {
  underline: MarkSpec
  strikethrough: MarkSpec
  subscript: MarkSpec
  superscript: MarkSpec
  textColor: MarkSpec
  fontSize: MarkSpec
  fontFamily: MarkSpec
}

export function normalizeColor(value: string | null | undefined, loose?: boolean): string | null
export function normalizeBorder(value: string | null | undefined): string | null

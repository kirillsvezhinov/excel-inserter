import type { Plugin, PluginKey } from 'prosemirror-state'
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
  source: SpreadsheetSource | 'tsv'
  width: number
  height: number
  /** Column widths in px (null for unknown columns), or null if none are known. */
  colWidths: (number | null)[] | null
  rows: CellModel[][]
}

export const excelPasteKey: PluginKey
export const DEFAULT_OPTIONS: Required<ExcelPasteOptions>

/** ProseMirror plugin; use together with `tableEditing()` from prosemirror-tables. */
export function excelPaste(options?: ExcelPasteOptions): Plugin

/** Rewrite spreadsheet clipboard HTML into clean table HTML; returns the input unchanged otherwise. */
export function transformSpreadsheetHTML(html: string, options?: ExcelPasteOptions): string

/** Parse spreadsheet clipboard HTML into a model, or null if it isn't a spreadsheet table. */
export function parseSpreadsheetHTML(html: string, options?: ExcelPasteOptions): TableModel | null

export function renderTableHTML(model: TableModel, options?: ExcelPasteOptions): string

export function parseTSV(text: string): string[][]
export function looksLikeTSV(rows: string[][]): boolean
export function tsvToModel(rows: string[][]): TableModel

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

import { resolveOptions } from '../options.js'
import { readXlsx } from './xlsx.js'
import { readDelimited } from './csv.js'

const KINDS = {
  xlsx: 'xlsx', xlsm: 'xlsx', xltx: 'xlsx', xltm: 'xlsx',
  csv: 'csv', tsv: 'tsv', tab: 'tsv',
  // Recognised so the paste is intercepted and reported, but not readable.
  xls: 'unsupported', xlsb: 'unsupported', xlt: 'unsupported', ods: 'unsupported', fods: 'unsupported', numbers: 'unsupported'
}
const MIME_KINDS = {
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'xlsx',
  'application/vnd.ms-excel.sheet.macroenabled.12': 'xlsx',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.template': 'xlsx',
  'text/csv': 'csv',
  'text/tab-separated-values': 'tsv',
  'application/vnd.ms-excel': 'unsupported',
  'application/vnd.ms-excel.sheet.binary.macroenabled.12': 'unsupported',
  'application/vnd.oasis.opendocument.spreadsheet': 'unsupported'
}

/** 'xlsx' | 'csv' | 'tsv' for readable files, 'unsupported' for other spreadsheet formats, else null. */
export function spreadsheetFileKind(file) {
  if (!file) return null
  const ext = /\.([a-z0-9]+)$/i.exec(file.name || '')?.[1]?.toLowerCase()
  return KINDS[ext] || MIME_KINDS[(file.type || '').toLowerCase()] || null
}

async function bytesOf(file) {
  if (file instanceof Uint8Array) return file
  if (file instanceof ArrayBuffer) return new Uint8Array(file)
  if (typeof file.arrayBuffer === 'function') return new Uint8Array(await file.arrayBuffer())
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(new Uint8Array(reader.result))
    reader.onerror = () => reject(reader.error)
    reader.readAsArrayBuffer(file)
  })
}

/**
 * Read a spreadsheet file into a table model. For workbooks only the first sheet is read.
 * @param {File|Blob|ArrayBuffer|Uint8Array} file
 * @param {object} [options] plugin options; `kind` overrides detection for raw buffers
 * @returns {Promise<{ name: string, model: object } | null>} null when there is no data
 */
export async function readSpreadsheetFile(file, options = {}) {
  const opts = resolveOptions(options)
  const kind = options.kind || spreadsheetFileKind(file)
  if (!kind) throw new Error(`Unsupported file type: ${file?.name || file?.type || 'unknown'}`)
  if (kind === 'unsupported') {
    throw new Error(`${file?.name || 'This file'}: only .xlsx and .csv can be pasted — save it as .xlsx`)
  }
  const bytes = await bytesOf(file)
  if (kind === 'xlsx') {
    if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) {
      throw new Error(bytes[0] === 0xd0 && bytes[1] === 0xcf
        ? 'Encrypted or legacy binary workbook (.xls) — save it as .xlsx'
        : 'Not an .xlsx file')
    }
    return readXlsx(bytes, opts)
  }
  const model = readDelimited(bytes, opts, kind === 'tsv' ? '\t' : null)
  return model ? { name: file.name || kind, model } : null
}

export { readXlsx, readDelimited }

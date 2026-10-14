export const DEFAULT_STYLES = {
  background: true,
  borders: true,
  textAlign: true,
  verticalAlign: true,
  bold: true,
  italic: true,
  underline: true,
  strikethrough: true,
  script: true,
  color: true,
  fontSize: true,
  fontFamily: true,
  links: true
}

export const DEFAULT_FILE_OPTIONS = {
  enabled: true,
  maxCells: 50000,
  conditionalFormatting: true,
  locale: undefined,
  encoding: undefined,
  onLoadStart: undefined,
  onLoad: undefined,
  onError: undefined
}

export const DEFAULT_OPTIONS = {
  sources: ['excel', 'google', 'libreoffice', 'numbers'],
  singleCell: 'inline',
  lineBreaks: 'br',
  columnWidths: true,
  skipHidden: true,
  tsv: false,
  styles: DEFAULT_STYLES,
  files: DEFAULT_FILE_OPTIONS
}

export function resolveOptions(options = {}) {
  return {
    ...DEFAULT_OPTIONS,
    ...options,
    styles: { ...DEFAULT_STYLES, ...(options.styles || {}) },
    files: options.files === false
      ? { ...DEFAULT_FILE_OPTIONS, enabled: false }
      : { ...DEFAULT_FILE_OPTIONS, ...(options.files || {}) }
  }
}

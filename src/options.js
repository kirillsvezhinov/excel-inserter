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

export const DEFAULT_OPTIONS = {
  sources: ['excel', 'google', 'libreoffice', 'numbers'],
  singleCell: 'inline',
  lineBreaks: 'br',
  columnWidths: true,
  skipHidden: true,
  tsv: false,
  styles: DEFAULT_STYLES
}

export function resolveOptions(options = {}) {
  return {
    ...DEFAULT_OPTIONS,
    ...options,
    styles: { ...DEFAULT_STYLES, ...(options.styles || {}) }
  }
}

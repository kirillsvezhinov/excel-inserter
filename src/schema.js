import { normalizeColor, parseDeclarations, resolveBorder, sanitizeFontFamily } from './css.js'

// Read the style attribute ourselves: `dom.style` re-serializes colors as rgb() differently per browser.
const declsOf = (dom) => parseDeclarations(dom.getAttribute('style') || '')

function appendStyle(attrs, css) {
  attrs.style = (attrs.style ? attrs.style.replace(/;?\s*$/, '; ') : '') + css
}

function styleAttr(read, cssProp) {
  return {
    default: null,
    getFromDOM: read,
    setDOMAttr(value, attrs) { if (value) appendStyle(attrs, `${cssProp}: ${value}`) }
  }
}

const borderAttr = (side) =>
  styleAttr((dom) => resolveBorder(declsOf(dom), side, null), `border-${side}`)

/**
 * Extra cell attributes for `tableNodes({ cellAttributes })` from prosemirror-tables.
 * Spread them in, pick a subset, or replace them with your own.
 */
export const excelCellAttributes = {
  background: styleAttr(
    (dom) => normalizeColor(declsOf(dom)['background-color']) || normalizeColor(dom.getAttribute('bgcolor')),
    'background-color'
  ),
  textAlign: styleAttr((dom) => {
    const v = (declsOf(dom)['text-align'] || dom.getAttribute('align') || '').toLowerCase()
    return /^(left|right|center|justify)$/.test(v) ? v : null
  }, 'text-align'),
  verticalAlign: styleAttr((dom) => {
    const v = (declsOf(dom)['vertical-align'] || dom.getAttribute('valign') || '').toLowerCase()
    return /^(top|middle|bottom)$/.test(v) ? v : null
  }, 'vertical-align'),
  borderTop: borderAttr('top'),
  borderRight: borderAttr('right'),
  borderBottom: borderAttr('bottom'),
  borderLeft: borderAttr('left')
}

const FONT_SIZE = /^\d+(\.\d+)?(pt|px)$/

/**
 * Mark specs for the inline formatting the paste produces. `strong`, `em` and `link`
 * come from prosemirror-schema-basic; add these alongside them.
 */
export const excelMarks = {
  underline: {
    parseDOM: [
      { tag: 'u' },
      { style: 'text-decoration', consuming: false, getAttrs: (v) => (/underline/.test(v) ? null : false) }
    ],
    toDOM: () => ['u', 0]
  },
  strikethrough: {
    parseDOM: [
      { tag: 's' }, { tag: 'strike' }, { tag: 'del' },
      { style: 'text-decoration', consuming: false, getAttrs: (v) => (/line-through/.test(v) ? null : false) }
    ],
    toDOM: () => ['s', 0]
  },
  subscript: {
    excludes: 'subscript superscript',
    parseDOM: [{ tag: 'sub' }, { style: 'vertical-align=sub' }],
    toDOM: () => ['sub', 0]
  },
  superscript: {
    excludes: 'subscript superscript',
    parseDOM: [{ tag: 'sup' }, { style: 'vertical-align=super' }],
    toDOM: () => ['sup', 0]
  },
  textColor: {
    attrs: { color: {} },
    parseDOM: [{
      style: 'color',
      getAttrs: (v) => {
        const color = normalizeColor(v)
        return color ? { color } : false
      }
    }],
    toDOM: (mark) => ['span', { style: `color: ${mark.attrs.color}` }, 0]
  },
  fontSize: {
    attrs: { size: {} },
    parseDOM: [{
      style: 'font-size',
      getAttrs: (v) => (FONT_SIZE.test(v.trim()) ? { size: v.trim() } : false)
    }],
    toDOM: (mark) => ['span', { style: `font-size: ${mark.attrs.size}` }, 0]
  },
  fontFamily: {
    attrs: { family: {} },
    parseDOM: [{
      style: 'font-family',
      getAttrs: (v) => {
        const family = sanitizeFontFamily(v)
        return family ? { family } : false
      }
    }],
    toDOM: (mark) => ['span', { style: `font-family: ${mark.attrs.family}` }, 0]
  }
}


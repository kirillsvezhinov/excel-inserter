import { describe, expect, it } from 'vitest'
import { StyleSheet, normalizeBorder, normalizeColor, parseDeclarations, toPt } from '../src/css.js'
import { parseTSV, looksLikeTSV } from '../src/tsv.js'

describe('css', () => {
  it('normalizes colors', () => {
    expect(normalizeColor('#ABC')).toBe('#aabbcc')
    expect(normalizeColor('windowtext')).toBe('#000000')
    expect(normalizeColor('rgb(255, 0, 0)')).toBe('#ff0000')
    expect(normalizeColor('rgba(0,0,0,0)')).toBe(null)
    expect(normalizeColor('auto')).toBe(null)
    expect(normalizeColor('rebeccapurple')).toBe('rebeccapurple')
    expect(normalizeColor('url(x)')).toBe(null)
  })

  it('keeps escaped semicolons in Excel number formats', () => {
    const d = parseDeclarations('mso-number-format:"0\\.00\\;\\[Red\\]\\-0\\.00";color:red')
    expect(d['mso-number-format']).toBe('"0\\.00\\;\\[Red\\]\\-0\\.00"')
    expect(d.color).toBe('red')
  })

  it('expands shorthands', () => {
    const d = parseDeclarations('border:.5pt solid windowtext;border-left:none;background:#FFC7CE;font:bold 12px "Helvetica Neue", sans-serif')
    expect(d['border-top-style']).toBe('solid')
    expect(d['border-left-style']).toBe('none')
    expect(d['background-color']).toBe('#FFC7CE')
    expect(d['font-weight']).toBe('bold')
    expect(d['font-size']).toBe('12px')
    expect(d['font-family']).toBe('"Helvetica Neue", sans-serif')
  })

  it('maps Excel border widths and styles', () => {
    expect(normalizeBorder('.5pt solid windowtext')).toBe('1px solid #000000')
    expect(normalizeBorder('1.0pt solid #FF0000')).toBe('2px solid #ff0000')
    expect(normalizeBorder('1.5pt double windowtext')).toBe('3px double #000000')
    expect(normalizeBorder('.5pt hairline windowtext')).toBe('1px solid #000000')
    expect(normalizeBorder('.5pt dot-dash windowtext')).toBe('1px dashed #000000')
    expect(normalizeBorder('none')).toBe(null)
  })

  it('converts units to points', () => {
    expect(toPt('11.0pt')).toBe(11)
    expect(toPt('16px')).toBe(12)
    expect(toPt('x-small')).toBe(7.5)
  })

  it('parses Excel stylesheets wrapped in HTML comments, with specificity', () => {
    const sheet = new StyleSheet([`<!--table {mso-displayed-decimal-separator:"\\,";}
      @page {margin:.75in;}
      td {color:black; font-weight:400}
      .xl65 {font-weight:700}
      td.special, p.p1 {color:red}
      tr > td {color:blue}
    -->`])
    expect(sheet.match('td', ['xl65'])).toMatchObject({ color: 'black', 'font-weight': '700' })
    expect(sheet.match('td', ['special']).color).toBe('red')
    expect(sheet.match('p', ['p1']).color).toBe('red')
  })
})

describe('tsv', () => {
  it('parses quoted fields with tabs, newlines and quotes', () => {
    expect(parseTSV('a\t"b\tc"\n"say ""hi"""\t"x\ny"\r\n')).toEqual([['a', 'b\tc'], ['say "hi"', 'x\ny']])
  })

  it('treats a lone quote as literal', () => {
    expect(parseTSV('5" pipe\tx')).toEqual([['5" pipe', 'x']])
  })

  it('keeps empty trailing cells', () => {
    expect(parseTSV('a\t\nb\t')).toEqual([['a', ''], ['b', '']])
  })

  it('detects tabular shape', () => {
    expect(looksLikeTSV(parseTSV('a\tb\nc\td'))).toBe(true)
    expect(looksLikeTSV(parseTSV('just text'))).toBe(false)
    expect(looksLikeTSV(parseTSV('a\tb\nc'))).toBe(false)
  })
})

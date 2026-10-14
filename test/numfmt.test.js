import { describe, expect, it } from 'vitest'
import { formatValue, serialToDate, builtinFormat, isDateFormat, localeInfo } from '../src/files/numfmt.js'

const en = (v, code) => formatValue(v, code, { locale: 'en-US' }).text
const ru = (v, code) => formatValue(v, code, { locale: 'ru-RU' }).text
const RU_GROUP = localeInfo('ru-RU').group

describe('number formats', () => {
  it('General', () => {
    expect(en(1234.5, 'General')).toBe('1234.5')
    expect(ru(1234.5, 'General')).toBe('1234,5')
    expect(en(-7, 'General')).toBe('-7')
    expect(en(1 / 3, 'General')).toBe('0.333333333')
    expect(en(123456789012345, 'General')).toBe('1.23457E+14')
    expect(en(0.1 + 0.2, 'General')).toBe('0.3')
  })

  it('fixed decimals, grouping and locale separators', () => {
    expect(en(3.14159, '0.00')).toBe('3.14')
    expect(en(1234567, '#,##0')).toBe('1,234,567')
    expect(ru(1234567.891, '#,##0.00')).toBe(`1${RU_GROUP}234${RU_GROUP}567,89`)
    expect(en(0.5, '#,##0.00')).toBe('0.50')
    expect(en(0.5, '#.##')).toBe('.5')
    expect(en(5, '#.##')).toBe('5.')
    expect(en(1.5, '0.0#')).toBe('1.5')
    expect(en(1.567, '0.0#')).toBe('1.57')
    expect(en(123, '00000')).toBe('00123')
    expect(en(2.5, '0')).toBe('3')
  })

  it('literals, currency and escapes', () => {
    expect(ru(1234.5, '#,##0.00\\ "₽"')).toBe(`1${RU_GROUP}234,50 ₽`)
    expect(en(1234.5, '[$€-407] #,##0.00')).toBe('€ 1,234.50')
    expect(en(1234.5, '"$"#,##0.00')).toBe('$1,234.50')
    expect(en(5551234567, '(000) 000-0000')).toBe('(555) 123-4567')
  })

  it('percent, scaling and scientific', () => {
    expect(en(0.256, '0%')).toBe('26%')
    expect(en(0.256, '0.00%')).toBe('25.60%')
    expect(en(1234567, '#,##0,')).toBe('1,235')
    expect(en(1234567, '0.0,,"M"')).toBe('1.2M')
    expect(en(12345, '0.00E+00')).toBe('1.23E+04')
    expect(en(0.00012, '0.00E+00')).toBe('1.20E-04')
  })

  it('sections and colors', () => {
    expect(formatValue(-5, '#,##0.00;[Red]-#,##0.00', { locale: 'en-US' })).toEqual({ text: '-5.00', color: '#ff0000', isDate: false })
    expect(en(-1234, '#,##0;(#,##0)')).toBe('(1,234)')
    expect(en(0, '0;-0;"zero"')).toBe('zero')
    expect(en(-3, '0')).toBe('-3')
    expect(en(1500, '[>=1000]#,##0,"K";0')).toBe('2K')
    expect(en(999, '[>=1000]#,##0,"K";0')).toBe('999')
    expect(en(0, builtinFormat(43)).trim()).toBe('-')
  })

  it('fractions', () => {
    expect(en(1.5, '# ?/?')).toBe('1 1/2')
    expect(en(0.333, '?/?')).toBe('1/3')
    expect(en(2.25, '# ?/4')).toBe('2 1/4')
  })

  it('text sections', () => {
    expect(en('x', '@')).toBe('x')
    expect(en('x', '0;0;0;"Name: "@')).toBe('Name: x')
    expect(en('plain', '0.00')).toBe('plain')
  })

  it('booleans are localized', () => {
    expect(en(true, 'General')).toBe('TRUE')
    expect(ru(false, 'General')).toBe('ЛОЖЬ')
  })
})

describe('date formats', () => {
  it('converts serials (1900 leap bug, 1904 system)', () => {
    expect(serialToDate(45000).toISOString()).toBe('2023-03-15T00:00:00.000Z')
    expect(serialToDate(1).toISOString()).toBe('1900-01-01T00:00:00.000Z')
    expect(serialToDate(61).toISOString()).toBe('1900-03-01T00:00:00.000Z')
    expect(serialToDate(0, true).toISOString()).toBe('1904-01-01T00:00:00.000Z')
  })

  it('formats dates and times', () => {
    expect(en(45000, 'dd.mm.yyyy')).toBe('15.03.2023')
    expect(en(45000.5, 'yyyy-mm-dd hh:mm:ss')).toBe('2023-03-15 12:00:00')
    expect(en(45000, 'mmmm d, yyyy')).toBe('March 15, 2023')
    expect(en(45000, 'ddd, d mmm yy')).toBe('Wed, 15 Mar 23')
    expect(ru(45000, 'd mmmm yyyy')).toMatch(/^15 март[а]? 2023$/)
    expect(en(0.75, 'h:mm AM/PM')).toBe('6:00 PM')
    expect(en(1.5, '[h]:mm:ss')).toBe('36:00:00')
    expect(en(0.0000063657, 'mm:ss.0')).toBe('00:00.6')
    expect(en(45000.25, 'm/d/yy h:mm')).toBe('3/15/23 6:00')
  })

  it('uses the viewer locale for built-in short dates', () => {
    expect(en(45000, builtinFormat(14))).toBe('03/15/2023')
    expect(ru(45000, builtinFormat(14))).toBe('15.03.2023')
    expect(isDateFormat(builtinFormat(14))).toBe(true)
    expect(isDateFormat('0.00')).toBe(false)
    expect(isDateFormat('[Red]0;"m"0')).toBe(false)
  })
})

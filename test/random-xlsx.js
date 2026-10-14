// Random but realistic workbook for clicking around the demo.
// Usage: node test/random-xlsx.js [seed] [rows]   → demo-files/random-<seed>.xlsx (+ .csv)
import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { writeXlsx, toSerial } from './xlsx-writer.js'

const seed = Number(process.argv[2]) || Math.floor(Math.random() * 1e6)
const rowCount = Number(process.argv[3]) || 20 + (seed % 21)

// mulberry32
let state = seed >>> 0
const rand = () => {
  state = (state + 0x6d2b79f5) >>> 0
  let t = state
  t = Math.imul(t ^ (t >>> 15), t | 1)
  t ^= t + Math.imul(t ^ (t >>> 7), t | 61)
  return ((t ^ (t >>> 14)) >>> 0) / 4294967296
}
const pick = (a) => a[Math.floor(rand() * a.length)]
const int = (lo, hi) => lo + Math.floor(rand() * (hi - lo + 1))

const regions = ['Москва', 'Санкт-Петербург', 'Казань', 'Новосибирск', 'Екатеринбург', 'Нижний Новгород', 'Самара']
const managers = ['Иванова А.', 'Петров С.', 'Сидоренко О.', 'Ким Д.', 'Абдуллаев Р.', 'Смирнова Е.']
const products = [['Ноутбук Pro 14', 89990], ['Монитор 27"', 24990], ['Клавиатура', 3490], ['Мышь беспроводная', 1990], ['Док-станция USB-C', 12990], ['Наушники', 7490], ['Веб-камера 4K', 9990]]
const statuses = ['Оплачен', 'Отгружен', 'В работе', 'Отменён', 'Ожидает оплаты']
const comments = [
  null, null, null,
  'Срочно',
  { rich: [{ text: 'Клиент просит ' }, { text: 'скидку', font: { b: true, color: 'C00000' } }] },
  'Доставка\nдо двери',
  { rich: [{ text: 'VIP', font: { b: true, color: { theme: 4 } } }, { text: ' — ' }, { text: 'не звонить после 18:00', font: { i: true } }] },
  'Повторный заказ'
]

const thin = { top: 'thin', bottom: 'thin', left: 'thin', right: 'thin' }
const head = { font: { b: true, color: 'FFFFFF' }, fill: { theme: 4 }, align: { h: 'center', v: 'center', wrap: true }, border: thin }
const cell = (extra = {}) => ({ border: thin, ...extra })
const zebra = (i, extra = {}) => cell(i % 2 ? { fill: { theme: 4, tint: 0.7999816888943144 }, ...extra } : extra)

const start = Date.UTC(2026, 0, 1)
const rows = []
let total = 0
let totalQty = 0
for (let i = 0; i < rowCount; i++) {
  const [name, price] = pick(products)
  const qty = int(1, 25)
  const discount = pick([0, 0, 0, 0.05, 0.1, 0.15])
  const sum = Math.round(qty * price * (1 - discount) * 100) / 100
  total += sum
  totalQty += qty
  const date = new Date(start + int(0, 280) * 86400000)
  const status = pick(statuses)
  const comment = pick(comments)
  const id = `ORD-${String(seed % 1000).padStart(3, '0')}-${String(i + 1).padStart(4, '0')}`
  rows.push([
    { v: i + 1, s: zebra(i, { align: { h: 'center' } }) },
    { v: id, s: zebra(i) }, // hidden column
    { v: toSerial(date), s: zebra(i, { numFmt: 'dd.mm.yyyy' }) },
    { v: pick(regions), s: zebra(i) },
    { v: pick(managers), s: zebra(i) },
    { v: name, s: zebra(i) },
    { v: qty, s: zebra(i, { numFmt: '#,##0' }) },
    { v: price, s: zebra(i, { numFmt: '#,##0.00\\ "₽"' }) },
    { v: discount, s: zebra(i, { numFmt: '0%' }) },
    { f: `G${i + 3}*H${i + 3}*(1-I${i + 3})`, v: sum, s: zebra(i, { numFmt: '#,##0.00\\ "₽";[Red]\\-#,##0.00\\ "₽"' }) },
    { v: status, s: zebra(i, { align: { h: 'center' } }) },
    comment == null ? { empty: true, s: zebra(i) } : typeof comment === 'string' ? { v: comment, s: zebra(i, { align: { wrap: true, v: 'top' } }) } : { ...comment, s: zebra(i) }
  ])
}
const last = rowCount + 2
const totals = [
  { v: 'Итого', s: { font: { b: true }, border: { top: 'double', bottom: 'thin', left: 'thin' } } },
  ...Array.from({ length: 5 }, () => ({ empty: true, s: { border: { top: 'double', bottom: 'thin' } } })),
  { f: `SUM(G3:G${last})`, v: totalQty, s: { font: { b: true }, numFmt: '#,##0', border: { top: 'double', bottom: 'thin' } } },
  { empty: true, s: { border: { top: 'double', bottom: 'thin' } } },
  { empty: true, s: { border: { top: 'double', bottom: 'thin' } } },
  { f: `SUM(J3:J${last})`, v: Math.round(total * 100) / 100, s: { font: { b: true }, numFmt: '#,##0.00\\ "₽"', border: { top: 'double', bottom: 'thin' } } },
  { empty: true, s: { border: { top: 'double', bottom: 'thin' } } },
  { empty: true, s: { border: { top: 'double', bottom: 'thin', right: 'thin' } } }
]

const workbook = writeXlsx({
  sheets: [
    {
      name: 'Продажи',
      cols: [
        { min: 1, width: 5 }, { min: 2, width: 16, hidden: true }, { min: 3, width: 11 }, { min: 4, width: 18 },
        { min: 5, width: 15 }, { min: 6, width: 20 }, { min: 7, width: 8 }, { min: 8, width: 13 },
        { min: 9, width: 8 }, { min: 10, width: 15 }, { min: 11, width: 15 }, { min: 12, width: 30 }
      ],
      merges: ['A1:L1'],
      rows: [
        [{ v: `Отчёт о продажах · выборка #${seed}`, s: { font: { b: true, sz: 16, color: { theme: 4, tint: -0.249977111117893 } }, align: { h: 'center' } } }],
        ['№', 'ID', 'Дата', 'Регион', 'Менеджер', 'Товар', 'Кол-во', 'Цена', 'Скидка', 'Сумма', 'Статус', 'Комментарий'].map((h) => ({ v: h, s: head })),
        ...rows,
        totals
      ],
      hyperlinks: Array.from({ length: Math.min(3, rowCount) }, (_, k) => ({ ref: `F${3 + k * 3}`, url: `https://example.com/catalog?item=${k + 1}` })),
      cf: (dxf) => {
        const ok = dxf({ fill: 'C6EFCE', font: { color: '006100' } })
        const bad = dxf({ fill: 'FFC7CE', font: { color: '9C0006' } })
        const warn = dxf({ fill: 'FFEB9C', font: { color: '9C5700' } })
        const top = dxf({ font: { b: true } })
        const r = `3:${last}`.split(':')
        return [
          `<conditionalFormatting sqref="K${r[0]}:K${r[1]}">` +
            `<cfRule type="containsText" dxfId="${ok}" priority="1" operator="containsText" text="Оплачен"><formula>NOT(ISERROR(SEARCH("Оплачен",K3)))</formula></cfRule>` +
            `<cfRule type="containsText" dxfId="${bad}" priority="2" operator="containsText" text="Отменён"><formula>NOT(ISERROR(SEARCH("Отменён",K3)))</formula></cfRule>` +
            `<cfRule type="containsText" dxfId="${warn}" priority="3" operator="containsText" text="Ожидает"><formula>NOT(ISERROR(SEARCH("Ожидает",K3)))</formula></cfRule>` +
          '</conditionalFormatting>',
          `<conditionalFormatting sqref="J${r[0]}:J${r[1]}"><cfRule type="colorScale" priority="4"><colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>`,
          `<conditionalFormatting sqref="G${r[0]}:G${r[1]}"><cfRule type="top10" dxfId="${top}" priority="5" rank="3"/></conditionalFormatting>`
        ].join('')
      }
    },
    {
      name: 'Справочник',
      rows: [['Товар', 'Цена'], ...products.map(([n, p]) => [n, { v: p, s: { numFmt: '#,##0.00\\ "₽"' } }])]
    }
  ]
})

const dir = join(process.cwd(), 'demo-files')
mkdirSync(dir, { recursive: true })
const xlsxPath = join(dir, `random-${seed}.xlsx`)
writeFileSync(xlsxPath, workbook)

// The same data as a Russian-Excel style CSV (semicolons, windows-1251-safe text, comma decimals).
const csvRows = [['№', 'Дата', 'Регион', 'Товар', 'Кол-во', 'Сумма', 'Статус']]
rows.forEach((r) => {
  const d = new Date(Date.UTC(1899, 11, 30) + r[2].v * 86400000)
  csvRows.push([r[0].v, d.toISOString().slice(0, 10).split('-').reverse().join('.'), r[3].v, r[5].v, r[6].v, r[9].v.toFixed(2).replace('.', ','), r[10].v])
})
const csv = csvRows.map((r) => r.map((v) => (/[;"\n]/.test(String(v)) ? `"${String(v).replace(/"/g, '""')}"` : v)).join(';')).join('\r\n') + '\r\n'
const csvPath = join(dir, `random-${seed}.csv`)
writeFileSync(csvPath, '﻿' + csv)

console.log(`seed ${seed}, ${rowCount} rows`)
console.log(xlsxPath)
console.log(csvPath)

// Builds a realistic .xlsx (the parts Excel writes) for tests, zipped with Node's zlib.
import zlib from 'node:zlib'

export function zip(files, { store = false } = {}) {
  const chunks = []
  const central = []
  let offset = 0
  for (const [name, content] of Object.entries(files)) {
    const data = Buffer.from(content)
    const comp = store ? data : zlib.deflateRawSync(data)
    const nameBuf = Buffer.from(name)
    const crc = zlib.crc32(data)
    const local = Buffer.alloc(30)
    local.writeUInt32LE(0x04034b50, 0)
    local.writeUInt16LE(20, 4)
    local.writeUInt16LE(0x0800, 6)
    local.writeUInt16LE(store ? 0 : 8, 8)
    local.writeUInt32LE(crc, 14)
    local.writeUInt32LE(comp.length, 18)
    local.writeUInt32LE(data.length, 22)
    local.writeUInt16LE(nameBuf.length, 26)
    chunks.push(local, nameBuf, comp)
    const cd = Buffer.alloc(46)
    cd.writeUInt32LE(0x02014b50, 0)
    cd.writeUInt16LE(20, 4)
    cd.writeUInt16LE(20, 6)
    cd.writeUInt16LE(0x0800, 8)
    cd.writeUInt16LE(store ? 0 : 8, 10)
    cd.writeUInt32LE(crc, 16)
    cd.writeUInt32LE(comp.length, 20)
    cd.writeUInt32LE(data.length, 24)
    cd.writeUInt16LE(nameBuf.length, 28)
    cd.writeUInt32LE(offset, 42)
    central.push(cd, nameBuf)
    offset += 30 + nameBuf.length + comp.length
  }
  const cdBuf = Buffer.concat(central)
  const eocd = Buffer.alloc(22)
  eocd.writeUInt32LE(0x06054b50, 0)
  eocd.writeUInt16LE(Object.keys(files).length, 8)
  eocd.writeUInt16LE(Object.keys(files).length, 10)
  eocd.writeUInt32LE(cdBuf.length, 12)
  eocd.writeUInt32LE(offset, 16)
  return new Uint8Array(Buffer.concat([...chunks, cdBuf, eocd]))
}

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>`

const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`

const theme = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme></a:themeElements></a:theme>`

// xf indexes used below:
// 0 default · 1 header (bold white on accent1, centered, thin borders, medium bottom)
// 2 currency "₽" with thin borders · 3 percent · 4 date dd.mm.yyyy · 5 short date (id 14)
// 6 wrap + top aligned · 7 italic red Times 14pt · 8 fill theme 4 tint 0.8 · 9 [Red] negative format
const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet ${NS}>
<numFmts count="3"><numFmt numFmtId="164" formatCode="#,##0.00\\ &quot;₽&quot;"/><numFmt numFmtId="165" formatCode="dd\\.mm\\.yyyy"/><numFmt numFmtId="166" formatCode="#,##0;[Red]\\-#,##0"/></numFmts>
<fonts count="4">
<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>
<font><b/><sz val="11"/><color theme="0"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>
<font><i/><sz val="14"/><color rgb="FFFF0000"/><name val="Times New Roman"/><family val="1"/></font>
<font><u/><sz val="11"/><color theme="10"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>
</fonts>
<fills count="4">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor theme="4"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor theme="4" tint="0.79998168889431442"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="3">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right><top style="thin"><color indexed="64"/></top><bottom style="medium"><color indexed="64"/></bottom><diagonal/></border>
<border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="11">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="164" fontId="0" fillId="0" borderId="2" xfId="0" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="10" fontId="0" fillId="0" borderId="2" xfId="0" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>
<xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>
</cellXfs>
<dxfs count="2">
<dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>
<dxf><font><b/><color rgb="FF006100"/></font><fill><patternFill><bgColor rgb="FFC6EFCE"/></patternFill></fill></dxf>
</dxfs>
</styleSheet>`

const sharedStrings = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst ${NS} count="12" uniqueCount="12">
<si><t>Регион</t></si>
<si><t>Выручка</t></si>
<si><t>Доля</t></si>
<si><t>Дата</t></si>
<si><t>Москва</t></si>
<si><t>Санкт-Петербург</t></si>
<si><r><t xml:space="preserve">Итого: </t></r><r><rPr><b/><sz val="11"/><color rgb="FFC00000"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t>важно</t></r><r><rPr><sz val="11"/><color theme="1"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t xml:space="preserve"> проверить</t></r></si>
<si><t>Скрытая строка</t></si>
<si><t>Строка 1
Строка 2</t></si>
<si><t>отчёт</t></si>
<si><t>H2O_x000D_</t></si>
<si><t>Казань</t></si>
</sst>`

// A1:E1 header (A1:A2 merged vertically, B1:C1 merged), C hidden? no: column D hidden.
const sheet1 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet ${NS}>
<dimension ref="B2:G9"/>
<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>
<sheetFormatPr defaultRowHeight="15"/>
<cols><col min="2" max="2" width="18.7109375" customWidth="1"/><col min="3" max="4" width="13.7109375" customWidth="1"/><col min="5" max="5" width="0" hidden="1" customWidth="1"/></cols>
<sheetData>
<row r="2"><c r="B2" s="1" t="s"><v>0</v></c><c r="C2" s="1" t="s"><v>1</v></c><c r="D2" s="1"/><c r="E2" s="1" t="inlineStr"><is><t>ID</t></is></c><c r="F2" s="1" t="s"><v>2</v></c><c r="G2" s="1" t="s"><v>3</v></c></row>
<row r="3"><c r="B3" s="6" t="s"><v>4</v></c><c r="C3" s="2"><v>1234.5</v></c><c r="D3" s="2"><f>C3*2</f><v>2469</v></c><c r="E3"><v>17</v></c><c r="F3" s="3"><v>-0.12</v></c><c r="G3" s="4"><v>45000</v></c></row>
<row r="4" hidden="1"><c r="B4" t="s"><v>7</v></c><c r="C4" s="2"><v>1</v></c></row>
<row r="5"><c r="B5" s="6" t="s"><v>5</v></c><c r="C5" s="2"><v>987654</v></c><c r="D5" s="2"><v>0</v></c><c r="E5"><v>19</v></c><c r="F5" s="3"><v>0.35</v></c><c r="G5" s="5"><v>45000.75</v></c></row>
<row r="6"><c r="B6" s="6" t="s"><v>6</v></c><c r="C6" s="9"><v>-500</v></c><c r="D6" t="b"><v>1</v></c><c r="F6" t="e"><v>#DIV/0!</v></c><c r="G6" s="10" t="s"><v>9</v></c></row>
<row r="7"><c r="B7" s="6" t="s"><v>8</v></c><c r="C7" s="7" t="str"><f>"x"</f><v>курсив</v></c><c r="D7" s="8"><v>3.14159</v></c><c r="F7" t="s"><v>10</v></c></row>
<row r="8"><c r="B8" t="s"><v>11</v></c><c r="C8" s="2"><v>10</v></c><c r="D8" s="2"><v>20</v></c><c r="F8"><v>0.5</v></c></row>
</sheetData>
<mergeCells count="2"><mergeCell ref="C2:D2"/><mergeCell ref="F7:G8"/></mergeCells>
<conditionalFormatting sqref="C3 C5 C8"><cfRule type="cellIs" dxfId="0" priority="2" operator="lessThan"><formula>100</formula></cfRule><cfRule type="cellIs" dxfId="1" priority="1" operator="greaterThanOrEqual"><formula>$C$5</formula></cfRule><cfRule type="expression" dxfId="0" priority="0"><formula>MOD(ROW(),2)=0</formula></cfRule></conditionalFormatting>
<conditionalFormatting sqref="F3:F5"><cfRule type="colorScale" priority="3"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>
<conditionalFormatting sqref="B3:B8"><cfRule type="containsText" dxfId="1" priority="4" operator="containsText" text="казань"><formula>NOT(ISERROR(SEARCH("казань",B3)))</formula></cfRule></conditionalFormatting>
<hyperlinks><hyperlink ref="G6" r:id="rId1"/><hyperlink ref="B5" location="Sheet2!A1" display="internal"/></hyperlinks>
<dataValidations count="1"><dataValidation type="list" allowBlank="1" sqref="B3:B8"><formula1>"Москва,Казань"</formula1></dataValidation></dataValidations>
<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>`

const sheet1Rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/hyperlink" Target="https://example.com/report?id=1" TargetMode="External"/></Relationships>`

const sheet2 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet ${NS}><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Второй лист</t></is></c><c r="B1"><v>42</v></c></row></sheetData></worksheet>`

const sheet3 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet ${NS}><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>secret</t></is></c></row></sheetData></worksheet>`

export function buildWorkbook({ activeTab = 0, store = false, hideFirst = false, parts = {} } = {}) {
  const workbook = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook ${NS}><workbookPr defaultThemeVersion="166925"/><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="12300" activeTab="${activeTab}"/></bookViews>
<sheets><sheet name="Отчёт" sheetId="1"${hideFirst ? ' state="hidden"' : ''} r:id="rId1"/><sheet name="Второй" sheetId="2" r:id="rId2"/><sheet name="Скрытый" sheetId="3" state="hidden" r:id="rId3"/></sheets></workbook>`
  const wbRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL}/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="${REL}/worksheet" Target="/xl/worksheets/sheet3.xml"/><Relationship Id="rId4" Type="${REL}/theme" Target="theme/theme1.xml"/><Relationship Id="rId5" Type="${REL}/styles" Target="styles.xml"/><Relationship Id="rId6" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/></Relationships>`
  return zip({
    '[Content_Types].xml': contentTypes,
    '_rels/.rels': rootRels,
    'xl/workbook.xml': workbook,
    'xl/_rels/workbook.xml.rels': wbRels,
    'xl/theme/theme1.xml': theme,
    'xl/styles.xml': styles,
    'xl/sharedStrings.xml': sharedStrings,
    'xl/worksheets/sheet1.xml': sheet1,
    'xl/worksheets/_rels/sheet1.xml.rels': sheet1Rels,
    'xl/worksheets/sheet2.xml': sheet2,
    'xl/worksheets/sheet3.xml': sheet3,
    ...parts
  }, { store })
}

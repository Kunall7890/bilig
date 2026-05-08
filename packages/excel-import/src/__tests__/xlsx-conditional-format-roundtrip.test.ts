import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'

import { exportXlsx, importXlsx } from '../index.js'

describe('xlsx conditional format roundtrip', () => {
  it('preserves imported basic conditional-format differential-format links', () => {
    const source = buildConditionalFormattingWorkbook()

    const imported = importXlsx(source, 'conditional-format-dxfs.xlsx')

    expect(imported.snapshot.sheets[0]?.metadata?.conditionalFormats).toHaveLength(3)

    const exported = exportXlsx(imported.snapshot)
    const exportedZip = unzipSync(exported)
    const exportedSheetXml = strFromU8(exportedZip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
    const exportedStylesXml = strFromU8(exportedZip['xl/styles.xml'] ?? new Uint8Array())

    expect(exportedSheetXml.match(/\bdxfId="/gu)).toHaveLength(3)
    expect(exportedSheetXml).toContain(
      '<cfRule type="cellIs" dxfId="1" priority="2" stopIfTrue="1" operator="equal"><formula>0</formula></cfRule>',
    )
    expect(exportedSheetXml).toContain('<cfRule type="expression" dxfId="2" priority="1"><formula>LEN(B1)&gt;0</formula></cfRule>')
    expect(exportedStylesXml).toContain('<dxfs count="3">')
    expect(exportedStylesXml).toContain('<color theme="5" tint="-0.249977111117893"/>')
    expect(exportedStylesXml).toContain('<border><bottom style="double"><color indexed="64"/></bottom></border>')
  })
})

function buildConditionalFormattingWorkbook(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([[0, 'Open']])
  XLSX.utils.book_append_sheet(workbook, sheet, 'Checks')

  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  zip['xl/worksheets/sheet1.xml'] = strToU8(conditionalFormattingWorksheetXml)
  zip['xl/styles.xml'] = strToU8(conditionalFormattingStylesXml)
  return zipSync(zip)
}

const conditionalFormattingWorksheetXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
  '<dimension ref="A1:B1"/>',
  '<sheetData><row r="1"><c r="A1"><v>0</v></c><c r="B1" t="str"><v>Open</v></c></row></sheetData>',
  '<conditionalFormatting sqref="A1">',
  '<cfRule type="cellIs" dxfId="1" priority="2" stopIfTrue="1" operator="equal"><formula>0</formula></cfRule>',
  '</conditionalFormatting>',
  '<conditionalFormatting sqref="B1">',
  '<cfRule type="expression" dxfId="2" priority="1"><formula>LEN(B1)&gt;0</formula></cfRule>',
  '<cfRule type="cellIs" dxfId="0" priority="3" operator="notEqual"><formula>&quot;Closed&quot;</formula></cfRule>',
  '</conditionalFormatting>',
  '</worksheet>',
].join('')

const conditionalFormattingStylesXml = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
  '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>',
  '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>',
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>',
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
  '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>',
  '<dxfs count="3">',
  '<dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/><bgColor indexed="64"/></patternFill></fill></dxf>',
  '<dxf><font><b/><color theme="5" tint="-0.249977111117893"/></font></dxf>',
  '<dxf><numFmt numFmtId="165" formatCode="0.00%;[Red]-0.00%"/><border><bottom style="double"><color indexed="64"/></bottom></border></dxf>',
  '</dxfs>',
  '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>',
  '</styleSheet>',
].join('')

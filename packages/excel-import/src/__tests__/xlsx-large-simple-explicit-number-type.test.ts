import { describe, expect, it } from 'vitest'
import { strToU8, unzipSync, zipSync } from 'fflate'

import { tryImportLargeSimpleXlsx } from '../xlsx-large-simple-import.js'

describe('large simple XLSX explicit number cell type import', () => {
  it('materializes cells marked t="n" as numeric values', () => {
    const bytes = buildWorkbookWithWorksheet(
      [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:B1"/>',
        '<sheetData><row r="1"><c r="A1" t="n"><v>13.5</v></c><c r="B1" t="inlineStr"><is><t>Label</t></is></c></row></sheetData>',
        '</worksheet>',
      ].join(''),
    )

    const imported = tryImportLargeSimpleXlsx(bytes, 'explicit-number-type.xlsx', unzipSync(bytes), { minByteLength: 0 })

    expect(imported?.stats?.cellCount).toBe(2)
    expect(imported?.stats?.valueCellCount).toBe(2)
    expect(imported?.snapshot.sheets[0]?.cells).toEqual([
      { address: 'A1', value: 13.5 },
      { address: 'B1', value: 'Label' },
    ])
  })
})

function buildWorkbookWithWorksheet(worksheetXml: string): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
    'xl/worksheets/sheet1.xml': strToU8(worksheetXml),
  })
}

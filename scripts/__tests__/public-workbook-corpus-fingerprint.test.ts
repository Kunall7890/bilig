import { strToU8, zipSync } from 'fflate'
import { afterEach, describe, expect, it, vi } from 'vitest'

describe('public workbook corpus fingerprinting', () => {
  afterEach(() => {
    vi.doUnmock('../../packages/excel-import/src/index.js')
    vi.resetModules()
  })

  it('fingerprints large-simple data-only workbooks without materializing public cell arrays', async () => {
    vi.doMock('../../packages/excel-import/src/index.js', () => ({
      importXlsx: () => {
        throw new Error('materialized import should not run for large-simple data-only fingerprints')
      },
    }))
    const { fingerprintWorkbookBytes } = await import('../public-workbook-corpus-workbook.ts')

    const first = fingerprintWorkbookBytes(buildLargeSimpleNumericWorkbookBytes(100_001), 'large-simple-data.xlsx')
    const second = fingerprintWorkbookBytes(buildLargeSimpleNumericWorkbookBytes(100_002), 'large-simple-data.xlsx')

    expect(first).toMatch(/^[a-f0-9]{64}$/u)
    expect(second).toMatch(/^[a-f0-9]{64}$/u)
    expect(first).not.toBe(second)
  })
})

function buildLargeSimpleNumericWorkbookBytes(rowCount: number): Uint8Array {
  const rows: string[] = []
  for (let row = 1; row <= rowCount; row += 1) {
    rows.push(`<row r="${String(row)}"><c r="A${String(row)}"><v>${String(row)}</v></c></row>`)
  }
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
    'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:A${String(rowCount)}"/>
  <sheetData>${rows.join('')}</sheetData>
</worksheet>`),
  })
}

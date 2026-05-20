import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import {
  prepareSheetJsParserXlsxBytes,
  stripNoOpEmptyRowsFromXlsx,
  stripStyleOnlyBlankCellsForSheetJs,
} from '../xlsx-style-only-blank-cells.js'

describe('stripStyleOnlyBlankCellsForSheetJs', () => {
  it('removes no-op empty rows while preserving meaningful row metadata', () => {
    const bytes = zipSync({
      'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetFormatPr defaultRowHeight="11.25"/>
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
    <row r="2" ht="11.25" customHeight="1"/>
    <row r="3" spans="1:6"/>
    <row r="4" ht="12" customHeight="1"/>
    <row r="5" hidden="1"/>
    <row r="6" ht="11.25" customHeight="1"></row>
    <row r="7"><c r="A7" s="2"/></row>
    <row r="8" s="2" customFormat="1"><c r="A8" s="2"/></row>
    <row r="9" ht="11.25" customHeight="1" x14ac:dyDescent="0.25"/>
  </sheetData>
</worksheet>`),
    })
    const stripped = stripStyleOnlyBlankCellsForSheetJs(bytes, unzipSync(bytes))
    const sheetXml = strFromU8(unzipSync(stripped)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())

    expect(sheetXml).toContain('<row r="1"><c r="A1"><v>1</v></c></row>')
    expect(sheetXml).not.toContain('r="2"')
    expect(sheetXml).not.toContain('r="3"')
    expect(sheetXml).toContain('<row r="4" ht="12" customHeight="1"/>')
    expect(sheetXml).toContain('<row r="5" hidden="1"/>')
    expect(sheetXml).not.toContain('r="6"')
    expect(sheetXml).not.toContain('r="7"')
    expect(sheetXml).toContain('<row r="8" s="2" customFormat="1"></row>')
    expect(sheetXml).not.toContain('r="9"')
  })

  it('can strip no-op rows while retaining blank style cells for artifact scans', () => {
    const bytes = zipSync({
      'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetFormatPr defaultRowHeight="11.25"/>
  <sheetData>
    <row r="1" ht="11.25" customHeight="1"/>
    <row r="2"><c r="A2" s="2"/></row>
  </sheetData>
</worksheet>`),
    })
    const stripped = stripNoOpEmptyRowsFromXlsx(bytes, unzipSync(bytes))
    const sheetXml = strFromU8(unzipSync(stripped)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())

    expect(sheetXml).not.toContain('r="1"')
    expect(sheetXml).toContain('<row r="2"><c r="A2" s="2"/></row>')
  })

  it('does not mutate the source zip while producing SheetJS parser bytes', () => {
    const bytes = zipSync({
      'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" s="2"/></row>
  </sheetData>
</worksheet>`),
    })
    const zip = unzipSync(bytes)
    const stripped = stripStyleOnlyBlankCellsForSheetJs(bytes, zip)
    const originalSheetXml = strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
    const strippedSheetXml = strFromU8(unzipSync(stripped)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())

    expect(originalSheetXml).toContain('<c r="A1" s="2"/>')
    expect(strippedSheetXml).not.toContain('<c r="A1" s="2"/>')
  })

  it('skips SheetJS parser rewrites when styled blank cells are below the configured threshold', () => {
    const bytes = zipSync({
      'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" s="2"/></row>
  </sheetData>
</worksheet>`),
    })
    const stripped = stripStyleOnlyBlankCellsForSheetJs(bytes, unzipSync(bytes), { minBlankCellCount: 2 })

    expect(stripped).toBe(bytes)
  })

  it('rewrites SheetJS parser bytes when styled blank cells reach the configured threshold', () => {
    const bytes = zipSync({
      'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" s="2"/></row>
    <row r="2"><c r="A2" s="2"/></row>
  </sheetData>
</worksheet>`),
    })
    const stripped = stripStyleOnlyBlankCellsForSheetJs(bytes, unzipSync(bytes), { minBlankCellCount: 2 })
    const strippedSheetXml = strFromU8(unzipSync(stripped)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())

    expect(stripped).not.toBe(bytes)
    expect(strippedSheetXml).not.toContain('<c r="A1" s="2"/>')
    expect(strippedSheetXml).not.toContain('<c r="A2" s="2"/>')
  })

  it('omits PowerPivot data model package parts from SheetJS parser bytes without reading lazy entries', () => {
    const bytes = zipSync({
      'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>`),
      'xl/model/item.data': new Uint8Array([1, 2, 3]),
      'xl/customData/item1.data': new Uint8Array([4, 5, 6]),
      'customXml/item1.xml': strToU8('<root/>'),
    })
    const zip = unzipSync(bytes)
    Object.defineProperty(zip, 'xl/model/item.data', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('data model bytes should stay out of the SheetJS parser package')
      },
    })
    const parserBytes = prepareSheetJsParserXlsxBytes(bytes, zip, { omitParserIgnoredPackageParts: true })
    const parserZip = unzipSync(parserBytes)

    expect(parserBytes).not.toBe(bytes)
    expect(parserZip['xl/worksheets/sheet1.xml']).toBeDefined()
    expect(parserZip['xl/model/item.data']).toBeUndefined()
    expect(parserZip['xl/customData/item1.data']).toBeUndefined()
    expect(parserZip['customXml/item1.xml']).toBeUndefined()
  })
})

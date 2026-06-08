import { describe, expect, it } from 'vitest'
import { strToU8, zipSync } from 'fflate'

import {
  externalWorkbookReferencesWarning,
  importWorkbookFile,
  importXlsx,
  inspectXlsx,
  XLSX_CONTENT_TYPE,
  XlsxImportSizeLimitExceededError,
} from '../index.js'
import { importXlsxFromZipByteSource } from '../xlsx-byte-source-import.js'

describe('large XLSX workbook routing', () => {
  it('keeps external defined names on the streaming path as a warning', () => {
    const bytes = buildWorkbook({
      definedNamesXml: '<definedNames><definedName name="ExternalList">[1]AdminName!$A$1</definedName></definedNames>',
      rowsXml: '<row r="1"><c r="A1"><v>1</v></c></row>',
      dimension: 'A1',
      paddingBytes: 1_100_000,
    })

    expect(() => inspectXlsx(bytes, 'external-defined-name-large-simple.xlsx')).toThrow(XlsxImportSizeLimitExceededError)
    expect(() => importXlsx(bytes, 'external-defined-name-large-simple.xlsx')).toThrow(XlsxImportSizeLimitExceededError)

    const imported = importXlsxFromZipByteSource(byteSourceFor(bytes), 'external-defined-name-large-simple.xlsx')

    expect(imported.stats?.cellCount).toBe(1)
    expect(imported.stats?.definedNameCount).toBe(0)
    expect(imported.warnings).toContain(externalWorkbookReferencesWarning)
    expect(imported.snapshot.sheets[0]?.cells).toEqual([{ address: 'A1', value: 1 }])

    const dispatched = importWorkbookFile(bytes, 'external-defined-name-large-simple.xlsx', XLSX_CONTENT_TYPE, {
      xlsx: { nativeOnly: true },
    })
    expect(dispatched.snapshot.sheets[0]?.cells).toEqual([{ address: 'A1', value: 1 }])
  })

  it('imports streaming-supported formula-heavy workbooks through the range-source path', () => {
    const bytes = buildFormulaWorkbook({ rowCount: 50_001, paddingBytes: 1_100_000 })

    const imported = importXlsxFromZipByteSource(byteSourceFor(bytes), 'streaming-formula-heavy.xlsx')

    expect(imported.stats?.formulaCellCount).toBe(50_001)
    expect(imported.snapshot.sheets[0]?.cells).toHaveLength(50_001)
  })

  it('keeps non-slicer worksheet extensions on the range-source streaming path instead of falling back to SheetJS', () => {
    const bytes = buildFormulaWorkbook({
      rowCount: 50_001,
      paddingBytes: 1_100_000,
      includeWorksheetExtensionList: true,
    })

    const imported = importXlsxFromZipByteSource(byteSourceFor(bytes), 'streaming-formula-heavy-extlst.xlsx')

    expect(imported.stats?.formulaCellCount).toBe(50_001)
    expect(imported.snapshot.sheets[0]?.cells).toHaveLength(50_001)
  })

  it('rejects unsafe SheetJS fallback materialization after the streaming path cannot own the package', () => {
    const bytes = buildFormulaWorkbook({
      rowCount: 50_001,
      paddingBytes: 1_100_000,
      includeUnsupportedPackagePart: true,
    })

    expect(() => importXlsxFromZipByteSource(byteSourceFor(bytes), 'unsupported-formula-heavy.xlsx')).toThrow(
      XlsxImportSizeLimitExceededError,
    )
  })

  it('rejects sparse large SheetJS fallback by source bytes before materialized reads', () => {
    const bytes = buildWorkbook({
      rowsXml: '<row r="1"><c r="A1"><v>1</v></c></row>',
      dimension: 'A1',
      paddingBytes: 1_100_000,
      includeUnsupportedPackagePart: true,
    })

    let thrown: unknown
    try {
      importXlsxFromZipByteSource(byteSourceFor(bytes), 'sparse-large-fallback.xlsx', {
        externalWorkbooks: [{ bytes: new Uint8Array([1]), fileName: 'external.xlsx' }],
      })
    } catch (error) {
      thrown = error
    }

    expect(bytes.byteLength).toBeGreaterThan(1_000_000)
    expect(thrown).toBeInstanceOf(XlsxImportSizeLimitExceededError)
    if (!(thrown instanceof XlsxImportSizeLimitExceededError)) {
      throw new Error('Expected XlsxImportSizeLimitExceededError')
    }
    expect(thrown.reason).toBe('source-byte-count')
    expect(thrown.sourceByteLength).toBe(bytes.byteLength)
  })

  it('does not treat limits false as an unbounded large SheetJS fallback opt-in', () => {
    const bytes = buildWorkbook({
      rowsXml: '<row r="1"><c r="A1"><v>1</v></c></row>',
      dimension: 'A1',
      paddingBytes: 1_100_000,
      includeUnsupportedPackagePart: true,
    })

    let thrown: unknown
    try {
      importXlsxFromZipByteSource(byteSourceFor(bytes), 'sparse-large-limits-false-fallback.xlsx', {
        externalWorkbooks: [{ bytes: new Uint8Array([1]), fileName: 'external.xlsx' }],
        limits: false,
      })
    } catch (error) {
      thrown = error
    }

    expect(thrown).toBeInstanceOf(XlsxImportSizeLimitExceededError)
    if (!(thrown instanceof XlsxImportSizeLimitExceededError)) {
      throw new Error('Expected XlsxImportSizeLimitExceededError')
    }
    expect(thrown.reason).toBe('source-byte-count')
    expect(thrown.sourceByteLength).toBe(bytes.byteLength)
  })
})

function byteSourceFor(bytes: Uint8Array): { readonly byteLength: number; readRange(start: number, end: number): Uint8Array } {
  return {
    byteLength: bytes.byteLength,
    readRange(start, end) {
      return bytes.subarray(start, end)
    },
  }
}

function buildFormulaWorkbook(options: {
  readonly rowCount: number
  readonly paddingBytes: number
  readonly includeUnsupportedPackagePart?: boolean
  readonly includeWorksheetExtensionList?: boolean
}): Uint8Array {
  const rows: string[] = []
  for (let row = 1; row <= options.rowCount; row += 1) {
    rows.push('<row r="' + String(row) + '"><c r="A' + String(row) + '"><f>B' + String(row) + '+1</f><v>' + String(row) + '</v></c></row>')
  }
  return buildWorkbook({
    rowsXml: rows.join(''),
    dimension: 'A1:A' + String(options.rowCount),
    paddingBytes: options.paddingBytes,
    ...(options.includeUnsupportedPackagePart !== undefined
      ? { includeUnsupportedPackagePart: options.includeUnsupportedPackagePart }
      : {}),
    ...(options.includeWorksheetExtensionList !== undefined
      ? { includeWorksheetExtensionList: options.includeWorksheetExtensionList }
      : {}),
  })
}

function buildWorkbook(options: {
  readonly rowsXml: string
  readonly dimension: string
  readonly definedNamesXml?: string
  readonly paddingBytes: number
  readonly includeUnsupportedPackagePart?: boolean
  readonly includeWorksheetExtensionList?: boolean
}): Uint8Array {
  return zipSync(
    {
      '[Content_Types].xml': strToU8(
        [
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
          '<Default Extension="bin" ContentType="application/octet-stream"/>',
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
          '<Default Extension="xml" ContentType="application/xml"/>',
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
          '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
          '</Types>',
        ].join(''),
      ),
      '_rels/.rels': strToU8(
        [
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>',
          '</Relationships>',
        ].join(''),
      ),
      'xl/workbook.xml': strToU8(
        [
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
          '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>',
          options.definedNamesXml ?? '',
          '</workbook>',
        ].join(''),
      ),
      'xl/_rels/workbook.xml.rels': strToU8(
        [
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>',
          '</Relationships>',
        ].join(''),
      ),
      'xl/worksheets/sheet1.xml': strToU8(
        [
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
          '<dimension ref="' + options.dimension + '"/>',
          '<sheetData>' + options.rowsXml + '</sheetData>',
          options.includeWorksheetExtensionList
            ? '<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:dataValidations count="1"/></ext></extLst>'
            : '',
          '</worksheet>',
        ].join(''),
      ),
      'docProps/padding.bin': deterministicBytes(options.paddingBytes),
      ...(options.includeUnsupportedPackagePart ? { 'xl/vbaProject.bin': deterministicBytes(64) } : {}),
    },
    { level: 0 },
  )
}

function deterministicBytes(length: number): Uint8Array {
  const bytes = new Uint8Array(length)
  for (let index = 0; index < bytes.length; index += 1) {
    bytes[index] = (index * 31 + 17) & 0xff
  }
  return bytes
}

import { strToU8, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { lazySheetCellMaterializationThreshold } from '../xlsx-large-simple-build-parsed-worksheet.js'
import { tryImportLargeSimpleXlsx } from '../xlsx-large-simple-import.js'
import { isLazyWorkbookSheetCells } from '../xlsx-large-simple-lazy-sheet-cells.js'
import { readLazyXlsxZipSourceByteLength, readXlsxZipEntriesLazy } from '../xlsx-zip.js'

describe('large simple XLSX import materialization lifetime', () => {
  it('releases ZIP source bytes before materializing independent sheets when ownership release is enabled', () => {
    const bytes = buildIndependentWorkbook([
      {
        name: 'First',
        path: 'xl/worksheets/sheet1.xml',
        xml: worksheetXml('A', 7),
      },
      {
        name: 'Second',
        path: 'xl/worksheets/sheet2.xml',
        xml: worksheetXml('B', 11),
      },
    ])
    const zip = readXlsxZipEntriesLazy(bytes)
    Object.defineProperty(zip, 'xl/worksheets/sheet1.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('sheet1 XML should be streamed instead of inflated')
      },
    })
    Object.defineProperty(zip, 'xl/worksheets/sheet2.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('sheet2 XML should be streamed instead of inflated')
      },
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'independent-sheets.xlsx', zip, {
      minByteLength: 0,
      releaseZipSource: true,
    })

    expect(imported?.snapshot.sheets.map((sheet) => sheet.cells)).toEqual([
      [
        { address: 'A1', value: 7 },
        { address: 'B1', value: 'A inline' },
      ],
      [
        { address: 'A1', value: 11 },
        { address: 'B1', value: 'B inline' },
      ],
    ])
    const phases = imported?.stats.phaseTelemetry.map((entry) => entry.phase) ?? []
    expect(phases).toEqual([
      'zip-setup',
      'worksheet-scan',
      'metadata-parsing',
      'shared-string-resolution',
      'style-parsing',
      'zip-source-release',
      'public-snapshot-materialization',
    ])
    expect(phases.indexOf('zip-source-release')).toBeLessThan(phases.indexOf('public-snapshot-materialization'))
  })

  it('finalizes independent sheets before ZIP source release when internal pre-release mode is enabled', () => {
    const bytes = buildIndependentWorkbook([
      {
        name: 'First',
        path: 'xl/worksheets/sheet1.xml',
        xml: worksheetXml('A', 7),
      },
      {
        name: 'Second',
        path: 'xl/worksheets/sheet2.xml',
        xml: worksheetXml('B', 11),
      },
    ])
    const zip = readXlsxZipEntriesLazy(bytes)

    const imported = tryImportLargeSimpleXlsx(bytes, 'independent-pre-release.xlsx', zip, {
      minByteLength: 0,
      releaseZipSource: true,
      allowPreReleaseSheetFinalization: true,
    })
    const phases = imported?.stats.phaseTelemetry.map((entry) => entry.phase) ?? []

    expect(imported?.snapshot.sheets.map((sheet) => sheet.name)).toEqual(['First', 'Second'])
    expect(imported?.snapshot.sheets.map((sheet) => sheet.cells)).toEqual([
      [
        { address: 'A1', value: 7 },
        { address: 'B1', value: 'A inline' },
      ],
      [
        { address: 'A1', value: 11 },
        { address: 'B1', value: 'B inline' },
      ],
    ])
    expect(phases.indexOf('public-snapshot-materialization')).toBeGreaterThanOrEqual(0)
    expect(phases.indexOf('public-snapshot-materialization')).toBeLessThan(phases.indexOf('zip-source-release'))
    expect(readLazyXlsxZipSourceByteLength(zip)).toBe(0)
  })

  it('keeps single-sheet imports on the default ZIP-release-before-materialization path', () => {
    const bytes = buildIndependentWorkbook([
      {
        name: 'Only',
        path: 'xl/worksheets/sheet1.xml',
        xml: worksheetXml('A', 7),
      },
    ])
    const zip = readXlsxZipEntriesLazy(bytes)

    const imported = tryImportLargeSimpleXlsx(bytes, 'single-sheet-pre-release.xlsx', zip, {
      minByteLength: 0,
      releaseZipSource: true,
      allowPreReleaseSheetFinalization: true,
    })
    const phases = imported?.stats.phaseTelemetry.map((entry) => entry.phase) ?? []

    expect(imported?.snapshot.sheets[0]?.cells).toEqual([
      { address: 'A1', value: 7 },
      { address: 'B1', value: 'A inline' },
    ])
    expect(phases.indexOf('zip-source-release')).toBeLessThan(phases.indexOf('public-snapshot-materialization'))
    expect(readLazyXlsxZipSourceByteLength(zip)).toBe(0)
  })

  it('finalizes shared-string sheets after string resolution in internal pre-release mode without dropping rich text', () => {
    const richStringXml = '<si><r><rPr><b/></rPr><t>Rich</t></r><r><t xml:space="preserve"> Value</t></r></si>'
    const bytes = buildSharedStringWorkbook(
      [
        {
          name: 'First',
          path: 'xl/worksheets/sheet1.xml',
          xml: sharedStringWorksheetXml([0, 1]),
        },
        {
          name: 'Second',
          path: 'xl/worksheets/sheet2.xml',
          xml: sharedStringWorksheetXml([2, 3]),
        },
      ],
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4">
  <si><t>Alpha</t></si>
  <si><t>Beta</t></si>
  <si><t>Gamma</t></si>
  ${richStringXml}
</sst>`,
      { 'xl/styles.xml': stylesXmlWithNumberFormat() },
    )
    const zip = readXlsxZipEntriesLazy(bytes)
    Object.defineProperty(zip, 'xl/sharedStrings.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('sharedStrings.xml should be streamed instead of fully inflated')
      },
    })
    Object.defineProperty(zip, 'xl/worksheets/sheet1.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('sheet1 XML should be streamed instead of inflated')
      },
    })
    Object.defineProperty(zip, 'xl/worksheets/sheet2.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('sheet2 XML should be streamed instead of inflated')
      },
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'shared-string-sheets.xlsx', zip, {
      minByteLength: 0,
      releaseZipSource: true,
      allowPreReleaseSheetFinalization: true,
    })

    expect(imported?.snapshot.sheets.map((sheet) => sheet.cells)).toEqual([
      [
        { address: 'A1', value: 'Alpha' },
        { address: 'B1', value: 'Beta' },
      ],
      [
        { address: 'A1', value: 'Gamma' },
        { address: 'B1', value: 'Rich Value' },
      ],
    ])
    expect(imported?.snapshot.sheets[1]?.metadata?.richTextArtifacts).toEqual({
      cells: [
        {
          address: 'B1',
          text: 'Rich Value',
          storage: 'sharedString',
          xml: richStringXml,
        },
      ],
    })
    expect(readLazyXlsxZipSourceByteLength(zip)).toBe(0)
    expect(imported?.stats.phaseTelemetry.map((entry) => entry.phase)).toEqual([
      'zip-setup',
      'worksheet-scan',
      'metadata-parsing',
      'shared-string-resolution',
      'public-snapshot-materialization',
      'style-parsing',
      'zip-source-release',
    ])
  })

  it('finalizes unstyled sheets before style parsing when the workbook has a style part', () => {
    const bytes = buildIndependentWorkbook(
      [
        {
          name: 'Plain',
          path: 'xl/worksheets/sheet1.xml',
          xml: worksheetXml('A', 7),
        },
        {
          name: 'Styled',
          path: 'xl/worksheets/sheet2.xml',
          xml: [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
            '<dimension ref="A1:B1"/>',
            '<sheetData><row r="1"><c r="A1" s="1"><v>11</v></c><c r="B1" t="inlineStr"><is><t>B inline</t></is></c></row></sheetData>',
            '</worksheet>',
          ].join(''),
        },
      ],
      { 'xl/styles.xml': stylesXmlWithNumberFormat() },
    )
    const zip = readXlsxZipEntriesLazy(bytes)

    const imported = tryImportLargeSimpleXlsx(bytes, 'unstyled-before-styles.xlsx', zip, {
      minByteLength: 0,
      releaseZipSource: true,
      allowPreReleaseSheetFinalization: true,
    })
    const phases = imported?.stats.phaseTelemetry.map((entry) => entry.phase) ?? []
    const firstMaterializationIndex = phases.indexOf('public-snapshot-materialization')
    const styleParsingIndex = phases.indexOf('style-parsing')

    expect(firstMaterializationIndex).toBeLessThan(styleParsingIndex)
    expect(imported?.snapshot.sheets.map((sheet) => sheet.cells)).toEqual([
      [
        { address: 'A1', value: 7 },
        { address: 'B1', value: 'A inline' },
      ],
      [
        { address: 'A1', value: 11, format: '00000' },
        { address: 'B1', value: 'B inline' },
      ],
    ])
    expect(readLazyXlsxZipSourceByteLength(zip)).toBe(0)
  })

  it('does not let relationship-backed package artifacts on one sheet pin independent sheet arenas', () => {
    const bytes = buildIndependentWorkbook(
      [
        {
          name: 'Plain',
          path: 'xl/worksheets/sheet1.xml',
          xml: worksheetXml('A', 7),
        },
        {
          name: 'Drawing',
          path: 'xl/worksheets/sheet2.xml',
          xml: [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
            '<dimension ref="A1:B1"/>',
            '<sheetData><row r="1"><c r="A1"><v>11</v></c><c r="B1" t="inlineStr"><is><t>B inline</t></is></c></row></sheetData>',
            '<drawing r:id="rIdDrawing"/>',
            '</worksheet>',
          ].join(''),
        },
      ],
      {
        'xl/worksheets/_rels/sheet2.xml.rels': [
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
          '<Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>',
          '</Relationships>',
        ].join(''),
        'xl/drawings/drawing1.xml':
          '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>',
      },
    )
    const zip = readXlsxZipEntriesLazy(bytes)

    const imported = tryImportLargeSimpleXlsx(bytes, 'relationship-backed-sheet.xlsx', zip, {
      minByteLength: 0,
      releaseZipSource: true,
      allowPreReleaseSheetFinalization: true,
    })
    const phases = imported?.stats.phaseTelemetry.map((entry) => entry.phase) ?? []
    const firstMaterializationIndex = phases.indexOf('public-snapshot-materialization')
    const sharedStringResolutionIndex = phases.indexOf('shared-string-resolution')

    expect(firstMaterializationIndex).toBeGreaterThanOrEqual(0)
    expect(firstMaterializationIndex).toBeLessThan(sharedStringResolutionIndex)
    expect(imported?.snapshot.sheets.map((sheet) => sheet.name)).toEqual(['Plain', 'Drawing'])
    expect(imported?.snapshot.sheets.map((sheet) => sheet.cells)).toEqual([
      [
        { address: 'A1', value: 7 },
        { address: 'B1', value: 'A inline' },
      ],
      [
        { address: 'A1', value: 11 },
        { address: 'B1', value: 'B inline' },
      ],
    ])
    expect(readLazyXlsxZipSourceByteLength(zip)).toBe(0)
  })

  it('resolves plain shared strings before exposing large lazy cell arrays', () => {
    const rowCount = 100_001
    const rows: string[] = []
    for (let row = 1; row <= rowCount; row += 1) {
      rows.push(`<row r="${String(row)}"><c r="A${String(row)}" t="s"><v>0</v></c></row>`)
    }
    const bytes = buildSharedStringWorkbook(
      [
        {
          name: 'Data',
          path: 'xl/worksheets/sheet1.xml',
          xml: [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
            `<dimension ref="A1:A${String(rowCount)}"/>`,
            `<sheetData>${rows.join('')}</sheetData>`,
            '</worksheet>',
          ].join(''),
        },
      ],
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${String(rowCount)}" uniqueCount="1">
  <si><t>Repeated shared label</t></si>
</sst>`,
    )
    const zip = readXlsxZipEntriesLazy(bytes)
    Object.defineProperty(zip, 'xl/sharedStrings.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('sharedStrings.xml should be streamed instead of fully inflated')
      },
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'large-plain-shared-strings.xlsx', zip, {
      minByteLength: 0,
      releaseZipSource: true,
    })
    const cells = imported?.snapshot.sheets[0]?.cells

    expect(cells).toHaveLength(rowCount)
    expect(cells?.[0]).toEqual({ address: 'A1', value: 'Repeated shared label' })
    expect(cells?.at(-1)).toEqual({ address: `A${String(rowCount)}`, value: 'Repeated shared label' })
    expect(imported?.preview.sheets[0]?.previewRows[0]?.[0]).toBe('Repeated shared label')
    expect(imported?.snapshot.sheets[0]?.metadata?.richTextArtifacts).toBeUndefined()
    expect(readLazyXlsxZipSourceByteLength(zip)).toBe(0)
  })

  it('uses lazy cells for medium-large unformatted sheets without changing values or order', () => {
    const rowCount = lazySheetCellMaterializationThreshold + 1
    const rows: string[] = []
    for (let row = 1; row <= rowCount; row += 1) {
      rows.push(`<row r="${String(row)}"><c r="A${String(row)}"><v>${String(row)}</v></c></row>`)
    }
    const bytes = buildIndependentWorkbook([
      {
        name: 'Data',
        path: 'xl/worksheets/sheet1.xml',
        xml: [
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
          `<dimension ref="A1:A${String(rowCount)}"/>`,
          `<sheetData>${rows.join('')}</sheetData>`,
          '</worksheet>',
        ].join(''),
      },
    ])
    const zip = readXlsxZipEntriesLazy(bytes)

    const imported = tryImportLargeSimpleXlsx(bytes, 'medium-large-unformatted.xlsx', zip, {
      minByteLength: 0,
      releaseZipSource: true,
    })
    const cells = imported?.snapshot.sheets[0]?.cells

    expect(isLazyWorkbookSheetCells(cells)).toBe(true)
    expect(cells).toHaveLength(rowCount)
    expect(cells?.[0]).toEqual({ address: 'A1', value: 1 })
    expect(cells?.at(-1)).toEqual({ address: `A${String(rowCount)}`, value: rowCount })
    expect(cells?.slice(0, 3)).toEqual([
      { address: 'A1', value: 1 },
      { address: 'A2', value: 2 },
      { address: 'A3', value: 3 },
    ])
    expect(imported?.preview.sheets[0]?.previewRows[0]?.[0]).toBe('1')
    expect(readLazyXlsxZipSourceByteLength(zip)).toBe(0)
  })
})

function worksheetXml(label: string, value: number): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<dimension ref="A1:B1"/>',
    `<sheetData><row r="1"><c r="A1"><v>${String(value)}</v></c><c r="B1" t="inlineStr"><is><t>${label} inline</t></is></c></row></sheetData>`,
    '</worksheet>',
  ].join('')
}

function buildIndependentWorkbook(
  sheets: readonly { readonly name: string; readonly path: string; readonly xml: string }[],
  extraEntries: Record<string, string> = {},
): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>${sheets
    .map((sheet, index) => `<sheet name="${sheet.name}" sheetId="${String(index + 1)}" r:id="rId${String(index + 1)}"/>`)
    .join('')}</sheets>
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${sheets
  .map(
    (sheet, index) =>
      `<Relationship Id="rId${String(index + 1)}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="${sheet.path.slice('xl/'.length)}"/>`,
  )
  .join('')}
</Relationships>`),
    ...Object.fromEntries(sheets.map((sheet) => [sheet.path, strToU8(sheet.xml)])),
    ...Object.fromEntries(Object.entries(extraEntries).map(([path, xml]) => [path, strToU8(xml)])),
  })
}

function sharedStringWorksheetXml(sharedStringIndexes: readonly number[]): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<dimension ref="A1:B1"/>',
    `<sheetData><row r="1"><c r="A1" t="s"><v>${String(sharedStringIndexes[0] ?? 0)}</v></c><c r="B1" t="s"><v>${String(
      sharedStringIndexes[1] ?? 0,
    )}</v></c></row></sheetData>`,
    '</worksheet>',
  ].join('')
}

function buildSharedStringWorkbook(
  sheets: readonly { readonly name: string; readonly path: string; readonly xml: string }[],
  sharedStringsXml: string,
  extraEntries: Record<string, string> = {},
): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>${sheets
    .map((sheet, index) => `<sheet name="${sheet.name}" sheetId="${String(index + 1)}" r:id="rId${String(index + 1)}"/>`)
    .join('')}</sheets>
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${sheets
  .map(
    (sheet, index) =>
      `<Relationship Id="rId${String(index + 1)}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="${sheet.path.slice('xl/'.length)}"/>`,
  )
  .join('')}
  <Relationship Id="rIdSharedStrings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`),
    'xl/sharedStrings.xml': strToU8(sharedStringsXml),
    ...Object.fromEntries(sheets.map((sheet) => [sheet.path, strToU8(sheet.xml)])),
    ...Object.fromEntries(Object.entries(extraEntries).map(([path, xml]) => [path, strToU8(xml)])),
  })
}

function stylesXmlWithNumberFormat(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="00000"/></numFmts>
  <fonts count="1"><font/></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>`
}

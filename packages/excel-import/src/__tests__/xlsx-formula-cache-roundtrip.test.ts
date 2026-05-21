import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import * as XLSX from 'xlsx'
import { describe, expect, it } from 'vitest'

import { attachRuntimeImage } from '@bilig/core'
import { ValueTag, type WorkbookSnapshot } from '@bilig/protocol'
import { exportXlsx, importXlsx } from '../index.js'

describe('formula cache roundtrip', () => {
  it('preserves cached formula result values through import and export', () => {
    const imported = importXlsx(buildFormulaCacheWorkbookBytes(), 'formula-cache-source.xlsx')
    const importedCells = new Map(imported.snapshot.sheets[0]?.cells.map((cell) => [cell.address, cell]) ?? [])

    expect(importedCells.get('B1')).toMatchObject({ formula: 'A1+A2', value: 1550 })
    expect(importedCells.get('B2')).toMatchObject({ formula: 'SUM(A1:A2)', value: 1550 })
    expect(importedCells.get('C2')).toMatchObject({ formula: 'B1/A1', value: 1.2916666666666667 })

    const exported = exportXlsx(imported.snapshot)
    const exportedSheetXml = strFromU8(unzipSync(exported)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())

    expect(cellXml(exportedSheetXml, 'B1')).toContain('<v>1550</v>')
    expect(cellXml(exportedSheetXml, 'B2')).toContain('<v>1550</v>')
    expect(cellXml(exportedSheetXml, 'C2')).toContain('<v>1.2916666666666667</v>')

    const reimported = importXlsx(exported, 'formula-cache-roundtrip.xlsx')
    const reimportedCells = new Map(reimported.snapshot.sheets[0]?.cells.map((cell) => [cell.address, cell]) ?? [])

    expect(reimportedCells.get('B1')).toMatchObject({ formula: 'A1+A2', value: 1550 })
    expect(reimportedCells.get('B2')).toMatchObject({ formula: 'SUM(A1:A2)', value: 1550 })
    expect(reimportedCells.get('C2')).toMatchObject({ formula: 'B1/A1', value: 1.2916666666666667 })
  })

  it('imports formula cells that omit cached result values', () => {
    const imported = importXlsx(buildFormulaWithoutCachedValueWorkbookBytes(), 'formula-without-cache.xlsx')
    const summary = imported.snapshot.sheets.find((sheet) => sheet.name === 'Summary')
    const cells = new Map(summary?.cells.map((cell) => [cell.address, cell]) ?? [])

    expect(cells.get('B2')).toMatchObject({ formula: 'Inputs!B2*Inputs!B3' })
    expect(cells.get('B2')?.value).toBeUndefined()
    expect(imported.preview.sheets.find((sheet) => sheet.name === 'Summary')?.previewRows[1]?.[1]).toBe('=Inputs!B2*Inputs!B3')
  })

  it('preserves string-literal references when expanding shared formulas', () => {
    const imported = importXlsx(buildSharedIndirectFormulaWorkbookBytes(), 'shared-indirect-formulas.xlsx')
    const contents = imported.snapshot.sheets.find((sheet) => sheet.name === 'Contents')
    const cells = new Map(contents?.cells.map((cell) => [cell.address, cell]) ?? [])

    expect(cells.get('B2')).toMatchObject({
      formula: `INDIRECT("'"&A2&"'!A2")`,
      value: 'First title',
    })
    expect(cells.get('B3')).toMatchObject({
      formula: `INDIRECT("'"&A3&"'!A2")`,
      value: 'Second title',
    })
    expect(cells.get('B3')?.formula).not.toContain("'!A3")

    const preview = imported.preview.sheets.find((sheet) => sheet.name === 'Contents')
    expect(preview?.previewRows[2]?.[1]).toBe(`=INDIRECT("'"&A3&"'!A2")`)
  })

  it('preserves leading formula whitespace across export round trips', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'formula-whitespace-export' },
      sheets: [
        {
          id: 1,
          name: 'A5',
          order: 0,
          cells: [
            {
              address: 'A1',
              formula: ' "   Canada Health Transfer (CHT) "&REPT(".",150)',
              value: '   Canada Health Transfer (CHT) ',
            },
          ],
        },
      ],
    }

    const reimported = importXlsx(exportXlsx(snapshot), 'formula-whitespace-export.xlsx')
    expect(reimported.snapshot.sheets[0]?.cells).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          address: 'A1',
          formula: ' "   Canada Health Transfer (CHT) "&REPT(".",150)',
        }),
      ]),
    )
  })

  it('exports future functions with Excel OOXML prefixes and normalizes them on import', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'future-function-export' },
      sheets: [
        {
          id: 1,
          name: 'Cases',
          order: 0,
          cells: [
            { address: 'A1', value: 'Text' },
            { address: 'A2', value: 'a' },
            { address: 'A3', value: '' },
            { address: 'A4', value: 'c' },
            { address: 'B1', value: 'Key' },
            { address: 'B2', value: 'a' },
            { address: 'B3', value: 'b' },
            { address: 'B4', value: 'c' },
            { address: 'C1', value: 'Value' },
            { address: 'C2', value: 10 },
            { address: 'C3', value: 20 },
            { address: 'C4', value: 30 },
            { address: 'D2', formula: 'TEXTJOIN("-",TRUE,A2:A4)' },
            { address: 'E2', formula: 'XLOOKUP("b",B2:B4,C2:C4)' },
            { address: 'F2', formula: 'XMATCH("b",B2:B4,0)' },
            { address: 'G2', formula: 'IF(TRUE,"XLOOKUP(",TEXTJOIN("-",TRUE,A2:A4))' },
            { address: 'H2', formula: '_xlfn.XLOOKUP("c",B2:B4,C2:C4)' },
            { address: 'I2', formula: 'FILTER(A2:A4,A2:A4<>"")' },
          ],
        },
      ],
    }

    const exported = exportXlsx(snapshot)
    const sheetXml = strFromU8(unzipSync(exported)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())

    expect(cellXml(sheetXml, 'D2')).toContain('<f>_xlfn.TEXTJOIN(&quot;-&quot;,TRUE,A2:A4)</f>')
    expect(cellXml(sheetXml, 'D2')).not.toContain('t="e"')
    expect(cellXml(sheetXml, 'E2')).toContain('<f>_xlfn.XLOOKUP(&quot;b&quot;,B2:B4,C2:C4)</f>')
    expect(cellXml(sheetXml, 'F2')).toContain('<f>_xlfn.XMATCH(&quot;b&quot;,B2:B4,0)</f>')
    expect(cellXml(sheetXml, 'G2')).toContain('&quot;XLOOKUP(&quot;,_xlfn.TEXTJOIN')
    expect(cellXml(sheetXml, 'H2')).toContain('<f>_xlfn.XLOOKUP(&quot;c&quot;,B2:B4,C2:C4)</f>')
    expect(cellXml(sheetXml, 'H2')).not.toContain('_xlfn._xlfn.')
    expect(cellXml(sheetXml, 'I2')).toContain('<f>_xlfn._xlws.FILTER(A2:A4,A2:A4&lt;&gt;&quot;&quot;)</f>')

    const reimported = importXlsx(exported, 'future-function-export.xlsx')
    const cells = new Map(reimported.snapshot.sheets[0]?.cells.map((cell) => [cell.address, cell]) ?? [])

    expect(cells.get('D2')).toMatchObject({ formula: 'TEXTJOIN("-",TRUE,A2:A4)' })
    expect(cells.get('E2')).toMatchObject({ formula: 'XLOOKUP("b",B2:B4,C2:C4)' })
    expect(cells.get('F2')).toMatchObject({ formula: 'XMATCH("b",B2:B4,0)' })
    expect(cells.get('G2')).toMatchObject({ formula: 'IF(TRUE,"XLOOKUP(",TEXTJOIN("-",TRUE,A2:A4))' })
    expect(cells.get('H2')).toMatchObject({ formula: 'XLOOKUP("c",B2:B4,C2:C4)' })
    expect(cells.get('I2')).toMatchObject({ formula: 'FILTER(A2:A4,A2:A4<>"")' })
  })

  it('exports native dynamic array spills with Desktop Excel metadata and cached children', () => {
    const snapshot = attachRuntimeImage(
      {
        version: 1,
        workbook: {
          name: 'lambda-spill-export',
          metadata: {
            spills: [{ sheetName: 'Sheet1', address: 'B1', rows: 3, cols: 1 }],
          },
        },
        sheets: [
          {
            id: 1,
            name: 'Sheet1',
            order: 0,
            cells: [
              { address: 'A1', value: 1 },
              { address: 'A2', value: 2 },
              { address: 'A3', value: 3 },
              { address: 'B1', formula: 'MAP(A1:A3,LAMBDA(x,x*2))' },
            ],
          },
        ],
      } satisfies WorkbookSnapshot,
      {
        version: 1,
        templateBank: [],
        formulaInstances: [],
        formulaValues: [],
        cellValues: [
          { sheetName: 'Sheet1', row: 0, col: 1, value: { tag: ValueTag.Number, value: 2 } },
          { sheetName: 'Sheet1', row: 1, col: 1, value: { tag: ValueTag.Number, value: 4 } },
          { sheetName: 'Sheet1', row: 2, col: 1, value: { tag: ValueTag.Number, value: 6 } },
        ],
      },
    )

    const exported = exportXlsx(snapshot)
    const zip = unzipSync(exported)
    const sheetXml = strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
    const metadataXml = strFromU8(zip['xl/metadata.xml'] ?? new Uint8Array())
    const workbookRelsXml = strFromU8(zip['xl/_rels/workbook.xml.rels'] ?? new Uint8Array())
    const contentTypesXml = strFromU8(zip['[Content_Types].xml'] ?? new Uint8Array())

    expect(cellXml(sheetXml, 'B1')).toContain('cm="1"')
    expect(cellXml(sheetXml, 'B1')).toContain('<f t="array" ref="B1:B3">_xlfn.MAP(A1:A3,_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2))</f>')
    expect(cellXml(sheetXml, 'B1')).toContain('<v>2</v>')
    expect(cellXml(sheetXml, 'B2')).toContain('<v>4</v>')
    expect(cellXml(sheetXml, 'B3')).toContain('<v>6</v>')
    expect(metadataXml).toContain('dynamicArrayProperties')
    expect(workbookRelsXml).toContain('relationships/sheetMetadata')
    expect(contentTypesXml).toContain('spreadsheetml.sheetMetadata+xml')

    const reimported = importXlsx(exported, 'lambda-spill-export.xlsx')
    const cells = new Map(reimported.snapshot.sheets[0]?.cells.map((cell) => [cell.address, cell]) ?? [])
    expect(cells.get('B1')).toMatchObject({ formula: 'MAP(A1:A3,LAMBDA(x,x*2))', value: 2 })
    expect(cells.get('B2')).toMatchObject({ value: 4 })
    expect(cells.get('B3')).toMatchObject({ value: 6 })
    expect(reimported.snapshot.workbook.metadata?.spills).toEqual([{ sheetName: 'Sheet1', address: 'B1', rows: 3, cols: 1 }])
  })
})

function buildFormulaCacheWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([
    [1200, null, null],
    [350, null, null],
  ])
  sheet.B1 = { t: 'n', f: 'A1+A2', v: 1550 }
  sheet.B2 = { t: 'n', f: 'SUM(A1:A2)', v: 1550 }
  sheet.C2 = { t: 'n', f: 'B1/A1', v: 1.2916666666666667 }
  sheet['!ref'] = 'A1:C2'

  XLSX.utils.book_append_sheet(workbook, sheet, 'FormulaCache')
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
}

function buildFormulaWithoutCachedValueWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ['Metric', 'Value'],
      ['Units', 40],
      ['Price', 1200],
    ]),
    'Inputs',
  )
  const summary = XLSX.utils.aoa_to_sheet([
    ['Metric', 'Value'],
    ['Revenue', null],
  ])
  summary.B2 = { t: 'n', f: 'Inputs!B2*Inputs!B3', v: 48_000 }
  summary['!ref'] = 'A1:B2'
  XLSX.utils.book_append_sheet(workbook, summary, 'Summary')

  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  const sheetXml = strFromU8(zip['xl/worksheets/sheet2.xml'] ?? new Uint8Array())
  zip['xl/worksheets/sheet2.xml'] = strToU8(replaceCellXml(sheetXml, 'B2', '<c r="B2"><f>Inputs!B2*Inputs!B3</f></c>'))
  return zipSync(zip)
}

function buildSharedIndirectFormulaWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const contents = XLSX.utils.aoa_to_sheet([
    ['Sheet', 'Title'],
    ['Data 1', null],
    ['Data 2', null],
  ])
  contents.B2 = { t: 'str', f: `INDIRECT("'"&A2&"'!A2")`, v: 'First title' }
  contents.B3 = { t: 'str', f: `INDIRECT("'"&A3&"'!A2")`, v: 'Second title' }
  contents['!ref'] = 'A1:B3'

  XLSX.utils.book_append_sheet(workbook, contents, 'Contents')
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([[], ['First title']]), 'Data 1')
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([[], ['Second title']]), 'Data 2')

  const zip = unzipSync(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))
  const sheetXml = strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
  zip['xl/worksheets/sheet1.xml'] = strToU8(
    replaceCellXml(
      replaceCellXml(
        sheetXml,
        'B2',
        '<c r="B2" t="str"><f t="shared" ref="B2:B3" si="0">INDIRECT(&quot;\'&quot;&amp;A2&amp;&quot;\'!A2&quot;)</f><v>First title</v></c>',
      ),
      'B3',
      '<c r="B3" t="str"><f t="shared" si="0"/><v>Second title</v></c>',
    ),
  )
  return zipSync(zip)
}

function cellXml(sheetXml: string, address: string): string {
  return sheetXml.match(new RegExp(`<c[^>]* r="${address}"[^>]*>[\\s\\S]*?<\\/c>`))?.[0] ?? ''
}

function replaceCellXml(sheetXml: string, address: string, replacement: string): string {
  const pattern = new RegExp(`<c\\b(?=[^>]*\\br="${address}")[\\s\\S]*?<\\/c>`, 'u')
  if (!pattern.test(sheetXml)) {
    throw new Error(`Missing fixture cell ${address}`)
  }
  return sheetXml.replace(pattern, replacement)
}

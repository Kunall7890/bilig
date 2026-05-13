import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import * as XLSX from 'xlsx'
import { describe, expect, it } from 'vitest'

import type { WorkbookSnapshot } from '@bilig/protocol'
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

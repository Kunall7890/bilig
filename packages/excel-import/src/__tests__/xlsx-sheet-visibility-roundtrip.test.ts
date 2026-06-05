import { describe, expect, it } from 'vitest'
import { writeSimpleXlsxWorkbook } from '@bilig/xlsx'

import { exportXlsx, importXlsx } from '../index.js'
import { patchXlsxTestZipText, readXlsxTestZipText } from './xlsx-test-helpers.js'

describe('sheet visibility roundtrip', () => {
  it('preserves hidden and very hidden worksheet state', () => {
    const imported = importXlsx(buildSheetVisibilityWorkbookBytes(), 'sheet-visibility.xlsx')

    expect(imported.snapshot.sheets.map((sheet) => ({ name: sheet.name, visibility: sheet.metadata?.visibility }))).toEqual([
      { name: 'Inputs', visibility: undefined },
      { name: 'Support', visibility: 'hidden' },
      { name: 'Audit', visibility: 'veryHidden' },
    ])

    const exportedWorkbookXml = readXlsxTestZipText(exportXlsx(imported.snapshot), 'xl/workbook.xml')
    expect(exportedWorkbookXml).toContain('<sheet name="Inputs"')
    expect(exportedWorkbookXml).toMatch(/<sheet\b[^>]*name="Support"[^>]*state="hidden"/u)
    expect(exportedWorkbookXml).toMatch(/<sheet\b[^>]*name="Audit"[^>]*state="veryHidden"/u)
  })
})

function buildSheetVisibilityWorkbookBytes(): Uint8Array {
  const bytes = writeSimpleXlsxWorkbook({
    sheets: [
      { name: 'Inputs', cells: [{ address: 'A1', row: 0, col: 0, value: 'visible' }] },
      { name: 'Support', cells: [{ address: 'A1', row: 0, col: 0, value: 'hidden' }] },
      { name: 'Audit', cells: [{ address: 'A1', row: 0, col: 0, value: 'very hidden' }] },
    ],
  })
  return patchXlsxTestZipText(bytes, 'xl/workbook.xml', (workbookXml) =>
    workbookXml
      .replace('<sheet name="Support" sheetId="2"', '<sheet name="Support" sheetId="2" state="hidden"')
      .replace('<sheet name="Audit" sheetId="3"', '<sheet name="Audit" sheetId="3" state="veryHidden"'),
  )
}

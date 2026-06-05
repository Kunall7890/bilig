import { describe, expect, it } from 'vitest'
import { writeSimpleXlsxWorkbook } from '@bilig/xlsx'
import { SpreadsheetEngine } from '@bilig/core'
import { ValueTag } from '@bilig/protocol'

import { exportXlsx, importXlsx } from '../index.js'
import { patchXlsxTestZipText, readXlsxTestZipText } from './xlsx-test-helpers.js'

describe('1904 date system import', () => {
  it('preserves workbookPr date1904 and evaluates date formulas in the workbook date system', () => {
    const imported = importXlsx(buildDate1904WorkbookBytes(), 'date1904-finance-dates.xlsx')

    expect(imported.snapshot.workbook.metadata?.calculationSettings).toMatchObject({
      dateSystem: '1904',
    })

    const engine = new SpreadsheetEngine({ workbookName: 'date1904-import' })
    engine.importSnapshot(imported.snapshot)

    expect(engine.getCellValue('Date1904', 'B2')).toEqual({ tag: ValueTag.Number, value: 1904 })
    expect(engine.getCellValue('Date1904', 'C2')).toMatchObject({
      tag: ValueTag.String,
      value: '1904-01-02',
    })

    expect(workbookXml(exportXlsx(engine.exportSnapshot()))).toContain('date1904="1"')
  })
})

function buildDate1904WorkbookBytes(): Uint8Array {
  const bytes = writeSimpleXlsxWorkbook({
    sheets: [
      {
        name: 'Date1904',
        cells: [
          { address: 'A1', row: 0, col: 0, value: 'Serial' },
          { address: 'B1', row: 0, col: 1, value: 'Year' },
          { address: 'C1', row: 0, col: 2, value: 'Text' },
          { address: 'A2', row: 1, col: 0, value: 1 },
          { address: 'B2', row: 1, col: 1, formula: 'YEAR(A2)', value: 1904 },
          { address: 'C2', row: 1, col: 2, formula: 'TEXT(A2,"yyyy-mm-dd")', value: '1904-01-02' },
        ],
      },
    ],
  })
  return patchXlsxTestZipText(bytes, 'xl/workbook.xml', (sourceWorkbookXml) =>
    /<workbookPr\b/u.test(sourceWorkbookXml)
      ? sourceWorkbookXml.replace(/<workbookPr\b([^>]*)\/>/u, '<workbookPr$1 date1904="1"/>')
      : sourceWorkbookXml.replace(/<sheets\b/u, '<workbookPr date1904="1"/><sheets'),
  )
}

function workbookXml(bytes: Uint8Array): string {
  return readXlsxTestZipText(bytes, 'xl/workbook.xml')
}

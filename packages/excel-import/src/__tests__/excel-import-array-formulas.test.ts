import { describe, expect, it } from 'vitest'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'

import { SpreadsheetEngine } from '@bilig/core'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import { writeSimpleXlsxWorkbook } from '@bilig/xlsx'
import { importXlsx } from '../index.js'

function buildArrayFormulaWorkbook(): Uint8Array {
  const zip = unzipSync(
    writeSimpleXlsxWorkbook({
      sheets: [
        {
          name: 'Sheet1',
          cells: [
            { address: 'A1', row: 0, col: 0, value: 1 },
            { address: 'B1', row: 0, col: 1, value: 0 },
            { address: 'D1', row: 0, col: 3, value: 1 },
            { address: 'F1', row: 0, col: 5, formula: 'MMULT(MINVERSE(A1:B2),D1:D2)', value: 1 },
            { address: 'H1', row: 0, col: 7, formula: 'INDEX(MMULT(MINVERSE(A1:B2),D1:D2),1,1)' },
            { address: 'A2', row: 1, col: 0, value: 0 },
            { address: 'B2', row: 1, col: 1, value: 1 },
            { address: 'D2', row: 1, col: 3, value: 2 },
            { address: 'F2', row: 1, col: 5, value: 2 },
            { address: 'H2', row: 1, col: 7, formula: 'INDEX(MMULT(MINVERSE(A1:B2),D1:D2),2,1)' },
          ],
        },
      ],
    }),
  )
  zip['xl/worksheets/sheet1.xml'] = strToU8(
    strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array()).replace(
      '<f>MMULT(MINVERSE(A1:B2),D1:D2)</f>',
      '<f t="array" ref="F1:F2">MMULT(MINVERSE(A1:B2),D1:D2)</f>',
    ),
  )
  return zipSync(zip)
}

describe('excel import array formulas', () => {
  it('imports legacy array-formula ranges so cached follower values do not block the leader', async () => {
    const imported = importXlsx(buildArrayFormulaWorkbook(), 'array-formula-import-evaluation.xlsx')

    expect(imported.snapshot.workbook.metadata?.spills).toEqual([{ sheetName: 'Sheet1', address: 'F1', rows: 2, cols: 1 }])

    const engine = new SpreadsheetEngine({ workbookName: 'issue-108-array-formula-import' })
    await engine.ready()
    engine.importSnapshot(imported.snapshot)

    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'F2')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'H1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'H2')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'F1')).not.toEqual({ tag: ValueTag.Error, code: ErrorCode.Blocked })
  })
})

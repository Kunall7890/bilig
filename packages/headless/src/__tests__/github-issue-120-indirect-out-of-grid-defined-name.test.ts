import { describe, expect, it } from 'vitest'
import { ValueTag, type CellValue, type WorkbookSnapshot } from '@bilig/protocol'
import { WorkPaper } from '../index.js'

function expectNumber(value: CellValue, expected: number): void {
  expect(value.tag).toBe(ValueTag.Number)
  if (value.tag !== ValueTag.Number) {
    throw new Error(`Expected number ${String(expected)}, received ${JSON.stringify(value)}`)
  }
  expect(value.value).toBeCloseTo(expected, 12)
}

describe('GitHub issue #120 INDIRECT out-of-grid defined names', () => {
  it('resolves names like change1 through INDIRECT instead of invalid far-right cells', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'issue-120-indirect-change-name',
        metadata: {
          definedNames: [
            { name: 'change1', value: { kind: 'range-ref', sheetName: 'YieldChanges', startAddress: 'A1', endAddress: 'A4' } },
            { name: 'change2', value: { kind: 'range-ref', sheetName: 'YieldChanges', startAddress: 'B1', endAddress: 'B4' } },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'YieldChanges',
          order: 0,
          cells: [
            { address: 'A1', value: 1 },
            { address: 'A2', value: 2 },
            { address: 'A3', value: 3 },
            { address: 'A4', value: 4 },
            { address: 'B1', value: 2 },
            { address: 'B2', value: 4 },
            { address: 'B3', value: 6 },
            { address: 'B4', value: 8 },
          ],
        },
        {
          id: 2,
          name: 'Main',
          order: 1,
          cells: [
            { address: 'I5', value: 'change1' },
            { address: 'J4', value: 'change2' },
            { address: 'J5', formula: 'COVARIANCE.P(INDIRECT($I5),INDIRECT(J$4))' },
            { address: 'T5', formula: 'CORREL(INDIRECT($I5),INDIRECT(J$4))' },
          ],
        },
      ],
    }

    const workbook = WorkPaper.buildFromSnapshot(snapshot, { maxRows: 32, maxColumns: 32, useColumnIndex: true })
    const mainSheetId = workbook.getSheetId('Main')!

    try {
      expectNumber(workbook.getCellValue({ sheet: mainSheetId, row: 4, col: 9 }), 2.5)
      expectNumber(workbook.getCellValue({ sheet: mainSheetId, row: 4, col: 19 }), 1)
    } finally {
      workbook.dispose()
    }
  })
})

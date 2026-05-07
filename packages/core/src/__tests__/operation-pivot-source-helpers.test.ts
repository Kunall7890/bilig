import { describe, expect, it } from 'vitest'
import { cellTouchesOperationPivotSource } from '../engine/services/operation-pivot-source-helpers.js'

describe('operation pivot source helpers', () => {
  it('does not list pivots when the workbook has none', () => {
    let listPivotsCalls = 0

    expect(
      cellTouchesOperationPivotSource({
        workbook: {
          hasPivots: () => false,
          listPivots: () => {
            listPivotsCalls += 1
            return []
          },
        },
        sheetName: 'Sheet1',
        row: 1,
        col: 1,
      }),
    ).toBe(false)
    expect(listPivotsCalls).toBe(0)
  })

  it('matches cells inside pivot source ranges on the same sheet only', () => {
    const workbook = {
      hasPivots: () => true,
      listPivots: () => [
        {
          source: {
            sheetName: 'Data',
            startAddress: 'B2',
            endAddress: 'D5',
          },
        },
      ],
    }

    expect(cellTouchesOperationPivotSource({ workbook, sheetName: 'Data', row: 1, col: 1 })).toBe(true)
    expect(cellTouchesOperationPivotSource({ workbook, sheetName: 'Data', row: 4, col: 3 })).toBe(true)
    expect(cellTouchesOperationPivotSource({ workbook, sheetName: 'Data', row: 5, col: 3 })).toBe(false)
    expect(cellTouchesOperationPivotSource({ workbook, sheetName: 'Other', row: 1, col: 1 })).toBe(false)
  })
})

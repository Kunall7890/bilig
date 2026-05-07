import { ValueTag, type CellValue } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'

type TestCell = string | number | null

function cellValue(workbook: WorkPaper, ref: string): CellValue {
  const address = workbook.simpleCellAddressFromString(ref)
  if (!address) {
    throw new Error(`Expected ${ref} to resolve`)
  }
  return workbook.getCellValue(address)
}

function issueWorkbook(formula: string, useColumnIndex: boolean): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      Sheet1: [
        [formula, '=A1', 10],
        [null, 'exclude', 'include'],
      ] satisfies TestCell[][],
    },
    { maxRows: 8, maxColumns: 8, useColumnIndex },
  )
}

describe('GitHub issue #132 SUMIFS excluded criteria cycle', () => {
  it.each([['=SUMIFS(B1:C1,B2:C2,"include")'], ['=IFERROR(SUMIFS(B1:C1,B2:C2,"include"),"")']])(
    'ignores excluded self-referential sum cells for %s',
    (formula) => {
      for (const useColumnIndex of [false, true]) {
        const workbook = issueWorkbook(formula, useColumnIndex)

        expect(cellValue(workbook, 'Sheet1!A1')).toEqual({
          tag: ValueTag.Number,
          value: 10,
        })
        expect(cellValue(workbook, 'Sheet1!B1')).toEqual({
          tag: ValueTag.Number,
          value: 10,
        })
        expect(workbook.getCellDisplayValue(workbook.simpleCellAddressFromString('Sheet1!A1')!)).toBe('10')

        workbook.dispose()
      }
    },
  )

  it.each([false, true])('compacts horizontal corpus-style criteria before cycle detection with useColumnIndex=%s', (useColumnIndex) => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Sheet1: [
          ['=IFERROR(SUMIFS(B1:D1,B2:D2,"<="&A4,B3:D3,"Amort\'n Expense"),"")', '=A1', 2958.333333, 4],
          [null, 1, 2, 3],
          [null, 'Excluded', "Amort'n Expense", 'Other'],
          [2, null, null, null],
        ],
      },
      { maxRows: 8, maxColumns: 8, useColumnIndex },
    )

    expect(cellValue(workbook, 'Sheet1!A1')).toEqual({
      tag: ValueTag.Number,
      value: 2958.333333,
    })
    expect(cellValue(workbook, 'Sheet1!B1')).toEqual({
      tag: ValueTag.Number,
      value: 2958.333333,
    })
    expect(workbook.getCellDisplayValue(workbook.simpleCellAddressFromString('Sheet1!A1')!)).toBe('2958.333333')

    workbook.dispose()
  })
})

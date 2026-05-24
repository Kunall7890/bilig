import { describe, expect, it } from 'vitest'
import { checkWorkbookRef, checkWorkbookRefData, isWorkbookRefData } from '../index.js'

class RangeData {
  readonly kind = 'range'
  readonly id = 'range_Sheet1_A1'
  readonly label = 'Sheet1!A1'
  readonly range = {
    sheetName: 'Sheet1',
    startAddress: 'A1',
    endAddress: 'A1',
  }
}

class CellRangeData {
  readonly sheetName = 'Sheet1'
  readonly startAddress = 'A1'
  readonly endAddress = 'A1'
}

class RowsWhereData {
  readonly column = 'Region'
  readonly op = 'eq'
  readonly value = 'West'
}

describe('@bilig/workbook ref data boundary', () => {
  it('rejects custom-prototype transported ref roots', () => {
    expect(checkWorkbookRefData(new RangeData())).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_type',
          path: 'ref',
          message: 'Workbook ref data ref must be an object record',
        },
      ],
    })
    expect(isWorkbookRefData(new RangeData())).toBe(false)
    expect(checkWorkbookRef(new RangeData())).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_type',
          path: 'ref',
          message: 'Workbook ref data ref must be an object record',
        },
      ],
    })
  })

  it('rejects custom-prototype nested range and row selector data', () => {
    expect(
      checkWorkbookRefData({
        kind: 'range',
        id: 'range_Sheet1_A1',
        label: 'Sheet1!A1',
        range: new CellRangeData(),
      }),
    ).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_type',
          path: 'ref.range',
          message: 'Workbook ref data ref.range must be an object record',
        },
      ],
    })

    expect(
      checkWorkbookRefData({
        kind: 'rows',
        id: 'rows_region_west',
        label: 'rows where Region eq West',
        where: new RowsWhereData(),
      }),
    ).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_type',
          path: 'ref.where',
          message: 'Workbook ref data ref.where must be an object record',
        },
      ],
    })
  })

  it('rejects custom-prototype nested table refs', () => {
    expect(
      checkWorkbookRefData({
        kind: 'column',
        id: 'sales_amount',
        label: 'Sales.Amount',
        name: 'Amount',
        table: {
          kind: 'table',
          id: 'sales',
          label: 'Sales',
          name: 'Sales',
        },
        rows: {
          kind: 'rows',
          id: 'sales_west',
          label: 'Sales rows where Region eq West',
          table: new (class SalesTableData {
            readonly kind = 'table'
            readonly id = 'sales'
            readonly label = 'Sales'
            readonly name = 'Sales'
          })(),
          where: {
            column: 'Region',
            op: 'eq',
            value: 'West',
          },
        },
      }),
    ).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_type',
          path: 'ref.rows.table',
          message: 'Workbook ref data ref.rows.table must be an object record',
        },
      ],
    })
  })
})

import { describe, expect, it } from 'vitest'
import {
  checkWorkbookRef,
  checkWorkbookRefData,
  collectWorkbookRefData,
  findRows,
  hydrateWorkbookRef,
  hydrateWorkbookRefs,
  isWorkbookRefData,
  toWorkbookRefData,
} from '../index.js'

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

function withThrowingObjectPrototypeGetters(keys: readonly string[], run: () => void): number {
  const originals = new Map<string, PropertyDescriptor | undefined>()
  let calls = 0
  for (const key of keys) {
    originals.set(key, Object.getOwnPropertyDescriptor(Object.prototype, key))
    // oxlint-disable-next-line eslint(no-extend-native) -- Prototype pollution is the exact boundary condition under test.
    Object.defineProperty(Object.prototype, key, {
      configurable: true,
      get() {
        calls += 1
        throw new Error(`Inherited ${key} getter must not run`)
      },
    })
  }
  try {
    run()
    return calls
  } finally {
    for (const [key, descriptor] of originals) {
      if (descriptor === undefined) {
        Reflect.deleteProperty(Object.prototype, key)
      } else {
        // oxlint-disable-next-line eslint(no-extend-native) -- Restore the native prototype descriptor changed by this regression test.
        Object.defineProperty(Object.prototype, key, descriptor)
      }
    }
  }
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

  it('normalizes transported refs without invoking inherited optional getters', () => {
    const table = {
      kind: 'table',
      id: 'sales',
      label: 'Sales',
    }
    const rows = {
      kind: 'rows',
      id: 'sales_west',
      label: 'Sales rows where Region eq West',
      where: {
        column: 'Region',
        op: 'eq',
        value: 'West',
      },
    }
    const column = {
      kind: 'column',
      id: 'sales_amount',
      label: 'Sales.Amount',
      table,
      name: 'Amount',
    }

    const calls = withThrowingObjectPrototypeGetters(['headers', 'name', 'rows', 'sheetName', 'table'], () => {
      expect(checkWorkbookRefData(table)).toMatchObject({ status: 'valid' })
      expect(toWorkbookRefData(table)).toEqual(table)
      expect(hydrateWorkbookRef(table)).toMatchObject(table)
      expect(collectWorkbookRefData({ column })).toEqual([toWorkbookRefData(column), toWorkbookRefData(table)])
      expect(hydrateWorkbookRefs({ rows })).toEqual({
        rows: expect.objectContaining(rows),
      })
    })

    expect(calls).toBe(0)
  })

  it('keeps row helper methods on own data when optional table is absent', () => {
    const rows = findRows({ sheetName: 'Sheet1', where: { column: 'Region', op: 'eq', value: 'West' } })

    const calls = withThrowingObjectPrototypeGetters(['table'], () => {
      expect(() => rows.column('Amount')).toThrowError('Rows column selection requires a table-backed row selector')
    })

    expect(calls).toBe(0)
  })
})

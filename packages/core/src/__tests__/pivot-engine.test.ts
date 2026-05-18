import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { materializePivotTable } from '../pivot-engine.js'

function stringValue(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 }
}

function numberValue(value: number): CellValue {
  return { tag: ValueTag.Number, value }
}

describe('materializePivotTable', () => {
  it('materializes column fields, page filters, and common Excel aggregate variants', () => {
    const result = materializePivotTable(
      {
        groupBy: ['Region'],
        columnFields: ['Quarter'],
        filters: [{ sourceColumn: 'Status', includedValues: ['Closed'] }],
        values: [
          { sourceColumn: 'Sales', summarizeBy: 'sum' },
          { sourceColumn: 'Sales', summarizeBy: 'average' },
          { sourceColumn: 'Units', summarizeBy: 'countNums' },
          { sourceColumn: 'Sales', summarizeBy: 'min' },
          { sourceColumn: 'Sales', summarizeBy: 'max' },
          { sourceColumn: 'Units', summarizeBy: 'product' },
        ],
      },
      [
        [stringValue('Region'), stringValue('Quarter'), stringValue('Status'), stringValue('Sales'), stringValue('Units')],
        [stringValue('East'), stringValue('Q1'), stringValue('Closed'), numberValue(10), numberValue(2)],
        [stringValue('East'), stringValue('Q1'), stringValue('Open'), numberValue(100), numberValue(9)],
        [stringValue('East'), stringValue('Q1'), stringValue('Closed'), numberValue(20), numberValue(3)],
        [stringValue('East'), stringValue('Q2'), stringValue('Closed'), numberValue(5), numberValue(4)],
        [stringValue('West'), stringValue('Q2'), stringValue('Closed'), numberValue(7), numberValue(5)],
      ],
    )

    expect(result).toEqual({
      kind: 'ok',
      rows: 3,
      cols: 13,
      values: [
        stringValue('Region'),
        stringValue('Q1 SUM of Sales'),
        stringValue('Q1 AVERAGE of Sales'),
        stringValue('Q1 COUNTNUMS of Units'),
        stringValue('Q1 MIN of Sales'),
        stringValue('Q1 MAX of Sales'),
        stringValue('Q1 PRODUCT of Units'),
        stringValue('Q2 SUM of Sales'),
        stringValue('Q2 AVERAGE of Sales'),
        stringValue('Q2 COUNTNUMS of Units'),
        stringValue('Q2 MIN of Sales'),
        stringValue('Q2 MAX of Sales'),
        stringValue('Q2 PRODUCT of Units'),
        stringValue('East'),
        numberValue(30),
        numberValue(15),
        numberValue(2),
        numberValue(10),
        numberValue(20),
        numberValue(6),
        numberValue(5),
        numberValue(5),
        numberValue(1),
        numberValue(5),
        numberValue(5),
        numberValue(4),
        stringValue('West'),
        numberValue(0),
        numberValue(0),
        numberValue(0),
        numberValue(0),
        numberValue(0),
        numberValue(0),
        numberValue(7),
        numberValue(7),
        numberValue(1),
        numberValue(7),
        numberValue(7),
        numberValue(5),
      ],
    })
  })

  it('groups rows by key columns and accumulates sum and count values', () => {
    const result = materializePivotTable(
      {
        groupBy: ['Region', 'Product'],
        values: [
          { sourceColumn: 'Sales', summarizeBy: 'sum' },
          { sourceColumn: 'Sales', summarizeBy: 'count', outputLabel: 'Rows' },
        ],
      },
      [
        [stringValue('Region'), stringValue('Product'), stringValue('Sales')],
        [stringValue('East'), stringValue('Widget'), numberValue(10)],
        [stringValue('East'), stringValue('Widget'), numberValue(5)],
        [stringValue('West'), stringValue('Gizmo'), numberValue(7)],
      ],
    )

    expect(result).toEqual({
      kind: 'ok',
      rows: 3,
      cols: 4,
      values: [
        stringValue('Region'),
        stringValue('Product'),
        stringValue('SUM of Sales'),
        stringValue('Rows'),
        stringValue('East'),
        stringValue('Widget'),
        numberValue(15),
        numberValue(2),
        stringValue('West'),
        stringValue('Gizmo'),
        numberValue(7),
        numberValue(1),
      ],
    })
  })

  it('returns a #VALUE pivot result when configured columns are missing', () => {
    const result = materializePivotTable(
      {
        groupBy: ['Region'],
        values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
      },
      [[stringValue('Category'), stringValue('Amount')]],
    )

    expect(result).toEqual({
      kind: 'error',
      code: ErrorCode.Value,
      rows: 1,
      cols: 1,
      values: [{ tag: ValueTag.Error, code: ErrorCode.Value }],
    })
  })

  it('skips fully empty rows, preserves boolean and error keys, and honors custom output labels', () => {
    const result = materializePivotTable(
      {
        groupBy: ['Group'],
        values: [
          { sourceColumn: 'Sales', summarizeBy: 'sum', outputLabel: 'Revenue ' },
          { sourceColumn: 'Tickets', summarizeBy: 'count' },
        ],
      },
      [
        [stringValue('Group'), stringValue('Sales'), stringValue('Tickets')],
        [{ tag: ValueTag.Empty }, { tag: ValueTag.Empty }, { tag: ValueTag.Empty }],
        [{ tag: ValueTag.Boolean, value: true }, { tag: ValueTag.Error, code: ErrorCode.Name }, { tag: ValueTag.Empty }],
        [{ tag: ValueTag.Error, code: ErrorCode.Ref }, numberValue(5), stringValue('x')],
      ],
    )

    expect(result).toEqual({
      kind: 'ok',
      rows: 3,
      cols: 3,
      values: [
        stringValue('Group'),
        stringValue('Revenue'),
        stringValue('COUNT of Tickets'),
        { tag: ValueTag.Boolean, value: true },
        numberValue(0),
        numberValue(0),
        { tag: ValueTag.Error, code: ErrorCode.Ref },
        numberValue(5),
        numberValue(1),
      ],
    })
  })

  it('rejects incomplete pivot configs and normalizes headers for count aggregations', () => {
    expect(
      materializePivotTable(
        {
          groupBy: ['Region'],
          values: [],
        },
        [[stringValue('Region'), stringValue('Sales')]],
      ),
    ).toEqual({
      kind: 'error',
      code: ErrorCode.Value,
      rows: 1,
      cols: 1,
      values: [{ tag: ValueTag.Error, code: ErrorCode.Value }],
    })

    expect(
      materializePivotTable(
        {
          groupBy: ['Region'],
          values: [{ sourceColumn: 'Sales', summarizeBy: 'count' }],
        },
        [],
      ),
    ).toEqual({
      kind: 'error',
      code: ErrorCode.Value,
      rows: 1,
      cols: 1,
      values: [{ tag: ValueTag.Error, code: ErrorCode.Value }],
    })

    expect(
      materializePivotTable(
        {
          groupBy: [' region '],
          values: [{ sourceColumn: 'sales', summarizeBy: 'count' }],
        },
        [
          [stringValue(' Region '), stringValue('Sales'), stringValue('sales')],
          [stringValue('East'), stringValue('ignored'), stringValue('ticket-1')],
          [{ tag: ValueTag.Empty }, { tag: ValueTag.Empty }, stringValue('ticket-2')],
        ],
      ),
    ).toEqual({
      kind: 'ok',
      rows: 2,
      cols: 2,
      values: [stringValue('Region'), stringValue('COUNT of Sales'), stringValue('East'), numberValue(1)],
    })
  })
})

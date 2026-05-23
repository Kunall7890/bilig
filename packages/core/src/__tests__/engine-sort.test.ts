import { describe, expect, it } from 'vitest'

import { ValueTag, type CellValue } from '@bilig/protocol'

import { SpreadsheetEngine } from '../engine.js'

describe('SpreadsheetEngine sort execution', () => {
  it('sorts range rows as bundles before recording sort metadata', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'sort-range' })
    await engine.ready()
    engine.createSheet('Ledger')
    engine.setRangeValues({ sheetName: 'Ledger', startAddress: 'A1', endAddress: 'D6' }, [
      ['Region', 'Amount', 'Memo', 'Double'],
      ['East', 10, 'e10', null],
      ['West', 5, 'w5a', null],
      ['West', 20, 'w20', null],
      ['East', 30, 'e30', null],
      ['West', 5, 'w5b', null],
    ])
    for (const row of [2, 3, 4, 5, 6]) {
      engine.setCellFormula('Ledger', `D${row}`, `B${row}*2`)
    }
    engine.setCellFormula('Ledger', 'E1', 'B2')
    engine.setCommentThread({
      threadId: 'thread-largest-invoice',
      sheetName: 'Ledger',
      address: 'C5',
      comments: [{ id: 'comment-largest-invoice', body: 'Largest East invoice needs review.' }],
    })
    engine.setNote({ sheetName: 'Ledger', address: 'A3', text: 'West row note' })
    engine.setRangeStyle({ sheetName: 'Ledger', startAddress: 'C5', endAddress: 'C5' }, { fill: { backgroundColor: '#fef3c7' } })
    engine.setRangeNumberFormat({ sheetName: 'Ledger', startAddress: 'B3', endAddress: 'B3' }, '0.00')
    engine.setDataValidation({
      range: { sheetName: 'Ledger', startAddress: 'B5', endAddress: 'B5' },
      rule: { kind: 'whole', operator: 'greaterThan', values: [0] },
    })

    expect(
      engine.sortRange(
        'Ledger',
        { sheetName: 'Ledger', startAddress: 'A1', endAddress: 'D6' },
        [
          { keyAddress: 'A1', direction: 'asc' },
          { keyAddress: 'B1', direction: 'desc' },
        ],
        { header: true },
      ),
    ).toBe(true)

    expect(displayRows(engine, 'A2', 'D6')).toEqual([
      ['East', 30, 'e30', 60],
      ['East', 10, 'e10', 20],
      ['West', 20, 'w20', 40],
      ['West', 5, 'w5a', 10],
      ['West', 5, 'w5b', 10],
    ])
    for (const row of [2, 3, 4, 5, 6]) {
      expect(engine.getCell('Ledger', `D${row}`).formula).toBe(`B${row}*2`)
    }
    expect(engine.getCellValue('Ledger', 'E1')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.getCommentThread('Ledger', 'C2')).toEqual({
      threadId: 'thread-largest-invoice',
      sheetName: 'Ledger',
      address: 'C2',
      comments: [{ id: 'comment-largest-invoice', body: 'Largest East invoice needs review.' }],
    })
    expect(engine.getCommentThread('Ledger', 'C5')).toBeUndefined()
    expect(engine.getNote('Ledger', 'A5')).toEqual({ sheetName: 'Ledger', address: 'A5', text: 'West row note' })
    expect(engine.getNote('Ledger', 'A3')).toBeUndefined()
    expect(engine.getCellStyle(engine.getCell('Ledger', 'C2').styleId)?.fill?.backgroundColor).toBe('#fef3c7')
    expect(engine.getCell('Ledger', 'C5').styleId).toBeUndefined()
    expect(engine.getCell('Ledger', 'B5').format).toBe('0.00')
    expect(engine.getCell('Ledger', 'B3').format).toBeUndefined()
    expect(engine.getDataValidation('Ledger', { sheetName: 'Ledger', startAddress: 'B5', endAddress: 'B5' })).toMatchObject({
      rule: { kind: 'whole', operator: 'greaterThan', values: [0] },
    })
    expect(engine.getDataValidation('Ledger', { sheetName: 'Ledger', startAddress: 'B2', endAddress: 'B2' })).toBeUndefined()
    expect(engine.getSorts('Ledger')).toEqual([
      {
        sheetName: 'Ledger',
        range: { sheetName: 'Ledger', startAddress: 'A2', endAddress: 'D6' },
        keys: [
          { keyAddress: 'A1', direction: 'asc' },
          { keyAddress: 'B1', direction: 'desc' },
        ],
      },
    ])
  })

  it('sorts only table bodies, preserves totals, and writes table sortState', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'sort-table' })
    await engine.ready()
    engine.createSheet('Ledger')
    engine.setRangeValues({ sheetName: 'Ledger', startAddress: 'A1', endAddress: 'D7' }, [
      ['Region', 'Amount', 'Memo', 'Double'],
      ['East', 10, 'e10', null],
      ['West', 40, 'w40', null],
      ['East', 30, 'e30', null],
      ['West', 20, 'w20', null],
      ['East', 50, 'e50', null],
      ['Total', 150, '', 300],
    ])
    for (const row of [2, 3, 4, 5, 6]) {
      engine.setCellFormula('Ledger', `D${row}`, `B${row}*2`)
    }
    engine.setCellFormula('Ledger', 'D7', 'SUM(D2:D6)')
    engine.setNote({ sheetName: 'Ledger', address: 'C6', text: 'largest invoice note' })
    engine.setTable({
      name: 'Sales',
      sheetName: 'Ledger',
      startAddress: 'A1',
      endAddress: 'D7',
      columnNames: ['Region', 'Amount', 'Memo', 'Double'],
      columns: [
        { name: 'Region', totalsRowLabel: 'Total' },
        { name: 'Amount', totalsRowFunction: 'sum' },
        { name: 'Memo' },
        { name: 'Double', totalsRowFunction: 'sum' },
      ],
      headerRow: true,
      totalsRow: true,
    })

    expect(engine.sortTable('Ledger', 'Sales', [{ keyAddress: 'B1', direction: 'desc' }])).toBe(true)

    expect(displayRows(engine, 'A1', 'D7')).toEqual([
      ['Region', 'Amount', 'Memo', 'Double'],
      ['East', 50, 'e50', 100],
      ['West', 40, 'w40', 80],
      ['East', 30, 'e30', 60],
      ['West', 20, 'w20', 40],
      ['East', 10, 'e10', 20],
      ['Total', 150, '', 300],
    ])
    expect(engine.getCell('Ledger', 'D2').formula).toBe('B2*2')
    expect(engine.getCell('Ledger', 'D6').formula).toBe('B6*2')
    expect(engine.getCell('Ledger', 'D7').formula).toBe('SUM(D2:D6)')
    expect(engine.getNote('Ledger', 'C2')).toEqual({ sheetName: 'Ledger', address: 'C2', text: 'largest invoice note' })
    expect(engine.getNote('Ledger', 'C6')).toBeUndefined()
    expect(engine.getSorts('Ledger')).toEqual([])
    expect(engine.getTable('Sales')).toMatchObject({
      name: 'Sales',
      sheetName: 'Ledger',
      startAddress: 'A1',
      endAddress: 'D7',
      sortState: '<sortState ref="A2:D6"><sortCondition descending="1" ref="B2:B6"/></sortState>',
      totalsRow: true,
    })
  })

  it('keeps blank sort keys at the bottom when sorting descending', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'sort-descending-blanks' })
    await engine.ready()
    engine.createSheet('Ledger')
    engine.setRangeValues({ sheetName: 'Ledger', startAddress: 'A1', endAddress: 'B4' }, [
      ['filled-low', 10],
      ['blank', null],
      ['filled-high', 30],
      ['filled-mid', 20],
    ])

    engine.sortRange('Ledger', { sheetName: 'Ledger', startAddress: 'A1', endAddress: 'B4' }, [{ keyAddress: 'B1', direction: 'desc' }])

    expect(displayRows(engine, 'A1', 'B4')).toEqual([
      ['filled-high', 30],
      ['filled-mid', 20],
      ['filled-low', 10],
      ['blank', null],
    ])
  })
})

function displayRows(engine: SpreadsheetEngine, startAddress: string, endAddress: string): unknown[][] {
  const values = engine.getRangeValues({ sheetName: 'Ledger', startAddress, endAddress })
  return values.map((row) => row.map(cellValue))
}

function cellValue(value: CellValue): unknown {
  switch (value.tag) {
    case ValueTag.Number:
    case ValueTag.String:
    case ValueTag.Boolean:
      return value.value
    case ValueTag.Empty:
    case ValueTag.Error:
      return null
    default:
      return null
  }
}

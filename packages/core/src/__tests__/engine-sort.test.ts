import { describe, expect, it } from 'vitest'

import { SpreadsheetEngine } from '../engine.js'
import { ValueTag } from '@bilig/protocol'

describe('SpreadsheetEngine sortRange', () => {
  it('mutates row order before recording sort metadata', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'sort-range' })
    await engine.ready()
    engine.createSheet('Ledger')
    engine.setRangeValues({ sheetName: 'Ledger', startAddress: 'A1', endAddress: 'C6' }, [
      ['Region', 'Amount', 'Memo'],
      ['East', 10, 'e10'],
      ['West', 5, 'w5a'],
      ['West', 20, 'w20'],
      ['East', 30, 'e30'],
      ['West', 5, 'w5b'],
    ])
    engine.setCellFormula('Ledger', 'E1', 'B2')

    expect(
      engine.sortRange(
        'Ledger',
        { sheetName: 'Ledger', startAddress: 'A1', endAddress: 'C6' },
        [
          { keyAddress: 'A1', direction: 'asc' },
          { keyAddress: 'B1', direction: 'desc' },
        ],
        { header: true },
      ),
    ).toBe(true)

    expect(displayRows(engine, 'A2', 'C6')).toEqual([
      ['East', 30, 'e30'],
      ['East', 10, 'e10'],
      ['West', 20, 'w20'],
      ['West', 5, 'w5a'],
      ['West', 5, 'w5b'],
    ])
    expect(engine.getCellValue('Ledger', 'E1')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.getSorts('Ledger')).toEqual([
      {
        sheetName: 'Ledger',
        range: { sheetName: 'Ledger', startAddress: 'A2', endAddress: 'C6' },
        keys: [
          { keyAddress: 'A1', direction: 'asc' },
          { keyAddress: 'B1', direction: 'desc' },
        ],
      },
    ])
  })

  it('sorts body ranges without a header row', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'sort-body-range' })
    await engine.ready()
    engine.createSheet('Ledger')
    engine.setRangeValues({ sheetName: 'Ledger', startAddress: 'A1', endAddress: 'B4' }, [
      ['Q3', 30],
      ['Q1', 10],
      ['Q4', 40],
      ['Q2', 20],
    ])

    engine.sortRange('Ledger', { sheetName: 'Ledger', startAddress: 'A1', endAddress: 'B4' }, [{ keyAddress: 'B1', direction: 'asc' }])

    expect(displayRows(engine, 'A1', 'B4')).toEqual([
      ['Q1', 10],
      ['Q2', 20],
      ['Q3', 30],
      ['Q4', 40],
    ])
    expect(engine.getSorts('Ledger')).toEqual([
      {
        sheetName: 'Ledger',
        range: { sheetName: 'Ledger', startAddress: 'A1', endAddress: 'B4' },
        keys: [{ keyAddress: 'B1', direction: 'asc' }],
      },
    ])
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
  return values.map((row) => row.map((cell) => ('value' in cell ? cell.value : null)))
}

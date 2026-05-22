import { describe, expect, it } from 'vitest'

import { SpreadsheetEngine } from '../engine.js'
import { ValueTag } from '@bilig/protocol'

describe('SpreadsheetEngine sortRange', () => {
  it('mutates row order before recording sort metadata', async () => {
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
    engine.setCellFormula('Ledger', 'D2', 'B2*2')
    engine.setCellFormula('Ledger', 'D3', 'B3*2')
    engine.setCellFormula('Ledger', 'D4', 'B4*2')
    engine.setCellFormula('Ledger', 'D5', 'B5*2')
    engine.setCellFormula('Ledger', 'D6', 'B6*2')
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
    expect(engine.getCell('Ledger', 'D2').formula).toBe('B2*2')
    expect(engine.getCell('Ledger', 'D3').formula).toBe('B3*2')
    expect(engine.getCell('Ledger', 'D4').formula).toBe('B4*2')
    expect(engine.getCell('Ledger', 'D5').formula).toBe('B5*2')
    expect(engine.getCell('Ledger', 'D6').formula).toBe('B6*2')
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

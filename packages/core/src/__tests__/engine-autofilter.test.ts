import { ValueTag } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'
import { SpreadsheetEngine } from '../engine.js'

describe('SpreadsheetEngine AutoFilter', () => {
  it('materializes value criteria into filtered row metadata and recalculates visible-row formulas', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'autofilter-local-materialization' })
    await engine.ready()
    engine.createSheet('Ledger')
    engine.setRangeValues({ sheetName: 'Ledger', startAddress: 'A1', endAddress: 'B6' }, [
      ['Region', 'Amount'],
      ['East', 10],
      ['West', 20],
      ['East', 30],
      ['West', 40],
      ['East', 50],
    ])
    engine.setCellFormula('Ledger', 'D1', 'SUBTOTAL(9,B2:B6)')
    engine.setCellFormula('Ledger', 'E1', 'SUBTOTAL(109,B2:B6)')
    engine.setCellFormula('Ledger', 'F1', 'AGGREGATE(9,4,B2:B6)')
    engine.setCellFormula('Ledger', 'G1', 'AGGREGATE(9,5,B2:B6)')

    engine.setFilter('Ledger', {
      sheetName: 'Ledger',
      startAddress: 'A1',
      endAddress: 'B6',
      criteria: [{ colId: 0, filters: { values: ['East'] } }],
    })

    expectRowFiltered(engine.getRowMetadata('Ledger'), 2)
    expectRowFiltered(engine.getRowMetadata('Ledger'), 4)
    expect(engine.getCellValue('Ledger', 'D1')).toEqual({ tag: ValueTag.Number, value: 90 })
    expect(engine.getCellValue('Ledger', 'E1')).toEqual({ tag: ValueTag.Number, value: 90 })
    expect(engine.getCellValue('Ledger', 'F1')).toEqual({ tag: ValueTag.Number, value: 150 })
    expect(engine.getCellValue('Ledger', 'G1')).toEqual({ tag: ValueTag.Number, value: 90 })

    engine.clearFilter('Ledger', { sheetName: 'Ledger', startAddress: 'A1', endAddress: 'B6' })

    expect(engine.getRowMetadata('Ledger').some((record) => record.filtered === true)).toBe(false)
    expect(engine.getCellValue('Ledger', 'D1')).toEqual({ tag: ValueTag.Number, value: 150 })
    expect(engine.getCellValue('Ledger', 'E1')).toEqual({ tag: ValueTag.Number, value: 150 })
    expect(engine.getCellValue('Ledger', 'F1')).toEqual({ tag: ValueTag.Number, value: 150 })
    expect(engine.getCellValue('Ledger', 'G1')).toEqual({ tag: ValueTag.Number, value: 150 })
  })

  it('materializes custom numeric AutoFilter criteria', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'autofilter-local-custom-materialization' })
    await engine.ready()
    engine.createSheet('Ledger')
    engine.setRangeValues({ sheetName: 'Ledger', startAddress: 'A1', endAddress: 'B6' }, [
      ['Region', 'Amount'],
      ['East', 10],
      ['West', 20],
      ['East', 30],
      ['West', 40],
      ['East', 50],
    ])
    engine.setCellFormula('Ledger', 'D1', 'SUBTOTAL(9,B2:B6)')

    engine.setFilter('Ledger', {
      sheetName: 'Ledger',
      startAddress: 'A1',
      endAddress: 'B6',
      criteria: [{ colId: 1, customFilters: { filters: [{ operator: 'greaterThan', value: '25' }] } }],
    })

    expectRowFiltered(engine.getRowMetadata('Ledger'), 1)
    expectRowFiltered(engine.getRowMetadata('Ledger'), 2)
    expect(engine.getCellValue('Ledger', 'D1')).toEqual({ tag: ValueTag.Number, value: 120 })
  })
})

function expectRowFiltered(
  records: ReadonlyArray<{ readonly start: number; readonly count: number; readonly filtered?: boolean | null }>,
  row: number,
): void {
  expect(records.some((record) => record.filtered === true && row >= record.start && row < record.start + record.count)).toBe(true)
}

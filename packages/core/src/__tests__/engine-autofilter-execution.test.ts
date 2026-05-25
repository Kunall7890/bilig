import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { SpreadsheetEngine } from '../index.js'

describe('engine AutoFilter execution', () => {
  it('applies table AutoFilter row visibility without conflating manual hidden rows', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'table-autofilter-execution' })
    await engine.ready()
    engine.createSheet('Ledger')
    engine.setRangeValues({ sheetName: 'Ledger', startAddress: 'A1', endAddress: 'B6' }, [
      ['Region', 'Amount'],
      ['East', 10],
      ['West', 40],
      ['East', 30],
      ['West', 20],
      ['East', 50],
    ])
    engine.setTable({
      name: 'Sales',
      sheetName: 'Ledger',
      startAddress: 'A1',
      endAddress: 'B6',
      columnNames: ['Region', 'Amount'],
      headerRow: true,
      totalsRow: false,
    })
    engine.setCellFormula('Ledger', 'D2', 'SUBTOTAL(9,B2:B6)')
    engine.setCellFormula('Ledger', 'D3', 'SUBTOTAL(109,B2:B6)')
    engine.updateRowMetadata('Ledger', 3, 1, null, true)
    engine.recalculateNow()

    expect(engine.getCellValue('Ledger', 'D2')).toEqual({ tag: ValueTag.Number, value: 150 })
    expect(engine.getCellValue('Ledger', 'D3')).toEqual({ tag: ValueTag.Number, value: 120 })

    expect(engine.applyTableAutoFilter('Ledger', 'Sales', [{ colId: 0, filters: { values: ['East'] } }])).toBe(true)

    expect(engine.getCellValue('Ledger', 'D2')).toEqual({ tag: ValueTag.Number, value: 90 })
    expect(engine.getCellValue('Ledger', 'D3')).toEqual({ tag: ValueTag.Number, value: 60 })
    expect(engine.getTable('Sales')?.autoFilter).toEqual({
      sheetName: 'Ledger',
      startAddress: 'A1',
      endAddress: 'B6',
      criteria: [{ colId: 0, filters: { values: ['East'] } }],
    })
    expect(rowFlags(engine.getRowMetadata('Ledger'))).toEqual([
      { start: 2, hidden: null, filterHidden: true },
      { start: 3, hidden: true, filterHidden: null },
      { start: 4, hidden: null, filterHidden: true },
    ])

    expect(engine.applyTableAutoFilter('Ledger', 'Sales', [])).toBe(true)

    expect(engine.getCellValue('Ledger', 'D2')).toEqual({ tag: ValueTag.Number, value: 150 })
    expect(engine.getCellValue('Ledger', 'D3')).toEqual({ tag: ValueTag.Number, value: 120 })
    expect(engine.getTable('Sales')?.autoFilter).toEqual({
      sheetName: 'Ledger',
      startAddress: 'A1',
      endAddress: 'B6',
    })
    expect(rowFlags(engine.getRowMetadata('Ledger'))).toEqual([{ start: 3, hidden: true, filterHidden: null }])
  })
})

function rowFlags(
  metadata: readonly { readonly start: number; readonly hidden: boolean | null; readonly filterHidden?: boolean | null }[],
): Array<{ start: number; hidden: boolean | null; filterHidden: boolean | null }> {
  return metadata.map((record) => ({
    start: record.start,
    hidden: record.hidden,
    filterHidden: record.filterHidden ?? null,
  }))
}

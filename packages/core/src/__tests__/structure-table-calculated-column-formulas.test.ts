import { describe, expect, it } from 'vitest'

import { ValueTag } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'

describe('structural table calculated-column formulas', () => {
  it('materializes calculated-column formulas when inserting rows inside a table', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'structure-table-calculated-column-insert' })
    await engine.ready()
    engine.createSheet('Sales')
    engine.setRangeValues({ sheetName: 'Sales', startAddress: 'A1', endAddress: 'C4' }, [
      ['Item', 'Qty', 'Total'],
      ['A', 2, null],
      ['B', 3, null],
      ['C', 4, null],
    ])
    engine.setCellFormula('Sales', 'C2', 'B2*10')
    engine.setCellFormula('Sales', 'C3', 'B3*10')
    engine.setCellFormula('Sales', 'C4', 'B4*10')
    engine.setTable({
      name: 'SalesTable',
      sheetName: 'Sales',
      startAddress: 'A1',
      endAddress: 'C4',
      columnNames: ['Item', 'Qty', 'Total'],
      columns: [{ name: 'Item' }, { name: 'Qty' }, { name: 'Total', calculatedColumnFormula: 'SalesTable[[#This Row],[Qty]]*10' }],
      headerRow: true,
      totalsRow: false,
    })

    engine.insertRows('Sales', 2, 1)
    engine.setCellValue('Sales', 'A3', 'New')
    engine.setCellValue('Sales', 'B3', 5)

    expect(engine.getTable('SalesTable')).toMatchObject({
      startAddress: 'A1',
      endAddress: 'C5',
      columnNames: ['Item', 'Qty', 'Total'],
      columns: [{ name: 'Item' }, { name: 'Qty' }, { name: 'Total', calculatedColumnFormula: 'SalesTable[[#This Row],[Qty]]*10' }],
    })
    expect(engine.getCell('Sales', 'C3').formula).toBe("'Sales'!B3*10")
    expect(engine.getCellValue('Sales', 'C3')).toEqual({ tag: ValueTag.Number, value: 50 })
    expect(engine.getCell('Sales', 'C4').formula).toBe('B4*10')
    expect(engine.getCellValue('Sales', 'C4')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.getCell('Sales', 'C5').formula).toBe('B5*10')
    expect(engine.getCellValue('Sales', 'C5')).toEqual({ tag: ValueTag.Number, value: 40 })
  })
})

import { ErrorCode, ValueTag } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'
import { SpreadsheetEngine } from '../engine.js'

describe('spaced table structured references', () => {
  it('rewrites formulas and defined names when renaming a spaced table header', async () => {
    const engine = await buildSpacedTableReferenceEngine('structure-spaced-table-header-rename')
    const initialSnapshot = engine.exportSnapshot()

    engine.setCellValue('Data', 'A1', 'Q1 Revenue')

    expect(engine.getTable('Sales')).toMatchObject({
      startAddress: 'A1',
      endAddress: 'B3',
      columnNames: ['Q1 Revenue', 'Units Sold'],
    })
    expect(engine.getCellValue('Data', 'A1')).toEqual({ tag: ValueTag.String, value: 'Q1 Revenue', stringId: expect.any(Number) })
    expect(engine.getCell('Data', 'D1').formula).toBe('SUM(Sales[Q1 Revenue])')
    expect(engine.getCellValue('Data', 'D1')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.getCell('Data', 'E1').formula).toBe('SUM(Sales[Units Sold])')
    expect(engine.getCellValue('Data', 'E1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getDefinedName('SalesQ1')).toEqual({
      name: 'SalesQ1',
      value: { kind: 'structured-ref', tableName: 'Sales', columnName: 'Q1 Revenue' },
    })
    expect(engine.getDefinedName('SalesUnitsFormula')).toEqual({
      name: 'SalesUnitsFormula',
      value: { kind: 'formula', formula: '=Sales[Units Sold]' },
    })
    expect(engine.getCellValue('Data', 'F1')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.getCellValue('Data', 'G1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.undo()).toBe(true)
    expect(engine.exportSnapshot()).toEqual(initialSnapshot)
  })

  it('rebinds dependents of surviving spaced-header defined names when table ranges move', async () => {
    const engine = await buildSpacedTableReferenceEngine('structure-spaced-table-column-delete')
    const initialSnapshot = engine.exportSnapshot()

    engine.deleteColumns('Data', 0, 1)

    expect(engine.getTable('Sales')).toMatchObject({
      startAddress: 'A1',
      endAddress: 'A3',
      columnNames: ['Units Sold'],
    })
    expect(engine.getDefinedName('SalesQ1')).toEqual({
      name: 'SalesQ1',
      value: { kind: 'formula', formula: '=#REF!' },
    })
    expect(engine.getDefinedName('SalesUnitsFormula')).toEqual({
      name: 'SalesUnitsFormula',
      value: { kind: 'formula', formula: '=Sales[Units Sold]' },
    })
    expect(engine.getCell('Data', 'C1').formula).toBe('SUM(#REF!)')
    expect(engine.getCellValue('Data', 'C1')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
    expect(engine.getCell('Data', 'D1').formula).toBe('SUM(Sales[Units Sold])')
    expect(engine.getCellValue('Data', 'D1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCell('Data', 'E1').formula).toBe('SUM(SalesQ1)')
    expect(engine.getCellValue('Data', 'E1')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
    expect(engine.getCell('Data', 'F1').formula).toBe('SUM(SalesUnitsFormula)')
    expect(engine.getCellValue('Data', 'F1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.undo()).toBe(true)
    expect(engine.exportSnapshot()).toEqual(initialSnapshot)
  })
})

async function buildSpacedTableReferenceEngine(workbookName: string): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({ workbookName })
  await engine.ready()
  engine.createSheet('Data')
  engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' }, [
    ['Q1 Sales', 'Units Sold'],
    [10, 2],
    [20, 3],
  ])
  engine.setTable({
    name: 'Sales',
    sheetName: 'Data',
    startAddress: 'A1',
    endAddress: 'B3',
    columnNames: ['Q1 Sales', 'Units Sold'],
    headerRow: true,
    totalsRow: false,
  })
  engine.setDefinedName('SalesQ1', { kind: 'structured-ref', tableName: 'Sales', columnName: 'Q1 Sales' })
  engine.setDefinedName('SalesUnitsFormula', { kind: 'formula', formula: '=Sales[Units Sold]' })
  engine.setCellFormula('Data', 'D1', 'SUM(Sales[Q1 Sales])')
  engine.setCellFormula('Data', 'E1', 'SUM(Sales[Units Sold])')
  engine.setCellFormula('Data', 'F1', 'SUM(SalesQ1)')
  engine.setCellFormula('Data', 'G1', 'SUM(SalesUnitsFormula)')
  return engine
}

import { ValueTag } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'
import { SpreadsheetEngine } from '../engine.js'

describe('table header canonicalization', () => {
  it('generates Excel-compatible names for blank table headers', async () => {
    const engine = await buildTableHeaderCanonicalizationEngine('structure-table-header-blank')
    const initialSnapshot = engine.exportSnapshot()

    engine.setCellValue('Data', 'B1', '')

    expect(engine.getTable('Sales')).toMatchObject({
      startAddress: 'A1',
      endAddress: 'C3',
      columnNames: ['Region', 'Column1', 'Margin'],
    })
    expect(engine.getCellValue('Data', 'B1')).toEqual({ tag: ValueTag.String, value: 'Column1', stringId: expect.any(Number) })
    expect(engine.getCell('Data', 'E1').formula).toBe('SUM(Sales[Column1])')
    expect(engine.getCellValue('Data', 'E1')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.getCell('Data', 'F1').formula).toBe('SUM(Sales[Margin])')
    expect(engine.getCellValue('Data', 'F1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getDefinedName('SalesAmount')).toEqual({
      name: 'SalesAmount',
      value: { kind: 'structured-ref', tableName: 'Sales', columnName: 'Column1' },
    })
    expect(engine.getDefinedName('SalesMarginFormula')).toEqual({
      name: 'SalesMarginFormula',
      value: { kind: 'formula', formula: '=Sales[Margin]' },
    })
    expect(engine.getCellValue('Data', 'G1')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.getCellValue('Data', 'H1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.undo()).toBe(true)
    expect(engine.exportSnapshot()).toEqual(initialSnapshot)
  })

  it('deduplicates table headers before rewriting structured references', async () => {
    const engine = await buildTableHeaderCanonicalizationEngine('structure-table-header-duplicate')
    const initialSnapshot = engine.exportSnapshot()

    engine.setCellValue('Data', 'C1', 'Amount')

    expect(engine.getTable('Sales')).toMatchObject({
      startAddress: 'A1',
      endAddress: 'C3',
      columnNames: ['Region', 'Amount', 'Amount2'],
    })
    expect(engine.getCellValue('Data', 'C1')).toEqual({ tag: ValueTag.String, value: 'Amount2', stringId: expect.any(Number) })
    expect(engine.getCell('Data', 'E1').formula).toBe('SUM(Sales[Amount])')
    expect(engine.getCellValue('Data', 'E1')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.getCell('Data', 'F1').formula).toBe('SUM(Sales[Amount2])')
    expect(engine.getCellValue('Data', 'F1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getDefinedName('SalesAmount')).toEqual({
      name: 'SalesAmount',
      value: { kind: 'structured-ref', tableName: 'Sales', columnName: 'Amount' },
    })
    expect(engine.getDefinedName('SalesMarginFormula')).toEqual({
      name: 'SalesMarginFormula',
      value: { kind: 'formula', formula: '=Sales[Amount2]' },
    })
    expect(engine.getCellValue('Data', 'G1')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.getCellValue('Data', 'H1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.undo()).toBe(true)
    expect(engine.exportSnapshot()).toEqual(initialSnapshot)
  })

  it('treats clearing a table header as an Excel blank-header rename', async () => {
    const engine = await buildTableHeaderCanonicalizationEngine('structure-table-header-clear')
    const initialSnapshot = engine.exportSnapshot()

    engine.clearCell('Data', 'B1')

    expect(engine.getTable('Sales')).toMatchObject({
      startAddress: 'A1',
      endAddress: 'C3',
      columnNames: ['Region', 'Column1', 'Margin'],
    })
    expect(engine.getCellValue('Data', 'B1')).toEqual({ tag: ValueTag.String, value: 'Column1', stringId: expect.any(Number) })
    expect(engine.getCell('Data', 'E1').formula).toBe('SUM(Sales[Column1])')
    expect(engine.getCellValue('Data', 'E1')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.getDefinedName('SalesAmount')).toEqual({
      name: 'SalesAmount',
      value: { kind: 'structured-ref', tableName: 'Sales', columnName: 'Column1' },
    })
    expect(engine.getCellValue('Data', 'G1')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.undo()).toBe(true)
    expect(engine.exportSnapshot()).toEqual(initialSnapshot)
  })
})

async function buildTableHeaderCanonicalizationEngine(workbookName: string): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({ workbookName })
  await engine.ready()
  engine.createSheet('Data')
  engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'C3' }, [
    ['Region', 'Amount', 'Margin'],
    ['East', 10, 2],
    ['West', 20, 3],
  ])
  engine.setTable({
    name: 'Sales',
    sheetName: 'Data',
    startAddress: 'A1',
    endAddress: 'C3',
    columnNames: ['Region', 'Amount', 'Margin'],
    headerRow: true,
    totalsRow: false,
  })
  engine.setDefinedName('SalesAmount', { kind: 'structured-ref', tableName: 'Sales', columnName: 'Amount' })
  engine.setDefinedName('SalesMarginFormula', { kind: 'formula', formula: '=Sales[Margin]' })
  engine.setCellFormula('Data', 'E1', 'SUM(Sales[Amount])')
  engine.setCellFormula('Data', 'F1', 'SUM(Sales[Margin])')
  engine.setCellFormula('Data', 'G1', 'SUM(SalesAmount)')
  engine.setCellFormula('Data', 'H1', 'SUM(SalesMarginFormula)')
  return engine
}

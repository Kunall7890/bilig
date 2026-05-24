import { ErrorCode, ValueTag } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { SpreadsheetEngine } from '../index.js'

describe('structural defined-name references', () => {
  it('preserves deleted range-ref names as Excel-compatible #REF! formulas', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'structural-defined-name-range-ref' })
    await engine.ready()
    engine.createSheet('Data')
    engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'A3' }, [[10], [20], [30]])
    engine.setDefinedName('SalesRange', { kind: 'range-ref', sheetName: 'Data', startAddress: 'A1', endAddress: 'A3' })
    engine.setCellFormula('Data', 'D5', 'SUM(SalesRange)')

    engine.deleteRows('Data', 0, 3)

    expect(engine.getDefinedName('SalesRange')).toEqual({
      name: 'SalesRange',
      value: { kind: 'formula', formula: '=Data!#REF!' },
    })
    expect(engine.getCell('Data', 'D2').formula).toBe('SUM(SalesRange)')
    expect(engine.getCellValue('Data', 'D2')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
  })

  it('preserves deleted cell-ref names as Excel-compatible #REF! formulas', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'structural-defined-name-cell-ref' })
    await engine.ready()
    engine.createSheet('Data')
    engine.setCellValue('Data', 'A1', 5)
    engine.setDefinedName('RateCell', { kind: 'cell-ref', sheetName: 'Data', address: 'A1' })
    engine.setCellFormula('Data', 'D3', 'RateCell*2')

    engine.deleteRows('Data', 0, 1)

    expect(engine.getDefinedName('RateCell')).toEqual({
      name: 'RateCell',
      value: { kind: 'formula', formula: '=Data!#REF!' },
    })
    expect(engine.getCell('Data', 'D2').formula).toBe('RateCell*2')
    expect(engine.getCellValue('Data', 'D2')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
  })

  it('quotes sheet names when structural deletes invalidate named references', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'structural-defined-name-quoted-sheet-ref' })
    await engine.ready()
    engine.createSheet('Q1 Data')
    engine.setCellValue('Q1 Data', 'A1', 5)
    engine.setDefinedName('RateCell', { kind: 'cell-ref', sheetName: 'Q1 Data', address: 'A1' })
    engine.setCellFormula('Q1 Data', 'D3', 'RateCell*2')

    engine.deleteRows('Q1 Data', 0, 1)

    expect(engine.getDefinedName('RateCell')).toEqual({
      name: 'RateCell',
      value: { kind: 'formula', formula: "='Q1 Data'!#REF!" },
    })
    expect(engine.getCellValue('Q1 Data', 'D2')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
  })

  it('preserves workbook-level names as #REF! and removes deleted-sheet scoped names on sheet delete', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'sheet-delete-defined-name-refs' })
    await engine.ready()
    engine.createSheet('Data')
    engine.createSheet('Report')
    engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'A3' }, [[5], [10], [20]])
    engine.setDefinedName('RateCell', { kind: 'cell-ref', sheetName: 'Data', address: 'A1' })
    engine.setDefinedName('SalesRange', { kind: 'range-ref', sheetName: 'Data', startAddress: 'A1', endAddress: 'A3' })
    engine.setDefinedName('FormulaRate', { kind: 'formula', formula: '=Data!$A$1' })
    engine.setDefinedName('FormulaSum', { kind: 'formula', formula: '=SUM(Data!$A$1:$A$3)' })
    engine.workbook.setDefinedName(
      '_xlnm.Print_Area',
      { kind: 'range-ref', sheetName: 'Data', startAddress: 'A1', endAddress: 'A3' },
      'Data',
    )
    engine.workbook.setDefinedName(
      '_xlnm.Print_Area',
      { kind: 'range-ref', sheetName: 'Report', startAddress: 'A1', endAddress: 'B4' },
      'Report',
    )
    engine.setCellFormula('Report', 'A1', 'RateCell*2')
    engine.setCellFormula('Report', 'A2', 'SUM(SalesRange)')
    engine.setCellFormula('Report', 'A3', 'FormulaRate+FormulaSum')
    expect(engine.getCellValue('Report', 'A1')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(engine.getCellValue('Report', 'A2')).toEqual({ tag: ValueTag.Number, value: 35 })
    expect(engine.getCellValue('Report', 'A3')).toEqual({ tag: ValueTag.Number, value: 40 })

    engine.deleteSheet('Data')

    expect(engine.getDefinedNames()).toEqual([
      {
        name: '_xlnm.Print_Area',
        scopeSheetName: 'Report',
        value: { kind: 'range-ref', sheetName: 'Report', startAddress: 'A1', endAddress: 'B4' },
      },
      { name: 'FormulaRate', value: { kind: 'formula', formula: '=#REF!' } },
      { name: 'FormulaSum', value: { kind: 'formula', formula: '=SUM(#REF!)' } },
      { name: 'RateCell', value: { kind: 'formula', formula: '=#REF!' } },
      { name: 'SalesRange', value: { kind: 'formula', formula: '=#REF!' } },
    ])
    expect(engine.getCellValue('Report', 'A1')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
    expect(engine.getCellValue('Report', 'A2')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
    expect(engine.getCellValue('Report', 'A3')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
  })
})

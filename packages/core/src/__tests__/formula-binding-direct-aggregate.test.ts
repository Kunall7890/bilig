import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'

function readRuntimeFormula(engine: SpreadsheetEngine, cellIndex: number): unknown {
  const formulas = Reflect.get(engine, 'formulas')
  if (typeof formulas !== 'object' || formulas === null || typeof Reflect.get(formulas, 'get') !== 'function') {
    throw new TypeError('Expected internal formulas store')
  }
  return Reflect.get(formulas, 'get').call(formulas, cellIndex)
}

function readRuntimeFormulaProperty(formula: unknown, property: string): unknown {
  if (typeof formula !== 'object' || formula === null) {
    throw new TypeError('Expected runtime formula')
  }
  return Reflect.get(formula, property)
}

describe('direct aggregate formula binding', () => {
  it('binds simple aggregate formulas without criteria metadata or managed plans', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'binding-direct-aggregate-without-criteria' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'B1', 2)
    engine.setCellFormula('Sheet1', 'C1', 'SUM(A1:B1)')

    const formulaCellIndex = engine.workbook.getCellIndex('Sheet1', 'C1')
    expect(formulaCellIndex).toBeDefined()
    const runtimeFormula = readRuntimeFormula(engine, formulaCellIndex!)

    expect(readRuntimeFormulaProperty(runtimeFormula, 'directAggregate')).toMatchObject({
      aggregateKind: 'sum',
      rowStart: 0,
      rowEnd: 0,
      col: 0,
      colEnd: 1,
      length: 2,
    })
    expect(readRuntimeFormulaProperty(runtimeFormula, 'directCriteria')).toBeUndefined()
    expect(readRuntimeFormulaProperty(runtimeFormula, 'runtimeProgram')).toHaveLength(0)
    expect(readRuntimeFormulaProperty(runtimeFormula, 'rangeDependencies')).toHaveLength(0)
    expect(readRuntimeFormulaProperty(runtimeFormula, 'graphRangeDependencies')).toHaveLength(0)
    expect(readRuntimeFormulaProperty(runtimeFormula, 'planId')).toBe(0)
  })

  it('rebases translated sheet-qualified aggregate templates before direct evaluation', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'binding-direct-aggregate-translated-sheet-range' })
    await engine.ready()

    engine.importSnapshot({
      version: 1,
      workbook: { name: 'binding-direct-aggregate-translated-sheet-range' },
      sheets: [
        {
          id: 1,
          name: 'Average Prices',
          order: 0,
          cells: [
            { address: 'C16', value: 48_000 },
            { address: 'D16', value: 74_000 },
            { address: 'C17', value: 200_000 },
            { address: 'D17', value: 250_000 },
          ],
        },
        {
          id: 2,
          name: 'Revenue',
          order: 1,
          cells: [
            { address: 'R7', formula: "AVERAGE('Average Prices'!C16:D16)", value: 61_000 },
            { address: 'R8', formula: "AVERAGE('Average Prices'!C17:D17)", value: 225_000 },
          ],
        },
      ],
    })
    engine.recalculateNow()

    expect(engine.getCellValue('Revenue', 'R7')).toEqual({ tag: ValueTag.Number, value: 61_000 })
    expect(engine.getCellValue('Revenue', 'R8')).toEqual({ tag: ValueTag.Number, value: 225_000 })
    const formulaCellIndex = engine.workbook.getCellIndex('Revenue', 'R8')
    expect(formulaCellIndex).toBeDefined()
    const runtimeFormula = readRuntimeFormula(engine, formulaCellIndex!)

    expect(readRuntimeFormulaProperty(runtimeFormula, 'directAggregate')).toMatchObject({
      sheetName: 'Average Prices',
      rowStart: 16,
      rowEnd: 16,
      col: 2,
      colEnd: 3,
      length: 2,
    })
  })
})

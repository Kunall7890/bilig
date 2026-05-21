import { describe, expect, it, vi } from 'vitest'
import { ValueTag } from '@bilig/protocol'

import { SpreadsheetEngine, type EngineCellMutationRef } from '../engine.js'

interface FormulaBindingServiceForTailAppendTest {
  readonly collectFormulaCellsReferencingSheetNow: (sheetName: string) => readonly number[]
  readonly forEachFormulaCellOwnedBySheetNow: (sheetName: string, fn: (cellIndex: number) => void) => void
}

function isFormulaBindingServiceForTailAppendTest(value: unknown): value is FormulaBindingServiceForTailAppendTest {
  return (
    typeof value === 'object' &&
    value !== null &&
    typeof Reflect.get(value, 'collectFormulaCellsReferencingSheetNow') === 'function' &&
    typeof Reflect.get(value, 'forEachFormulaCellOwnedBySheetNow') === 'function'
  )
}

function getFormulaBindingService(engine: SpreadsheetEngine): FormulaBindingServiceForTailAppendTest {
  const runtime = Reflect.get(engine, 'runtime')
  if (typeof runtime !== 'object' || runtime === null) {
    throw new TypeError('Expected engine runtime')
  }
  const binding = Reflect.get(runtime, 'binding')
  if (!isFormulaBindingServiceForTailAppendTest(binding)) {
    throw new TypeError('Expected engine formula binding service')
  }
  return binding
}

describe('structural tail append direct aggregates', () => {
  it('does not defer unchanged existing row aggregate formulas after tail row inserts', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'tail-append-direct-aggregates' })
    await engine.ready()
    engine.createSheet('Sheet1')

    const existingRows = 12
    const appendRows = 40
    const inputCols = 4
    for (let row = 1; row <= existingRows; row += 1) {
      for (let col = 0; col < inputCols; col += 1) {
        engine.setCellValue('Sheet1', `${String.fromCharCode(65 + col)}${row}`, row * (col + 1))
      }
      engine.setCellFormula('Sheet1', `E${row}`, `SUM(A${row}:D${row})`)
    }

    engine.resetPerformanceCounters()
    const positionSpy = vi.spyOn(engine.workbook, 'getCellPosition')
    try {
      engine.insertRows('Sheet1', existingRows, appendRows)
      expect(positionSpy).not.toHaveBeenCalled()
    } finally {
      positionSpy.mockRestore()
    }

    expect([...engine.state.formulas.values()].filter((formula) => formula.structuralSourceTransform !== undefined)).toHaveLength(0)
    expect(engine.getPerformanceCounters()).toMatchObject({
      structuralFormulaImpactCandidates: 0,
      structuralFormulaRebindInputs: 0,
      structuralTransactions: 1,
    })

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const valueRefs: EngineCellMutationRef[] = []
    const formulaRefs: EngineCellMutationRef[] = []
    for (let row = 0; row < appendRows; row += 1) {
      const rowIndex = existingRows + row
      const rowNumber = rowIndex + 1
      for (let col = 0; col < inputCols; col += 1) {
        valueRefs.push({
          sheetId,
          mutation: { kind: 'setCellValue', row: rowIndex, col, value: rowNumber * (col + 1) },
        })
      }
      formulaRefs.push({
        sheetId,
        mutation: { kind: 'setCellFormula', row: rowIndex, col: inputCols, formula: `SUM(A${rowNumber}:D${rowNumber})` },
      })
    }

    engine.applyCellMutationsAt(valueRefs, valueRefs.length)
    engine.resetPerformanceCounters()
    engine.applyCellMutationsAt(formulaRefs, formulaRefs.length)

    expect(engine.getCellValue('Sheet1', 'E13')).toEqual({ tag: ValueTag.Number, value: 130 })
    expect(engine.getCellValue('Sheet1', 'E20')).toEqual({ tag: ValueTag.Number, value: 200 })
    engine.setCellValue('Sheet1', 'A13', 1000)
    expect(engine.getCellValue('Sheet1', 'E13')).toEqual({ tag: ValueTag.Number, value: 1117 })
    expect(engine.getPerformanceCounters()).toMatchObject({
      directAggregateScanCells: 0,
      directAggregateScanEvaluations: 0,
      formulasBound: 0,
      structuralFormulaRebindInputs: 0,
    })
  })

  it('uses known owner positions when binding appended direct aggregate formulas', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'tail-append-owner-position-fast-path' })
    await engine.ready()
    engine.createSheet('Sheet1')

    const existingRows = 12
    const appendRows = 40
    const inputCols = 4
    for (let row = 1; row <= existingRows; row += 1) {
      for (let col = 0; col < inputCols; col += 1) {
        engine.setCellValue('Sheet1', `${String.fromCharCode(65 + col)}${row}`, row * (col + 1))
      }
      engine.setCellFormula('Sheet1', `E${row}`, `SUM(A${row}:D${row})`)
    }

    engine.insertRows('Sheet1', existingRows, appendRows)

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const valueRefs: EngineCellMutationRef[] = []
    const formulaRefs: EngineCellMutationRef[] = []
    for (let row = 0; row < appendRows; row += 1) {
      const rowIndex = existingRows + row
      const rowNumber = rowIndex + 1
      for (let col = 0; col < inputCols; col += 1) {
        valueRefs.push({
          sheetId,
          mutation: { kind: 'setCellValue', row: rowIndex, col, value: rowNumber * (col + 1) },
        })
      }
      formulaRefs.push({
        sheetId,
        mutation: { kind: 'setCellFormula', row: rowIndex, col: inputCols, formula: `SUM(A${rowNumber}:D${rowNumber})` },
      })
    }

    engine.applyCellMutationsAt(valueRefs, valueRefs.length)

    const positionSpy = vi.spyOn(engine.workbook, 'getCellPosition')
    engine.applyCellMutationsAt(formulaRefs, formulaRefs.length)

    expect(positionSpy).not.toHaveBeenCalled()
    expect(engine.getCellValue('Sheet1', 'E13')).toEqual({ tag: ValueTag.Number, value: 130 })
    expect(engine.getCellValue('Sheet1', 'E52')).toEqual({ tag: ValueTag.Number, value: 520 })
  })

  it('falls back to structural formula scanning when existing formulas reference appended rows', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'tail-append-reference-overlap' })
    await engine.ready()
    engine.createSheet('Sheet1')

    for (let row = 1; row <= 20; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A20)')

    const binding = getFormulaBindingService(engine)
    const ownedFormulaScan = vi.spyOn(binding, 'forEachFormulaCellOwnedBySheetNow')
    try {
      engine.insertRows('Sheet1', 12, 2)
      expect(ownedFormulaScan).toHaveBeenCalled()
    } finally {
      ownedFormulaScan.mockRestore()
    }

    expect(engine.getCell('Sheet1', 'B1').formula).toBe('SUM(A1:A22)')
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 210 })
  })
})

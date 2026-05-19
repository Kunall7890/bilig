import { FormulaMode, ValueTag } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'
import { SpreadsheetEngine } from '../index.js'

describe('SpreadsheetEngine wasm initialization', () => {
  it('keeps the wasm kernel lazy until ready is requested', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })

    expect(engine.wasm.ready).toBe(false)

    await engine.ready()

    expect(engine.wasm.ready).toBe(true)

    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellFormula('Sheet1', 'B1', 'SIN(A1)+1')

    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: Math.sin(10) + 1 })
    expect(engine.getLastMetrics().wasmFormulaCount).toBeGreaterThan(0)

    engine.setCellValue('Sheet1', 'A1', 12)

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: Math.sin(12) + 1 })
  })

  it('keeps exact vector MATCH bindings on the JS lookup path', () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })

    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'Key')
    engine.setCellValue('Sheet1', 'A2', 'KEY-00001')
    engine.setCellValue('Sheet1', 'A3', 'KEY-00002')
    engine.setCellValue('Sheet1', 'D1', 'KEY-00002')
    engine.setCellFormula('Sheet1', 'E1', 'MATCH(D1,A2:A3,0)')

    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.JsOnly)
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.wasm.ready).toBe(false)
  })

  it('keeps direct-only wasm-mode formulas off the wasm startup path', () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })

    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellFormula('Sheet1', 'B1', 'A1+1')

    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.wasm.ready).toBe(false)

    engine.clearCell('Sheet1', 'B1')

    expect(engine.wasm.ready).toBe(false)
  })

  it('flushes deferred JS-only edits before the first wasm formula evaluation', () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })

    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellValue('Sheet1', 'A1', 12)
    engine.setCellFormula('Sheet1', 'B1', 'SIN(A1)+1')

    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: Math.sin(12) + 1 })

    engine.setCellValue('Sheet1', 'A1', 20)

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: Math.sin(20) + 1 })
  })
})

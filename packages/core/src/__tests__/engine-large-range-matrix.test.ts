import { describe, expect, it } from 'vitest'
import { indexToColumn } from '@bilig/formula'
import { MAX_WASM_RANGE_CELLS, ValueTag } from '@bilig/protocol'
import { SpreadsheetEngine } from '../index.js'

function expectNumber(value: ReturnType<SpreadsheetEngine['getCellValue']>, expected: number): void {
  expect(value.tag).toBe(ValueTag.Number)
  if (value.tag !== ValueTag.Number) {
    throw new Error(`Expected ${expected}, got ${JSON.stringify(value)}`)
  }
  expect(value.value).toBe(expected)
}

describe('large bounded matrix ranges', () => {
  it('evaluates MMULT ranges above the wasm cap and tracks later blank-to-value edits', async () => {
    const leftRows = 1_001
    const leftCols = 100
    expect(leftRows * leftCols).toBeGreaterThan(MAX_WASM_RANGE_CELLS)

    const engine = new SpreadsheetEngine({ workbookName: 'large-matrix-range' })
    await engine.ready()
    engine.createSheet('Sheet1')

    const leftEndCol = indexToColumn(leftCols - 1)
    const rightCol = indexToColumn(leftCols)
    const outputCol = indexToColumn(leftCols + 2)

    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', `${rightCol}1`, 3)
    engine.setCellFormula('Sheet1', `${outputCol}1`, `MMULT(A1:${leftEndCol}${leftRows},${rightCol}1:${rightCol}${leftCols})`)

    expectNumber(engine.getCellValue('Sheet1', `${outputCol}1`), 6)
    expectNumber(engine.getCellValue('Sheet1', `${outputCol}2`), 0)
    expectNumber(engine.getCellValue('Sheet1', `${outputCol}${leftRows}`), 0)

    engine.setCellValue('Sheet1', 'B2', 4)
    engine.setCellValue('Sheet1', `${rightCol}2`, 5)

    expectNumber(engine.getCellValue('Sheet1', `${outputCol}2`), 20)
  })
})

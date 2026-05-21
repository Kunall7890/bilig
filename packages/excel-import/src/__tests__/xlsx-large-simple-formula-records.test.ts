import { describe, expect, it } from 'vitest'

import { ImportedWorkbookArena } from '../xlsx-large-simple-arena.js'
import { LargeSimpleFormulaRecords, readLargeSimpleFormulaTypeCode } from '../xlsx-large-simple-formula-records.js'

describe('large simple XLSX formula records', () => {
  it('resolves pooled formula text and shared formulas into the import arena', () => {
    const arena = new ImportedWorkbookArena()
    const records = new LargeSimpleFormulaRecords()
    const directCell = arena.addCell({ sheetIndex: 0, row: 0, column: 0, value: 1 })
    const sharedBaseCell = arena.addCell({ sheetIndex: 0, row: 1, column: 0, value: 2 })
    const sharedFollowerCell = arena.addCell({ sheetIndex: 0, row: 2, column: 0, value: 3 })

    records.add(directCell, 0, 0, readLargeSimpleFormulaTypeCode(null), null, 'IF(B1&lt;3,&quot;yes&quot;,&quot;no&quot;)')
    records.add(sharedBaseCell, 1, 0, readLargeSimpleFormulaTypeCode('shared'), 0, 'B1+1')
    records.add(sharedFollowerCell, 2, 0, readLargeSimpleFormulaTypeCode('shared'), 0, '')

    expect(records.resolveIntoArena(arena)).toBe(true)
    expect(arena.materializeSheetCells(0)).toEqual([
      { address: 'A1', value: 1, formula: 'IF(B1<3,"yes","no")' },
      { address: 'A2', value: 2, formula: 'B1+1' },
      { address: 'A3', value: 3, formula: 'B2+1' },
    ])
  })

  it('pools repeated raw formula text when unsupported formula text is allowed', () => {
    const arena = new ImportedWorkbookArena()
    const records = new LargeSimpleFormulaRecords(true)
    const formula = "'[external.xlsx]Sheet1'!A1"
    const firstCell = arena.addCell({ sheetIndex: 0, row: 0, column: 0, value: 1 })
    const secondCell = arena.addCell({ sheetIndex: 0, row: 1, column: 0, value: 2 })
    const thirdCell = arena.addCell({ sheetIndex: 0, row: 2, column: 0, value: 3 })

    records.add(firstCell, 0, 0, readLargeSimpleFormulaTypeCode(null), null, formula)
    records.add(secondCell, 1, 0, readLargeSimpleFormulaTypeCode(null), null, formula)
    records.add(thirdCell, 2, 0, readLargeSimpleFormulaTypeCode(null), null, formula)

    expect(records.rawFormulaPoolCount).toBe(1)
    expect(records.resolveIntoArena(arena)).toBe(true)
    expect(arena.materializeSheetCells(0)).toEqual([
      { address: 'A1', value: 1, formula },
      { address: 'A2', value: 2, formula },
      { address: 'A3', value: 3, formula },
    ])
  })

  it('keeps shared-formula coordinate storage lazy for ordinary formula-heavy sheets', () => {
    const arena = new ImportedWorkbookArena()
    const records = new LargeSimpleFormulaRecords()
    const formulaCount = 1_024

    for (let index = 0; index < formulaCount; index += 1) {
      const cell = arena.addCell({ sheetIndex: 0, row: index, column: 0, value: index })
      records.add(cell, index, 0, readLargeSimpleFormulaTypeCode(null), null, `B${String(index + 1)}+1`)
    }

    expect(records.retainedStorageByteLength()).toBeLessThanOrEqual(formulaCount * 8)
    expect(records.resolveIntoArena(arena)).toBe(true)
    expect(arena.materializeSheetCells(0).at(-1)).toEqual({
      address: `A${String(formulaCount)}`,
      value: formulaCount - 1,
      formula: `B${String(formulaCount)}+1`,
    })

    records.release()
    expect(records.rawFormulaPoolCount).toBe(0)
    expect(records.retainedStorageByteLength()).toBe(0)
  })
})

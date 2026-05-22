import { describe, expect, it } from 'vitest'

import { createWorksheetImportStorageSync } from '../index.js'

describe('worksheet import storage kernel', () => {
  it('owns compact cell, style, and formula records in wasm memory', () => {
    const storage = createWorksheetImportStorageSync()
    storage.reset(1, 1, 1)

    const firstCell = storage.addNumberCell(0, 0, 42)
    const secondCell = storage.addSharedStringCell(1, 2, 7)
    const formulaOnlyCell = storage.addFormulaOnlyCell(2, 3)
    storage.addStyle(0, 0, 3)
    storage.addStyle(1, 2, 4)
    storage.addFormulaRecord(firstCell, 0, 0, 0, null)
    storage.addFormulaRecord(formulaOnlyCell, 2, 3, 1, 9)

    const snapshot = storage.snapshot()

    expect(secondCell).toBe(1)
    expect([...snapshot.rows]).toEqual([0, 1, 2])
    expect([...snapshot.columns]).toEqual([0, 2, 3])
    expect([...snapshot.valueKinds]).toEqual([storage.valueKindNumber, storage.valueKindSharedString, storage.valueKindFormulaOnly])
    expect([...snapshot.numbers].map((value) => (Number.isNaN(value) ? 'NaN' : value))).toEqual([42, 'NaN', 'NaN'])
    expect([...snapshot.sharedStringIds]).toEqual([storage.noSharedFormulaIndex, 7, storage.noSharedFormulaIndex])
    expect([...snapshot.styleRows]).toEqual([0, 1])
    expect([...snapshot.styleColumns]).toEqual([0, 2])
    expect([...snapshot.styleIds]).toEqual([3, 4])
    expect([...snapshot.formulaCellIndexes]).toEqual([0, 2])
    expect([...snapshot.formulaRows]).toEqual([0, 2])
    expect([...snapshot.formulaColumns]).toEqual([0, 3])
    expect([...snapshot.formulaTypeCodes]).toEqual([0, 1])
    expect([...snapshot.formulaSharedIndexes]).toEqual([storage.noSharedFormulaIndex, 9])
  })
})

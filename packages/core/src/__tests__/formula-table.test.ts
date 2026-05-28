import { describe, expect, it } from 'vitest'
import { CellStore } from '../cell-store.js'
import { FormulaTable, MULTIPLE_DIRECT_SCALAR_OUTPUTS } from '../formula-table.js'

describe('FormulaTable', () => {
  it('stores formulas in a compact vector and writes real formula ids into the cell store', () => {
    const store = new CellStore()
    const firstCell = store.allocate(0, 0, 0)
    const secondCell = store.allocate(0, 0, 5)
    const formulas = new FormulaTable<{ cellIndex: number; source: string }>(store)

    const firstId = formulas.set(secondCell, { cellIndex: secondCell, source: 'A1*2' })
    const secondId = formulas.set(firstCell, { cellIndex: firstCell, source: 'B1*3' })

    expect(firstId).toBe(1)
    expect(secondId).toBe(2)
    expect(store.formulaIds[secondCell]).toBe(1)
    expect(store.formulaIds[firstCell]).toBe(2)
    expect([...formulas.keys()]).toEqual([secondCell, firstCell])
  })

  it('reuses freed formula slots instead of growing forever', () => {
    const store = new CellStore()
    const a = store.allocate(0, 0, 0)
    const b = store.allocate(0, 0, 1)
    const c = store.allocate(0, 0, 2)
    const formulas = new FormulaTable<{ cellIndex: number; source: string }>(store)

    formulas.set(a, { cellIndex: a, source: '1' })
    formulas.set(b, { cellIndex: b, source: '2' })
    const removed = formulas.delete(a)
    const reusedId = formulas.set(c, { cellIndex: c, source: '3' })

    expect(removed?.source).toBe('1')
    expect(reusedId).toBe(1)
    expect(store.formulaIds[a]).toBe(0)
    expect(store.formulaIds[c]).toBe(1)
    expect(formulas.size).toBe(2)
    expect([...formulas.values()].map((formula) => formula.source)).toEqual(['3', '2'])
  })

  it('updates an existing formula record in place without allocating a new id', () => {
    const store = new CellStore()
    const cellIndex = store.allocate(0, 0, 0)
    const formulas = new FormulaTable<{ cellIndex: number; source: string }>(store)

    const originalId = formulas.set(cellIndex, { cellIndex, source: 'A1*2' })
    const updatedId = formulas.set(cellIndex, { cellIndex, source: 'A1*3' })

    expect(updatedId).toBe(originalId)
    expect(formulas.get(cellIndex)).toEqual({ cellIndex, source: 'A1*3' })
    expect(formulas.size).toBe(1)
  })

  it('tracks and clears direct scalar delta-input metadata', () => {
    const store = new CellStore()
    const cellIndex = store.allocate(0, 0, 0)
    const deltaInputCellIndices: Array<number | undefined> = []
    const formulas = new FormulaTable<{ cellIndex: number; deltaInputCellIndex?: number }>(store, {
      deltaInputCellIndices,
      readDeltaInputCellIndex: (record) => record.deltaInputCellIndex,
    })

    formulas.set(cellIndex, { cellIndex, deltaInputCellIndex: 7 })
    expect(deltaInputCellIndices[cellIndex]).toBe(7)

    formulas.get(cellIndex)!.deltaInputCellIndex = 9
    formulas.refreshTrackedMetadata(cellIndex)
    expect(deltaInputCellIndices[cellIndex]).toBe(9)

    formulas.set(cellIndex, { cellIndex })
    expect(deltaInputCellIndices[cellIndex]).toBeUndefined()

    formulas.set(cellIndex, { cellIndex, deltaInputCellIndex: 11 })
    expect(deltaInputCellIndices[cellIndex]).toBe(11)

    formulas.delete(cellIndex)
    expect(deltaInputCellIndices[cellIndex]).toBeUndefined()

    formulas.set(cellIndex, { cellIndex, deltaInputCellIndex: 13 })
    formulas.clear()
    expect(deltaInputCellIndices).toHaveLength(0)
  })

  it('tracks unambiguous direct scalar output cells by input cell', () => {
    const store = new CellStore()
    const firstFormula = store.allocate(0, 0, 1)
    const secondFormula = store.allocate(0, 1, 1)
    const deltaInputCellIndices: Array<number | undefined> = []
    const singleOutputCellIndicesByInput: Array<number | undefined> = []
    const formulas = new FormulaTable<{ cellIndex: number; deltaInputCellIndex?: number }>(store, {
      deltaInputCellIndices,
      singleOutputCellIndicesByInput,
      readDeltaInputCellIndex: (record) => record.deltaInputCellIndex,
    })

    formulas.set(firstFormula, { cellIndex: firstFormula, deltaInputCellIndex: 7 })
    expect(singleOutputCellIndicesByInput[7]).toBe(firstFormula)

    formulas.set(firstFormula, { cellIndex: firstFormula, deltaInputCellIndex: 8 })
    expect(singleOutputCellIndicesByInput[7]).toBeUndefined()
    expect(singleOutputCellIndicesByInput[8]).toBe(firstFormula)

    formulas.set(secondFormula, { cellIndex: secondFormula, deltaInputCellIndex: 8 })
    expect(singleOutputCellIndicesByInput[8]).toBe(MULTIPLE_DIRECT_SCALAR_OUTPUTS)

    formulas.delete(firstFormula)
    expect(singleOutputCellIndicesByInput[8]).toBe(MULTIPLE_DIRECT_SCALAR_OUTPUTS)

    formulas.clear()
    expect(singleOutputCellIndicesByInput).toHaveLength(0)
  })
})

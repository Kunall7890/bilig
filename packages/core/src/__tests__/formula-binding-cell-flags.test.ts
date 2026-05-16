import { describe, expect, it } from 'vitest'
import { FormulaMode } from '@bilig/protocol'
import { CellFlags, CellStore } from '../cell-store.js'
import { clearFormulaRuntimeFlags, markFormulaCellBound } from '../engine/services/formula-binding-cell-flags.js'

describe('formula binding cell flags', () => {
  it('marks JS-only formula cells while clearing derived output ownership flags', () => {
    const store = new CellStore()
    const cellIndex = store.allocate(1, 0, 0)
    store.flags[cellIndex] = CellFlags.Materialized | CellFlags.SpillChild | CellFlags.PivotOutput

    markFormulaCellBound(store, cellIndex, FormulaMode.JsOnly)

    expect(store.flags[cellIndex]).toBe(CellFlags.Materialized | CellFlags.HasFormula | CellFlags.JsOnly)
  })

  it('marks WASM formula cells and clears stale JS-only flags', () => {
    const store = new CellStore()
    const cellIndex = store.allocate(1, 0, 0)
    store.flags[cellIndex] = CellFlags.Materialized | CellFlags.HasFormula | CellFlags.JsOnly

    markFormulaCellBound(store, cellIndex, FormulaMode.WasmFastPath)

    expect(store.flags[cellIndex]).toBe(CellFlags.Materialized | CellFlags.HasFormula)
  })

  it('clears runtime formula flags without dropping unrelated materialization state', () => {
    const store = new CellStore()
    const cellIndex = store.allocate(1, 0, 0)
    store.flags[cellIndex] =
      CellFlags.Materialized |
      CellFlags.HasFormula |
      CellFlags.JsOnly |
      CellFlags.InCycle |
      CellFlags.SpillChild |
      CellFlags.PivotOutput |
      CellFlags.AuthoredBlank

    clearFormulaRuntimeFlags(store, cellIndex)

    expect(store.flags[cellIndex]).toBe(CellFlags.Materialized | CellFlags.AuthoredBlank)
  })
})

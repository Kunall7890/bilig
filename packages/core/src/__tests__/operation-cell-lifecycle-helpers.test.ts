import { describe, expect, it } from 'vitest'
import { CellFlags, CellStore } from '../cell-store.js'
import { makeCellEntity, makeRangeEntity } from '../entity-ids.js'
import { createEngineCounters } from '../perf/engine-counters.js'
import { StringPool } from '../string-pool.js'
import {
  hasOperationCycleMembers,
  markOperationCycleMemberInputsChanged,
  normalizeOperationHistoryDependencyPlaceholder,
  pruneOperationCellIfOrphaned,
  refreshDependentRangesAndRebindOperationFormulaDependents,
} from '../engine/services/operation-cell-lifecycle-helpers.js'

function createCellStoreWithEmptyCell(): CellStore {
  const cellStore = new CellStore()
  cellStore.allocate(1, 0, 0)
  return cellStore
}

describe('operation cell lifecycle helpers', () => {
  it('prunes cells only after formula dependents are gone', () => {
    const pruned: number[] = []
    const workbook = { pruneCellIfEmpty: (cellIndex: number) => pruned.push(cellIndex) }

    pruneOperationCellIfOrphaned({
      workbook,
      cellIndex: 4,
      collectFormulaDependents: () => Uint32Array.of(9),
    })
    pruneOperationCellIfOrphaned({
      workbook,
      cellIndex: 5,
      collectFormulaDependents: () => new Uint32Array(),
    })

    expect(pruned).toEqual([5])
  })

  it('normalizes undo and restore placeholders only for empty dependent cells', () => {
    const cellStore = createCellStoreWithEmptyCell()
    const strings = new StringPool()
    cellStore.versions[0] = 12

    normalizeOperationHistoryDependencyPlaceholder({
      state: {
        workbook: { cellStore, getCellFormat: () => undefined },
        strings,
      },
      source: 'undo',
      cellIndex: 0,
      collectFormulaDependents: (entityId) => (entityId === makeCellEntity(0) ? Uint32Array.of(10) : new Uint32Array()),
    })

    expect(cellStore.versions[0]).toBe(0)

    cellStore.versions[0] = 12
    cellStore.flags[0] = CellFlags.PendingDelete
    normalizeOperationHistoryDependencyPlaceholder({
      state: {
        workbook: { cellStore, getCellFormat: () => undefined },
        strings,
      },
      source: 'restore',
      cellIndex: 0,
      collectFormulaDependents: () => Uint32Array.of(10),
    })

    expect(cellStore.versions[0]).toBe(12)
  })

  it('marks cycle members as changed inputs and records cycle scans', () => {
    const cellStore = createCellStoreWithEmptyCell()
    cellStore.allocate(1, 1, 0)
    cellStore.flags[1] = CellFlags.InCycle
    const counters = createEngineCounters()
    const formulas = new Map<number, unknown>([
      [0, {}],
      [1, {}],
    ])
    const marked: number[] = []

    expect(hasOperationCycleMembers({ counters, formulas, cellStore })).toBe(true)
    expect(counters.cycleFormulaScans).toBe(1)
    expect(
      markOperationCycleMemberInputsChanged({
        formulas,
        cellStore,
        changedInputCount: 3,
        markInputChanged(cellIndex, changedInputCount) {
          marked.push(cellIndex)
          return changedInputCount + 1
        },
      }),
    ).toBe(4)
    expect(marked).toEqual([1])
  })

  it('refreshes range dependents and rebinds formula dependents except the edited cell', () => {
    const refreshed: number[][] = []
    const rebound: { formulas: number[]; formulaChangedCount: number }[] = []

    const result = refreshDependentRangesAndRebindOperationFormulaDependents({
      cellIndex: 7,
      formulaChangedCount: 2,
      getEntityDependents: () => Uint32Array.of(makeRangeEntity(3), makeCellEntity(8), makeRangeEntity(5)),
      collectFormulaDependents: () => Uint32Array.of(7, 11, 13),
      refreshRangeDependencies: (ranges) => refreshed.push([...ranges]),
      rebindFormulaCells(formulas, formulaChangedCount) {
        rebound.push({ formulas: [...formulas], formulaChangedCount })
        return formulaChangedCount + formulas.length
      },
    })

    expect(result).toBe(4)
    expect(refreshed).toEqual([[3, 5]])
    expect(rebound).toEqual([{ formulas: [11, 13], formulaChangedCount: 2 }])
  })
})

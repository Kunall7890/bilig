import { describe, expect, it } from 'vitest'
import { CellStore } from '../cell-store.js'
import { makeRangeEntity } from '../entity-ids.js'
import {
  canSkipOperationDirtyTraversalForChangedInputs,
  operationChangedInputsNeedRegionQueryIndices,
  type OperationPostRecalcDirectFormulaIndexAccess,
} from '../engine/services/operation-dirty-traversal-helpers.js'

function postRecalcIndex(
  input: {
    readonly cells?: readonly number[]
    readonly coveredRangeInputs?: readonly number[]
    readonly coveredFormulaInputs?: readonly number[]
  } = {},
): OperationPostRecalcDirectFormulaIndexAccess {
  const cells = new Set(input.cells ?? [])
  const coveredRangeInputs = new Set(input.coveredRangeInputs ?? [])
  const coveredFormulaInputs = new Set(input.coveredFormulaInputs ?? [])
  return {
    has: (cellIndex) => cells.has(cellIndex),
    hasCoveredDirectRangeInput: (cellIndex) => coveredRangeInputs.has(cellIndex),
    hasCoveredDirectFormulaInput: (cellIndex) => coveredFormulaInputs.has(cellIndex),
  }
}

function createCellStore(): CellStore {
  const cellStore = new CellStore()
  cellStore.allocate(1, 4, 2)
  cellStore.allocate(1, 5, 2)
  return cellStore
}

function createAccess(overrides: Partial<Parameters<typeof canSkipOperationDirtyTraversalForChangedInputs>[0]['access']> = {}) {
  const cellStore = createCellStore()
  return {
    cellStore,
    access: {
      workbook: {
        cellStore,
        getSheetNameById: (sheetId: number) => (sheetId === 1 ? 'Sheet1' : undefined),
        getSheetById: () => undefined,
      },
      getSingleEntityDependent: () => -1,
      getEntityDependents: () => new Uint32Array(),
      collectRegionFormulaDependentsForCell: () => new Uint32Array(),
      collectAffectedDirectRangeDependents: () => [],
      hasTrackedExactLookupDependents: () => false,
      hasTrackedSortedLookupDependents: () => false,
      hasTrackedDirectRangeDependents: () => false,
      ...overrides,
    },
  }
}

describe('operation dirty traversal helpers', () => {
  it('allows dirty traversal skip when all direct formula dependents are post-recalc covered', () => {
    const { access } = createAccess({ getSingleEntityDependent: () => 10 })

    expect(
      canSkipOperationDirtyTraversalForChangedInputs({
        changedInputCellIndices: Uint32Array.of(0),
        changedInputCount: 1,
        postRecalcDirectFormulaIndices: postRecalcIndex({ cells: [10] }),
        access,
      }),
    ).toBe(true)
  })

  it('rejects dirty traversal skip for uncovered formula and range dependents', () => {
    const directFormula = createAccess({ getSingleEntityDependent: () => 10 })
    expect(
      canSkipOperationDirtyTraversalForChangedInputs({
        changedInputCellIndices: Uint32Array.of(0),
        changedInputCount: 1,
        postRecalcDirectFormulaIndices: postRecalcIndex({ cells: [11] }),
        access: directFormula.access,
      }),
    ).toBe(false)

    const range = createAccess({
      getSingleEntityDependent: () => makeRangeEntity(3),
      hasTrackedDirectRangeDependents: () => false,
    })
    expect(
      canSkipOperationDirtyTraversalForChangedInputs({
        changedInputCellIndices: Uint32Array.of(0),
        changedInputCount: 1,
        postRecalcDirectFormulaIndices: postRecalcIndex(),
        access: range.access,
      }),
    ).toBe(false)
  })

  it('accepts range dependents when tracked region and direct range dependents are covered', () => {
    const { access } = createAccess({
      getSingleEntityDependent: () => makeRangeEntity(3),
      hasTrackedDirectRangeDependents: () => true,
      collectRegionFormulaDependentsForCell: () => Uint32Array.of(20),
      collectAffectedDirectRangeDependents: () => [21],
    })

    expect(
      canSkipOperationDirtyTraversalForChangedInputs({
        changedInputCellIndices: Uint32Array.of(0),
        changedInputCount: 1,
        postRecalcDirectFormulaIndices: postRecalcIndex({ cells: [20, 21] }),
        access,
      }),
    ).toBe(true)
  })

  it('requires lookup dependents to be covered unless the lookup input was already handled', () => {
    const { access } = createAccess({
      getSingleEntityDependent: () => 10,
      hasTrackedExactLookupDependents: () => true,
      getEntityDependents: () => Uint32Array.of(30),
    })

    expect(
      canSkipOperationDirtyTraversalForChangedInputs({
        changedInputCellIndices: Uint32Array.of(0),
        changedInputCount: 1,
        postRecalcDirectFormulaIndices: postRecalcIndex({ cells: [10] }),
        access,
      }),
    ).toBe(false)
    expect(
      canSkipOperationDirtyTraversalForChangedInputs({
        changedInputCellIndices: Uint32Array.of(0),
        changedInputCount: 1,
        postRecalcDirectFormulaIndices: postRecalcIndex({ cells: [10] }),
        options: { lookupHandledInputCellIndices: [0] },
        access,
      }),
    ).toBe(true)
  })

  it('detects when changed inputs require region query indices', () => {
    const { access } = createAccess({ hasTrackedDirectRangeDependents: () => true })

    expect(
      operationChangedInputsNeedRegionQueryIndices({
        changedInputCellIndices: Uint32Array.of(0),
        changedInputCount: 1,
        postRecalcDirectFormulaIndices: postRecalcIndex(),
        access,
      }),
    ).toBe(true)
    expect(
      operationChangedInputsNeedRegionQueryIndices({
        changedInputCellIndices: Uint32Array.of(0),
        changedInputCount: 1,
        postRecalcDirectFormulaIndices: postRecalcIndex({ coveredRangeInputs: [0] }),
        access,
      }),
    ).toBe(false)
  })
})

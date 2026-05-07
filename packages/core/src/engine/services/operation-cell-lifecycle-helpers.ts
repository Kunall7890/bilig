import type { EntityId } from '@bilig/protocol'
import { ValueTag } from '@bilig/protocol'
import { CellFlags, type CellStore } from '../../cell-store.js'
import { entityPayload, isRangeEntity, makeCellEntity } from '../../entity-ids.js'
import { addEngineCounter, type EngineCounters } from '../../perf/engine-counters.js'
import type { StringPool } from '../../string-pool.js'

type OperationMutationSource = 'local' | 'remote' | 'restore' | 'undo' | 'redo'

type OperationFormulaDependentsReader = (entityId: EntityId) => Uint32Array

export function pruneOperationCellIfOrphaned(input: {
  readonly workbook: { readonly pruneCellIfEmpty: (cellIndex: number) => void }
  readonly cellIndex: number
  readonly collectFormulaDependents: OperationFormulaDependentsReader
}): void {
  if (input.collectFormulaDependents(makeCellEntity(input.cellIndex)).length > 0) {
    return
  }
  input.workbook.pruneCellIfEmpty(input.cellIndex)
}

export function normalizeOperationHistoryDependencyPlaceholder(input: {
  readonly state: {
    readonly workbook: {
      readonly cellStore: CellStore
      readonly getCellFormat: (cellIndex: number) => unknown
    }
    readonly strings: StringPool
  }
  readonly source: OperationMutationSource
  readonly cellIndex: number
  readonly collectFormulaDependents: OperationFormulaDependentsReader
}): void {
  if (input.source !== 'undo' && input.source !== 'restore') {
    return
  }
  if (input.state.workbook.getCellFormat(input.cellIndex) !== undefined) {
    return
  }
  const flags = input.state.workbook.cellStore.flags[input.cellIndex] ?? 0
  if (
    (flags &
      (CellFlags.HasFormula |
        CellFlags.JsOnly |
        CellFlags.InCycle |
        CellFlags.SpillChild |
        CellFlags.PivotOutput |
        CellFlags.PendingDelete)) !==
    0
  ) {
    return
  }
  const value = input.state.workbook.cellStore.getValue(input.cellIndex, (id) => input.state.strings.get(id))
  if (value.tag !== ValueTag.Empty) {
    return
  }
  if (input.collectFormulaDependents(makeCellEntity(input.cellIndex)).length === 0) {
    return
  }
  input.state.workbook.cellStore.versions[input.cellIndex] = 0
}

export function markOperationCycleMemberInputsChanged(input: {
  readonly formulas: { readonly forEach: (callback: (formula: unknown, cellIndex: number) => void) => void }
  readonly cellStore: CellStore
  readonly changedInputCount: number
  readonly markInputChanged: (cellIndex: number, changedInputCount: number) => number
}): number {
  let changedInputCount = input.changedInputCount
  input.formulas.forEach((_formula, cellIndex) => {
    if (((input.cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) === 0) {
      return
    }
    changedInputCount = input.markInputChanged(cellIndex, changedInputCount)
  })
  return changedInputCount
}

export function hasOperationCycleMembers(input: {
  readonly counters: EngineCounters
  readonly formulas: { readonly forEach: (callback: (formula: unknown, cellIndex: number) => void) => void }
  readonly cellStore: CellStore
}): boolean {
  addEngineCounter(input.counters, 'cycleFormulaScans')
  let found = false
  input.formulas.forEach((_formula, cellIndex) => {
    if (((input.cellStore.flags[cellIndex] ?? 0) & CellFlags.InCycle) !== 0) {
      found = true
    }
  })
  return found
}

export function refreshDependentRangesAndRebindOperationFormulaDependents(input: {
  readonly cellIndex: number
  readonly formulaChangedCount: number
  readonly getEntityDependents: (entityId: EntityId) => Uint32Array
  readonly collectFormulaDependents: OperationFormulaDependentsReader
  readonly refreshRangeDependencies: (rangeIndices: readonly number[]) => void
  readonly rebindFormulaCells: (formulas: readonly number[], formulaChangedCount: number) => number
}): number {
  const directDependents = input.getEntityDependents(makeCellEntity(input.cellIndex))
  const rangeIndices: number[] = []
  for (let index = 0; index < directDependents.length; index += 1) {
    const dependent = directDependents[index]!
    if (isRangeEntity(dependent)) {
      rangeIndices.push(entityPayload(dependent))
    }
  }
  if (rangeIndices.length > 0) {
    input.refreshRangeDependencies(rangeIndices)
  }
  const formulas = Array.from(input.collectFormulaDependents(makeCellEntity(input.cellIndex))).filter(
    (candidate) => candidate !== input.cellIndex,
  )
  if (formulas.length === 0) {
    return input.formulaChangedCount
  }
  return input.rebindFormulaCells(formulas, input.formulaChangedCount)
}

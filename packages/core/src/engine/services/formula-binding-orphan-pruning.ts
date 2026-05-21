import { CellFlags } from '../../cell-store.js'
import type { EdgeSlice } from '../../edge-arena.js'
import { makeCellEntity } from '../../entity-ids.js'
import type { EngineRuntimeState } from '../runtime-state.js'

export function pruneFormulaBindingTrackedDependencyCell(args: {
  readonly state: Pick<EngineRuntimeState, 'workbook'>
  readonly getReverseEdgeSlice: (entityId: number) => EdgeSlice | undefined
  readonly cellIndex: number
  readonly ownerCellIndex: number
}): void {
  if (args.cellIndex === args.ownerCellIndex || !canPruneFormulaBindingDependencyCell(args)) {
    return
  }
  args.state.workbook.pruneCellIfEmpty(args.cellIndex)
}

export function pruneFormulaBindingOrphanedDependencyCells(args: {
  readonly state: Pick<EngineRuntimeState, 'workbook'>
  readonly getReverseEdgeSlice: (entityId: number) => EdgeSlice | undefined
  readonly cellIndices: readonly number[]
}): void {
  args.cellIndices.forEach((cellIndex) => {
    if (canPruneFormulaBindingDependencyCell({ ...args, cellIndex })) {
      args.state.workbook.pruneCellIfEmpty(cellIndex)
    }
  })
}

function canPruneFormulaBindingDependencyCell(args: {
  readonly state: Pick<EngineRuntimeState, 'workbook'>
  readonly getReverseEdgeSlice: (entityId: number) => EdgeSlice | undefined
  readonly cellIndex: number
}): boolean {
  return (
    !args.getReverseEdgeSlice(makeCellEntity(args.cellIndex)) &&
    ((args.state.workbook.cellStore.flags[args.cellIndex] ?? 0) & CellFlags.AuthoredBlank) === 0
  )
}

import type { EngineOp } from '@bilig/workbook'
import { spillDependencyKey } from '../../engine-metadata-utils.js'
import { batchOpOrder, createBatch, type OpOrder, type ReplicaState } from '../../replica-state.js'
import type { WorkbookPivotRecord, WorkbookSpillRecord } from '../../workbook-store.js'
import { collectTrackedDependents } from './direct-formula-recalc-helpers.js'
import { assertNever } from './operation-change-helpers.js'

type OperationSpillRangeOp = Extract<EngineOp, { kind: 'upsertSpillRange' | 'deleteSpillRange' }>
type OperationPivotUpsertOp = Extract<EngineOp, { kind: 'upsertPivotTable' }>
type OperationPivotDeleteOp = Extract<EngineOp, { kind: 'deletePivotTable' }>

export type OperationDerivedOp = OperationSpillRangeOp | OperationPivotUpsertOp | OperationPivotDeleteOp

export interface OperationDerivedWorkbookAccess {
  readonly setSpill: (sheetName: string, address: string, rows: number, cols: number) => WorkbookSpillRecord
  readonly deleteSpill: (sheetName: string, address: string) => boolean
  readonly setPivot: (record: WorkbookPivotRecord) => WorkbookPivotRecord
  readonly getPivot: (sheetName: string, address: string) => WorkbookPivotRecord | undefined
  readonly deletePivot: (sheetName: string, address: string) => boolean
}

export interface OperationDerivedOpApplier {
  readonly applySpillRangeOp: (op: OperationSpillRangeOp, order: OpOrder) => number[]
  readonly applyPivotUpsertOp: (op: OperationPivotUpsertOp, order: OpOrder) => number[]
  readonly applyPivotDeleteOp: (op: OperationPivotDeleteOp, order: OpOrder) => number[]
  readonly applyDerivedOpNow: (op: OperationDerivedOp) => number[]
}

export function createOperationDerivedOpApplier(args: {
  readonly state: {
    readonly workbook: OperationDerivedWorkbookAccess
    readonly replicaState: ReplicaState
  }
  readonly reverseSpillEdges: Map<string, Set<number>>
  readonly setEntityVersionForOp: (op: OperationDerivedOp, order: OpOrder) => void
  readonly materializePivot: (pivot: WorkbookPivotRecord) => number[]
  readonly clearOwnedPivot: (pivot: WorkbookPivotRecord) => number[]
  readonly rebindFormulaCells: (candidates: readonly number[], formulaChangedCount: number) => number
}): OperationDerivedOpApplier {
  const applySpillRangeOp = (op: OperationSpillRangeOp, order: OpOrder): number[] => {
    if (op.kind === 'upsertSpillRange') {
      args.state.workbook.setSpill(op.sheetName, op.address, op.rows, op.cols)
    } else {
      args.state.workbook.deleteSpill(op.sheetName, op.address)
    }
    args.setEntityVersionForOp(op, order)
    return collectTrackedDependents(args.reverseSpillEdges, [spillDependencyKey(op.sheetName, op.address)])
  }

  const applyPivotUpsertOp = (op: OperationPivotUpsertOp, order: OpOrder): number[] => {
    const pivot = {
      name: op.name,
      sheetName: op.sheetName,
      address: op.address,
      source: op.source,
      groupBy: op.groupBy,
      values: op.values,
      rows: op.rows,
      cols: op.cols,
    } satisfies WorkbookPivotRecord
    args.state.workbook.setPivot(pivot)
    args.setEntityVersionForOp(op, order)
    return args.materializePivot(pivot)
  }

  const applyPivotDeleteOp = (op: OperationPivotDeleteOp, order: OpOrder): number[] => {
    const pivot = args.state.workbook.getPivot(op.sheetName, op.address)
    if (!pivot) {
      args.setEntityVersionForOp(op, order)
      return []
    }
    const changedPivotOutputs = args.clearOwnedPivot(pivot)
    args.state.workbook.deletePivot(op.sheetName, op.address)
    args.setEntityVersionForOp(op, order)
    return changedPivotOutputs
  }

  return {
    applySpillRangeOp,
    applyPivotUpsertOp,
    applyPivotDeleteOp,
    applyDerivedOpNow(op) {
      const batch = createBatch(args.state.replicaState, [op])
      const order = batchOpOrder(batch, 0)
      switch (op.kind) {
        case 'upsertSpillRange':
        case 'deleteSpillRange': {
          const candidates = applySpillRangeOp(op, order)
          args.rebindFormulaCells(candidates, 0)
          return candidates
        }
        case 'upsertPivotTable':
          return applyPivotUpsertOp(op, order)
        case 'deletePivotTable':
          return applyPivotDeleteOp(op, order)
        default:
          return assertNever(op)
      }
    },
  }
}

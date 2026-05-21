import type { EngineOp } from '@bilig/workbook'

type StructuralAxisOp = Extract<
  EngineOp,
  { kind: 'insertRows' | 'deleteRows' | 'moveRows' | 'insertColumns' | 'deleteColumns' | 'moveColumns' }
>

export interface StructuralNoValueFinalizationEligibility {
  readonly isRestore: boolean
  readonly topologyChanged: boolean
  readonly formulaChangedCount: number
  readonly explicitChangedCount: number
  readonly precomputedKernelSyncCellCount: number
  readonly invalidatedRangeCount: number
  readonly invalidatedRowCount: number
  readonly invalidatedColumnCount: number
  readonly activeFormulaCount: number
  readonly hasVolatileFormulas: boolean | undefined
  readonly hasActivePivots: boolean
}

export function isStructuralAxisOp(op: EngineOp): op is StructuralAxisOp {
  return (
    op.kind === 'insertRows' ||
    op.kind === 'deleteRows' ||
    op.kind === 'moveRows' ||
    op.kind === 'insertColumns' ||
    op.kind === 'deleteColumns' ||
    op.kind === 'moveColumns'
  )
}

export function canFinalizeStructuralNoValueMutationWithoutRecalc(args: StructuralNoValueFinalizationEligibility): boolean {
  return (
    !args.isRestore &&
    !args.topologyChanged &&
    args.formulaChangedCount === 0 &&
    args.explicitChangedCount === 0 &&
    args.precomputedKernelSyncCellCount === 0 &&
    args.invalidatedRangeCount === 0 &&
    (args.invalidatedRowCount > 0 || args.invalidatedColumnCount > 0) &&
    !args.hasActivePivots &&
    (args.activeFormulaCount === 0 || args.hasVolatileFormulas === false)
  )
}

import { parseCellAddress } from '@bilig/formula'
import type { EngineOp } from '@bilig/workbook-domain'
import { CellFlags } from '../../cell-store.js'
import type { OpOrder } from '../../replica-state.js'
import type { PreparedCellAddress } from '../runtime-state.js'
import type { DirectFormulaIndexCollection } from './direct-formula-index-collection.js'
import type { DirectFormulaMetricCounts } from './operation-post-recalc-direct-formulas.js'
import { cellTouchesOperationPivotSource } from './operation-pivot-source-helpers.js'
import type { OperationDirectRangeDependentService } from './operation-direct-range-dependents.js'
import type { CreateEngineOperationServiceArgs } from './operation-service-types.js'

type BatchSetCellFormulaOp = Extract<EngineOp, { kind: 'setCellFormula' }>

interface OperationBatchPreparedCells {
  readonly getExistingCellIndex: (sheetName: string, address: string, preparedCellAddress: PreparedCellAddress | null) => number | undefined
  readonly ensureCellTracked: (sheetName: string, address: string, preparedCellAddress: PreparedCellAddress | null) => number
}

interface BatchCellFormulaMutationCounts {
  readonly changedInputCount: number
  readonly formulaChangedCount: number
  readonly explicitChangedCount: number
  readonly topologyChanged: boolean
  readonly refreshAllPivots: boolean
  readonly compileMs: number
}

interface ApplyBatchSetCellFormulaOpArgs extends BatchCellFormulaMutationCounts {
  readonly serviceArgs: CreateEngineOperationServiceArgs
  readonly op: BatchSetCellFormulaOp
  readonly order: OpOrder
  readonly isRestore: boolean
  readonly preparedCellAddress: PreparedCellAddress | null
  readonly preparedCells: OperationBatchPreparedCells
  readonly postRecalcDirectFormulaIndices: DirectFormulaIndexCollection
  readonly postRecalcDirectFormulaMetrics: DirectFormulaMetricCounts
  readonly setEntityVersionForOp: (op: EngineOp, order: OpOrder) => void
  readonly hasTrackedExactLookupDependents: (sheetId: number, col: number) => boolean
  readonly hasTrackedSortedLookupDependents: (sheetId: number, col: number) => boolean
  readonly hasTrackedDirectRangeDependents: (sheetId: number, col: number) => boolean
  readonly readExactNumericValueForLookup: (cellIndex: number | undefined) => number | undefined
  readonly tryApplyFormulaReplacementAsDirectScalarDeltaRoot: (request: {
    readonly cellIndex: number
    readonly oldNumber: number | undefined
    readonly changedTopology: boolean
    readonly postRecalcDirectFormulaIndices: DirectFormulaIndexCollection
    readonly postRecalcDirectFormulaMetrics: DirectFormulaMetricCounts
  }) => boolean
  readonly refreshDependentRangesAndRebindFormulaDependents: (cellIndex: number, formulaChangedCount: number) => number
  readonly collectAffectedDirectRangeDependents: OperationDirectRangeDependentService['collectAffectedDirectRangeDependents']
  readonly clearLookupImpactCaches: () => void
}

export function applyBatchSetCellFormulaOp(request: ApplyBatchSetCellFormulaOpArgs): BatchCellFormulaMutationCounts {
  const args = request.serviceArgs
  const { op, preparedCellAddress } = request
  const parsedAddress = parseCellAddress(op.address, op.sheetName)
  const sheetId = args.state.workbook.getSheet(op.sheetName)?.id
  let changedInputCount = request.changedInputCount
  let formulaChangedCount = request.formulaChangedCount
  let explicitChangedCount = request.explicitChangedCount
  let topologyChanged = request.topologyChanged
  let refreshAllPivots = request.refreshAllPivots
  let compileMs = request.compileMs

  if (
    !request.isRestore &&
    cellTouchesOperationPivotSource({
      workbook: args.state.workbook,
      sheetName: op.sheetName,
      row: parsedAddress.row,
      col: parsedAddress.col,
    })
  ) {
    refreshAllPivots = true
  }
  args.invalidateExactLookupColumn({ sheetName: op.sheetName, col: parsedAddress.col })
  args.invalidateSortedLookupColumn({ sheetName: op.sheetName, col: parsedAddress.col })
  if (!request.isRestore) {
    const existingIndex = request.preparedCells.getExistingCellIndex(op.sheetName, op.address, preparedCellAddress)
    if (existingIndex !== undefined) {
      changedInputCount = args.markPivotRootsChanged(args.clearPivotForCell(existingIndex), changedInputCount)
    }
  }
  const cellIndex = request.preparedCells.ensureCellTracked(op.sheetName, op.address, preparedCellAddress)
  const priorHadFormula = args.state.formulas.get(cellIndex) !== undefined
  const oldFormulaNumber = !request.isRestore && priorHadFormula ? request.readExactNumericValueForLookup(cellIndex) : undefined
  args.state.workbook.cellStore.flags[cellIndex] = (args.state.workbook.cellStore.flags[cellIndex] ?? 0) & ~CellFlags.AuthoredBlank
  if (!request.isRestore) {
    changedInputCount = args.markSpillRootsChanged(args.clearOwnedSpill(cellIndex), changedInputCount)
  }
  const compileStarted = request.isRestore ? 0 : performance.now()
  const hasFormulaColumnAggregateDependents = sheetId !== undefined && request.hasTrackedDirectRangeDependents(sheetId, parsedAddress.col)
  try {
    const priorDirectScalarFormula = args.state.formulas.get(cellIndex)?.directScalar !== undefined
    const canRewriteFormulaPreservingBinding =
      !request.isRestore &&
      priorDirectScalarFormula &&
      sheetId !== undefined &&
      !request.hasTrackedExactLookupDependents(sheetId, parsedAddress.col) &&
      !request.hasTrackedSortedLookupDependents(sheetId, parsedAddress.col) &&
      !hasFormulaColumnAggregateDependents &&
      args.rewriteFormulaSourcePreservingBinding !== undefined
    const changedTopology = canRewriteFormulaPreservingBinding
      ? args.rewriteFormulaSourcePreservingBinding(cellIndex, op.sheetName, op.formula)
        ? false
        : args.bindFormula(cellIndex, op.sheetName, op.formula)
      : args.bindFormula(cellIndex, op.sheetName, op.formula)
    if (hasFormulaColumnAggregateDependents) {
      args.invalidateAggregateColumn({ sheetName: op.sheetName, col: parsedAddress.col })
    }
    request.clearLookupImpactCaches()
    if (!request.isRestore) {
      compileMs += performance.now() - compileStarted
    }
    changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
    const handledFormulaReplacementAsDirectDelta =
      priorHadFormula &&
      sheetId !== undefined &&
      !request.hasTrackedExactLookupDependents(sheetId, parsedAddress.col) &&
      !request.hasTrackedSortedLookupDependents(sheetId, parsedAddress.col) &&
      !hasFormulaColumnAggregateDependents &&
      request.tryApplyFormulaReplacementAsDirectScalarDeltaRoot({
        cellIndex,
        oldNumber: oldFormulaNumber,
        changedTopology,
        postRecalcDirectFormulaIndices: request.postRecalcDirectFormulaIndices,
        postRecalcDirectFormulaMetrics: request.postRecalcDirectFormulaMetrics,
      })
    if (!handledFormulaReplacementAsDirectDelta) {
      formulaChangedCount = args.markFormulaChanged(cellIndex, formulaChangedCount)
    }
    topologyChanged = topologyChanged || changedTopology
    if (!priorHadFormula) {
      formulaChangedCount = request.refreshDependentRangesAndRebindFormulaDependents(cellIndex, formulaChangedCount)
      topologyChanged = true
    }
    const aggregateDependents = hasFormulaColumnAggregateDependents
      ? request
          .collectAffectedDirectRangeDependents({
            sheetName: op.sheetName,
            row: parsedAddress.row,
            col: parsedAddress.col,
          })
          .filter((candidate) => candidate !== cellIndex)
      : []
    if (aggregateDependents.length > 0) {
      formulaChangedCount = args.rebindFormulaCells(aggregateDependents, formulaChangedCount)
      for (let index = 0; index < aggregateDependents.length; index += 1) {
        request.postRecalcDirectFormulaIndices.add(aggregateDependents[index]!)
        formulaChangedCount = args.markFormulaChanged(aggregateDependents[index]!, formulaChangedCount)
        changedInputCount = args.markInputChanged(aggregateDependents[index]!, changedInputCount)
      }
      topologyChanged = true
    }
  } catch {
    if (!request.isRestore) {
      compileMs += performance.now() - compileStarted
    }
    topologyChanged = args.removeFormula(cellIndex) || topologyChanged
    args.setInvalidFormulaValue(cellIndex)
    changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
  }
  if (!request.isRestore) {
    explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
    request.setEntityVersionForOp(op, request.order)
  }
  return { changedInputCount, formulaChangedCount, explicitChangedCount, topologyChanged, refreshAllPivots, compileMs }
}

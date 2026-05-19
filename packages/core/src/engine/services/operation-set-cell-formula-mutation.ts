import { formatAddress } from '@bilig/formula'
import type { EngineCellMutationRef } from '../../cell-mutations-at.js'
import type { OpOrder } from '../../replica-state.js'
import type { DirectFormulaIndexCollection, DirectScalarCurrentOperand } from './direct-formula-index-collection.js'
import type { OperationTrackedColumnDependencyFlags } from './operation-column-dependency-tracker.js'
import type { OperationDirectRangeDependentService } from './operation-direct-range-dependents.js'
import {
  analyzeFreshDirectAggregateFormula,
  bindFreshTemplateFormula,
  markFreshDirectAggregateInputsCovered,
} from './operation-fresh-direct-aggregate.js'
import type { CreateEngineOperationServiceArgs } from './operation-service-types.js'

type SetCellFormulaMutation = Extract<EngineCellMutationRef['mutation'], { kind: 'setCellFormula' }>

interface SetCellFormulaMutationCounts {
  readonly changedInputCount: number
  readonly formulaChangedCount: number
  readonly explicitChangedCount: number
  readonly topologyChanged: boolean
  readonly compileMs: number
}

interface ApplySetCellFormulaMutationArgs extends SetCellFormulaMutationCounts {
  readonly serviceArgs: CreateEngineOperationServiceArgs
  readonly sheetId: number
  readonly sheetName: string
  readonly mutation: SetCellFormulaMutation
  readonly existingIndex: number | undefined
  readonly isRestore: boolean
  readonly trackExplicitChanges: boolean
  readonly order: OpOrder | undefined
  readonly dependencyFlags: OperationTrackedColumnDependencyFlags
  readonly postRecalcDirectFormulaIndices: DirectFormulaIndexCollection
  readonly batchMayNeedFreshAggregateInputCoverage: boolean
  readonly setCellEntityVersion: (sheetName: string, address: string, order: OpOrder) => void
  readonly readExactNumericValueForLookup: (cellIndex: number) => number | undefined
  readonly tryApplyFormulaReplacementAsDirectScalarDeltaRoot: (request: {
    readonly cellIndex: number
    readonly oldNumber: number | undefined
    readonly changedTopology: boolean
    readonly postRecalcDirectFormulaIndices: DirectFormulaIndexCollection
  }) => boolean
  readonly collectAffectedDirectRangeDependents: OperationDirectRangeDependentService['collectAffectedDirectRangeDependents']
  readonly clearTrackedColumnDependencyFlagCache: () => void
  readonly applyDirectFormulaCurrentResult: (cellIndex: number, result: DirectScalarCurrentOperand) => boolean
}

export function applySetCellFormulaMutation(request: ApplySetCellFormulaMutationArgs): SetCellFormulaMutationCounts {
  const args = request.serviceArgs
  const { existingIndex, isRestore, mutation, sheetId, sheetName } = request
  const { hasAggregateDependents, hasExactLookupDependents, hasSortedLookupDependents } = request.dependencyFlags
  let changedInputCount = request.changedInputCount
  let formulaChangedCount = request.formulaChangedCount
  let explicitChangedCount = request.explicitChangedCount
  let topologyChanged = request.topologyChanged
  let compileMs = request.compileMs

  if (hasExactLookupDependents) {
    args.invalidateExactLookupColumn({ sheetName, col: mutation.col })
  }
  if (hasSortedLookupDependents) {
    args.invalidateSortedLookupColumn({ sheetName, col: mutation.col })
  }
  if (!isRestore && existingIndex !== undefined) {
    changedInputCount = args.markPivotRootsChanged(args.clearPivotForCell(existingIndex), changedInputCount)
  }
  const cellIndex = args.state.workbook.ensureCellAt(sheetId, mutation.row, mutation.col).cellIndex
  if (!isRestore && existingIndex !== undefined) {
    changedInputCount = args.markSpillRootsChanged(args.clearOwnedSpill(cellIndex), changedInputCount)
  }
  const priorHadFormula = args.state.formulas.get(cellIndex) !== undefined
  const oldFormulaNumber = !isRestore && priorHadFormula ? request.readExactNumericValueForLookup(cellIndex) : undefined
  const compileStarted = isRestore ? 0 : performance.now()
  try {
    const priorDirectScalarFormula = args.state.formulas.get(cellIndex)?.directScalar !== undefined
    const rewriteFormulaSourcePreservingBinding = args.rewriteFormulaSourcePreservingBinding
    const canRewriteFormulaPreservingBinding =
      !isRestore &&
      priorDirectScalarFormula &&
      !hasExactLookupDependents &&
      !hasSortedLookupDependents &&
      !hasAggregateDependents &&
      rewriteFormulaSourcePreservingBinding !== undefined
    const canAssumeFreshFormula =
      !isRestore &&
      existingIndex === undefined &&
      !priorHadFormula &&
      args.state.workbook.metadata.definedNames.size === 0 &&
      args.bindPreparedFormula !== undefined &&
      args.compileTemplateFormula !== undefined
    const changedTopology = canRewriteFormulaPreservingBinding
      ? rewriteFormulaSourcePreservingBinding(cellIndex, sheetName, mutation.formula)
        ? false
        : args.bindFormula(cellIndex, sheetName, mutation.formula)
      : canAssumeFreshFormula
        ? bindFreshTemplateFormula(args, cellIndex, sheetName, mutation)
        : args.bindFormula(cellIndex, sheetName, mutation.formula)
    const runtimeFormula = args.state.formulas.get(cellIndex)
    if (hasAggregateDependents) {
      args.invalidateAggregateColumn({ sheetName, col: mutation.col })
    }
    if (!isRestore) {
      compileMs += performance.now() - compileStarted
    }
    request.clearTrackedColumnDependencyFlagCache()
    changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
    const freshDirectAggregateAnalysis = analyzeFreshDirectAggregateFormula(args, {
      priorHadFormula,
      formulaCellIndex: cellIndex,
      formula: runtimeFormula,
    })
    const canSkipTopoRepair = freshDirectAggregateAnalysis.canSkipTopoRepair
    const freshDirectFormulaResult = freshDirectAggregateAnalysis.currentResult
    const evaluatedFreshDirectFormula =
      freshDirectFormulaResult !== undefined
        ? (() => {
            request.postRecalcDirectFormulaIndices.addCurrentResult(cellIndex, freshDirectFormulaResult)
            const applied = request.applyDirectFormulaCurrentResult(cellIndex, freshDirectFormulaResult)
            if (applied && request.batchMayNeedFreshAggregateInputCoverage) {
              markFreshDirectAggregateInputsCovered(args, {
                formulaCellIndex: cellIndex,
                formula: runtimeFormula,
                postRecalcDirectFormulaIndices: request.postRecalcDirectFormulaIndices,
              })
            }
            return applied
          })()
        : canSkipTopoRepair && args.evaluateDirectFormula(cellIndex) !== undefined
    const handledFormulaReplacementAsDirectDelta =
      priorHadFormula &&
      !hasExactLookupDependents &&
      !hasSortedLookupDependents &&
      !hasAggregateDependents &&
      request.tryApplyFormulaReplacementAsDirectScalarDeltaRoot({
        cellIndex,
        oldNumber: oldFormulaNumber,
        changedTopology,
        postRecalcDirectFormulaIndices: request.postRecalcDirectFormulaIndices,
      })
    const replacementDirectAggregateResult =
      !handledFormulaReplacementAsDirectDelta && !evaluatedFreshDirectFormula && priorHadFormula
        ? analyzeFreshDirectAggregateFormula(args, {
            priorHadFormula: false,
            formulaCellIndex: cellIndex,
            formula: runtimeFormula,
          }).currentResult
        : undefined
    const evaluatedReplacementDirectAggregate =
      replacementDirectAggregateResult !== undefined
        ? (() => {
            request.postRecalcDirectFormulaIndices.addCurrentResult(cellIndex, replacementDirectAggregateResult)
            return request.applyDirectFormulaCurrentResult(cellIndex, replacementDirectAggregateResult)
          })()
        : false
    if (!handledFormulaReplacementAsDirectDelta && !evaluatedFreshDirectFormula && !evaluatedReplacementDirectAggregate) {
      formulaChangedCount = args.markFormulaChanged(cellIndex, formulaChangedCount)
    }
    topologyChanged = topologyChanged || (changedTopology && !canSkipTopoRepair)
    const aggregateDependents = hasAggregateDependents
      ? request
          .collectAffectedDirectRangeDependents({
            sheetName,
            row: mutation.row,
            col: mutation.col,
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
    if (!isRestore) {
      compileMs += performance.now() - compileStarted
    }
    const removedFormula = args.removeFormula(cellIndex)
    topologyChanged = removedFormula || topologyChanged
    request.clearTrackedColumnDependencyFlagCache()
    args.setInvalidFormulaValue(cellIndex)
    changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
  }
  if (request.trackExplicitChanges) {
    explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
  }
  if (!isRestore && args.state.trackReplicaVersions) {
    request.setCellEntityVersion(sheetName, formatAddress(mutation.row, mutation.col), request.order!)
  }
  return { changedInputCount, formulaChangedCount, explicitChangedCount, topologyChanged, compileMs }
}

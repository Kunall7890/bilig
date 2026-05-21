import type { CellValue } from '@bilig/protocol'
import { makeCellEntity } from '../../entity-ids.js'
import type { RuntimeDirectAggregateDescriptor, RuntimeDirectCriteriaDescriptor } from '../runtime-state.js'
import type { DirectFormulaIndexCollection } from './direct-formula-index-collection.js'
import {
  aggregateColumnDependencyKey,
  directAggregateNumericContribution,
  directCriteriaTouchesPoint,
} from './direct-formula-recalc-helpers.js'
import {
  collectIndexedDirectAggregateColumnDependentsForRow,
  visitIndexedDirectAggregateColumnDependentsForRow,
} from './formula-binding-dependency-helpers.js'

export interface OperationDirectRangePoint {
  readonly sheetName: string
  readonly row: number
  readonly col: number
  readonly sheetId?: number
}

export interface OperationDirectRangeDirtyPoint extends OperationDirectRangePoint {
  readonly oldValue?: CellValue
  readonly newValue?: CellValue
  readonly inputCellIndex?: number
}

interface DirectRangeFormulaRecord {
  readonly directAggregate: RuntimeDirectAggregateDescriptor | undefined
  readonly directCriteria: RuntimeDirectCriteriaDescriptor | undefined
  readonly dependencyIndices: { readonly length: number }
}

interface OperationDirectRangeFormulaAccess {
  get(cellIndex: number): DirectRangeFormulaRecord | undefined
}

function directAggregateTouchesPoint(directAggregate: RuntimeDirectAggregateDescriptor, request: OperationDirectRangePoint): boolean {
  return (
    directAggregate.sheetName === request.sheetName &&
    request.col >= directAggregate.col &&
    request.col <= directAggregate.colEnd &&
    request.row >= directAggregate.rowStart &&
    request.row <= directAggregate.rowEnd
  )
}

function formulaTouchesDirectRange(formula: DirectRangeFormulaRecord | undefined, request: OperationDirectRangePoint): boolean {
  const directAggregate = formula?.directAggregate
  if (directAggregate !== undefined && directAggregateTouchesPoint(directAggregate, request)) {
    return true
  }
  const directCriteria = formula?.directCriteria
  return directCriteria !== undefined && directCriteriaTouchesPoint(directCriteria, request)
}

function mergeSingleDirectRangeDependent(current: number, candidate: number): number {
  if (candidate < 0) {
    return candidate === -2 ? -2 : current
  }
  if (current === -1) {
    return candidate
  }
  return current === candidate ? current : -2
}

export interface OperationDirectRangeDependentService {
  readonly collectAffectedDirectRangeDependents: (request: OperationDirectRangePoint) => number[]
  readonly collectSingleAffectedDirectRangeDependent: (request: OperationDirectRangePoint) => number
  readonly canApplyDirectAggregateLiteralDelta: (formulaCellIndex: number) => boolean
  readonly canApplyDirectAggregateLiteralDeltaForRequest: (formulaCellIndex: number, request: OperationDirectRangePoint) => boolean
  readonly collectSingleApplicableDirectAggregateDependent: (request: OperationDirectRangePoint) => number
  readonly markAffectedDirectRangeDependents: (
    request: OperationDirectRangeDirtyPoint,
    formulaChangedCount: number,
    postRecalcDirectFormulaIndices?: DirectFormulaIndexCollection,
  ) => number
}

export function createOperationDirectRangeDependentService(args: {
  readonly workbook: {
    getSheet(sheetName: string): { readonly id: number } | undefined
  }
  readonly formulas: OperationDirectRangeFormulaAccess
  readonly reverseAggregateColumnEdges: Map<number, Set<number>>
  readonly collectRegionFormulaDependentsForCell: (sheetName: string, row: number, col: number) => Uint32Array
  readonly collectSingleRegionFormulaDependentForCell: (sheetName: string, row: number, col: number) => number
  readonly collectSingleRegionFormulaDependentForCellAt?: ((sheetId: number, row: number, col: number) => number) | undefined
  readonly hasNoCellDependents: (cellIndex: number) => boolean
  readonly getSingleEntityDependent: (entityId: number) => number
  readonly markFormulaChanged: (cellIndex: number, count: number) => number
  readonly tryDirectCriteriaSumDelta: (
    directCriteria: RuntimeDirectCriteriaDescriptor,
    request: OperationDirectRangeDirtyPoint,
  ) => number | undefined
  readonly postRecalcLimit: number
}): OperationDirectRangeDependentService {
  const collectAffectedDirectRangeDependents = (request: OperationDirectRangePoint): number[] => {
    const sheetId = args.workbook.getSheet(request.sheetName)?.id
    const dependents = args.collectRegionFormulaDependentsForCell(request.sheetName, request.row, request.col)
    const affected: number[] = []
    const seen = new Set<number>()
    const consider = (formulaCellIndex: number): void => {
      if (seen.has(formulaCellIndex)) {
        return
      }
      seen.add(formulaCellIndex)
      if (formulaTouchesDirectRange(args.formulas.get(formulaCellIndex), request)) {
        affected.push(formulaCellIndex)
      }
    }
    for (let dependentIndex = 0; dependentIndex < dependents.length; dependentIndex += 1) {
      consider(dependents[dependentIndex]!)
    }
    if (sheetId !== undefined) {
      const aggregateDependents = args.reverseAggregateColumnEdges.get(aggregateColumnDependencyKey(sheetId, request.col))
      if (aggregateDependents !== undefined) {
        const indexedDependents = collectIndexedDirectAggregateColumnDependentsForRow(aggregateDependents, request.row)
        const dependentsToVisit = indexedDependents ?? aggregateDependents
        dependentsToVisit.forEach(consider)
      }
    }
    return affected
  }

  const collectSingleAffectedDirectRangeDependent = (request: OperationDirectRangePoint): number => {
    const regionDependent =
      request.sheetId !== undefined && args.collectSingleRegionFormulaDependentForCellAt
        ? args.collectSingleRegionFormulaDependentForCellAt(request.sheetId, request.row, request.col)
        : args.collectSingleRegionFormulaDependentForCell(request.sheetName, request.row, request.col)
    if (regionDependent === -2) {
      return regionDependent
    }
    let singleDependent = -1
    if (regionDependent >= 0 && formulaTouchesDirectRange(args.formulas.get(regionDependent), request)) {
      singleDependent = regionDependent
    }
    const sheetId = request.sheetId ?? args.workbook.getSheet(request.sheetName)?.id
    if (sheetId === undefined) {
      return singleDependent
    }
    const aggregateDependents = args.reverseAggregateColumnEdges.get(aggregateColumnDependencyKey(sheetId, request.col))
    if (!aggregateDependents || aggregateDependents.size === 0) {
      return singleDependent
    }
    for (const candidate of aggregateDependents) {
      if (!formulaTouchesDirectRange(args.formulas.get(candidate), request)) {
        continue
      }
      singleDependent = mergeSingleDirectRangeDependent(singleDependent, candidate)
      if (singleDependent === -2) {
        return -2
      }
    }
    return singleDependent
  }

  const canApplyDirectAggregateLiteralDelta = (formulaCellIndex: number): boolean => {
    const formula = args.formulas.get(formulaCellIndex)
    return (
      formula !== undefined &&
      formula.directAggregate?.aggregateKind === 'sum' &&
      formula.dependencyIndices.length === 0 &&
      args.hasNoCellDependents(formulaCellIndex)
    )
  }

  const canApplyDirectAggregateLiteralDeltaForRequest = (formulaCellIndex: number, request: OperationDirectRangePoint): boolean => {
    const formula = args.formulas.get(formulaCellIndex)
    const directAggregate = formula?.directAggregate
    return (
      directAggregate !== undefined &&
      canApplyDirectAggregateLiteralDelta(formulaCellIndex) &&
      directAggregateTouchesPoint(directAggregate, request)
    )
  }

  const collectSingleApplicableDirectAggregateDependent = (request: OperationDirectRangePoint): number => {
    const regionDependent =
      request.sheetId !== undefined && args.collectSingleRegionFormulaDependentForCellAt
        ? args.collectSingleRegionFormulaDependentForCellAt(request.sheetId, request.row, request.col)
        : args.collectSingleRegionFormulaDependentForCell(request.sheetName, request.row, request.col)
    if (regionDependent === -2) {
      return regionDependent
    }
    let singleAggregateDependent = -1
    if (regionDependent >= 0 && formulaTouchesDirectRange(args.formulas.get(regionDependent), request)) {
      if (!canApplyDirectAggregateLiteralDeltaForRequest(regionDependent, request)) {
        return -2
      }
      singleAggregateDependent = regionDependent
    }
    const sheetId = request.sheetId ?? args.workbook.getSheet(request.sheetName)?.id
    if (sheetId === undefined) {
      return singleAggregateDependent
    }
    const aggregateDependents = args.reverseAggregateColumnEdges.get(aggregateColumnDependencyKey(sheetId, request.col))
    if (!aggregateDependents || aggregateDependents.size === 0) {
      return singleAggregateDependent
    }
    const usedIndexedDependents = visitIndexedDirectAggregateColumnDependentsForRow(aggregateDependents, request.row, (candidate) => {
      const formula = args.formulas.get(candidate)
      if (!formulaTouchesDirectRange(formula, request)) {
        return true
      }
      if (!canApplyDirectAggregateLiteralDeltaForRequest(candidate, request)) {
        singleAggregateDependent = -2
        return false
      }
      singleAggregateDependent = mergeSingleDirectRangeDependent(singleAggregateDependent, candidate)
      return singleAggregateDependent !== -2
    })
    if (usedIndexedDependents) {
      return singleAggregateDependent
    }
    for (const candidate of aggregateDependents) {
      const formula = args.formulas.get(candidate)
      if (!formulaTouchesDirectRange(formula, request)) {
        continue
      }
      if (!canApplyDirectAggregateLiteralDeltaForRequest(candidate, request)) {
        return -2
      }
      singleAggregateDependent = mergeSingleDirectRangeDependent(singleAggregateDependent, candidate)
      if (singleAggregateDependent === -2) {
        return -2
      }
    }
    return singleAggregateDependent
  }

  const markAffectedDirectRangeDependents = (
    request: OperationDirectRangeDirtyPoint,
    formulaChangedCount: number,
    postRecalcDirectFormulaIndices?: DirectFormulaIndexCollection,
  ): number => {
    const singleAffected = collectSingleAffectedDirectRangeDependent(request)
    const oldContribution = request.oldValue ? directAggregateNumericContribution(request.oldValue) : undefined
    const newContribution = request.newValue ? directAggregateNumericContribution(request.newValue) : undefined
    const contributionDelta = oldContribution === undefined || newContribution === undefined ? undefined : newContribution - oldContribution
    if (singleAffected === -1) {
      return formulaChangedCount
    }
    if (singleAffected >= 0) {
      const formula = args.formulas.get(singleAffected)
      const canUsePostRecalc =
        postRecalcDirectFormulaIndices !== undefined &&
        (formula?.directAggregate !== undefined || formula?.directCriteria !== undefined) &&
        args.getSingleEntityDependent(makeCellEntity(singleAffected)) === -1
      if (canUsePostRecalc) {
        if (
          contributionDelta !== undefined &&
          formula?.directAggregate?.aggregateKind === 'sum' &&
          formula.dependencyIndices.length === 0
        ) {
          postRecalcDirectFormulaIndices.addDelta(singleAffected, contributionDelta)
        } else if (formula?.directCriteria !== undefined) {
          const criteriaDelta = args.tryDirectCriteriaSumDelta(formula.directCriteria, request)
          if (criteriaDelta !== undefined) {
            postRecalcDirectFormulaIndices.addDelta(singleAffected, criteriaDelta)
          } else {
            postRecalcDirectFormulaIndices.add(singleAffected)
          }
        } else {
          postRecalcDirectFormulaIndices.add(singleAffected)
        }
        if (request.inputCellIndex !== undefined) {
          postRecalcDirectFormulaIndices.markDirectRangeInputCovered(request.inputCellIndex)
        }
        return formulaChangedCount
      }
      if (postRecalcDirectFormulaIndices && (formula?.dependencyIndices.length ?? 0) > 0) {
        postRecalcDirectFormulaIndices.add(singleAffected)
      }
      return args.markFormulaChanged(singleAffected, formulaChangedCount)
    }
    const dependents = collectAffectedDirectRangeDependents(request)
    const canUsePostRecalc =
      postRecalcDirectFormulaIndices !== undefined &&
      dependents.length > 0 &&
      dependents.length <= args.postRecalcLimit &&
      dependents.every((formulaCellIndex) => {
        const formula = args.formulas.get(formulaCellIndex)
        return (
          (formula?.directAggregate !== undefined || formula?.directCriteria !== undefined) &&
          args.getSingleEntityDependent(makeCellEntity(formulaCellIndex)) === -1
        )
      })
    const canUseDeltaPostRecalc =
      canUsePostRecalc &&
      contributionDelta !== undefined &&
      dependents.every((formulaCellIndex) => {
        const formula = args.formulas.get(formulaCellIndex)
        return formula?.directAggregate?.aggregateKind === 'sum' && formula.dependencyIndices.length === 0
      })
    let canUseCriteriaDeltaPostRecalc = false
    let criteriaDelta: number | undefined
    if (canUsePostRecalc && !canUseDeltaPostRecalc) {
      canUseCriteriaDeltaPostRecalc = true
      for (let index = 0; index < dependents.length; index += 1) {
        const formula = args.formulas.get(dependents[index]!)
        if (formula?.directCriteria === undefined) {
          canUseCriteriaDeltaPostRecalc = false
          break
        }
        const nextDelta = args.tryDirectCriteriaSumDelta(formula.directCriteria, request)
        if (nextDelta === undefined) {
          canUseCriteriaDeltaPostRecalc = false
          break
        }
        if (criteriaDelta === undefined) {
          criteriaDelta = nextDelta
        } else if (!Object.is(criteriaDelta, nextDelta)) {
          canUseCriteriaDeltaPostRecalc = false
          break
        }
      }
    }
    if (canUsePostRecalc && canUseDeltaPostRecalc) {
      postRecalcDirectFormulaIndices.appendConstantDelta(dependents, contributionDelta)
      if (request.inputCellIndex !== undefined) {
        postRecalcDirectFormulaIndices.markDirectRangeInputCovered(request.inputCellIndex)
      }
      return formulaChangedCount
    }
    if (canUsePostRecalc && canUseCriteriaDeltaPostRecalc && criteriaDelta !== undefined) {
      postRecalcDirectFormulaIndices.appendConstantDelta(dependents, criteriaDelta)
      if (request.inputCellIndex !== undefined) {
        postRecalcDirectFormulaIndices.markDirectRangeInputCovered(request.inputCellIndex)
      }
      return formulaChangedCount
    }
    for (let index = 0; index < dependents.length; index += 1) {
      const formulaCellIndex = dependents[index]!
      if (canUsePostRecalc) {
        postRecalcDirectFormulaIndices.add(formulaCellIndex)
        continue
      }
      if (postRecalcDirectFormulaIndices && (args.formulas.get(formulaCellIndex)?.dependencyIndices.length ?? 0) > 0) {
        postRecalcDirectFormulaIndices.add(formulaCellIndex)
      }
      formulaChangedCount = args.markFormulaChanged(formulaCellIndex, formulaChangedCount)
    }
    if (canUsePostRecalc && request.inputCellIndex !== undefined) {
      postRecalcDirectFormulaIndices.markDirectRangeInputCovered(request.inputCellIndex)
    }
    return formulaChangedCount
  }

  return {
    collectAffectedDirectRangeDependents,
    collectSingleAffectedDirectRangeDependent,
    canApplyDirectAggregateLiteralDelta,
    canApplyDirectAggregateLiteralDeltaForRequest,
    collectSingleApplicableDirectAggregateDependent,
    markAffectedDirectRangeDependents,
  }
}

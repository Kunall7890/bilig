import type { LiteralInput, RecalcMetrics } from '@bilig/protocol'
import type { EngineExistingNumericCellMutationResult } from '../../cell-mutations-at.js'
import type { SheetRecord } from '../../workbook-store.js'
import { composeSingleDisjointExplicitEventChanges } from './direct-formula-recalc-helpers.js'
import {
  makeCompactExistingNumericMutationResult,
  makeExistingNumericMutationResult,
  tagTrustedPhysicalTrackedChanges,
} from './operation-change-helpers.js'
import { emitOperationTrackedCellsBatch } from './operation-tracked-event-helpers.js'
import type { CreateEngineOperationServiceArgs } from './operation-service-types.js'

export interface OperationDirectAggregateLiteralMutationRequest {
  readonly existingIndex: number
  readonly sheetId?: number
  readonly sheetName: string
  readonly row: number
  readonly col: number
  readonly value: LiteralInput
  readonly delta: number
  readonly emitTracked: boolean
  readonly singleRangeEntityDependent?: number
}

export interface OperationTrustedRangeDirectAggregateExistingNumericMutationRequest {
  readonly existingIndex: number
  readonly rangeEntityDependent: number
  readonly sheet: SheetRecord
  readonly sheetId: number
  readonly col: number
  readonly value: number
  readonly delta: number
  readonly hasExactLookupDependents: boolean
  readonly hasSortedLookupDependents: boolean
}

export interface OperationTrustedColumnDirectAggregateExistingNumericMutationRequest {
  readonly existingIndex: number
  readonly sheet: SheetRecord
  readonly sheetId: number
  readonly sheetName: string
  readonly row: number
  readonly col: number
  readonly value: number
  readonly delta: number
  readonly hasExactLookupDependents: boolean
  readonly hasSortedLookupDependents: boolean
}

export interface OperationDirectAggregateLiteralFastPathArgs {
  readonly state: Pick<CreateEngineOperationServiceArgs['state'], 'workbook' | 'counters' | 'events' | 'setLastMetrics'>
  readonly directRangePostRecalcLimit: number
  readonly getSingleEntityDependent: (entityId: number) => number
  readonly collectAffectedDirectRangeDependents: (request: {
    readonly sheetName: string
    readonly row: number
    readonly col: number
  }) => readonly number[]
  readonly collectSingleApplicableDirectAggregateDependent: (request: {
    readonly sheetName: string
    readonly sheetId?: number
    readonly row: number
    readonly col: number
  }) => number
  readonly canApplyDirectAggregateLiteralDeltaForRequest: (
    formulaCellIndex: number,
    request: {
      readonly sheetName: string
      readonly sheetId?: number
      readonly row: number
      readonly col: number
    },
  ) => boolean
  readonly canApplyDirectAggregateLiteralDelta: (formulaCellIndex: number) => boolean
  readonly writeFastPathLiteralToExistingCell: (existingIndex: number, value: LiteralInput) => void
  readonly writeTrustedExistingNumericLiteralToCell: (existingIndex: number, sheet: SheetRecord, col: number, value: number) => void
  readonly applyTerminalDirectFormulaNumericDeltaAndReturn: (formulaCellIndex: number, delta: number) => number | undefined
  readonly applyDirectFormulaNumericDelta: (formulaCellIndex: number, delta: number) => boolean
  readonly applyDirectFormulaNumericDeltaBatch: (formulaCellIndices: readonly number[] | Uint32Array, delta: number) => boolean
  readonly cellsShareVersionColumn: (leftCellIndex: number, rightCellIndex: number) => boolean
  readonly withOptionalColumnVersionBatch: (enabled: boolean, apply: () => void) => void
  readonly deferSingleCellKernelSync: (cellIndex: number) => void
  readonly makeSingleLiteralSkipMetrics: () => RecalcMetrics
}

export function tryApplySingleDirectAggregateLiteralMutationFastPath(
  args: OperationDirectAggregateLiteralFastPathArgs,
  request: OperationDirectAggregateLiteralMutationRequest,
): EngineExistingNumericCellMutationResult | null {
  let singleAffected = -2
  if (request.singleRangeEntityDependent !== undefined) {
    const rangeDependent = args.getSingleEntityDependent(request.singleRangeEntityDependent)
    if (rangeDependent < -1) {
      return null
    }
    if (rangeDependent >= 0) {
      singleAffected = rangeDependent
    }
  }
  if (singleAffected >= 0 && !args.canApplyDirectAggregateLiteralDeltaForRequest(singleAffected, request)) {
    return null
  }
  if (singleAffected < -1) {
    singleAffected = args.collectSingleApplicableDirectAggregateDependent({
      sheetName: request.sheetName,
      ...(request.sheetId === undefined ? {} : { sheetId: request.sheetId }),
      row: request.row,
      col: request.col,
    })
  }
  if (singleAffected === -1) {
    args.writeFastPathLiteralToExistingCell(request.existingIndex, request.value)
    args.deferSingleCellKernelSync(request.existingIndex)
    const lastMetrics = args.makeSingleLiteralSkipMetrics()
    args.state.setLastMetrics(lastMetrics)
    if (request.emitTracked) {
      const changed = Uint32Array.of(request.existingIndex)
      emitOperationTrackedCellsBatch({
        events: args.state.events,
        changedCellIndices: changed,
        metrics: lastMetrics,
      })
      return makeExistingNumericMutationResult(changed, 1)
    }
    return makeCompactExistingNumericMutationResult(request.existingIndex, undefined, 1)
  }

  let singleAggregateCellIndex = -1
  let affected: readonly number[] | undefined
  if (singleAffected >= 0) {
    singleAggregateCellIndex = singleAffected
  } else {
    const collected = args.collectAffectedDirectRangeDependents({
      sheetName: request.sheetName,
      row: request.row,
      col: request.col,
    })
    if (collected.length === 0 || collected.length > args.directRangePostRecalcLimit) {
      return null
    }
    for (let index = 0; index < collected.length; index += 1) {
      if (!args.canApplyDirectAggregateLiteralDelta(collected[index]!)) {
        return null
      }
    }
    affected = collected
  }
  const affectedCount = singleAggregateCellIndex >= 0 ? 1 : (affected?.length ?? 0)
  const sharesSingleAggregateVersionColumn =
    singleAggregateCellIndex >= 0 &&
    request.sheetId !== undefined &&
    args.state.workbook.cellStore.sheetIds[singleAggregateCellIndex] === request.sheetId &&
    args.state.workbook.cellStore.cols[singleAggregateCellIndex] === request.col
  const shouldBatchColumnVersions =
    affectedCount > 1 ||
    (singleAggregateCellIndex >= 0 &&
      (request.sheetId === undefined
        ? args.cellsShareVersionColumn(request.existingIndex, singleAggregateCellIndex)
        : sharesSingleAggregateVersionColumn))

  let singleAggregateNumericValue: number | undefined
  if (!shouldBatchColumnVersions && singleAggregateCellIndex >= 0) {
    args.writeFastPathLiteralToExistingCell(request.existingIndex, request.value)
    singleAggregateNumericValue = args.applyTerminalDirectFormulaNumericDeltaAndReturn(singleAggregateCellIndex, request.delta)
    if (singleAggregateNumericValue === undefined) {
      throw new Error('Failed to apply direct aggregate delta')
    }
  } else {
    args.withOptionalColumnVersionBatch(shouldBatchColumnVersions, () => {
      args.writeFastPathLiteralToExistingCell(request.existingIndex, request.value)
      if (singleAggregateCellIndex >= 0) {
        if (!args.applyDirectFormulaNumericDelta(singleAggregateCellIndex, request.delta)) {
          throw new Error('Failed to apply direct aggregate delta')
        }
      } else {
        if (!args.applyDirectFormulaNumericDeltaBatch(affected!, request.delta)) {
          throw new Error('Failed to apply direct aggregate delta')
        }
      }
    })
  }
  args.state.counters.directAggregateDeltaApplications += affectedCount
  args.state.counters.directAggregateDeltaOnlyRecalcSkips += 1
  args.deferSingleCellKernelSync(request.existingIndex)
  const lastMetrics = args.makeSingleLiteralSkipMetrics()
  args.state.setLastMetrics(lastMetrics)
  if (request.emitTracked) {
    const changed =
      singleAggregateCellIndex >= 0
        ? Uint32Array.of(request.existingIndex, singleAggregateCellIndex)
        : affectedCount === 0
          ? Uint32Array.of(request.existingIndex)
          : composeSingleDisjointExplicitEventChanges(request.existingIndex, Uint32Array.from(affected!))
    if (singleAggregateCellIndex >= 0 && changed.length > 4 && request.sheetId !== undefined) {
      tagTrustedPhysicalTrackedChanges(changed, request.sheetId, 1)
    }
    emitOperationTrackedCellsBatch({
      events: args.state.events,
      changedCellIndices: changed,
      metrics: lastMetrics,
    })
    return makeExistingNumericMutationResult(changed, 1)
  }
  if (singleAggregateCellIndex >= 0) {
    const cellStore = args.state.workbook.cellStore
    return makeCompactExistingNumericMutationResult(
      request.existingIndex,
      singleAggregateCellIndex,
      1,
      singleAggregateNumericValue,
      cellStore.rows[singleAggregateCellIndex] ?? 0,
      cellStore.cols[singleAggregateCellIndex] ?? 0,
    )
  }
  if (affectedCount === 0) {
    return makeCompactExistingNumericMutationResult(request.existingIndex, undefined, 1)
  }
  const changed = composeSingleDisjointExplicitEventChanges(request.existingIndex, Uint32Array.from(affected!))
  if (singleAggregateCellIndex >= 0 && changed.length > 4 && request.sheetId !== undefined) {
    tagTrustedPhysicalTrackedChanges(changed, request.sheetId, 1)
  }
  return makeExistingNumericMutationResult(changed, 1)
}

export function tryApplyTrustedSingleRangeDirectAggregateExistingNumericMutation(
  args: OperationDirectAggregateLiteralFastPathArgs,
  request: OperationTrustedRangeDirectAggregateExistingNumericMutationRequest,
): EngineExistingNumericCellMutationResult | null {
  if (request.hasExactLookupDependents || request.hasSortedLookupDependents) {
    return null
  }
  const formulaCellIndex = args.getSingleEntityDependent(request.rangeEntityDependent)
  if (formulaCellIndex < 0 || !args.canApplyDirectAggregateLiteralDelta(formulaCellIndex)) {
    return null
  }
  args.writeTrustedExistingNumericLiteralToCell(request.existingIndex, request.sheet, request.col, request.value)
  const aggregateNumericValue = args.applyTerminalDirectFormulaNumericDeltaAndReturn(formulaCellIndex, request.delta)
  if (aggregateNumericValue === undefined) {
    throw new Error('Failed to apply direct aggregate delta')
  }
  args.state.counters.directAggregateDeltaApplications += 1
  args.state.counters.directAggregateDeltaOnlyRecalcSkips += 1
  args.deferSingleCellKernelSync(request.existingIndex)
  args.state.setLastMetrics(args.makeSingleLiteralSkipMetrics())
  const cellStore = args.state.workbook.cellStore
  return makeCompactExistingNumericMutationResult(
    request.existingIndex,
    formulaCellIndex,
    1,
    aggregateNumericValue,
    cellStore.rows[formulaCellIndex] ?? 0,
    cellStore.cols[formulaCellIndex] ?? 0,
  )
}

export function tryApplyTrustedColumnDirectAggregateExistingNumericMutation(
  args: OperationDirectAggregateLiteralFastPathArgs,
  request: OperationTrustedColumnDirectAggregateExistingNumericMutationRequest,
): EngineExistingNumericCellMutationResult | null {
  if (request.hasExactLookupDependents || request.hasSortedLookupDependents) {
    return null
  }
  const formulaCellIndex = args.collectSingleApplicableDirectAggregateDependent({
    sheetName: request.sheetName,
    sheetId: request.sheetId,
    row: request.row,
    col: request.col,
  })
  if (formulaCellIndex < 0) {
    return null
  }
  args.writeTrustedExistingNumericLiteralToCell(request.existingIndex, request.sheet, request.col, request.value)
  const aggregateNumericValue = args.applyTerminalDirectFormulaNumericDeltaAndReturn(formulaCellIndex, request.delta)
  if (aggregateNumericValue === undefined) {
    throw new Error('Failed to apply direct aggregate delta')
  }
  args.state.counters.directAggregateDeltaApplications += 1
  args.state.counters.directAggregateDeltaOnlyRecalcSkips += 1
  args.deferSingleCellKernelSync(request.existingIndex)
  args.state.setLastMetrics(args.makeSingleLiteralSkipMetrics())
  const cellStore = args.state.workbook.cellStore
  return makeCompactExistingNumericMutationResult(
    request.existingIndex,
    formulaCellIndex,
    1,
    aggregateNumericValue,
    cellStore.rows[formulaCellIndex] ?? 0,
    cellStore.cols[formulaCellIndex] ?? 0,
  )
}

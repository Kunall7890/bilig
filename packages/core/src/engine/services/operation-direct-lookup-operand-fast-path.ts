import { ErrorCode, ValueTag, type LiteralInput, type RecalcMetrics } from '@bilig/protocol'
import type { EngineExistingNumericCellMutationResult } from '../../cell-mutations-at.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type { SheetRecord } from '../../workbook-store.js'
import type { RuntimeDirectLookupDescriptor } from '../runtime-state.js'
import { makeCompactExistingNumericMutationResult } from './operation-change-helpers.js'
import type { DirectScalarCurrentOperand } from './direct-formula-index-collection.js'
import { emitOperationTrackedCellsBatch } from './operation-tracked-event-helpers.js'
import type { CreateEngineOperationServiceArgs } from './operation-service-types.js'

export interface OperationDirectLookupOperandMutationRequest {
  readonly existingIndex: number
  readonly formulaCellIndex: number
  readonly value: LiteralInput
  readonly exactLookupValue: number | undefined
  readonly approximateLookupValue: number | undefined
  readonly emitTracked: boolean
  readonly lookupSheetHint?: SheetRecord | undefined
  readonly trustedInputSheet?: SheetRecord | undefined
  readonly trustedInputCol?: number | undefined
}

export interface OperationDirectLookupOperandFastPathArgs {
  readonly state: Pick<CreateEngineOperationServiceArgs['state'], 'workbook' | 'formulas' | 'counters' | 'events' | 'setLastMetrics'>
  readonly hasNoCellDependents: (cellIndex: number) => boolean
  readonly directScalarNumericResultMatchesCell: (cellIndex: number, result: number) => boolean
  readonly directScalarCurrentResultMatchesCell: (cellIndex: number, result: DirectScalarCurrentOperand) => boolean
  readonly tryDirectUniformLookupNumericResultFromDescriptor: (
    directLookup: RuntimeDirectLookupDescriptor | undefined,
    exactLookupValue: number | undefined,
    approximateLookupValue: number | undefined,
    lookupSheetHint?: SheetRecord,
  ) => number | undefined
  readonly tryDirectApproximateLookupCurrentResultFromNumeric: (
    directLookup: Extract<RuntimeDirectLookupDescriptor, { kind: 'approximate' }>,
    lookupValue: number,
  ) => DirectScalarCurrentOperand | undefined
  readonly tryDirectUniformLookupCurrentResultFromNumeric: (
    formulaCellIndex: number,
    exactLookupValue: number | undefined,
    approximateLookupValue: number | undefined,
    lookupSheetHint?: SheetRecord,
  ) => DirectScalarCurrentOperand | undefined
  readonly writeTrustedExistingNumericLiteralToCell: (existingIndex: number, sheet: SheetRecord, col: number, value: number) => void
  readonly writeNumericLiteralToExistingCell: (existingIndex: number, value: number) => void
  readonly writeFastPathLiteralToExistingCell: (existingIndex: number, value: LiteralInput) => void
  readonly applyTerminalDirectFormulaNumericResult: (formulaCellIndex: number, result: number) => void
  readonly applyDirectFormulaCurrentResult: (formulaCellIndex: number, result: DirectScalarCurrentOperand) => boolean
  readonly cellsShareVersionColumn: (leftCellIndex: number, rightCellIndex: number) => boolean
  readonly withOptionalColumnVersionBatch: (enabled: boolean, apply: () => void) => void
  readonly deferSingleCellKernelSync: (cellIndex: number) => void
  readonly makeSingleLiteralSkipMetrics: () => RecalcMetrics
  readonly evaluateDirectFormula: (formulaCellIndex: number) => void
}

export function tryApplySingleDirectLookupOperandMutationFastPath(
  args: OperationDirectLookupOperandFastPathArgs,
  request: OperationDirectLookupOperandMutationRequest,
): EngineExistingNumericCellMutationResult | null {
  const formulaCellIndex = request.formulaCellIndex
  if (formulaCellIndex < 0 || !args.hasNoCellDependents(formulaCellIndex)) {
    return null
  }
  const formula = args.state.formulas.get(formulaCellIndex)
  const directLookup = formula?.directLookup
  const compactDirectLookupResult = (
    resultChanged: boolean,
    secondChangedNumericValue?: number,
  ): EngineExistingNumericCellMutationResult => {
    if (!resultChanged) {
      return makeCompactExistingNumericMutationResult(request.existingIndex, undefined, 1)
    }
    const cellStore = args.state.workbook.cellStore
    return makeCompactExistingNumericMutationResult(request.existingIndex, formulaCellIndex, 1, secondChangedNumericValue, {
      row: cellStore.rows[formulaCellIndex] ?? 0,
      col: cellStore.cols[formulaCellIndex] ?? 0,
    })
  }
  const numericResult = args.tryDirectUniformLookupNumericResultFromDescriptor(
    directLookup,
    request.exactLookupValue,
    request.approximateLookupValue,
    request.lookupSheetHint,
  )
  if (numericResult !== undefined) {
    const resultChanged = !args.directScalarNumericResultMatchesCell(formulaCellIndex, numericResult)
    const writeInput = (): void => {
      if (typeof request.value === 'number' && request.trustedInputSheet !== undefined && request.trustedInputCol !== undefined) {
        args.writeTrustedExistingNumericLiteralToCell(
          request.existingIndex,
          request.trustedInputSheet,
          request.trustedInputCol,
          request.value,
        )
      } else if (typeof request.value === 'number') {
        args.writeNumericLiteralToExistingCell(request.existingIndex, request.value)
      } else {
        args.writeFastPathLiteralToExistingCell(request.existingIndex, request.value)
      }
    }
    const apply = (): void => {
      writeInput()
      if (resultChanged) {
        args.applyTerminalDirectFormulaNumericResult(formulaCellIndex, numericResult)
      }
    }
    if (resultChanged && args.cellsShareVersionColumn(request.existingIndex, formulaCellIndex)) {
      args.withOptionalColumnVersionBatch(true, apply)
    } else {
      apply()
    }
    addEngineCounter(args.state.counters, resultChanged ? 'directFormulaKernelSyncOnlyRecalcSkips' : 'kernelSyncOnlyRecalcSkips')
    args.deferSingleCellKernelSync(request.existingIndex)
    const lastMetrics = args.makeSingleLiteralSkipMetrics()
    args.state.setLastMetrics(lastMetrics)
    if (request.emitTracked) {
      const changedCellIndices = resultChanged
        ? Uint32Array.of(request.existingIndex, formulaCellIndex)
        : Uint32Array.of(request.existingIndex)
      emitOperationTrackedCellsBatch({
        events: args.state.events,
        changedCellIndices,
        metrics: lastMetrics,
      })
    }
    return compactDirectLookupResult(resultChanged, resultChanged ? numericResult : undefined)
  }
  if (
    typeof request.value === 'number' &&
    directLookup?.kind === 'approximate' &&
    directLookup.operandCellIndex === request.existingIndex &&
    request.approximateLookupValue !== undefined
  ) {
    const numericValue = request.value
    const result = args.tryDirectApproximateLookupCurrentResultFromNumeric(directLookup, request.approximateLookupValue)
    if (result !== undefined) {
      const approximateNumericResult = result.kind === 'number' ? result.value : undefined
      const resultChanged =
        approximateNumericResult === undefined
          ? !args.directScalarCurrentResultMatchesCell(formulaCellIndex, result)
          : !args.directScalarNumericResultMatchesCell(formulaCellIndex, approximateNumericResult)
      const writeInput = (): void => {
        if (request.trustedInputSheet !== undefined && request.trustedInputCol !== undefined) {
          args.writeTrustedExistingNumericLiteralToCell(
            request.existingIndex,
            request.trustedInputSheet,
            request.trustedInputCol,
            numericValue,
          )
        } else {
          args.writeNumericLiteralToExistingCell(request.existingIndex, numericValue)
        }
      }
      const apply = (): void => {
        writeInput()
        if (resultChanged) {
          if (approximateNumericResult !== undefined) {
            args.applyTerminalDirectFormulaNumericResult(formulaCellIndex, approximateNumericResult)
          } else if (!args.applyDirectFormulaCurrentResult(formulaCellIndex, result)) {
            throw new Error('Failed to apply direct lookup result')
          }
        }
      }
      if (resultChanged && args.cellsShareVersionColumn(request.existingIndex, formulaCellIndex)) {
        args.withOptionalColumnVersionBatch(true, apply)
      } else {
        apply()
      }
      addEngineCounter(args.state.counters, resultChanged ? 'directFormulaKernelSyncOnlyRecalcSkips' : 'kernelSyncOnlyRecalcSkips')
      args.deferSingleCellKernelSync(request.existingIndex)
      const lastMetrics = args.makeSingleLiteralSkipMetrics()
      args.state.setLastMetrics(lastMetrics)
      if (request.emitTracked) {
        const changedCellIndices = resultChanged
          ? Uint32Array.of(request.existingIndex, formulaCellIndex)
          : Uint32Array.of(request.existingIndex)
        emitOperationTrackedCellsBatch({
          events: args.state.events,
          changedCellIndices,
          metrics: lastMetrics,
        })
      }
      return compactDirectLookupResult(resultChanged, approximateNumericResult)
    }
  }
  if (
    typeof request.value === 'number' &&
    directLookup?.operandCellIndex === request.existingIndex &&
    ((directLookup.kind === 'exact' && request.exactLookupValue !== undefined) ||
      (directLookup.kind === 'approximate' && request.approximateLookupValue !== undefined))
  ) {
    const cellStore = args.state.workbook.cellStore
    const beforeTag = cellStore.tags[formulaCellIndex]
    const beforeNumber = cellStore.numbers[formulaCellIndex] ?? 0
    const beforeStringId = cellStore.stringIds[formulaCellIndex] ?? 0
    const beforeError = cellStore.errors[formulaCellIndex] ?? ErrorCode.None
    const numericValue = request.value
    const apply = (): void => {
      if (request.trustedInputSheet !== undefined && request.trustedInputCol !== undefined) {
        args.writeTrustedExistingNumericLiteralToCell(
          request.existingIndex,
          request.trustedInputSheet,
          request.trustedInputCol,
          numericValue,
        )
      } else {
        args.writeNumericLiteralToExistingCell(request.existingIndex, numericValue)
      }
      args.evaluateDirectFormula(formulaCellIndex)
    }
    args.withOptionalColumnVersionBatch(args.cellsShareVersionColumn(request.existingIndex, formulaCellIndex), apply)
    const afterTag = cellStore.tags[formulaCellIndex]
    const afterNumber = cellStore.numbers[formulaCellIndex] ?? 0
    const afterStringId = cellStore.stringIds[formulaCellIndex] ?? 0
    const afterError = cellStore.errors[formulaCellIndex] ?? ErrorCode.None
    const resultChanged =
      beforeTag !== afterTag ||
      (afterTag === ValueTag.Number && !Object.is(beforeNumber, afterNumber)) ||
      (afterTag === ValueTag.Boolean && beforeNumber !== afterNumber) ||
      (afterTag === ValueTag.String && beforeStringId !== afterStringId) ||
      (afterTag === ValueTag.Error && beforeError !== afterError)
    addEngineCounter(args.state.counters, resultChanged ? 'directFormulaKernelSyncOnlyRecalcSkips' : 'kernelSyncOnlyRecalcSkips')
    args.deferSingleCellKernelSync(request.existingIndex)
    const lastMetrics = args.makeSingleLiteralSkipMetrics()
    args.state.setLastMetrics(lastMetrics)
    if (request.emitTracked) {
      const changedCellIndices = resultChanged
        ? Uint32Array.of(request.existingIndex, formulaCellIndex)
        : Uint32Array.of(request.existingIndex)
      emitOperationTrackedCellsBatch({
        events: args.state.events,
        changedCellIndices,
        metrics: lastMetrics,
      })
    }
    return compactDirectLookupResult(resultChanged, resultChanged && afterTag === ValueTag.Number ? afterNumber : undefined)
  }
  const result = args.tryDirectUniformLookupCurrentResultFromNumeric(
    formulaCellIndex,
    request.exactLookupValue,
    request.approximateLookupValue,
    request.lookupSheetHint,
  )
  if (result === undefined) {
    return null
  }
  const resultChanged = !args.directScalarCurrentResultMatchesCell(formulaCellIndex, result)
  const writeInput = (): void => {
    if (typeof request.value === 'number' && request.trustedInputSheet !== undefined && request.trustedInputCol !== undefined) {
      args.writeTrustedExistingNumericLiteralToCell(
        request.existingIndex,
        request.trustedInputSheet,
        request.trustedInputCol,
        request.value,
      )
    } else if (typeof request.value === 'number') {
      args.writeNumericLiteralToExistingCell(request.existingIndex, request.value)
    } else {
      args.writeFastPathLiteralToExistingCell(request.existingIndex, request.value)
    }
  }
  const apply = (): void => {
    writeInput()
    if (resultChanged && !args.applyDirectFormulaCurrentResult(formulaCellIndex, result)) {
      throw new Error('Failed to apply direct lookup result')
    }
  }
  if (resultChanged && args.cellsShareVersionColumn(request.existingIndex, formulaCellIndex)) {
    args.withOptionalColumnVersionBatch(true, apply)
  } else {
    apply()
  }
  addEngineCounter(args.state.counters, resultChanged ? 'directFormulaKernelSyncOnlyRecalcSkips' : 'kernelSyncOnlyRecalcSkips')
  args.deferSingleCellKernelSync(request.existingIndex)
  const lastMetrics = args.makeSingleLiteralSkipMetrics()
  args.state.setLastMetrics(lastMetrics)
  if (request.emitTracked) {
    const changedCellIndices = resultChanged
      ? Uint32Array.of(request.existingIndex, formulaCellIndex)
      : Uint32Array.of(request.existingIndex)
    emitOperationTrackedCellsBatch({
      events: args.state.events,
      changedCellIndices,
      metrics: lastMetrics,
    })
  }
  return compactDirectLookupResult(resultChanged)
}

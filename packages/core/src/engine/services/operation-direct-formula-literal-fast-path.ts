import { ValueTag, type LiteralInput, type RecalcMetrics } from '@bilig/protocol'
import { CellFlags } from '../../cell-store.js'
import { writeLiteralToCellStore } from '../../engine-value-utils.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type { RuntimeDirectLookupDescriptor, RuntimeDirectScalarDescriptor } from '../runtime-state.js'
import type { DirectScalarCurrentOperand } from './direct-formula-index-collection.js'
import type { CreateEngineOperationServiceArgs } from './operation-service-types.js'

export interface OperationDirectFormulaLiteralMutationRequest {
  readonly existingIndex: number
  readonly formulaCellIndex: number
  readonly value: LiteralInput
  readonly oldNumber: number
  readonly newNumber: number
  readonly exactLookupValue: number | undefined
  readonly approximateLookupValue: number | undefined
}

export interface OperationDirectFormulaLiteralFastPathArgs {
  readonly state: Pick<CreateEngineOperationServiceArgs['state'], 'workbook' | 'strings' | 'formulas' | 'counters' | 'setLastMetrics'>
  readonly hasNoCellDependents: (formulaCellIndex: number) => boolean
  readonly tryDirectUniformLookupNumericResultFromDescriptor: (
    directLookup: RuntimeDirectLookupDescriptor | undefined,
    exactLookupValue: number | undefined,
    approximateLookupValue: number | undefined,
  ) => number | undefined
  readonly directScalarNumericResultMatchesCell: (formulaCellIndex: number, result: number) => boolean
  readonly tryDirectUniformLookupCurrentResultFromNumeric: (
    formulaCellIndex: number,
    exactLookupValue: number | undefined,
    approximateLookupValue: number | undefined,
  ) => DirectScalarCurrentOperand | undefined
  readonly directScalarCurrentResultMatchesCell: (formulaCellIndex: number, result: DirectScalarCurrentOperand) => boolean
  readonly canUseDirectFormulaPostRecalc: (formulaCellIndex: number) => boolean
  readonly tryDirectScalarNumericDeltaFromNumbers: (
    directScalar: RuntimeDirectScalarDescriptor,
    changedCellIndex: number,
    oldChangedNumber: number,
    newChangedNumber: number,
  ) => number | undefined
  readonly writeFastPathLiteralToExistingCell: (existingIndex: number, value: LiteralInput) => void
  readonly applyTerminalDirectFormulaNumericResult: (formulaCellIndex: number, result: number) => void
  readonly applyDirectFormulaCurrentResult: (formulaCellIndex: number, result: DirectScalarCurrentOperand) => boolean
  readonly applyDirectFormulaNumericDelta: (formulaCellIndex: number, delta: number) => boolean
  readonly cellsShareVersionColumn: (leftCellIndex: number, rightCellIndex: number) => boolean
  readonly withOptionalColumnVersionBatch: (enabled: boolean, apply: () => void) => void
  readonly deferSingleCellKernelSync: (cellIndex: number) => void
  readonly makeSingleLiteralSkipMetrics: () => RecalcMetrics
}

export function tryApplySingleDirectFormulaLiteralMutationWithoutEvents(
  args: OperationDirectFormulaLiteralFastPathArgs,
  request: OperationDirectFormulaLiteralMutationRequest,
): boolean {
  const formulaCellIndex = request.formulaCellIndex
  if (formulaCellIndex < 0) {
    return false
  }
  const formula = args.state.formulas.get(formulaCellIndex)
  if (!formula || ((args.state.workbook.cellStore.flags[formulaCellIndex] ?? 0) & CellFlags.InCycle) !== 0) {
    return false
  }
  if (formula.directLookup !== undefined) {
    return tryApplyDirectLookupFormulaLiteralMutation(args, request, formula.directLookup)
  }
  return tryApplyDirectScalarFormulaLiteralMutation(args, request, formula.directScalar)
}

function tryApplyDirectLookupFormulaLiteralMutation(
  args: OperationDirectFormulaLiteralFastPathArgs,
  request: OperationDirectFormulaLiteralMutationRequest,
  directLookup: RuntimeDirectLookupDescriptor,
): boolean {
  const formulaCellIndex = request.formulaCellIndex
  if (!args.hasNoCellDependents(formulaCellIndex)) {
    return false
  }
  const numericResult = args.tryDirectUniformLookupNumericResultFromDescriptor(
    directLookup,
    request.exactLookupValue,
    request.approximateLookupValue,
  )
  if (numericResult !== undefined) {
    const resultChanged = !args.directScalarNumericResultMatchesCell(formulaCellIndex, numericResult)
    args.withOptionalColumnVersionBatch(resultChanged && args.cellsShareVersionColumn(request.existingIndex, formulaCellIndex), () => {
      args.writeFastPathLiteralToExistingCell(request.existingIndex, request.value)
      if (resultChanged) {
        args.applyTerminalDirectFormulaNumericResult(formulaCellIndex, numericResult)
      }
    })
    addEngineCounter(args.state.counters, resultChanged ? 'directFormulaKernelSyncOnlyRecalcSkips' : 'kernelSyncOnlyRecalcSkips')
    args.deferSingleCellKernelSync(request.existingIndex)
    args.state.setLastMetrics(args.makeSingleLiteralSkipMetrics())
    return true
  }
  const result = args.tryDirectUniformLookupCurrentResultFromNumeric(
    formulaCellIndex,
    request.exactLookupValue,
    request.approximateLookupValue,
  )
  if (result === undefined) {
    return false
  }
  const resultChanged = !args.directScalarCurrentResultMatchesCell(formulaCellIndex, result)
  args.withOptionalColumnVersionBatch(resultChanged && args.cellsShareVersionColumn(request.existingIndex, formulaCellIndex), () => {
    args.writeFastPathLiteralToExistingCell(request.existingIndex, request.value)
    if (resultChanged && !args.applyDirectFormulaCurrentResult(formulaCellIndex, result)) {
      throw new Error('Failed to apply direct lookup result')
    }
  })
  addEngineCounter(args.state.counters, resultChanged ? 'directFormulaKernelSyncOnlyRecalcSkips' : 'kernelSyncOnlyRecalcSkips')
  args.deferSingleCellKernelSync(request.existingIndex)
  args.state.setLastMetrics(args.makeSingleLiteralSkipMetrics())
  return true
}

function tryApplyDirectScalarFormulaLiteralMutation(
  args: OperationDirectFormulaLiteralFastPathArgs,
  request: OperationDirectFormulaLiteralMutationRequest,
  directScalar: RuntimeDirectScalarDescriptor | undefined,
): boolean {
  const formulaCellIndex = request.formulaCellIndex
  if (!args.canUseDirectFormulaPostRecalc(formulaCellIndex)) {
    return false
  }
  if (directScalar === undefined || args.state.workbook.cellStore.tags[formulaCellIndex] !== ValueTag.Number) {
    return false
  }
  const delta = args.tryDirectScalarNumericDeltaFromNumbers(directScalar, request.existingIndex, request.oldNumber, request.newNumber)
  if (delta === undefined) {
    return false
  }
  args.withOptionalColumnVersionBatch(args.cellsShareVersionColumn(request.existingIndex, formulaCellIndex), () => {
    writeLiteralToCellStore(args.state.workbook.cellStore, request.existingIndex, request.value, args.state.strings)
    args.state.workbook.notifyCellValueWritten(request.existingIndex)
    if (!args.applyDirectFormulaNumericDelta(formulaCellIndex, delta)) {
      throw new Error('Failed to apply direct scalar delta')
    }
  })
  addEngineCounter(args.state.counters, 'directScalarDeltaApplications')
  addEngineCounter(args.state.counters, 'directScalarDeltaOnlyRecalcSkips')
  args.deferSingleCellKernelSync(request.existingIndex)
  args.state.setLastMetrics(args.makeSingleLiteralSkipMetrics())
  return true
}

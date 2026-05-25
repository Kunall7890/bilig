import { ValueTag, type LiteralInput, type RecalcMetrics } from '@bilig/protocol'
import { CellFlags } from '../../cell-store.js'
import { writeLiteralToCellStore } from '../../engine-value-utils.js'
import { makeCellEntity } from '../../entity-ids.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type { RuntimeDirectScalarDescriptor, U32 } from '../runtime-state.js'
import type { CreateEngineOperationServiceArgs } from './operation-service-types.js'

export interface OperationDirectScalarLiteralMutationRequest {
  readonly existingIndex: number
  readonly value: LiteralInput
  readonly oldNumber: number
  readonly newNumber: number
}

export interface OperationDirectScalarLiteralFastPathArgs {
  readonly state: Pick<CreateEngineOperationServiceArgs['state'], 'workbook' | 'strings' | 'formulas' | 'counters' | 'setLastMetrics'>
  readonly getSingleEntityDependent: (entityId: number) => number
  readonly getEntityDependents: (entityId: number) => U32
  readonly canSkipDirectFormulaColumnVersion: (cellIndex: number) => boolean
  readonly canUseDirectFormulaPostRecalc: (formulaCellIndex: number) => boolean
  readonly tryDirectScalarNumericDeltaFromNumbers: (
    directScalar: RuntimeDirectScalarDescriptor,
    changedCellIndex: number,
    oldChangedNumber: number,
    newChangedNumber: number,
  ) => number | undefined
  readonly applyDirectFormulaNumericDelta: (formulaCellIndex: number, delta: number) => boolean
  readonly deferSingleCellKernelSync: (cellIndex: number) => void
  readonly makeSingleLiteralSkipMetrics: () => RecalcMetrics
}

export function tryApplySingleDirectScalarLiteralMutationWithoutEvents(
  args: OperationDirectScalarLiteralFastPathArgs,
  request: OperationDirectScalarLiteralMutationRequest,
): boolean {
  return tryApplySingleDirectScalarLiteralMutationWithoutEventsAndReturnChanged(args, request) !== null
}

export function tryApplySingleDirectScalarLiteralMutationWithoutEventsAndReturnChanged(
  args: OperationDirectScalarLiteralFastPathArgs,
  request: OperationDirectScalarLiteralMutationRequest,
): U32 | null {
  const dependencyEntity = makeCellEntity(request.existingIndex)
  const singleDependent = args.getSingleEntityDependent(dependencyEntity)
  if (singleDependent === -1) {
    return null
  }

  let singleFormulaCellIndex = -1
  let dependents: U32 | undefined
  if (singleDependent >= 0) {
    singleFormulaCellIndex = singleDependent
  } else {
    dependents = args.getEntityDependents(dependencyEntity)
    if (dependents.length === 0) {
      return null
    }
  }

  let commonDelta = 0
  let hasCommonDelta = false
  let canUseCleanTerminalWrites = true
  const validateDependent = (formulaCellIndex: number): boolean => {
    if (!args.canUseDirectFormulaPostRecalc(formulaCellIndex)) {
      return false
    }
    const formula = args.state.formulas.get(formulaCellIndex)
    const delta =
      formula?.directScalar === undefined
        ? undefined
        : args.tryDirectScalarNumericDeltaFromNumbers(formula.directScalar, request.existingIndex, request.oldNumber, request.newNumber)
    if (
      delta === undefined ||
      ((args.state.workbook.cellStore.flags[formulaCellIndex] ?? 0) & CellFlags.InCycle) !== 0 ||
      args.state.workbook.cellStore.tags[formulaCellIndex] !== ValueTag.Number
    ) {
      return false
    }
    if (
      canUseCleanTerminalWrites &&
      (!args.canSkipDirectFormulaColumnVersion(formulaCellIndex) ||
        (args.state.workbook.cellStore.stringIds[formulaCellIndex] ?? 0) !== 0 ||
        (args.state.workbook.cellStore.errors[formulaCellIndex] ?? 0) !== 0 ||
        ((args.state.workbook.cellStore.flags[formulaCellIndex] ?? 0) & (CellFlags.SpillChild | CellFlags.PivotOutput)) !== 0)
    ) {
      canUseCleanTerminalWrites = false
    }
    if (!hasCommonDelta) {
      commonDelta = delta
      hasCommonDelta = true
      return true
    }
    return Object.is(commonDelta, delta)
  }

  if (singleFormulaCellIndex >= 0) {
    if (!validateDependent(singleFormulaCellIndex)) {
      return null
    }
  } else {
    for (let index = 0; index < dependents!.length; index += 1) {
      if (!validateDependent(dependents![index]!)) {
        return null
      }
    }
  }

  const changedFormulaIndices = singleFormulaCellIndex >= 0 ? Uint32Array.of(singleFormulaCellIndex) : dependents!

  args.state.workbook.withBatchedColumnVersionUpdates(() => {
    writeLiteralToCellStore(args.state.workbook.cellStore, request.existingIndex, request.value, args.state.strings)
    args.state.workbook.notifyCellValueWritten(request.existingIndex)
    if (canUseCleanTerminalWrites) {
      const numbers = args.state.workbook.cellStore.numbers
      const versions = args.state.workbook.cellStore.versions
      for (let index = 0; index < changedFormulaIndices.length; index += 1) {
        const cellIndex = changedFormulaIndices[index]!
        numbers[cellIndex] = numbers[cellIndex]! + commonDelta
        versions[cellIndex] = versions[cellIndex]! + 1
      }
      return
    }
    for (let index = 0; index < changedFormulaIndices.length; index += 1) {
      if (!args.applyDirectFormulaNumericDelta(changedFormulaIndices[index]!, commonDelta)) {
        throw new Error('Failed to apply direct scalar delta')
      }
    }
  })
  const applicationCount = singleFormulaCellIndex >= 0 ? 1 : dependents!.length
  addEngineCounter(args.state.counters, 'directScalarDeltaApplications', applicationCount)
  addEngineCounter(args.state.counters, 'directScalarDeltaOnlyRecalcSkips')
  args.deferSingleCellKernelSync(request.existingIndex)
  args.state.setLastMetrics(args.makeSingleLiteralSkipMetrics())
  return changedFormulaIndices
}

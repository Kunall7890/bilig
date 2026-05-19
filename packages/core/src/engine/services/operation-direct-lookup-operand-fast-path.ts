import { ErrorCode, ValueTag, type CellValue, type LiteralInput, type RecalcMetrics } from '@bilig/protocol'
import type { EngineExistingNumericCellMutationResult } from '../../cell-mutations-at.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type { SheetRecord } from '../../workbook-store.js'
import type { RuntimeDirectCriteriaDescriptor, RuntimeDirectCriteriaOperand, RuntimeDirectLookupDescriptor } from '../runtime-state.js'
import { makeCompactExistingNumericMutationResult } from './operation-change-helpers.js'
import type { DirectScalarCurrentOperand } from './direct-formula-index-collection.js'
import {
  approximateUniformLookupNumericResult,
  directLookupVersionMatches,
  exactUniformLookupNumericResult,
} from './direct-lookup-helpers.js'
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
  readonly tryDirectExactLookupCurrentResult: (
    directLookup: Extract<RuntimeDirectLookupDescriptor, { kind: 'exact' }>,
    lookupValue: CellValue,
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

type DirectLookupCellStore = OperationDirectLookupOperandFastPathArgs['state']['workbook']['cellStore']

type DirectIndexOffsetResult = DirectScalarCurrentOperand & {
  readonly sourceCellIndex?: number
}

function exactLookupCellValueFromLiteral(value: LiteralInput): CellValue {
  if (value === null) {
    return { tag: ValueTag.Empty }
  }
  if (typeof value === 'number') {
    return { tag: ValueTag.Number, value }
  }
  if (typeof value === 'boolean') {
    return { tag: ValueTag.Boolean, value }
  }
  return { tag: ValueTag.String, value, stringId: 0 }
}

function directIndexOffsetFromLiteral(value: LiteralInput): DirectScalarCurrentOperand {
  if (value === null) {
    return { kind: 'number', value: 0 }
  }
  if (typeof value === 'number') {
    return Number.isFinite(value) ? { kind: 'number', value } : { kind: 'error', code: ErrorCode.Value }
  }
  if (typeof value === 'boolean') {
    return { kind: 'number', value: value ? 1 : 0 }
  }
  return { kind: 'error', code: ErrorCode.Value }
}

function tryDirectIndexOffsetOperandResult(
  args: OperationDirectLookupOperandFastPathArgs,
  request: OperationDirectLookupOperandMutationRequest,
): DirectIndexOffsetResult | undefined {
  const formula = args.state.formulas.get(request.formulaCellIndex)
  const directCriteria = formula?.directCriteria
  const aggregateRange = directCriteria?.aggregateRange
  if (
    formula?.compiled.producesSpill ||
    directCriteria?.aggregateKind !== 'first' ||
    directCriteria.offsetOperand?.kind !== 'cell' ||
    directCriteria.offsetOperand.cellIndex !== request.existingIndex ||
    directCriteria.criteriaPairs.length !== 0 ||
    directCriteria.resultTransforms !== undefined ||
    aggregateRange === undefined
  ) {
    return undefined
  }
  const offset = directIndexOffsetFromLiteral(request.value)
  if (offset.kind === 'error') {
    return offset
  }
  const rowOffset = Math.trunc(offset.value)
  if (rowOffset < 1 || rowOffset > aggregateRange.length) {
    return { kind: 'error', code: ErrorCode.Ref }
  }
  const sheet = args.state.workbook.getSheet(aggregateRange.sheetName)
  if (!sheet) {
    return undefined
  }
  const row = aggregateRange.rowStart + rowOffset - 1
  const sourceCellIndex =
    sheet.structureVersion === 1 ? sheet.grid.getPhysical(row, aggregateRange.col) : sheet.grid.get(row, aggregateRange.col)
  if (sourceCellIndex === -1) {
    return undefined
  }
  const cellStore = args.state.workbook.cellStore
  const tag = (cellStore.tags[sourceCellIndex] as ValueTag | undefined) ?? ValueTag.Empty
  if (tag === ValueTag.Number) {
    return { kind: 'number', value: cellStore.numbers[sourceCellIndex] ?? 0, sourceCellIndex }
  }
  if (tag === ValueTag.Error) {
    return { kind: 'error', code: (cellStore.errors[sourceCellIndex] as ErrorCode | undefined) ?? ErrorCode.None, sourceCellIndex }
  }
  return undefined
}

function writeDirectLookupOperandInput(
  args: OperationDirectLookupOperandFastPathArgs,
  request: OperationDirectLookupOperandMutationRequest,
): void {
  if (typeof request.value === 'number' && request.trustedInputSheet !== undefined && request.trustedInputCol !== undefined) {
    args.writeTrustedExistingNumericLiteralToCell(request.existingIndex, request.trustedInputSheet, request.trustedInputCol, request.value)
  } else if (typeof request.value === 'number') {
    args.writeNumericLiteralToExistingCell(request.existingIndex, request.value)
  } else {
    args.writeFastPathLiteralToExistingCell(request.existingIndex, request.value)
  }
}

function makeCompactDirectLookupMutationResult(
  cellStore: DirectLookupCellStore,
  existingIndex: number,
  formulaCellIndex: number,
  resultChanged: boolean,
  secondChangedNumericValue?: number,
): EngineExistingNumericCellMutationResult {
  if (!resultChanged) {
    return makeCompactExistingNumericMutationResult(existingIndex, undefined, 1)
  }
  return makeCompactExistingNumericMutationResult(
    existingIndex,
    formulaCellIndex,
    1,
    secondChangedNumericValue,
    cellStore.rows[formulaCellIndex] ?? 0,
    cellStore.cols[formulaCellIndex] ?? 0,
  )
}

function directCriteriaOperandUsesCell(operand: RuntimeDirectCriteriaOperand, cellIndex: number): boolean {
  switch (operand.kind) {
    case 'cell':
    case 'cell-string-concat':
    case 'cell-month-boundary-string-concat':
      return operand.cellIndex === cellIndex
    case 'literal':
      return false
  }
}

function directCriteriaUsesOperandCell(directCriteria: RuntimeDirectCriteriaDescriptor, cellIndex: number): boolean {
  const offsetOperand = directCriteria.offsetOperand
  if (offsetOperand?.kind === 'cell' && offsetOperand.cellIndex === cellIndex) {
    return true
  }
  for (let index = 0; index < directCriteria.criteriaPairs.length; index += 1) {
    if (directCriteriaOperandUsesCell(directCriteria.criteriaPairs[index]!.criterion, cellIndex)) {
      return true
    }
  }
  const transforms = directCriteria.resultTransforms
  if (transforms === undefined) {
    return false
  }
  for (let index = 0; index < transforms.length; index += 1) {
    const transform = transforms[index]!
    if (transform.kind === 'if-empty-cell' && transform.cellIndex === cellIndex) {
      return true
    }
  }
  return false
}

function tryApplyDirectCriteriaOperandMutationFastPath(
  args: OperationDirectLookupOperandFastPathArgs,
  request: OperationDirectLookupOperandMutationRequest,
): EngineExistingNumericCellMutationResult | null {
  const formulaCellIndex = request.formulaCellIndex
  const formula = args.state.formulas.get(formulaCellIndex)
  const directCriteria = formula?.directCriteria
  if (
    formula?.compiled.producesSpill ||
    directCriteria === undefined ||
    !directCriteriaUsesOperandCell(directCriteria, request.existingIndex)
  ) {
    return null
  }
  const cellStore = args.state.workbook.cellStore
  const beforeTag = cellStore.tags[formulaCellIndex]
  const beforeNumber = cellStore.numbers[formulaCellIndex] ?? 0
  const beforeStringId = cellStore.stringIds[formulaCellIndex] ?? 0
  const beforeError = cellStore.errors[formulaCellIndex] ?? ErrorCode.None
  const apply = (): void => {
    writeDirectLookupOperandInput(args, request)
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
  return makeCompactDirectLookupMutationResult(
    cellStore,
    request.existingIndex,
    formulaCellIndex,
    resultChanged,
    resultChanged && afterTag === ValueTag.Number ? afterNumber : undefined,
  )
}

function tryApplyDirectIndexOffsetOperandMutationFastPath(
  args: OperationDirectLookupOperandFastPathArgs,
  request: OperationDirectLookupOperandMutationRequest,
): EngineExistingNumericCellMutationResult | null {
  const result = tryDirectIndexOffsetOperandResult(args, request)
  if (result === undefined) {
    return null
  }
  const formulaCellIndex = request.formulaCellIndex
  const resultChanged = !args.directScalarCurrentResultMatchesCell(formulaCellIndex, result)
  const apply = (): void => {
    writeDirectLookupOperandInput(args, request)
    if (!resultChanged) {
      return
    }
    if (result.kind === 'number') {
      args.applyTerminalDirectFormulaNumericResult(formulaCellIndex, result.value)
    } else if (!args.applyDirectFormulaCurrentResult(formulaCellIndex, result)) {
      throw new Error('Failed to apply direct INDEX offset result')
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
  return makeCompactDirectLookupMutationResult(
    args.state.workbook.cellStore,
    request.existingIndex,
    formulaCellIndex,
    resultChanged,
    result.kind === 'number' ? result.value : undefined,
  )
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
  if (directLookup?.operandCellIndex !== request.existingIndex) {
    const directIndexOffsetResult = tryApplyDirectIndexOffsetOperandMutationFastPath(args, request)
    if (directIndexOffsetResult !== null) {
      return directIndexOffsetResult
    }
    const directCriteriaResult = tryApplyDirectCriteriaOperandMutationFastPath(args, request)
    if (directCriteriaResult !== null) {
      return directCriteriaResult
    }
  }
  let numericResult: number | undefined
  if (directLookup?.kind === 'exact-uniform-numeric' && request.exactLookupValue !== undefined) {
    const lookupSheet =
      request.lookupSheetHint?.id === directLookup.sheetId
        ? request.lookupSheetHint
        : args.state.workbook.getSheetById(directLookup.sheetId)
    numericResult = directLookupVersionMatches(lookupSheet, directLookup)
      ? exactUniformLookupNumericResult(directLookup, request.exactLookupValue)
      : undefined
  } else if (directLookup?.kind === 'approximate-uniform-numeric' && request.approximateLookupValue !== undefined) {
    const lookupSheet =
      request.lookupSheetHint?.id === directLookup.sheetId
        ? request.lookupSheetHint
        : args.state.workbook.getSheetById(directLookup.sheetId)
    numericResult = directLookupVersionMatches(lookupSheet, directLookup)
      ? approximateUniformLookupNumericResult(directLookup, request.approximateLookupValue)
      : undefined
  }
  if (numericResult !== undefined) {
    const cellStore = args.state.workbook.cellStore
    const resultChanged =
      cellStore.tags[formulaCellIndex] !== ValueTag.Number || !Object.is(cellStore.numbers[formulaCellIndex] ?? 0, numericResult)
    if (resultChanged && args.cellsShareVersionColumn(request.existingIndex, formulaCellIndex)) {
      args.withOptionalColumnVersionBatch(true, () => {
        writeDirectLookupOperandInput(args, request)
        args.applyTerminalDirectFormulaNumericResult(formulaCellIndex, numericResult)
      })
    } else {
      writeDirectLookupOperandInput(args, request)
      if (resultChanged) {
        args.applyTerminalDirectFormulaNumericResult(formulaCellIndex, numericResult)
      }
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
    return makeCompactDirectLookupMutationResult(
      cellStore,
      request.existingIndex,
      formulaCellIndex,
      resultChanged,
      resultChanged ? numericResult : undefined,
    )
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
      return makeCompactDirectLookupMutationResult(
        args.state.workbook.cellStore,
        request.existingIndex,
        formulaCellIndex,
        resultChanged,
        approximateNumericResult,
      )
    }
  }
  if (directLookup?.kind === 'exact' && directLookup.operandCellIndex === request.existingIndex) {
    const result = args.tryDirectExactLookupCurrentResult(directLookup, exactLookupCellValueFromLiteral(request.value))
    if (result !== undefined) {
      const exactNumericResult = result.kind === 'number' ? result.value : undefined
      const resultChanged =
        exactNumericResult === undefined
          ? !args.directScalarCurrentResultMatchesCell(formulaCellIndex, result)
          : !args.directScalarNumericResultMatchesCell(formulaCellIndex, exactNumericResult)
      const apply = (): void => {
        writeDirectLookupOperandInput(args, request)
        if (resultChanged) {
          if (exactNumericResult !== undefined) {
            args.applyTerminalDirectFormulaNumericResult(formulaCellIndex, exactNumericResult)
          } else if (!args.applyDirectFormulaCurrentResult(formulaCellIndex, result)) {
            throw new Error('Failed to apply direct exact lookup result')
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
      return makeCompactDirectLookupMutationResult(
        args.state.workbook.cellStore,
        request.existingIndex,
        formulaCellIndex,
        resultChanged,
        exactNumericResult,
      )
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
    return makeCompactDirectLookupMutationResult(
      args.state.workbook.cellStore,
      request.existingIndex,
      formulaCellIndex,
      resultChanged,
      resultChanged && afterTag === ValueTag.Number ? afterNumber : undefined,
    )
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
  return makeCompactDirectLookupMutationResult(args.state.workbook.cellStore, request.existingIndex, formulaCellIndex, resultChanged)
}

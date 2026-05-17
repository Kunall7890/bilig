import type { EngineOpBatch } from '@bilig/workbook'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import type { EngineCellMutationRef } from '../../cell-mutations-at.js'
import { CellFlags } from '../../cell-store.js'
import { makeCellEntity } from '../../entity-ids.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import { markBatchApplied } from '../../replica-state.js'
import type { TransactionRecord } from '../runtime-state.js'
import type { OperationDirectScalarBatchFastPathArgs } from './operation-direct-scalar-batch-fast-paths.js'
import {
  directScalarCellNumber,
  evaluateDirectScalarWithReplacementNumbers,
  singleInputAffineDirectScalar,
} from './direct-scalar-helpers.js'
import { tagTrustedPhysicalTrackedChanges } from './operation-change-helpers.js'
import { emitCellMutationFastPathBatchResult } from './operation-fast-path-batch-result.js'

const EMPTY_CHANGED_CELLS = new Uint32Array(0)

export function createOperationDirectScalarSingleColumnBatchFastPaths(args: OperationDirectScalarBatchFastPathArgs): {
  readonly tryApplyDenseSingleColumnAffineDirectScalarLiteralBatch: (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    potentialNewCells?: number,
  ) => boolean
  readonly tryApplyDenseSingleColumnDirectScalarLiteralBatch: (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    potentialNewCells?: number,
  ) => boolean
  readonly tryApplyDenseSingleColumnAffineExistingNumericBatch: (
    record: Extract<TransactionRecord, { kind: 'existing-numeric-cell-mutations' }>,
    batch: EngineOpBatch | null,
  ) => boolean
} {
  const {
    emitBatch,
    hasTrackedExactLookupDependents,
    hasTrackedSortedLookupDependents,
    hasTrackedDirectRangeDependents,
    canFastPathLiteralOverwrite,
    canUseDirectFormulaPostRecalc,
    canSkipFormulaColumnVersion,
    applyTerminalDirectFormulaNumericResults,
  } = args

  const tryApplyDenseSingleColumnDirectScalarLiteralBatch = (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    potentialNewCells?: number,
  ): boolean => {
    const firstRef = refs[0]
    if (firstRef === undefined || refs.length < 32 || (potentialNewCells ?? 0) !== 0) {
      return false
    }
    const firstMutation = firstRef.mutation
    if (firstMutation.kind !== 'setCellValue' || typeof firstMutation.value !== 'number' || Object.is(firstMutation.value, -0)) {
      return false
    }
    const secondMutation = refs[1]?.mutation
    if (secondMutation?.kind !== 'setCellValue') {
      return false
    }
    const rowOrder = secondMutation.row > firstMutation.row ? 1 : secondMutation.row < firstMutation.row ? -1 : 0
    if (rowOrder === 0) {
      return false
    }
    const firstSheet = args.state.workbook.getSheetById(firstRef.sheetId)
    if (
      !firstSheet ||
      firstSheet.structureVersion !== 1 ||
      hasTrackedExactLookupDependents(firstRef.sheetId, firstMutation.col) ||
      hasTrackedSortedLookupDependents(firstRef.sheetId, firstMutation.col) ||
      hasTrackedDirectRangeDependents(firstRef.sheetId, firstMutation.col)
    ) {
      return false
    }
    const inputCellIndices = new Uint32Array(refs.length)
    const formulaCellIndices = new Uint32Array(refs.length)
    const inputNumericValues = new Float64Array(refs.length)
    const formulaNumericResults = new Float64Array(refs.length)
    const cellStore = args.state.workbook.cellStore
    const readDirectScalarNumber = (cellIndex: number): number | undefined => directScalarCellNumber(cellStore, cellIndex)
    let previousRow = firstMutation.row
    for (let refIndex = 0; refIndex < refs.length; refIndex += 1) {
      const ref = refs[refIndex]!
      const mutation = ref.mutation
      if (refIndex > 0) {
        if ((rowOrder > 0 && mutation.row <= previousRow) || (rowOrder < 0 && mutation.row >= previousRow)) {
          return false
        }
        previousRow = mutation.row
      }
      if (
        ref.sheetId !== firstRef.sheetId ||
        mutation.kind !== 'setCellValue' ||
        mutation.col !== firstMutation.col ||
        typeof mutation.value !== 'number' ||
        Object.is(mutation.value, -0)
      ) {
        return false
      }
      const existingIndex =
        ref.cellIndex !== undefined &&
        args.state.workbook.cellStore.sheetIds[ref.cellIndex] === ref.sheetId &&
        args.state.workbook.cellStore.rows[ref.cellIndex] === mutation.row &&
        args.state.workbook.cellStore.cols[ref.cellIndex] === mutation.col
          ? ref.cellIndex
          : firstSheet.grid.getPhysical(mutation.row, mutation.col)
      if (existingIndex === -1 || !canFastPathLiteralOverwrite(existingIndex)) {
        return false
      }
      const singleDependent = args.getSingleEntityDependent(makeCellEntity(existingIndex))
      if (singleDependent < 0 || !canUseDirectFormulaPostRecalc(singleDependent) || !canSkipFormulaColumnVersion(singleDependent)) {
        return false
      }
      const formula = args.state.formulas.get(singleDependent)
      const result =
        formula?.directScalar === undefined
          ? undefined
          : evaluateDirectScalarWithReplacementNumbers(formula.directScalar, existingIndex, mutation.value, readDirectScalarNumber)
      if (result === undefined) {
        return false
      }
      const outputIndex = rowOrder < 0 ? refs.length - 1 - refIndex : refIndex
      inputCellIndices[outputIndex] = existingIndex
      formulaCellIndices[outputIndex] = singleDependent
      inputNumericValues[outputIndex] = mutation.value
      formulaNumericResults[outputIndex] = result
    }

    args.materializeDeferredStructuralFormulaSources()

    const hasGeneralEventListeners = args.state.events.hasListeners()
    const hasTrackedEventListeners = args.state.events.hasTrackedListeners()
    const hasWatchedCellListeners = args.state.events.hasCellListeners()
    const requiresChangedSet = hasGeneralEventListeners || hasTrackedEventListeners || hasWatchedCellListeners
    const changedInputCount = inputCellIndices.length
    const explicitChangedCount = requiresChangedSet ? inputCellIndices.length : 0
    const flags = cellStore.flags
    const versions = cellStore.versions
    const stringIds = cellStore.stringIds
    const tags = cellStore.tags
    const numbers = cellStore.numbers
    const errors = cellStore.errors
    const formulaOutputFlags = CellFlags.SpillChild | CellFlags.PivotOutput
    const clearFormulaOutputFlags = ~formulaOutputFlags
    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      for (let index = 0; index < refs.length; index += 1) {
        const cellIndex = inputCellIndices[index]!
        const currentFlags = flags[cellIndex] ?? 0
        if ((currentFlags & formulaOutputFlags) !== 0) {
          flags[cellIndex] = currentFlags & clearFormulaOutputFlags
        }
        versions[cellIndex] = (versions[cellIndex] ?? 0) + 1
        stringIds[cellIndex] = 0
        tags[cellIndex] = ValueTag.Number
        numbers[cellIndex] = inputNumericValues[index]!
        errors[cellIndex] = ErrorCode.None
      }
      args.state.workbook.notifyColumnsWritten(firstRef.sheetId, Uint32Array.of(firstMutation.col))
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
    }
    if (batch) {
      markBatchApplied(args.state.replicaState, batch)
    }
    args.deferKernelSync(inputCellIndices)
    const formulaChanged = requiresChangedSet ? new Uint32Array(refs.length) : EMPTY_CHANGED_CELLS
    if (requiresChangedSet) {
      for (let index = 0; index < refs.length; index += 1) {
        formulaChanged[index] = formulaCellIndices[index]!
      }
    }
    applyTerminalDirectFormulaNumericResults(formulaCellIndices, formulaNumericResults)
    addEngineCounter(args.state.counters, 'directScalarDeltaApplications', refs.length)
    addEngineCounter(args.state.counters, 'directScalarDeltaOnlyRecalcSkips')
    const changed = requiresChangedSet ? composeDenseInputFormulaChanges(inputCellIndices, formulaChanged) : EMPTY_CHANGED_CELLS
    if (hasTrackedEventListeners && changed.length > 4 && explicitChangedCount > 0 && explicitChangedCount < changed.length) {
      tagTrustedPhysicalTrackedChanges(changed, firstRef.sheetId, explicitChangedCount)
    }
    emitCellMutationFastPathBatchResult({
      state: args.state,
      changed,
      changedInputCount,
      explicitChangedCount,
      hasGeneralEventListeners,
      hasTrackedEventListeners,
      hasWatchedCellListeners,
      captureChangedCells: args.captureChangedCells,
      batch,
      emitBatch,
    })
    return true
  }

  const tryApplyDenseSingleColumnAffineDirectScalarLiteralBatch = (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    potentialNewCells?: number,
  ): boolean => {
    const firstRef = refs[0]
    if (firstRef === undefined || refs.length < 32 || (potentialNewCells ?? 0) !== 0) {
      return false
    }
    const firstMutation = firstRef.mutation
    if (firstMutation.kind !== 'setCellValue' || typeof firstMutation.value !== 'number' || Object.is(firstMutation.value, -0)) {
      return false
    }
    const secondMutation = refs[1]?.mutation
    if (secondMutation?.kind !== 'setCellValue') {
      return false
    }
    const rowOrder = secondMutation.row > firstMutation.row ? 1 : secondMutation.row < firstMutation.row ? -1 : 0
    if (rowOrder === 0) {
      return false
    }
    const firstSheet = args.state.workbook.getSheetById(firstRef.sheetId)
    if (
      !firstSheet ||
      firstSheet.structureVersion !== 1 ||
      hasTrackedExactLookupDependents(firstRef.sheetId, firstMutation.col) ||
      hasTrackedSortedLookupDependents(firstRef.sheetId, firstMutation.col) ||
      hasTrackedDirectRangeDependents(firstRef.sheetId, firstMutation.col)
    ) {
      return false
    }

    const inputCellIndices = new Uint32Array(refs.length)
    const inputNumericValues = rowOrder < 0 ? new Float64Array(refs.length) : undefined
    const formulaCellIndices = new Uint32Array(refs.length)
    const cellStore = args.state.workbook.cellStore
    let previousRow = firstMutation.row
    let previousFormulaRow = -1
    let previousFormulaCol = -1
    let affineScale: number | undefined
    let affineOffset: number | undefined
    for (let refIndex = 0; refIndex < refs.length; refIndex += 1) {
      const ref = refs[refIndex]!
      const mutation = ref.mutation
      if (refIndex > 0) {
        if ((rowOrder > 0 && mutation.row <= previousRow) || (rowOrder < 0 && mutation.row >= previousRow)) {
          return false
        }
        previousRow = mutation.row
      }
      if (
        ref.sheetId !== firstRef.sheetId ||
        mutation.kind !== 'setCellValue' ||
        mutation.col !== firstMutation.col ||
        typeof mutation.value !== 'number' ||
        Object.is(mutation.value, -0) ||
        ref.cellIndex === undefined ||
        cellStore.sheetIds[ref.cellIndex] !== ref.sheetId ||
        cellStore.rows[ref.cellIndex] !== mutation.row ||
        cellStore.cols[ref.cellIndex] !== mutation.col
      ) {
        return false
      }
      const existingIndex = ref.cellIndex
      if (!canFastPathLiteralOverwrite(existingIndex)) {
        return false
      }
      const singleDependent = args.getSingleEntityDependent(makeCellEntity(existingIndex))
      if (singleDependent < 0 || !canUseDirectFormulaPostRecalc(singleDependent) || !canSkipFormulaColumnVersion(singleDependent)) {
        return false
      }
      const formula = args.state.formulas.get(singleDependent)
      if (
        !formula ||
        formula.directScalar === undefined ||
        cellStore.sheetIds[singleDependent] !== firstRef.sheetId ||
        cellStore.rows[singleDependent] !== mutation.row
      ) {
        return false
      }
      const formulaRow = cellStore.rows[singleDependent] ?? 0
      const formulaCol = cellStore.cols[singleDependent] ?? 0
      if (
        refIndex > 0 &&
        ((rowOrder > 0 && (formulaRow < previousFormulaRow || (formulaRow === previousFormulaRow && formulaCol <= previousFormulaCol))) ||
          (rowOrder < 0 && (formulaRow > previousFormulaRow || (formulaRow === previousFormulaRow && formulaCol >= previousFormulaCol))))
      ) {
        return false
      }
      const affine = singleInputAffineDirectScalar(formula.directScalar, existingIndex)
      if (affine === null) {
        return false
      }
      if (affineScale === undefined) {
        affineScale = affine.scale
        affineOffset = affine.offset
      } else if (!Object.is(affineScale, affine.scale) || !Object.is(affineOffset, affine.offset)) {
        return false
      }
      const outputIndex = rowOrder < 0 ? refs.length - 1 - refIndex : refIndex
      inputCellIndices[outputIndex] = existingIndex
      if (inputNumericValues !== undefined) {
        inputNumericValues[outputIndex] = mutation.value
      }
      formulaCellIndices[outputIndex] = singleDependent
      previousFormulaRow = formulaRow
      previousFormulaCol = formulaCol
    }
    if (affineScale === undefined || affineOffset === undefined) {
      return false
    }

    args.materializeDeferredStructuralFormulaSources()

    const hasGeneralEventListeners = args.state.events.hasListeners()
    const hasTrackedEventListeners = args.state.events.hasTrackedListeners()
    const hasWatchedCellListeners = args.state.events.hasCellListeners()
    const requiresChangedSet = hasGeneralEventListeners || hasTrackedEventListeners || hasWatchedCellListeners
    const changedInputCount = inputCellIndices.length
    const explicitChangedCount = requiresChangedSet ? inputCellIndices.length : 0
    const flags = cellStore.flags
    const versions = cellStore.versions
    const stringIds = cellStore.stringIds
    const tags = cellStore.tags
    const numbers = cellStore.numbers
    const errors = cellStore.errors
    const formulaOutputFlags = CellFlags.SpillChild | CellFlags.PivotOutput
    const clearFormulaOutputFlags = ~formulaOutputFlags
    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      for (let index = 0; index < refs.length; index += 1) {
        const cellIndex = inputCellIndices[index]!
        const value = inputNumericValues === undefined ? readValidatedNumericMutationValue(refs[index]!) : inputNumericValues[index]!
        const inputFlags = flags[cellIndex] ?? 0
        if ((inputFlags & formulaOutputFlags) !== 0) {
          flags[cellIndex] = inputFlags & clearFormulaOutputFlags
        }
        versions[cellIndex] = (versions[cellIndex] ?? 0) + 1
        stringIds[cellIndex] = 0
        tags[cellIndex] = ValueTag.Number
        numbers[cellIndex] = value
        errors[cellIndex] = ErrorCode.None

        const formulaCellIndex = formulaCellIndices[index]!
        const currentFlags = flags[formulaCellIndex] ?? 0
        if ((currentFlags & formulaOutputFlags) !== 0) {
          flags[formulaCellIndex] = currentFlags & clearFormulaOutputFlags
        }
        versions[formulaCellIndex] = (versions[formulaCellIndex] ?? 0) + 1
        stringIds[formulaCellIndex] = 0
        tags[formulaCellIndex] = ValueTag.Number
        numbers[formulaCellIndex] = value * affineScale + affineOffset
        errors[formulaCellIndex] = ErrorCode.None
      }
      args.state.workbook.notifyColumnsWritten(firstRef.sheetId, Uint32Array.of(firstMutation.col))
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
    }
    if (batch) {
      markBatchApplied(args.state.replicaState, batch)
    }
    args.deferKernelSync(inputCellIndices)
    addEngineCounter(args.state.counters, 'directScalarDeltaApplications', refs.length)
    addEngineCounter(args.state.counters, 'directScalarDeltaOnlyRecalcSkips')
    const changed = requiresChangedSet ? composeDenseInputFormulaChanges(inputCellIndices, formulaCellIndices) : EMPTY_CHANGED_CELLS
    if (hasTrackedEventListeners && changed.length > 4 && explicitChangedCount > 0 && explicitChangedCount < changed.length) {
      tagTrustedPhysicalTrackedChanges(changed, firstRef.sheetId, explicitChangedCount)
    }
    emitCellMutationFastPathBatchResult({
      state: args.state,
      changed,
      changedInputCount,
      explicitChangedCount,
      hasGeneralEventListeners,
      hasTrackedEventListeners,
      hasWatchedCellListeners,
      captureChangedCells: args.captureChangedCells,
      batch,
      emitBatch,
    })
    return true
  }

  const tryApplyDenseSingleColumnAffineExistingNumericBatch = (
    record: Extract<TransactionRecord, { kind: 'existing-numeric-cell-mutations' }>,
    batch: EngineOpBatch | null,
  ): boolean => {
    const count = record.sheetIds.length
    if (count < 32 || (record.potentialNewCells ?? 0) !== 0) {
      return false
    }
    const firstSheetId = record.sheetIds[0]!
    const firstRow = record.rows[0]!
    const firstCol = record.cols[0]!
    const secondRow = record.rows[1]!
    const rowOrder = secondRow > firstRow ? 1 : secondRow < firstRow ? -1 : 0
    if (rowOrder === 0) {
      return false
    }
    const firstSheet = args.state.workbook.getSheetById(firstSheetId)
    if (
      !firstSheet ||
      firstSheet.structureVersion !== 1 ||
      hasTrackedExactLookupDependents(firstSheetId, firstCol) ||
      hasTrackedSortedLookupDependents(firstSheetId, firstCol) ||
      hasTrackedDirectRangeDependents(firstSheetId, firstCol)
    ) {
      return false
    }

    const inputCellIndices = new Uint32Array(count)
    const inputNumericValues = rowOrder < 0 ? new Float64Array(count) : undefined
    const formulaCellIndices = new Uint32Array(count)
    const cellStore = args.state.workbook.cellStore
    let previousRow = firstRow
    let previousFormulaRow = -1
    let previousFormulaCol = -1
    let affineScale: number | undefined
    let affineOffset: number | undefined
    for (let refIndex = 0; refIndex < count; refIndex += 1) {
      const row = record.rows[refIndex]!
      const col = record.cols[refIndex]!
      const sheetId = record.sheetIds[refIndex]!
      const cellIndexPlusOne = record.cellIndexPlusOnes[refIndex]!
      const value = record.numbers[refIndex]!
      if (refIndex > 0) {
        if ((rowOrder > 0 && row <= previousRow) || (rowOrder < 0 && row >= previousRow)) {
          return false
        }
        previousRow = row
      }
      if (sheetId !== firstSheetId || col !== firstCol || Object.is(value, -0) || cellIndexPlusOne === 0) {
        return false
      }
      const existingIndex = cellIndexPlusOne - 1
      if (
        cellStore.sheetIds[existingIndex] !== sheetId ||
        cellStore.rows[existingIndex] !== row ||
        cellStore.cols[existingIndex] !== col ||
        !canFastPathLiteralOverwrite(existingIndex)
      ) {
        return false
      }
      const singleDependent = args.getSingleEntityDependent(makeCellEntity(existingIndex))
      if (singleDependent < 0 || !canUseDirectFormulaPostRecalc(singleDependent) || !canSkipFormulaColumnVersion(singleDependent)) {
        return false
      }
      const formula = args.state.formulas.get(singleDependent)
      if (
        !formula ||
        formula.directScalar === undefined ||
        cellStore.sheetIds[singleDependent] !== firstSheetId ||
        cellStore.rows[singleDependent] !== row
      ) {
        return false
      }
      const formulaRow = cellStore.rows[singleDependent] ?? 0
      const formulaCol = cellStore.cols[singleDependent] ?? 0
      if (
        refIndex > 0 &&
        ((rowOrder > 0 && (formulaRow < previousFormulaRow || (formulaRow === previousFormulaRow && formulaCol <= previousFormulaCol))) ||
          (rowOrder < 0 && (formulaRow > previousFormulaRow || (formulaRow === previousFormulaRow && formulaCol >= previousFormulaCol))))
      ) {
        return false
      }
      const affine = singleInputAffineDirectScalar(formula.directScalar, existingIndex)
      if (affine === null) {
        return false
      }
      if (affineScale === undefined) {
        affineScale = affine.scale
        affineOffset = affine.offset
      } else if (!Object.is(affineScale, affine.scale) || !Object.is(affineOffset, affine.offset)) {
        return false
      }
      const outputIndex = rowOrder < 0 ? count - 1 - refIndex : refIndex
      inputCellIndices[outputIndex] = existingIndex
      if (inputNumericValues !== undefined) {
        inputNumericValues[outputIndex] = value
      }
      formulaCellIndices[outputIndex] = singleDependent
      previousFormulaRow = formulaRow
      previousFormulaCol = formulaCol
    }
    if (affineScale === undefined || affineOffset === undefined) {
      return false
    }

    args.materializeDeferredStructuralFormulaSources()

    const hasGeneralEventListeners = args.state.events.hasListeners()
    const hasTrackedEventListeners = args.state.events.hasTrackedListeners()
    const hasWatchedCellListeners = args.state.events.hasCellListeners()
    const requiresChangedSet = hasGeneralEventListeners || hasTrackedEventListeners || hasWatchedCellListeners
    const changedInputCount = inputCellIndices.length
    const explicitChangedCount = requiresChangedSet ? inputCellIndices.length : 0
    const flags = cellStore.flags
    const versions = cellStore.versions
    const stringIds = cellStore.stringIds
    const tags = cellStore.tags
    const numbers = cellStore.numbers
    const errors = cellStore.errors
    const formulaOutputFlags = CellFlags.SpillChild | CellFlags.PivotOutput
    const clearFormulaOutputFlags = ~formulaOutputFlags
    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      for (let index = 0; index < count; index += 1) {
        const cellIndex = inputCellIndices[index]!
        const value = inputNumericValues === undefined ? record.numbers[index]! : inputNumericValues[index]!
        const inputFlags = flags[cellIndex] ?? 0
        if ((inputFlags & formulaOutputFlags) !== 0) {
          flags[cellIndex] = inputFlags & clearFormulaOutputFlags
        }
        versions[cellIndex] = (versions[cellIndex] ?? 0) + 1
        stringIds[cellIndex] = 0
        tags[cellIndex] = ValueTag.Number
        numbers[cellIndex] = value
        errors[cellIndex] = ErrorCode.None

        const formulaCellIndex = formulaCellIndices[index]!
        const currentFlags = flags[formulaCellIndex] ?? 0
        if ((currentFlags & formulaOutputFlags) !== 0) {
          flags[formulaCellIndex] = currentFlags & clearFormulaOutputFlags
        }
        versions[formulaCellIndex] = (versions[formulaCellIndex] ?? 0) + 1
        stringIds[formulaCellIndex] = 0
        tags[formulaCellIndex] = ValueTag.Number
        numbers[formulaCellIndex] = value * affineScale + affineOffset
        errors[formulaCellIndex] = ErrorCode.None
      }
      args.state.workbook.notifyColumnsWritten(firstSheetId, Uint32Array.of(firstCol))
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
    }
    if (batch) {
      markBatchApplied(args.state.replicaState, batch)
    }
    args.deferKernelSync(inputCellIndices)
    addEngineCounter(args.state.counters, 'directScalarDeltaApplications', count)
    addEngineCounter(args.state.counters, 'directScalarDeltaOnlyRecalcSkips')
    const changed = requiresChangedSet ? composeDenseInputFormulaChanges(inputCellIndices, formulaCellIndices) : EMPTY_CHANGED_CELLS
    if (hasTrackedEventListeners && changed.length > 4 && explicitChangedCount > 0 && explicitChangedCount < changed.length) {
      tagTrustedPhysicalTrackedChanges(changed, firstSheetId, explicitChangedCount)
    }
    emitCellMutationFastPathBatchResult({
      state: args.state,
      changed,
      changedInputCount,
      explicitChangedCount,
      hasGeneralEventListeners,
      hasTrackedEventListeners,
      hasWatchedCellListeners,
      captureChangedCells: args.captureChangedCells,
      batch,
      emitBatch,
    })
    return true
  }

  return {
    tryApplyDenseSingleColumnAffineDirectScalarLiteralBatch,
    tryApplyDenseSingleColumnDirectScalarLiteralBatch,
    tryApplyDenseSingleColumnAffineExistingNumericBatch,
  }
}

function composeDenseInputFormulaChanges(inputCellIndices: Uint32Array, formulaCellIndices: Uint32Array): Uint32Array {
  const changed = new Uint32Array(inputCellIndices.length + formulaCellIndices.length)
  changed.set(inputCellIndices)
  changed.set(formulaCellIndices, inputCellIndices.length)
  return changed
}

function readValidatedNumericMutationValue(ref: EngineCellMutationRef): number {
  const mutation = ref.mutation
  if (mutation.kind !== 'setCellValue' || typeof mutation.value !== 'number') {
    throw new Error('Expected affine direct scalar batch to contain numeric literal writes')
  }
  return mutation.value
}

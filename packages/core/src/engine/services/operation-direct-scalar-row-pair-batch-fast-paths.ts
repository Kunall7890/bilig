import type { EngineChangedCell } from '@bilig/protocol'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import type { EngineOpBatch } from '@bilig/workbook'
import type { EngineCellMutationRef } from '../../cell-mutations-at.js'
import { CellFlags } from '../../cell-store.js'
import { makeCellEntity } from '../../entity-ids.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import { markBatchApplied } from '../../replica-state.js'
import type { EngineRuntimeState, U32 } from '../runtime-state.js'
import {
  directScalarCellNumber,
  evaluateDirectScalarWithReplacementNumbers,
  evaluateRowPairDirectScalarCode,
  rowPairDirectScalarCode,
} from './direct-scalar-helpers.js'
import { tagTrustedPhysicalTrackedChanges } from './operation-change-helpers.js'
import { emitCellMutationFastPathBatchResult } from './operation-fast-path-batch-result.js'

const EMPTY_CHANGED_CELLS = new Uint32Array(0)

type FastPathState = Pick<
  EngineRuntimeState,
  | 'workbook'
  | 'strings'
  | 'events'
  | 'formulas'
  | 'counters'
  | 'replicaState'
  | 'getLastMetrics'
  | 'setLastMetrics'
  | 'getSyncClientConnection'
>

interface OperationDirectScalarRowPairBatchFastPathArgs {
  readonly state: FastPathState
  readonly emitBatch: (batch: EngineOpBatch) => void
  readonly hasTrackedExactLookupDependents: (sheetId: number, col: number) => boolean
  readonly hasTrackedSortedLookupDependents: (sheetId: number, col: number) => boolean
  readonly hasTrackedDirectRangeDependents: (sheetId: number, col: number) => boolean
  readonly canFastPathLiteralOverwrite: (cellIndex: number) => boolean
  readonly canUseDirectFormulaPostRecalc: (cellIndex: number) => boolean
  readonly canSkipFormulaColumnVersion: (cellIndex: number) => boolean
  readonly applyTerminalDirectFormulaNumericResults: (
    cellIndices: readonly number[] | Uint32Array,
    values: ArrayLike<number>,
    count?: number,
  ) => void
  readonly getEntityDependents: (entityId: number) => Uint32Array
  readonly materializeDeferredStructuralFormulaSources: () => void
  readonly beginMutationCollection: () => void
  readonly ensureRecalcScratchCapacity: (size: number) => void
  readonly resetMaterializedCellScratch: (expectedSize: number) => void
  readonly getBatchMutationDepth: () => number
  readonly setBatchMutationDepth: (next: number) => void
  readonly markInputChanged: (cellIndex: number, count: number) => number
  readonly markExplicitChanged: (cellIndex: number, count: number) => number
  readonly getChangedInputBuffer: () => U32
  readonly deferKernelSync: (cellIndices: readonly number[] | U32) => void
  readonly composeDisjointEventChanges: (recalculated: U32, explicitChangedCount: number) => U32
  readonly captureChangedCells: (changedCellIndices: readonly number[] | U32) => readonly EngineChangedCell[]
}

export function createOperationDirectScalarRowPairBatchFastPaths(args: OperationDirectScalarRowPairBatchFastPathArgs): {
  readonly tryApplyDenseRowPairSimpleDirectScalarLiteralBatch: (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    potentialNewCells?: number,
  ) => boolean
  readonly tryApplyDenseRowPairDirectScalarLiteralBatch: (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    potentialNewCells?: number,
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

  const tryApplyDenseRowPairSimpleDirectScalarLiteralBatch = (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    potentialNewCells?: number,
  ): boolean => {
    if (refs.length < 32 || refs.length % 2 !== 0 || (potentialNewCells ?? 0) !== 0) {
      return false
    }
    const firstRef = refs[0]!
    const secondRef = refs[1]!
    const firstMutation = firstRef.mutation
    const secondMutation = secondRef.mutation
    if (
      firstMutation.kind !== 'setCellValue' ||
      secondMutation.kind !== 'setCellValue' ||
      firstRef.sheetId !== secondRef.sheetId ||
      firstMutation.row !== secondMutation.row ||
      firstMutation.col >= secondMutation.col ||
      typeof firstMutation.value !== 'number' ||
      typeof secondMutation.value !== 'number' ||
      Object.is(firstMutation.value, -0) ||
      Object.is(secondMutation.value, -0)
    ) {
      return false
    }
    const sheet = args.state.workbook.getSheetById(firstRef.sheetId)
    if (
      !sheet ||
      sheet.structureVersion !== 1 ||
      hasTrackedExactLookupDependents(firstRef.sheetId, firstMutation.col) ||
      hasTrackedExactLookupDependents(firstRef.sheetId, secondMutation.col) ||
      hasTrackedSortedLookupDependents(firstRef.sheetId, firstMutation.col) ||
      hasTrackedSortedLookupDependents(firstRef.sheetId, secondMutation.col) ||
      hasTrackedDirectRangeDependents(firstRef.sheetId, firstMutation.col) ||
      hasTrackedDirectRangeDependents(firstRef.sheetId, secondMutation.col)
    ) {
      return false
    }

    const inputCellIndices = new Uint32Array(refs.length)
    const formulaCellIndices = new Uint32Array(refs.length)
    const formulaNumericResults = new Float64Array(refs.length)
    const cellStore = args.state.workbook.cellStore
    let formulaCount = 0
    let previousRow = firstMutation.row - 1
    let previousFormulaRow = -1
    let previousFormulaCol = -1
    for (let refIndex = 0; refIndex < refs.length; refIndex += 2) {
      const leftRef = refs[refIndex]!
      const rightRef = refs[refIndex + 1]!
      const leftMutation = leftRef.mutation
      const rightMutation = rightRef.mutation
      if (
        leftRef.sheetId !== firstRef.sheetId ||
        rightRef.sheetId !== firstRef.sheetId ||
        leftMutation.kind !== 'setCellValue' ||
        rightMutation.kind !== 'setCellValue' ||
        leftMutation.row !== rightMutation.row ||
        leftMutation.row <= previousRow ||
        leftMutation.col !== firstMutation.col ||
        rightMutation.col !== secondMutation.col ||
        typeof leftMutation.value !== 'number' ||
        typeof rightMutation.value !== 'number' ||
        Object.is(leftMutation.value, -0) ||
        Object.is(rightMutation.value, -0) ||
        leftRef.cellIndex === undefined ||
        rightRef.cellIndex === undefined ||
        cellStore.sheetIds[leftRef.cellIndex] !== leftRef.sheetId ||
        cellStore.rows[leftRef.cellIndex] !== leftMutation.row ||
        cellStore.cols[leftRef.cellIndex] !== leftMutation.col ||
        cellStore.sheetIds[rightRef.cellIndex] !== rightRef.sheetId ||
        cellStore.rows[rightRef.cellIndex] !== rightMutation.row ||
        cellStore.cols[rightRef.cellIndex] !== rightMutation.col
      ) {
        return false
      }
      previousRow = leftMutation.row
      const leftValue = leftMutation.value
      const rightValue = rightMutation.value
      const leftIndex = leftRef.cellIndex
      const rightIndex = rightRef.cellIndex
      if (!canFastPathLiteralOverwrite(leftIndex) || !canFastPathLiteralOverwrite(rightIndex)) {
        return false
      }
      const rowFormulaStart = formulaCount
      const considerDependent = (formulaCellIndex: number): boolean => {
        for (let index = rowFormulaStart; index < formulaCount; index += 1) {
          if (formulaCellIndices[index] === formulaCellIndex) {
            return true
          }
        }
        const formula = args.state.formulas.get(formulaCellIndex)
        if (
          !formula ||
          formula.directScalar === undefined ||
          cellStore.sheetIds[formulaCellIndex] !== firstRef.sheetId ||
          cellStore.rows[formulaCellIndex] !== leftMutation.row ||
          !canUseDirectFormulaPostRecalc(formulaCellIndex) ||
          !canSkipFormulaColumnVersion(formulaCellIndex)
        ) {
          return false
        }
        const formulaRow = cellStore.rows[formulaCellIndex] ?? 0
        const formulaCol = cellStore.cols[formulaCellIndex] ?? 0
        if (formulaRow < previousFormulaRow || (formulaRow === previousFormulaRow && formulaCol <= previousFormulaCol)) {
          return false
        }
        const code = rowPairDirectScalarCode(formula.directScalar, leftIndex, rightIndex)
        const result = code === 0 ? undefined : evaluateRowPairDirectScalarCode(code, leftValue, rightValue)
        if (result === undefined || formulaCount >= formulaCellIndices.length) {
          return false
        }
        formulaCellIndices[formulaCount] = formulaCellIndex
        formulaNumericResults[formulaCount] = result
        formulaCount += 1
        previousFormulaRow = formulaRow
        previousFormulaCol = formulaCol
        return true
      }
      const leftDependents = args.getEntityDependents(makeCellEntity(leftIndex))
      const rightDependents = args.getEntityDependents(makeCellEntity(rightIndex))
      if (leftDependents.length === 0 || rightDependents.length === 0) {
        return false
      }
      for (let index = 0; index < leftDependents.length; index += 1) {
        if (!considerDependent(leftDependents[index]!)) {
          return false
        }
      }
      if (!sameDependentOrder(leftDependents, rightDependents)) {
        for (let index = 0; index < rightDependents.length; index += 1) {
          if (!considerDependent(rightDependents[index]!)) {
            return false
          }
        }
      }
      if (formulaCount !== rowFormulaStart + 2) {
        return false
      }
      inputCellIndices[refIndex] = leftIndex
      inputCellIndices[refIndex + 1] = rightIndex
    }
    if (formulaCount !== refs.length) {
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
    const clearAuthoredBlankFlag = ~CellFlags.AuthoredBlank
    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      for (let index = 0; index < refs.length; index += 1) {
        const ref = refs[index]!
        const cellIndex = ref.cellIndex!
        const mutation = ref.mutation
        if (mutation.kind !== 'setCellValue' || typeof mutation.value !== 'number') {
          throw new Error('Expected dense row-pair batch to contain only numeric literal writes')
        }
        const currentFlags = flags[cellIndex] ?? 0
        if ((currentFlags & CellFlags.AuthoredBlank) !== 0) {
          flags[cellIndex] = currentFlags & clearAuthoredBlankFlag
        }
        tags[cellIndex] = ValueTag.Number
        errors[cellIndex] = ErrorCode.None
        stringIds[cellIndex] = 0
        numbers[cellIndex] = mutation.value
        versions[cellIndex] = (versions[cellIndex] ?? 0) + 1
      }
      args.state.workbook.notifyColumnsWritten(firstRef.sheetId, Uint32Array.of(firstMutation.col, secondMutation.col))
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
    }
    if (batch) {
      markBatchApplied(args.state.replicaState, batch)
    }
    args.deferKernelSync(inputCellIndices)
    applyTerminalDirectFormulaNumericResults(formulaCellIndices, formulaNumericResults, formulaCount)
    addEngineCounter(args.state.counters, 'directScalarDeltaApplications', formulaCount)
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

  const tryApplyDenseRowPairDirectScalarLiteralBatch = (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    potentialNewCells?: number,
  ): boolean => {
    if (refs.length < 32 || refs.length % 2 !== 0 || (potentialNewCells ?? 0) !== 0) {
      return false
    }
    const firstRef = refs[0]!
    const secondRef = refs[1]!
    const firstMutation = firstRef.mutation
    const secondMutation = secondRef.mutation
    if (
      firstMutation.kind !== 'setCellValue' ||
      secondMutation.kind !== 'setCellValue' ||
      firstRef.sheetId !== secondRef.sheetId ||
      firstMutation.row !== secondMutation.row ||
      firstMutation.col >= secondMutation.col ||
      typeof firstMutation.value !== 'number' ||
      typeof secondMutation.value !== 'number' ||
      Object.is(firstMutation.value, -0) ||
      Object.is(secondMutation.value, -0)
    ) {
      return false
    }
    const sheet = args.state.workbook.getSheetById(firstRef.sheetId)
    if (
      !sheet ||
      sheet.structureVersion !== 1 ||
      hasTrackedExactLookupDependents(firstRef.sheetId, firstMutation.col) ||
      hasTrackedExactLookupDependents(firstRef.sheetId, secondMutation.col) ||
      hasTrackedSortedLookupDependents(firstRef.sheetId, firstMutation.col) ||
      hasTrackedSortedLookupDependents(firstRef.sheetId, secondMutation.col) ||
      hasTrackedDirectRangeDependents(firstRef.sheetId, firstMutation.col) ||
      hasTrackedDirectRangeDependents(firstRef.sheetId, secondMutation.col)
    ) {
      return false
    }
    const inputCellIndices = new Uint32Array(refs.length)
    const inputNumericValues = new Float64Array(refs.length)
    const formulaCellIndices = new Uint32Array(refs.length * 2)
    const formulaNumericResults = new Float64Array(refs.length * 2)
    const cellStore = args.state.workbook.cellStore
    let formulaCount = 0
    let previousRow = firstMutation.row - 1
    let previousFormulaRow = -1
    let previousFormulaCol = -1

    const readDirectScalarNumber = (cellIndex: number): number | undefined => directScalarCellNumber(cellStore, cellIndex)

    for (let refIndex = 0; refIndex < refs.length; refIndex += 2) {
      const leftRef = refs[refIndex]!
      const rightRef = refs[refIndex + 1]!
      const leftMutation = leftRef.mutation
      const rightMutation = rightRef.mutation
      if (
        leftRef.sheetId !== firstRef.sheetId ||
        rightRef.sheetId !== firstRef.sheetId ||
        leftMutation.kind !== 'setCellValue' ||
        rightMutation.kind !== 'setCellValue' ||
        leftMutation.row !== rightMutation.row ||
        leftMutation.row <= previousRow ||
        leftMutation.col !== firstMutation.col ||
        rightMutation.col !== secondMutation.col ||
        typeof leftMutation.value !== 'number' ||
        typeof rightMutation.value !== 'number' ||
        Object.is(leftMutation.value, -0) ||
        Object.is(rightMutation.value, -0)
      ) {
        return false
      }
      const leftValue = leftMutation.value
      const rightValue = rightMutation.value
      previousRow = leftMutation.row
      const leftIndex =
        leftRef.cellIndex !== undefined &&
        cellStore.sheetIds[leftRef.cellIndex] === leftRef.sheetId &&
        cellStore.rows[leftRef.cellIndex] === leftMutation.row &&
        cellStore.cols[leftRef.cellIndex] === leftMutation.col
          ? leftRef.cellIndex
          : sheet.grid.getPhysical(leftMutation.row, leftMutation.col)
      const rightIndex =
        rightRef.cellIndex !== undefined &&
        cellStore.sheetIds[rightRef.cellIndex] === rightRef.sheetId &&
        cellStore.rows[rightRef.cellIndex] === rightMutation.row &&
        cellStore.cols[rightRef.cellIndex] === rightMutation.col
          ? rightRef.cellIndex
          : sheet.grid.getPhysical(rightMutation.row, rightMutation.col)
      if (leftIndex === -1 || rightIndex === -1 || !canFastPathLiteralOverwrite(leftIndex) || !canFastPathLiteralOverwrite(rightIndex)) {
        return false
      }
      const rowFormulaStart = formulaCount
      const considerDependent = (formulaCellIndex: number): boolean => {
        for (let index = rowFormulaStart; index < formulaCount; index += 1) {
          if (formulaCellIndices[index] === formulaCellIndex) {
            return true
          }
        }
        const formula = args.state.formulas.get(formulaCellIndex)
        if (
          !formula ||
          formula.directScalar === undefined ||
          cellStore.sheetIds[formulaCellIndex] !== firstRef.sheetId ||
          cellStore.rows[formulaCellIndex] !== leftMutation.row ||
          !canUseDirectFormulaPostRecalc(formulaCellIndex) ||
          !canSkipFormulaColumnVersion(formulaCellIndex)
        ) {
          return false
        }
        const formulaRow = cellStore.rows[formulaCellIndex] ?? 0
        const formulaCol = cellStore.cols[formulaCellIndex] ?? 0
        if (formulaRow < previousFormulaRow || (formulaRow === previousFormulaRow && formulaCol <= previousFormulaCol)) {
          return false
        }
        const result = evaluateDirectScalarWithReplacementNumbers(
          formula.directScalar,
          leftIndex,
          leftValue,
          readDirectScalarNumber,
          rightIndex,
          rightValue,
        )
        if (result === undefined || formulaCount >= formulaCellIndices.length) {
          return false
        }
        formulaCellIndices[formulaCount] = formulaCellIndex
        formulaNumericResults[formulaCount] = result
        formulaCount += 1
        previousFormulaRow = formulaRow
        previousFormulaCol = formulaCol
        return true
      }
      const leftDependents = args.getEntityDependents(makeCellEntity(leftIndex))
      const rightDependents = args.getEntityDependents(makeCellEntity(rightIndex))
      if (leftDependents.length === 0 || rightDependents.length === 0) {
        return false
      }
      for (let index = 0; index < leftDependents.length; index += 1) {
        if (!considerDependent(leftDependents[index]!)) {
          return false
        }
      }
      if (!sameDependentOrder(leftDependents, rightDependents)) {
        for (let index = 0; index < rightDependents.length; index += 1) {
          if (!considerDependent(rightDependents[index]!)) {
            return false
          }
        }
      }
      inputCellIndices[refIndex] = leftIndex
      inputCellIndices[refIndex + 1] = rightIndex
      inputNumericValues[refIndex] = leftValue
      inputNumericValues[refIndex + 1] = rightValue
    }
    if (formulaCount === 0) {
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
    const clearAuthoredBlankFlag = ~CellFlags.AuthoredBlank
    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      for (let index = 0; index < refs.length; index += 1) {
        const cellIndex = inputCellIndices[index]!
        const currentFlags = flags[cellIndex] ?? 0
        if ((currentFlags & CellFlags.AuthoredBlank) !== 0) {
          flags[cellIndex] = currentFlags & clearAuthoredBlankFlag
        }
        tags[cellIndex] = ValueTag.Number
        errors[cellIndex] = ErrorCode.None
        stringIds[cellIndex] = 0
        numbers[cellIndex] = inputNumericValues[index]!
        versions[cellIndex] = (versions[cellIndex] ?? 0) + 1
      }
      args.state.workbook.notifyColumnsWritten(firstRef.sheetId, Uint32Array.of(firstMutation.col, secondMutation.col))
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
    }
    if (batch) {
      markBatchApplied(args.state.replicaState, batch)
    }
    args.deferKernelSync(inputCellIndices)
    const formulaChanged = requiresChangedSet ? new Uint32Array(formulaCount) : EMPTY_CHANGED_CELLS
    for (let index = 0; index < formulaCount; index += 1) {
      if (requiresChangedSet) {
        formulaChanged[index] = formulaCellIndices[index]!
      }
    }
    applyTerminalDirectFormulaNumericResults(formulaCellIndices, formulaNumericResults, formulaCount)
    addEngineCounter(args.state.counters, 'directScalarDeltaApplications', formulaCount)
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

  return { tryApplyDenseRowPairSimpleDirectScalarLiteralBatch, tryApplyDenseRowPairDirectScalarLiteralBatch }
}

function composeDenseInputFormulaChanges(inputCellIndices: Uint32Array, formulaCellIndices: Uint32Array): Uint32Array {
  const changed = new Uint32Array(inputCellIndices.length + formulaCellIndices.length)
  changed.set(inputCellIndices)
  changed.set(formulaCellIndices, inputCellIndices.length)
  return changed
}

function sameDependentOrder(left: Uint32Array, right: Uint32Array): boolean {
  if (left === right) {
    return true
  }
  if (left.length !== right.length) {
    return false
  }
  for (let index = 0; index < left.length; index += 1) {
    if (left[index] !== right[index]) {
      return false
    }
  }
  return true
}

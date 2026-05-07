import type { EngineChangedCell } from '@bilig/protocol'
import type { EngineOpBatch } from '@bilig/workbook-domain'
import type { EngineCellMutationRef } from '../../cell-mutations-at.js'
import { makeCellEntity } from '../../entity-ids.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import { markBatchApplied } from '../../replica-state.js'
import type { EngineRuntimeState, U32 } from '../runtime-state.js'
import {
  directScalarCellNumber,
  evaluateDirectScalarWithReplacementNumbers,
  evaluateRowPairDirectScalarCode,
  rowPairDirectScalarCode,
  rowPairDirectScalarCodeNeedsZeroGuard,
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
  readonly writeNumericLiteralToCellStore: (cellIndex: number, value: number) => void
  readonly applyTerminalDirectFormulaNumericResult: (cellIndex: number, value: number) => void
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
    writeNumericLiteralToCellStore,
    applyTerminalDirectFormulaNumericResult,
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

    const formulaCellIndices = new Uint32Array(refs.length)
    const formulaCodes = new Uint8Array(refs.length)
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
        if (
          code === 0 ||
          (rowPairDirectScalarCodeNeedsZeroGuard(code) && evaluateRowPairDirectScalarCode(code, leftValue, rightValue) === undefined) ||
          formulaCount >= formulaCellIndices.length
        ) {
          return false
        }
        formulaCellIndices[formulaCount] = formulaCellIndex
        formulaCodes[formulaCount] = code
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
      for (let index = 0; index < rightDependents.length; index += 1) {
        if (!considerDependent(rightDependents[index]!)) {
          return false
        }
      }
      if (formulaCount !== rowFormulaStart + 2) {
        return false
      }
    }
    if (formulaCount !== refs.length) {
      return false
    }

    args.materializeDeferredStructuralFormulaSources()
    args.beginMutationCollection()
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + refs.length * 2 + 1)
    args.resetMaterializedCellScratch(0)

    const hasGeneralEventListeners = args.state.events.hasListeners()
    const hasTrackedEventListeners = args.state.events.hasTrackedListeners()
    const hasWatchedCellListeners = args.state.events.hasCellListeners()
    const requiresChangedSet = hasGeneralEventListeners || hasTrackedEventListeners || hasWatchedCellListeners
    let changedInputCount = 0
    let explicitChangedCount = 0
    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      for (let index = 0; index < refs.length; index += 1) {
        const ref = refs[index]!
        const cellIndex = ref.cellIndex!
        const mutation = ref.mutation
        if (mutation.kind !== 'setCellValue' || typeof mutation.value !== 'number') {
          throw new Error('Expected dense row-pair batch to contain only numeric literal writes')
        }
        writeNumericLiteralToCellStore(cellIndex, mutation.value)
        changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
        if (requiresChangedSet) {
          explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
        }
      }
      args.state.workbook.notifyColumnsWritten(firstRef.sheetId, Uint32Array.of(firstMutation.col, secondMutation.col))
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
    }
    if (batch) {
      markBatchApplied(args.state.replicaState, batch)
    }
    const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
    args.deferKernelSync(changedInputArray)
    for (let refIndex = 0; refIndex < refs.length; refIndex += 2) {
      const leftMutation = refs[refIndex]!.mutation
      const rightMutation = refs[refIndex + 1]!.mutation
      if (
        leftMutation.kind !== 'setCellValue' ||
        rightMutation.kind !== 'setCellValue' ||
        typeof leftMutation.value !== 'number' ||
        typeof rightMutation.value !== 'number'
      ) {
        throw new Error('Expected dense row-pair batch to contain only numeric literal writes')
      }
      const leftValue = leftMutation.value
      const rightValue = rightMutation.value
      for (let formulaIndex = refIndex; formulaIndex < refIndex + 2; formulaIndex += 1) {
        const result = evaluateRowPairDirectScalarCode(formulaCodes[formulaIndex]!, leftValue, rightValue)
        if (result === undefined) {
          throw new Error('Failed to apply direct row-pair scalar result')
        }
        applyTerminalDirectFormulaNumericResult(formulaCellIndices[formulaIndex]!, result)
      }
    }
    addEngineCounter(args.state.counters, 'directScalarDeltaApplications', formulaCount)
    addEngineCounter(args.state.counters, 'directScalarDeltaOnlyRecalcSkips')
    const changed = requiresChangedSet ? args.composeDisjointEventChanges(formulaCellIndices, explicitChangedCount) : EMPTY_CHANGED_CELLS
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
    if (refs.length < 32 || refs.length % 2 !== 0) {
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
      for (let index = 0; index < rightDependents.length; index += 1) {
        if (!considerDependent(rightDependents[index]!)) {
          return false
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
    args.beginMutationCollection()
    const reservedNewCells = potentialNewCells ?? 0
    args.state.workbook.cellStore.ensureCapacity(args.state.workbook.cellStore.size + reservedNewCells)
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + refs.length + formulaCount + 1)
    args.resetMaterializedCellScratch(reservedNewCells)

    const hasGeneralEventListeners = args.state.events.hasListeners()
    const hasTrackedEventListeners = args.state.events.hasTrackedListeners()
    const hasWatchedCellListeners = args.state.events.hasCellListeners()
    const requiresChangedSet = hasGeneralEventListeners || hasTrackedEventListeners || hasWatchedCellListeners
    let changedInputCount = 0
    let explicitChangedCount = 0
    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      for (let index = 0; index < refs.length; index += 1) {
        const cellIndex = inputCellIndices[index]!
        writeNumericLiteralToCellStore(cellIndex, inputNumericValues[index]!)
        changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
        if (requiresChangedSet) {
          explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
        }
      }
      args.state.workbook.notifyColumnsWritten(firstRef.sheetId, Uint32Array.of(firstMutation.col, secondMutation.col))
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
    }
    if (batch) {
      markBatchApplied(args.state.replicaState, batch)
    }
    const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
    args.deferKernelSync(changedInputArray)
    const formulaChanged = requiresChangedSet ? new Uint32Array(formulaCount) : EMPTY_CHANGED_CELLS
    for (let index = 0; index < formulaCount; index += 1) {
      const formulaCellIndex = formulaCellIndices[index]!
      applyTerminalDirectFormulaNumericResult(formulaCellIndex, formulaNumericResults[index]!)
      if (requiresChangedSet) {
        formulaChanged[index] = formulaCellIndex
      }
    }
    addEngineCounter(args.state.counters, 'directScalarDeltaApplications', formulaCount)
    addEngineCounter(args.state.counters, 'directScalarDeltaOnlyRecalcSkips')
    const changed = requiresChangedSet ? args.composeDisjointEventChanges(formulaChanged, explicitChangedCount) : EMPTY_CHANGED_CELLS
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

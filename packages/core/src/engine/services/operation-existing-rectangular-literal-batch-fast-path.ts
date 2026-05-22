import { ValueTag, type EngineChangedCell } from '@bilig/protocol'
import type { EngineOpBatch } from '@bilig/workbook'
import type { EngineCellMutationRef } from '../../cell-mutations-at.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import { markBatchApplied } from '../../replica-state.js'
import type { EngineRuntimeState, U32 } from '../runtime-state.js'
import { tagTrustedPhysicalTrackedChanges } from './operation-change-helpers.js'
import { emitCellMutationFastPathBatchResult } from './operation-fast-path-batch-result.js'
import { isTableHeaderCell } from './operation-table-header-rename.js'

type FastPathState = Pick<
  EngineRuntimeState,
  'workbook' | 'events' | 'formulas' | 'counters' | 'replicaState' | 'getLastMetrics' | 'setLastMetrics' | 'getSyncClientConnection'
>

export interface OperationExistingRectangularLiteralBatchFastPathArgs {
  readonly state: FastPathState
  readonly emitBatch: (batch: EngineOpBatch) => void
  readonly canFastPathLiteralOverwrite: (cellIndex: number) => boolean
  readonly writeNumericLiteralToCellStore: (cellIndex: number, value: number) => void
  readonly materializeDeferredStructuralFormulaSources: () => void
  readonly getBatchMutationDepth: () => number
  readonly setBatchMutationDepth: (next: number) => void
  readonly deferKernelSync: (cellIndices: readonly number[] | U32) => void
  readonly captureChangedCells: (changedCellIndices: readonly number[] | U32) => readonly EngineChangedCell[]
}

interface ExistingNumericRectangle {
  readonly sheetId: number
  readonly firstCol: number
  readonly colCount: number
  readonly cellIndices: Uint32Array
  readonly values: Float64Array
}

const EMPTY_CHANGED_CELLS = new Uint32Array(0)

export function createOperationExistingRectangularLiteralBatchFastPath(args: OperationExistingRectangularLiteralBatchFastPathArgs): {
  readonly tryApplyExistingDenseRectangularNumericLiteralBatch: (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    potentialNewCells?: number,
  ) => boolean
} {
  const tryApplyExistingDenseRectangularNumericLiteralBatch = (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    potentialNewCells?: number,
  ): boolean => {
    if (refs.length < 32 || (potentialNewCells ?? 0) !== 0 || args.state.workbook.hasPivots()) {
      return false
    }
    if (args.state.formulas.size !== 0 || args.state.workbook.listDefinedNames().length !== 0) {
      return false
    }
    args.materializeDeferredStructuralFormulaSources()
    const rectangle = collectExistingDenseNumericRectangle(args, refs)
    if (rectangle === null) {
      return false
    }

    const hasGeneralEventListeners = args.state.events.hasListeners()
    const hasTrackedEventListeners = args.state.events.hasTrackedListeners()
    const hasWatchedCellListeners = args.state.events.hasCellListeners()
    const requiresChangedSet = hasGeneralEventListeners || hasTrackedEventListeners || hasWatchedCellListeners
    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      for (let index = 0; index < rectangle.cellIndices.length; index += 1) {
        args.writeNumericLiteralToCellStore(rectangle.cellIndices[index]!, rectangle.values[index]!)
      }
      const writtenColumns = new Uint32Array(rectangle.colCount)
      for (let colOffset = 0; colOffset < rectangle.colCount; colOffset += 1) {
        writtenColumns[colOffset] = rectangle.firstCol + colOffset
      }
      args.state.workbook.notifyColumnsWritten(rectangle.sheetId, writtenColumns)
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
    }

    if (batch) {
      markBatchApplied(args.state.replicaState, batch)
    }
    args.deferKernelSync(rectangle.cellIndices)
    addEngineCounter(args.state.counters, 'kernelSyncOnlyRecalcSkips')

    const changed = requiresChangedSet ? rectangle.cellIndices : EMPTY_CHANGED_CELLS
    if (requiresChangedSet && hasTrackedEventListeners) {
      tagTrustedPhysicalTrackedChanges(changed, rectangle.sheetId, rectangle.cellIndices.length)
    }
    emitCellMutationFastPathBatchResult({
      state: args.state,
      changed,
      changedInputCount: rectangle.cellIndices.length,
      explicitChangedCount: requiresChangedSet ? rectangle.cellIndices.length : 0,
      hasGeneralEventListeners,
      hasTrackedEventListeners,
      hasWatchedCellListeners,
      captureChangedCells: args.captureChangedCells,
      batch,
      emitBatch: args.emitBatch,
    })
    return true
  }

  return { tryApplyExistingDenseRectangularNumericLiteralBatch }
}

function collectExistingDenseNumericRectangle(
  args: OperationExistingRectangularLiteralBatchFastPathArgs,
  refs: readonly EngineCellMutationRef[],
): ExistingNumericRectangle | null {
  const firstRef = refs[0]
  const firstMutation = firstRef?.mutation
  if (
    firstRef === undefined ||
    firstMutation === undefined ||
    firstMutation.kind !== 'setCellValue' ||
    typeof firstMutation.value !== 'number' ||
    Object.is(firstMutation.value, -0)
  ) {
    return null
  }
  const sheet = args.state.workbook.getSheetById(firstRef.sheetId)
  if (!sheet || sheet.structureVersion !== 1) {
    return null
  }

  const cellStore = args.state.workbook.cellStore
  const tables = args.state.workbook.hasTables() ? args.state.workbook.listTables() : []
  const cellIndices = new Uint32Array(refs.length)
  const values = new Float64Array(refs.length)
  const firstRow = firstMutation.row
  const firstCol = firstMutation.col
  let currentRow = firstRow
  let currentWidth = 0
  let colCount = 0
  for (let refIndex = 0; refIndex < refs.length; refIndex += 1) {
    const ref = refs[refIndex]!
    const mutation = ref.mutation
    if (
      ref.sheetId !== firstRef.sheetId ||
      mutation.kind !== 'setCellValue' ||
      typeof mutation.value !== 'number' ||
      Object.is(mutation.value, -0)
    ) {
      return null
    }
    if (mutation.row === currentRow) {
      if (mutation.col !== firstCol + currentWidth) {
        return null
      }
      currentWidth += 1
    } else {
      if (mutation.row !== currentRow + 1 || mutation.col !== firstCol || currentWidth === 0) {
        return null
      }
      if (colCount === 0) {
        colCount = currentWidth
      } else if (currentWidth !== colCount) {
        return null
      }
      currentRow = mutation.row
      currentWidth = 1
    }

    const candidate = ref.cellIndex
    const cellIndex =
      candidate !== undefined &&
      cellStore.sheetIds[candidate] === ref.sheetId &&
      cellStore.rows[candidate] === mutation.row &&
      cellStore.cols[candidate] === mutation.col
        ? candidate
        : sheet.grid.getPhysical(mutation.row, mutation.col)
    if (cellIndex === -1 || !args.canFastPathLiteralOverwrite(cellIndex) || cellStore.tags[cellIndex] !== ValueTag.Number) {
      return null
    }
    if (isTableHeaderCell(tables, sheet.name, mutation.row, mutation.col)) {
      return null
    }
    cellIndices[refIndex] = cellIndex
    values[refIndex] = mutation.value
  }
  if (colCount === 0) {
    colCount = currentWidth
  } else if (currentWidth !== colCount) {
    return null
  }
  if (colCount === 0) {
    return null
  }
  return { sheetId: firstRef.sheetId, firstCol, colCount, cellIndices, values }
}

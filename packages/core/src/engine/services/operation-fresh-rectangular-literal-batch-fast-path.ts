import type { EngineChangedCell } from '@bilig/protocol'
import type { EngineOpBatch } from '@bilig/workbook-domain'
import type { EngineCellMutationRef } from '../../cell-mutations-at.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import { markBatchApplied } from '../../replica-state.js'
import type { SheetRecord } from '../../workbook-store.js'
import type {
  EngineRuntimeState,
  RuntimeDirectAggregateDescriptor,
  RuntimeDirectCriteriaDescriptor,
  RuntimeFormula,
  U32,
} from '../runtime-state.js'
import { emitCellMutationFastPathBatchResult } from './operation-fast-path-batch-result.js'

type FastPathState = Pick<
  EngineRuntimeState,
  | 'workbook'
  | 'ranges'
  | 'events'
  | 'formulas'
  | 'counters'
  | 'replicaState'
  | 'getLastMetrics'
  | 'setLastMetrics'
  | 'getSyncClientConnection'
>

export interface OperationFreshRectangularLiteralBatchFastPathArgs {
  readonly state: FastPathState
  readonly emitBatch: (batch: EngineOpBatch) => void
  readonly hasTrackedExactLookupDependents: (sheetId: number, col: number) => boolean
  readonly hasTrackedSortedLookupDependents: (sheetId: number, col: number) => boolean
  readonly writeNumericLiteralToCellStore: (cellIndex: number, value: number) => void
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
  readonly captureChangedCells: (changedCellIndices: readonly number[] | U32) => readonly EngineChangedCell[]
}

interface FreshNumericRectangle {
  readonly sheet: NonNullable<ReturnType<FastPathState['workbook']['getSheetById']>>
  readonly sheetId: number
  readonly firstRow: number
  readonly rowCount: number
  readonly firstCol: number
  readonly colCount: number
  readonly values: Float64Array
}

const EMPTY_CHANGED_CELLS = new Uint32Array(0)

export function createOperationFreshRectangularLiteralBatchFastPath(args: OperationFreshRectangularLiteralBatchFastPathArgs): {
  readonly tryApplyFreshDenseRectangularNumericLiteralBatch: (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    potentialNewCells?: number,
  ) => boolean
} {
  const tryApplyFreshDenseRectangularNumericLiteralBatch = (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    potentialNewCells?: number,
  ): boolean => {
    const firstRef = refs[0]
    if (firstRef === undefined || refs.length < 32 || potentialNewCells !== refs.length) {
      return false
    }
    const rectangle = collectFreshDenseNumericRectangle(args, refs, firstRef)
    if (rectangle === null || rectangleOverlapsFormulaDependencies(args, rectangle)) {
      return false
    }
    for (let col = rectangle.firstCol; col < rectangle.firstCol + rectangle.colCount; col += 1) {
      if (args.hasTrackedExactLookupDependents(rectangle.sheetId, col) || args.hasTrackedSortedLookupDependents(rectangle.sheetId, col)) {
        return false
      }
    }

    args.materializeDeferredStructuralFormulaSources()
    args.beginMutationCollection()
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + refs.length + 1)
    args.resetMaterializedCellScratch(refs.length)

    const hasGeneralEventListeners = args.state.events.hasListeners()
    const hasTrackedEventListeners = args.state.events.hasTrackedListeners()
    const hasWatchedCellListeners = args.state.events.hasCellListeners()
    const requiresChangedSet = hasGeneralEventListeners || hasTrackedEventListeners || hasWatchedCellListeners
    let changedInputCount = 0
    let explicitChangedCount = 0
    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    const ensureRowId = args.state.workbook.createLogicalAxisIdEnsurer(rectangle.sheetId, 'row')
    const ensureColId = args.state.workbook.createLogicalAxisIdEnsurer(rectangle.sheetId, 'column')
    const rowIds = materializeAxisIds(rectangle.rowCount, rectangle.firstRow, ensureRowId)
    const columnIds = materializeAxisIds(rectangle.colCount, rectangle.firstCol, ensureColId)
    const firstCellIndex = args.state.workbook.cellStore.allocateDenseRowMajorAtReserved(
      rectangle.sheetId,
      rectangle.firstRow,
      rectangle.rowCount,
      rectangle.firstCol,
      rectangle.colCount,
    )
    attachFreshDenseNumericCells(rectangle.sheet, firstCellIndex, rectangle.firstRow, rectangle.firstCol, rowIds, columnIds)
    try {
      let valueIndex = 0
      for (let rowOffset = 0; rowOffset < rectangle.rowCount; rowOffset += 1) {
        for (let colOffset = 0; colOffset < rectangle.colCount; colOffset += 1) {
          const cellIndex = firstCellIndex + rowOffset * rectangle.colCount + colOffset
          args.writeNumericLiteralToCellStore(cellIndex, rectangle.values[valueIndex]!)
          changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
          if (requiresChangedSet) {
            explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
          }
          valueIndex += 1
        }
      }
      const writtenColumns = new Uint32Array(rectangle.colCount)
      for (let index = 0; index < rectangle.colCount; index += 1) {
        writtenColumns[index] = rectangle.firstCol + index
      }
      args.state.workbook.notifyColumnsWritten(rectangle.sheetId, writtenColumns)
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
    }
    if (batch) {
      markBatchApplied(args.state.replicaState, batch)
    }
    const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
    args.deferKernelSync(changedInputArray)
    addEngineCounter(args.state.counters, 'kernelSyncOnlyRecalcSkips')

    emitCellMutationFastPathBatchResult({
      state: args.state,
      changed: requiresChangedSet ? changedInputArray : EMPTY_CHANGED_CELLS,
      changedInputCount,
      explicitChangedCount,
      hasGeneralEventListeners,
      hasTrackedEventListeners,
      hasWatchedCellListeners,
      captureChangedCells: args.captureChangedCells,
      batch,
      emitBatch: args.emitBatch,
    })
    return true
  }

  return { tryApplyFreshDenseRectangularNumericLiteralBatch }
}

function collectFreshDenseNumericRectangle(
  args: OperationFreshRectangularLiteralBatchFastPathArgs,
  refs: readonly EngineCellMutationRef[],
  firstRef: EngineCellMutationRef,
): FreshNumericRectangle | null {
  const firstMutation = firstRef.mutation
  if (firstMutation.kind !== 'setCellValue' || typeof firstMutation.value !== 'number' || Object.is(firstMutation.value, -0)) {
    return null
  }
  const sheet = args.state.workbook.getSheetById(firstRef.sheetId)
  if (!sheet) {
    return null
  }
  const values = new Float64Array(refs.length)
  const firstRow = firstMutation.row
  const firstCol = firstMutation.col
  let currentRow = firstRow
  let currentWidth = 0
  let colCount = 0
  let rowOffset = 0
  for (let refIndex = 0; refIndex < refs.length; refIndex += 1) {
    const ref = refs[refIndex]!
    const mutation = ref.mutation
    if (
      ref.sheetId !== firstRef.sheetId ||
      ref.cellIndex !== undefined ||
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
      rowOffset += 1
    }
    if (physicalSheetCellExists(sheet, mutation.row, mutation.col)) {
      return null
    }
    values[refIndex] = mutation.value
  }
  if (colCount === 0) {
    colCount = currentWidth
  } else if (currentWidth !== colCount) {
    return null
  }
  const rowCount = rowOffset + 1
  if (rowCount < 2 || colCount < 2 || rowCount * colCount !== refs.length) {
    return null
  }
  return { sheet, sheetId: firstRef.sheetId, firstRow, rowCount, firstCol, colCount, values }
}

function physicalSheetCellExists(sheet: SheetRecord, row: number, col: number): boolean {
  if (sheet.structureVersion === 1) {
    return sheet.grid.getPhysical(row, col) !== -1
  }
  return sheet.grid.getPhysical(row, col) !== -1 || sheet.logical.getVisibleCell(row, col) !== undefined
}

function materializeAxisIds(count: number, start: number, ensureAxisId: (index: number) => string): string[] {
  const axisIds: string[] = []
  axisIds.length = count
  for (let offset = 0; offset < count; offset += 1) {
    axisIds[offset] = ensureAxisId(start + offset)
  }
  return axisIds
}

function attachFreshDenseNumericCells(
  sheet: SheetRecord,
  firstCellIndex: number,
  firstRow: number,
  firstCol: number,
  rowIds: readonly string[],
  colIds: readonly string[],
): void {
  sheet.logical.deferVisibleCellPageRebuild()
  sheet.logical.setFreshVisibleDenseRowMajorIdentitiesWithAxisIdsDeferred(firstCellIndex, rowIds, colIds)
  sheet.grid.setDenseRowMajor(firstRow, firstCol, rowIds.length, colIds.length, firstCellIndex)
}

function rectangleOverlapsFormulaDependencies(
  args: OperationFreshRectangularLiteralBatchFastPathArgs,
  rectangle: FreshNumericRectangle,
): boolean {
  const rowEnd = rectangle.firstRow + rectangle.rowCount - 1
  const colEnd = rectangle.firstCol + rectangle.colCount - 1
  let overlaps = false
  args.state.formulas.forEach((formula) => {
    if (overlaps) {
      return
    }
    overlaps =
      directAggregateOverlapsRectangle(
        formula.directAggregate,
        rectangle.sheet.name,
        rectangle.firstRow,
        rowEnd,
        rectangle.firstCol,
        colEnd,
      ) ||
      directCriteriaOverlapsRectangle(
        formula.directCriteria,
        rectangle.sheet.name,
        rectangle.firstRow,
        rowEnd,
        rectangle.firstCol,
        colEnd,
      ) ||
      rangeDependenciesOverlapRectangle(args, formula, rectangle.sheetId, rectangle.firstRow, rowEnd, rectangle.firstCol, colEnd)
  })
  return overlaps
}

function directAggregateOverlapsRectangle(
  aggregate: RuntimeDirectAggregateDescriptor | undefined,
  sheetName: string,
  rowStart: number,
  rowEnd: number,
  colStart: number,
  colEnd: number,
): boolean {
  return (
    aggregate !== undefined &&
    aggregate.sheetName === sheetName &&
    aggregate.rowEnd >= rowStart &&
    aggregate.rowStart <= rowEnd &&
    aggregate.colEnd >= colStart &&
    aggregate.col <= colEnd
  )
}

function directCriteriaOverlapsRectangle(
  criteria: RuntimeDirectCriteriaDescriptor | undefined,
  sheetName: string,
  rowStart: number,
  rowEnd: number,
  colStart: number,
  colEnd: number,
): boolean {
  if (criteria === undefined) {
    return false
  }
  const rangeOverlaps = (range: {
    readonly sheetName: string
    readonly rowStart: number
    readonly rowEnd: number
    readonly col: number
  }): boolean =>
    range.sheetName === sheetName && range.rowEnd >= rowStart && range.rowStart <= rowEnd && range.col >= colStart && range.col <= colEnd
  return (
    (criteria.aggregateRange !== undefined && rangeOverlaps(criteria.aggregateRange)) ||
    criteria.criteriaPairs.some((pair) => rangeOverlaps(pair.range))
  )
}

function rangeDependenciesOverlapRectangle(
  args: OperationFreshRectangularLiteralBatchFastPathArgs,
  formula: RuntimeFormula,
  sheetId: number,
  rowStart: number,
  rowEnd: number,
  colStart: number,
  colEnd: number,
): boolean {
  for (let index = 0; index < formula.rangeDependencies.length; index += 1) {
    if (rangeDescriptorOverlapsRectangle(args, formula.rangeDependencies[index]!, sheetId, rowStart, rowEnd, colStart, colEnd)) {
      return true
    }
  }
  for (let index = 0; index < formula.graphRangeDependencies.length; index += 1) {
    if (rangeDescriptorOverlapsRectangle(args, formula.graphRangeDependencies[index]!, sheetId, rowStart, rowEnd, colStart, colEnd)) {
      return true
    }
  }
  return false
}

function rangeDescriptorOverlapsRectangle(
  args: OperationFreshRectangularLiteralBatchFastPathArgs,
  rangeIndex: number,
  sheetId: number,
  rowStart: number,
  rowEnd: number,
  colStart: number,
  colEnd: number,
): boolean {
  const descriptor = args.state.ranges.getDescriptor(rangeIndex)
  return (
    descriptor.sheetId === sheetId &&
    descriptor.row2 >= rowStart &&
    descriptor.row1 <= rowEnd &&
    descriptor.col2 >= colStart &&
    descriptor.col1 <= colEnd
  )
}

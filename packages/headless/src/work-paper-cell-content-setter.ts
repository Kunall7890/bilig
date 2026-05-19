import type { EngineCellMutationRef, EngineExistingNumericCellMutationResult, SheetRecord, SpreadsheetEngine } from '@bilig/core'
import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import { WorkPaperOperationError } from './work-paper-errors.js'
import {
  assertRowAndColumn,
  isBlankRawCellContent,
  isParsableFormulaContent,
  isWorkPaperSheetMatrix,
} from './work-paper-runtime-helpers.js'
import { buildWorkPaperLiteralCellValueMutation, buildWorkPaperRawCellMutation } from './work-paper-literal-mutation-queue.js'
import type { WorkPaperCellMutationApplyOptions } from './work-paper-cell-mutation-refs.js'
import type {
  RawCellContent,
  WorkPaperCellAddress,
  WorkPaperCellValueUpdate,
  WorkPaperChange,
  WorkPaperConfig,
  WorkPaperSheet,
  WorkPaperSheetCellValueUpdate,
  WorkPaperSheetRangeValues,
} from './work-paper-types.js'

interface ExistingNumericMutationEngine {
  readonly tryApplyExistingNumericCellMutationAt?: (request: {
    readonly sheetId: number
    readonly row: number
    readonly col: number
    readonly cellIndex: number
    readonly value: number
  }) => EngineExistingNumericCellMutationResult | null
}

export interface WorkPaperSetCellContentsRuntime {
  readonly assertNotDisposed: () => void
  readonly getConfig: () => Pick<WorkPaperConfig, 'maxRows' | 'maxColumns'>
  readonly getEngine: () => SpreadsheetEngine
  readonly sheetRecord: (sheetId: number) => SheetRecord
  readonly getVisibleCellIndexInSheet: (sheet: SheetRecord, row: number, col: number) => number | undefined
  readonly isEvaluationSuspended: () => boolean
  readonly getBatchDepth: () => number
  readonly enqueueSuspendedLiteralMutation: (
    sheetId: number,
    row: number,
    col: number,
    content: RawCellContent,
    cellIndex: number | undefined,
  ) => boolean
  readonly enqueueDeferredBatchLiteral: (
    sheetId: number,
    row: number,
    col: number,
    content: RawCellContent,
    cellIndex: number | undefined,
  ) => boolean
  readonly getCellSerialized: (address: WorkPaperCellAddress) => RawCellContent
  readonly trySetExistingNumericCellContentsWithTrackedFastPath: (args: {
    readonly sheet: SheetRecord
    readonly address: WorkPaperCellAddress
    readonly cellIndex: number
    readonly value: number
  }) => WorkPaperChange[] | null
  readonly trySetExistingLiteralCellContentsWithTrackedFastPath: (args: {
    readonly sheet: SheetRecord
    readonly address: WorkPaperCellAddress
    readonly cellIndex: number
    readonly value: Exclude<RawCellContent, null>
  }) => WorkPaperChange[] | null
  readonly flushPendingBatchOps: () => void
  readonly rewriteFormulaForStorage: (formula: string, ownerSheetId: number) => string
  readonly applyCellMutationRefs: (refs: readonly EngineCellMutationRef[], options: WorkPaperCellMutationApplyOptions) => void
  readonly canUseTrackedMutationFastPath: () => boolean
  readonly isTrackedBatchFastPathActive: () => boolean
  readonly captureTrackedChangesWithoutVisibilityCache: (
    mutate: () => void,
    options: {
      readonly singleLiteralChange?: {
        readonly address: WorkPaperCellAddress
        readonly cellIndex?: number
        readonly isPhysicalSheet: boolean
        readonly sheetName: string
        readonly value: RawCellContent
      }
    },
  ) => WorkPaperChange[]
  readonly captureChanges: (mutate: () => void) => WorkPaperChange[]
  readonly isItPossibleToSetCellContents: (address: WorkPaperCellAddress, content: RawCellContent | WorkPaperSheet) => boolean
  readonly applyMatrixContents: (address: WorkPaperCellAddress, content: WorkPaperSheet) => void
}

export function setWorkPaperCellContents(
  runtime: WorkPaperSetCellContentsRuntime,
  address: WorkPaperCellAddress,
  content: RawCellContent | WorkPaperSheet,
): WorkPaperChange[] {
  runtime.assertNotDisposed()
  const sheet = runtime.sheetRecord(address.sheet)
  assertRowAndColumn(address.row, 'address.row')
  assertRowAndColumn(address.col, 'address.col')
  if (!isWorkPaperSheetMatrix(content)) {
    const config = runtime.getConfig()
    if (address.row >= (config.maxRows ?? MAX_ROWS) || address.col >= (config.maxColumns ?? MAX_COLS)) {
      throw new WorkPaperOperationError('Cell contents cannot be set')
    }
    const visibleCellIndex = runtime.getVisibleCellIndexInSheet(sheet, address.row, address.col)
    if (
      !runtime.isEvaluationSuspended() &&
      runtime.getBatchDepth() === 0 &&
      rawCellContentsEqual(runtime.getCellSerialized(address), content)
    ) {
      return []
    }
    if (
      runtime.isEvaluationSuspended() &&
      runtime.enqueueSuspendedLiteralMutation(address.sheet, address.row, address.col, content, visibleCellIndex)
    ) {
      return []
    }
    if (
      runtime.getBatchDepth() !== 0 &&
      runtime.enqueueDeferredBatchLiteral(address.sheet, address.row, address.col, content, visibleCellIndex)
    ) {
      return []
    }
    if (typeof content === 'number' && visibleCellIndex !== undefined) {
      const fastPathChanges = runtime.trySetExistingNumericCellContentsWithTrackedFastPath({
        sheet,
        address,
        cellIndex: visibleCellIndex,
        value: content,
      })
      if (fastPathChanges !== null) {
        return fastPathChanges
      }
    }
    if (content !== null && !isParsableFormulaContent(content) && typeof content !== 'number' && visibleCellIndex !== undefined) {
      const fastPathChanges = runtime.trySetExistingLiteralCellContentsWithTrackedFastPath({
        sheet,
        address,
        cellIndex: visibleCellIndex,
        value: content,
      })
      if (fastPathChanges !== null) {
        return fastPathChanges
      }
    }
    const mutate = () => {
      runtime.flushPendingBatchOps()
      const existingNumericMutationEngine = runtime.getEngine() as ExistingNumericMutationEngine
      if (
        typeof content === 'number' &&
        visibleCellIndex !== undefined &&
        sheet.structureVersion === 1 &&
        existingNumericMutationEngine.tryApplyExistingNumericCellMutationAt?.({
          sheetId: address.sheet,
          row: address.row,
          col: address.col,
          cellIndex: visibleCellIndex,
          value: content,
        })
      ) {
        return
      }
      const mutation = buildWorkPaperRawCellMutation({
        row: address.row,
        col: address.col,
        content,
        rewriteFormulaForStorage: (formula) => runtime.rewriteFormulaForStorage(formula, address.sheet),
      })
      runtime.applyCellMutationRefs(
        [{ sheetId: address.sheet, mutation, ...(visibleCellIndex !== undefined ? { cellIndex: visibleCellIndex } : {}) }],
        {
          captureUndo: true,
          potentialNewCells: content === null || visibleCellIndex !== undefined ? 0 : 1,
          source: 'local',
          returnUndoOps: false,
          reuseRefs: true,
        },
      )
    }
    if (runtime.canUseTrackedMutationFastPath()) {
      const captureOptions: {
        singleLiteralChange?: {
          address: WorkPaperCellAddress
          cellIndex?: number
          isPhysicalSheet: boolean
          sheetName: string
          value: RawCellContent
        }
      } = {}
      if (!isParsableFormulaContent(content)) {
        captureOptions.singleLiteralChange = {
          address: { sheet: address.sheet, row: address.row, col: address.col },
          ...(visibleCellIndex === undefined ? {} : { cellIndex: visibleCellIndex }),
          isPhysicalSheet: sheet.structureVersion === 1,
          sheetName: sheet.name,
          value: content,
        }
      }
      return runtime.captureTrackedChangesWithoutVisibilityCache(mutate, captureOptions)
    }
    return runtime.captureChanges(mutate)
  }
  if (!runtime.isItPossibleToSetCellContents(address, content)) {
    throw new WorkPaperOperationError('Cell contents cannot be set')
  }
  if (runtime.isTrackedBatchFastPathActive()) {
    runtime.flushPendingBatchOps()
    runtime.applyMatrixContents(address, content)
    return []
  }
  return runtime.captureChanges(() => {
    runtime.flushPendingBatchOps()
    runtime.applyMatrixContents(address, content)
  })
}

export function setWorkPaperCellValues(
  runtime: WorkPaperSetCellContentsRuntime,
  updates: readonly WorkPaperCellValueUpdate[],
): WorkPaperChange[] {
  runtime.assertNotDisposed()
  if (updates.length === 0) {
    return []
  }
  const refs: EngineCellMutationRef[] = []
  refs.length = updates.length
  let potentialNewCells = 0
  const config = runtime.getConfig()
  let currentSheetId: number | undefined
  let currentSheet: SheetRecord | undefined
  for (let index = 0; index < updates.length; index += 1) {
    const { address, value } = updates[index]!
    if (currentSheet === undefined || currentSheetId !== address.sheet) {
      currentSheetId = address.sheet
      currentSheet = runtime.sheetRecord(address.sheet)
    }
    assertRowAndColumn(address.row, 'address.row')
    assertRowAndColumn(address.col, 'address.col')
    if (address.row >= (config.maxRows ?? MAX_ROWS) || address.col >= (config.maxColumns ?? MAX_COLS)) {
      throw new WorkPaperOperationError('Cell contents cannot be set')
    }
    if (isWorkPaperSheetMatrix(value) || isParsableFormulaContent(value)) {
      throw new WorkPaperOperationError('Bulk cell value updates require literal values')
    }
    const visibleCellIndex = runtime.getVisibleCellIndexInSheet(currentSheet, address.row, address.col)
    refs[index] = {
      sheetId: address.sheet,
      mutation: buildWorkPaperLiteralCellValueMutation({
        row: address.row,
        col: address.col,
        content: value,
      }),
      ...(visibleCellIndex !== undefined ? { cellIndex: visibleCellIndex } : {}),
    }
    if (!isBlankRawCellContent(value) && visibleCellIndex === undefined) {
      potentialNewCells += 1
    }
  }
  return applyBulkWorkPaperCellValueRefs(runtime, refs, potentialNewCells)
}

export function setWorkPaperSheetCellValues(
  runtime: WorkPaperSetCellContentsRuntime,
  sheetId: number,
  updates: readonly WorkPaperSheetCellValueUpdate[],
): WorkPaperChange[] {
  runtime.assertNotDisposed()
  if (updates.length === 0) {
    return []
  }
  const sheet = runtime.sheetRecord(sheetId)
  const refs: EngineCellMutationRef[] = []
  refs.length = updates.length
  let potentialNewCells = 0
  const config = runtime.getConfig()
  for (let index = 0; index < updates.length; index += 1) {
    const { row, col, value } = updates[index]!
    assertRowAndColumn(row, 'row')
    assertRowAndColumn(col, 'col')
    if (row >= (config.maxRows ?? MAX_ROWS) || col >= (config.maxColumns ?? MAX_COLS)) {
      throw new WorkPaperOperationError('Cell contents cannot be set')
    }
    if (isWorkPaperSheetMatrix(value) || isParsableFormulaContent(value)) {
      throw new WorkPaperOperationError('Bulk cell value updates require literal values')
    }
    const visibleCellIndex = runtime.getVisibleCellIndexInSheet(sheet, row, col)
    refs[index] = {
      sheetId,
      mutation: buildWorkPaperLiteralCellValueMutation({
        row,
        col,
        content: value,
      }),
      ...(visibleCellIndex !== undefined ? { cellIndex: visibleCellIndex } : {}),
    }
    if (!isBlankRawCellContent(value) && visibleCellIndex === undefined) {
      potentialNewCells += 1
    }
  }
  return applyBulkWorkPaperCellValueRefs(runtime, refs, potentialNewCells)
}

export function setWorkPaperSheetRangeValues(
  runtime: WorkPaperSetCellContentsRuntime,
  sheetId: number,
  startRow: number,
  startCol: number,
  values: WorkPaperSheetRangeValues,
): WorkPaperChange[] {
  runtime.assertNotDisposed()
  if (values.length === 0) {
    return []
  }
  assertRowAndColumn(startRow, 'startRow')
  assertRowAndColumn(startCol, 'startCol')

  const config = runtime.getConfig()
  const maxRows = config.maxRows ?? MAX_ROWS
  const maxColumns = config.maxColumns ?? MAX_COLS
  let refCount = 0
  for (let rowOffset = 0; rowOffset < values.length; rowOffset += 1) {
    const row = values[rowOffset]
    if (row === undefined || row.length === 0) {
      continue
    }
    const destinationRow = startRow + rowOffset
    if (destinationRow >= maxRows || startCol + row.length > maxColumns) {
      throw new WorkPaperOperationError('Cell contents cannot be set')
    }
    refCount += row.length
  }
  if (refCount === 0) {
    return []
  }

  const sheet = runtime.sheetRecord(sheetId)
  const refs: EngineCellMutationRef[] = []
  refs.length = refCount
  let potentialNewCells = 0
  let refIndex = 0
  for (let rowOffset = 0; rowOffset < values.length; rowOffset += 1) {
    const row = values[rowOffset]
    if (row === undefined || row.length === 0) {
      continue
    }
    const destinationRow = startRow + rowOffset
    for (let colOffset = 0; colOffset < row.length; colOffset += 1) {
      const destinationCol = startCol + colOffset
      const value = row[colOffset] ?? null
      if (Array.isArray(value) || isParsableFormulaContent(value)) {
        throw new WorkPaperOperationError('Bulk cell value updates require literal values')
      }
      const visibleCellIndex = runtime.getVisibleCellIndexInSheet(sheet, destinationRow, destinationCol)
      refs[refIndex] = {
        sheetId,
        mutation: buildWorkPaperLiteralCellValueMutation({
          row: destinationRow,
          col: destinationCol,
          content: value,
        }),
        ...(visibleCellIndex !== undefined ? { cellIndex: visibleCellIndex } : {}),
      }
      refIndex += 1
      if (!isBlankRawCellContent(value) && visibleCellIndex === undefined) {
        potentialNewCells += 1
      }
    }
  }
  return applyBulkWorkPaperCellValueRefs(runtime, refs, potentialNewCells)
}

function applyBulkWorkPaperCellValueRefs(
  runtime: WorkPaperSetCellContentsRuntime,
  refs: readonly EngineCellMutationRef[],
  potentialNewCells: number,
): WorkPaperChange[] {
  const mutate = () => {
    runtime.flushPendingBatchOps()
    runtime.applyCellMutationRefs(refs, {
      captureUndo: true,
      potentialNewCells,
      source: 'local',
      returnUndoOps: false,
      reuseRefs: true,
    })
  }
  if (runtime.isEvaluationSuspended() || runtime.getBatchDepth() !== 0) {
    mutate()
    return []
  }
  if (runtime.canUseTrackedMutationFastPath()) {
    return runtime.captureTrackedChangesWithoutVisibilityCache(mutate, {})
  }
  return runtime.captureChanges(mutate)
}

function rawCellContentsEqual(left: RawCellContent, right: RawCellContent): boolean {
  return typeof left === 'number' && typeof right === 'number' ? Object.is(left, right) : left === right
}

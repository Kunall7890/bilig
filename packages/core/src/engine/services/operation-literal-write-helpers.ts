import { ErrorCode, ValueTag, type LiteralInput } from '@bilig/protocol'
import { CellFlags, type CellStore } from '../../cell-store.js'
import { writeLiteralToCellStore } from '../../engine-value-utils.js'
import type { StringPool } from '../../string-pool.js'
import type { SheetRecord } from '../../workbook-store.js'

export const FAST_OPERATION_LITERAL_OVERWRITE_FLAGS =
  CellFlags.HasFormula | CellFlags.JsOnly | CellFlags.InCycle | CellFlags.SpillChild | CellFlags.PivotOutput

export interface OperationLiteralVisibleCellSheet {
  readonly structureVersion: number
  readonly logical: {
    readonly getCellVisiblePosition: (cellIndex: number) => { readonly row: number; readonly col: number } | undefined
  }
}

export interface OperationLiteralWriteWorkbook {
  readonly cellStore: CellStore
  readonly notifyCellValueWritten: (cellIndex: number) => void
  readonly getSheetById: (sheetId: number) => OperationLiteralVisibleCellSheet | undefined
}

export interface OperationLiteralPredicateState {
  readonly workbook: {
    readonly cellStore: CellStore
    readonly getCellFormat: (cellIndex: number) => unknown
  }
  readonly strings: StringPool
  readonly formulas: {
    readonly get: (cellIndex: number) => unknown
  }
}

export function canFastPathOperationLiteralOverwrite(input: {
  readonly cellStore: CellStore
  readonly formulas: { readonly get: (cellIndex: number) => unknown }
  readonly cellIndex: number
}): boolean {
  const flags = input.cellStore.flags[input.cellIndex] ?? 0
  return (flags & FAST_OPERATION_LITERAL_OVERWRITE_FLAGS) === 0 && input.formulas.get(input.cellIndex) === undefined
}

export function isOperationNullLiteralWriteNoOp(input: {
  readonly state: OperationLiteralPredicateState
  readonly cellIndex: number
}): boolean {
  if (input.state.formulas.get(input.cellIndex) !== undefined) {
    return false
  }
  if (input.state.workbook.getCellFormat(input.cellIndex) !== undefined) {
    return false
  }
  const flags = input.state.workbook.cellStore.flags[input.cellIndex] ?? 0
  if ((flags & FAST_OPERATION_LITERAL_OVERWRITE_FLAGS) !== 0) {
    return false
  }
  const value = input.state.workbook.cellStore.getValue(input.cellIndex, (id) => input.state.strings.get(id))
  return value.tag === ValueTag.Empty
}

export function isOperationClearCellNoOp(input: { readonly state: OperationLiteralPredicateState; readonly cellIndex: number }): boolean {
  if (((input.state.workbook.cellStore.flags[input.cellIndex] ?? 0) & CellFlags.AuthoredBlank) !== 0) {
    return false
  }
  return isOperationNullLiteralWriteNoOp(input)
}

export function writeOperationLiteralToExistingCell(input: {
  readonly workbook: Pick<OperationLiteralWriteWorkbook, 'cellStore' | 'notifyCellValueWritten'>
  readonly strings: StringPool
  readonly cellIndex: number
  readonly value: LiteralInput
}): void {
  const cellStore = input.workbook.cellStore
  const hasSetValueHook = cellStore.onSetValue !== null
  writeLiteralToCellStore(cellStore, input.cellIndex, input.value, input.strings)
  if (!hasSetValueHook) {
    input.workbook.notifyCellValueWritten(input.cellIndex)
  }
}

export function writeOperationNumericLiteralToExistingCell(input: {
  readonly workbook: Pick<OperationLiteralWriteWorkbook, 'cellStore' | 'notifyCellValueWritten'>
  readonly cellIndex: number
  readonly value: number
}): void {
  const cellStore = input.workbook.cellStore
  const flags = cellStore.flags[input.cellIndex] ?? 0
  const hasSetValueHook = cellStore.onSetValue !== null
  cellStore.tags[input.cellIndex] = ValueTag.Number
  cellStore.errors[input.cellIndex] = ErrorCode.None
  cellStore.stringIds[input.cellIndex] = 0
  cellStore.numbers[input.cellIndex] = input.value
  if ((flags & CellFlags.AuthoredBlank) !== 0) {
    cellStore.flags[input.cellIndex] = flags & ~CellFlags.AuthoredBlank
  }
  cellStore.versions[input.cellIndex] = (cellStore.versions[input.cellIndex] ?? 0) + 1
  cellStore.onSetValue?.(input.cellIndex)
  if (!hasSetValueHook) {
    input.workbook.notifyCellValueWritten(input.cellIndex)
  }
}

export function writeOperationNumericLiteralToExistingCellWithoutColumnNotify(input: {
  readonly cellStore: CellStore
  readonly cellIndex: number
  readonly value: number
}): void {
  const flags = input.cellStore.flags[input.cellIndex] ?? 0
  input.cellStore.tags[input.cellIndex] = ValueTag.Number
  input.cellStore.errors[input.cellIndex] = ErrorCode.None
  input.cellStore.stringIds[input.cellIndex] = 0
  input.cellStore.numbers[input.cellIndex] = input.value
  if ((flags & CellFlags.AuthoredBlank) !== 0) {
    input.cellStore.flags[input.cellIndex] = flags & ~CellFlags.AuthoredBlank
  }
  input.cellStore.versions[input.cellIndex] = (input.cellStore.versions[input.cellIndex] ?? 0) + 1
}

export function writeTrustedOperationExistingNumericLiteralToCell(input: {
  readonly cellStore: CellStore
  readonly cellIndex: number
  readonly sheet: Pick<SheetRecord, 'columnVersions'>
  readonly col: number
  readonly value: number
}): void {
  const flags = input.cellStore.flags[input.cellIndex] ?? 0
  input.cellStore.numbers[input.cellIndex] = input.value
  if ((flags & CellFlags.AuthoredBlank) !== 0) {
    input.cellStore.flags[input.cellIndex] = flags & ~CellFlags.AuthoredBlank
  }
  input.cellStore.versions[input.cellIndex] = (input.cellStore.versions[input.cellIndex] ?? 0) + 1
  const onSetValue = input.cellStore.onSetValue
  if (onSetValue) {
    onSetValue(input.cellIndex)
    return
  }
  input.sheet.columnVersions[input.col] = (input.sheet.columnVersions[input.col] ?? 0) + 1
}

export function writeFastPathOperationLiteralToExistingCell(input: {
  readonly workbook: Pick<OperationLiteralWriteWorkbook, 'cellStore' | 'notifyCellValueWritten'>
  readonly strings: StringPool
  readonly cellIndex: number
  readonly value: LiteralInput
}): void {
  if (typeof input.value === 'number') {
    writeOperationNumericLiteralToExistingCell({
      workbook: input.workbook,
      cellIndex: input.cellIndex,
      value: input.value,
    })
    return
  }
  writeOperationLiteralToExistingCell(input)
}

function getVisibleOperationCellColumn(
  workbook: Pick<OperationLiteralWriteWorkbook, 'cellStore' | 'getSheetById'>,
  sheetId: number,
  cellIndex: number,
): number | undefined {
  const sheet = workbook.getSheetById(sheetId)
  if (!sheet || sheet.structureVersion === 1) {
    return workbook.cellStore.cols[cellIndex]
  }
  return sheet.logical.getCellVisiblePosition(cellIndex)?.col ?? workbook.cellStore.cols[cellIndex]
}

export function canSkipOperationFormulaColumnVersion(input: {
  readonly workbook: Pick<OperationLiteralWriteWorkbook, 'cellStore' | 'getSheetById'>
  readonly cellIndex: number
  readonly hasTrackedColumnDependents: (sheetId: number, col: number) => boolean
}): boolean {
  const sheetId = input.workbook.cellStore.sheetIds[input.cellIndex]
  if (sheetId === undefined) {
    return false
  }
  const col = getVisibleOperationCellColumn(input.workbook, sheetId, input.cellIndex)
  return col !== undefined && !input.hasTrackedColumnDependents(sheetId, col)
}

export function cellsShareOperationVersionColumn(input: {
  readonly workbook: Pick<OperationLiteralWriteWorkbook, 'cellStore' | 'getSheetById'>
  readonly leftCellIndex: number
  readonly rightCellIndex: number
}): boolean {
  const cellStore = input.workbook.cellStore
  const leftSheetId = cellStore.sheetIds[input.leftCellIndex]
  if (leftSheetId === undefined || leftSheetId !== cellStore.sheetIds[input.rightCellIndex]) {
    return false
  }
  const leftCol = getVisibleOperationCellColumn(input.workbook, leftSheetId, input.leftCellIndex)
  const rightCol = getVisibleOperationCellColumn(input.workbook, leftSheetId, input.rightCellIndex)
  return leftCol !== undefined && leftCol === rightCol
}

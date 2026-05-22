import { formatAddress } from '@bilig/formula'
import { ValueTag, type LiteralInput, type WorkbookAutoFilterColumnSnapshot, type WorkbookAutoFilterSnapshot } from '@bilig/protocol'
import type { EngineOp } from '@bilig/workbook'
import { normalizeRange } from '../engine-range-utils.js'
import type { WorkbookStore, WorkbookFilterRecord } from '../workbook-store.js'

export interface AutoFilterMaterializationState {
  readonly workbook: WorkbookStore
  readonly strings: {
    get(id: number): string
  }
}

interface RowFilterUpdate {
  readonly row: number
  readonly size: number | null
  readonly hidden: boolean | null
  readonly filtered: boolean | null
}

export function workbookAutoFiltersEqual(
  left: WorkbookFilterRecord | WorkbookAutoFilterSnapshot | undefined,
  right: WorkbookAutoFilterSnapshot,
): boolean {
  const leftRange = isWorkbookFilterRecord(left) ? left.range : left
  return leftRange !== undefined && JSON.stringify(leftRange) === JSON.stringify(right)
}

function isWorkbookFilterRecord(value: WorkbookFilterRecord | WorkbookAutoFilterSnapshot | undefined): value is WorkbookFilterRecord {
  return value !== undefined && 'range' in value
}

export function buildAutoFilterRowMetadataOps(
  state: AutoFilterMaterializationState,
  sheetName: string,
  previous: WorkbookAutoFilterSnapshot | undefined,
  next: WorkbookAutoFilterSnapshot | undefined,
): EngineOp[] {
  const rows = affectedBodyRows(sheetName, previous, next)
  if (rows.length === 0) {
    return []
  }
  const updates: RowFilterUpdate[] = []
  for (const row of rows) {
    const existing = state.workbook.getRowMetadata(sheetName, row, 1)
    const desiredFiltered =
      next !== undefined && filterContainsBodyRow(sheetName, next, row) && !rowMatchesAutoFilter(state, next, row) ? true : null
    if ((existing?.filtered ?? null) === desiredFiltered) {
      continue
    }
    updates.push({
      row,
      size: existing?.size ?? null,
      hidden: existing?.hidden ?? null,
      filtered: desiredFiltered,
    })
  }
  return groupRowFilterUpdates(sheetName, updates)
}

function affectedBodyRows(
  sheetName: string,
  previous: WorkbookAutoFilterSnapshot | undefined,
  next: WorkbookAutoFilterSnapshot | undefined,
): number[] {
  const rows = new Set<number>()
  addFilterBodyRows(rows, sheetName, previous)
  addFilterBodyRows(rows, sheetName, next)
  return [...rows].toSorted((left, right) => left - right)
}

function addFilterBodyRows(rows: Set<number>, sheetName: string, filter: WorkbookAutoFilterSnapshot | undefined): void {
  if (filter === undefined || filter.sheetName !== sheetName) {
    return
  }
  const bounds = normalizeRange(filter)
  for (let row = bounds.startRow + 1; row <= bounds.endRow; row += 1) {
    rows.add(row)
  }
}

function filterContainsBodyRow(sheetName: string, filter: WorkbookAutoFilterSnapshot, row: number): boolean {
  if (filter.sheetName !== sheetName) {
    return false
  }
  const bounds = normalizeRange(filter)
  return row > bounds.startRow && row <= bounds.endRow
}

function rowMatchesAutoFilter(state: AutoFilterMaterializationState, filter: WorkbookAutoFilterSnapshot, row: number): boolean {
  const criteria = filter.criteria?.filter((criterion) => criterion.filters !== undefined || criterion.customFilters !== undefined) ?? []
  if (criteria.length === 0) {
    return true
  }
  const bounds = normalizeRange(filter)
  return criteria.every((criterion) => {
    const col = bounds.startCol + criterion.colId
    if (col < bounds.startCol || col > bounds.endCol) {
      return true
    }
    return columnCriterionMatches(state, filter.sheetName, row, col, criterion)
  })
}

function columnCriterionMatches(
  state: AutoFilterMaterializationState,
  sheetName: string,
  row: number,
  col: number,
  criterion: WorkbookAutoFilterColumnSnapshot,
): boolean {
  const value = readCellLiteral(state, sheetName, row, col)
  if (criterion.filters && !valueFilterMatches(criterion.filters.blank === true, criterion.filters.values, value)) {
    return false
  }
  if (criterion.customFilters && !customFiltersMatch(criterion.customFilters.and === true, criterion.customFilters.filters, value)) {
    return false
  }
  return true
}

function readCellLiteral(state: AutoFilterMaterializationState, sheetName: string, row: number, col: number): LiteralInput | undefined {
  const cellIndex = state.workbook.getCellIndex(sheetName, formatAddress(row, col))
  if (cellIndex === undefined) {
    return null
  }
  const cellStore = state.workbook.cellStore
  const tag = (cellStore.tags[cellIndex] ?? ValueTag.Empty) as ValueTag
  switch (tag) {
    case ValueTag.Empty:
      return null
    case ValueTag.Number:
      return cellStore.numbers[cellIndex] ?? 0
    case ValueTag.Boolean:
      return (cellStore.numbers[cellIndex] ?? 0) !== 0
    case ValueTag.String: {
      const stringId = cellStore.stringIds[cellIndex] ?? 0
      return stringId === 0 ? '' : state.strings.get(stringId)
    }
    case ValueTag.Error:
      return undefined
  }
  return undefined
}

function valueFilterMatches(blankMatches: boolean, acceptedValues: readonly string[], value: LiteralInput | undefined): boolean {
  if (value === undefined || value === null || value === '') {
    return blankMatches
  }
  const normalized = normalizeFilterValue(value)
  return acceptedValues.some((accepted) => normalizeFilterValue(accepted) === normalized)
}

function customFiltersMatch(
  requireAll: boolean,
  filters: ReadonlyArray<NonNullable<WorkbookAutoFilterColumnSnapshot['customFilters']>['filters'][number]>,
  value: LiteralInput | undefined,
): boolean {
  if (filters.length === 0) {
    return true
  }
  const matches = filters.map((filter) => customFilterMatches(filter.operator ?? 'equal', filter.value, value))
  return requireAll ? matches.every(Boolean) : matches.some(Boolean)
}

function customFilterMatches(operator: string, expected: string, value: LiteralInput | undefined): boolean {
  if (value === undefined || value === null) {
    return false
  }
  const actualNumber = typeof value === 'number' ? value : Number(String(value))
  const expectedNumber = Number(expected)
  if (Number.isFinite(actualNumber) && Number.isFinite(expectedNumber)) {
    switch (operator) {
      case 'equal':
        return actualNumber === expectedNumber
      case 'notEqual':
        return actualNumber !== expectedNumber
      case 'lessThan':
        return actualNumber < expectedNumber
      case 'lessThanOrEqual':
        return actualNumber <= expectedNumber
      case 'greaterThan':
        return actualNumber > expectedNumber
      case 'greaterThanOrEqual':
        return actualNumber >= expectedNumber
      default:
        return false
    }
  }
  const actual = normalizeFilterValue(value)
  const expectedText = normalizeFilterValue(expected)
  switch (operator) {
    case 'equal':
      return actual === expectedText
    case 'notEqual':
      return actual !== expectedText
    default:
      return false
  }
}

function normalizeFilterValue(value: NonNullable<LiteralInput> | string): string {
  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE'
  }
  return String(value)
}

function groupRowFilterUpdates(sheetName: string, updates: readonly RowFilterUpdate[]): EngineOp[] {
  const ops: EngineOp[] = []
  for (const update of updates) {
    const previous = ops[ops.length - 1]
    if (
      previous?.kind === 'updateRowMetadata' &&
      previous.sheetName === sheetName &&
      previous.start + previous.count === update.row &&
      previous.size === update.size &&
      previous.hidden === update.hidden &&
      (previous.filtered ?? null) === update.filtered
    ) {
      previous.count += 1
      continue
    }
    ops.push({
      kind: 'updateRowMetadata',
      sheetName,
      start: update.row,
      count: 1,
      size: update.size,
      hidden: update.hidden,
      filtered: update.filtered,
    })
  }
  return ops
}

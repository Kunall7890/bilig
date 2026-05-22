import type {
  WorkbookAxisEntrySnapshot,
  WorkbookAutoFilterColumnSnapshot,
  WorkbookAutoFilterSnapshot,
  WorkbookSnapshot,
} from '@bilig/protocol'
import { decodeA1CellRef, decodeA1RangeRef } from './xlsx-a1-utils.js'

type SheetCellSnapshot = WorkbookSnapshot['sheets'][number]['cells'][number]

export function applyImportedAutoFilterVisibility(
  sheetName: string,
  cells: readonly SheetCellSnapshot[],
  rows: readonly WorkbookAxisEntrySnapshot[] | undefined,
  filters: readonly WorkbookAutoFilterSnapshot[] | undefined,
): WorkbookAxisEntrySnapshot[] | undefined {
  if (!filters || filters.length === 0 || cells.length === 0) {
    return rows && rows.length > 0 ? [...rows] : undefined
  }
  const filteredRows = new Set<number>()
  const cellValues = buildCellValueMap(cells)
  for (const filter of filters) {
    const rowsForFilter = filteredRowsForCriteria(sheetName, filter, cellValues)
    if (!rowsForFilter) {
      continue
    }
    for (const row of rowsForFilter) {
      filteredRows.add(row)
    }
  }
  if (filteredRows.size === 0) {
    return rows && rows.length > 0 ? [...rows] : undefined
  }

  const merged = new Map<number, WorkbookAxisEntrySnapshot>()
  for (const row of rows ?? []) {
    merged.set(row.index, { ...row })
  }
  for (const rowIndex of filteredRows) {
    const existing = merged.get(rowIndex)
    merged.set(rowIndex, {
      id: existing?.id ?? `row:${String(rowIndex)}`,
      index: rowIndex,
      ...(existing?.size !== undefined ? { size: existing.size } : {}),
      hidden: true,
      filtered: true,
      ...(existing?.styleIndex !== undefined ? { styleIndex: existing.styleIndex } : {}),
      ...(existing?.xlsxHeight !== undefined ? { xlsxHeight: existing.xlsxHeight } : {}),
      ...(existing?.customFormat !== undefined ? { customFormat: existing.customFormat } : {}),
      ...(existing?.customHeight !== undefined ? { customHeight: existing.customHeight } : {}),
      ...(existing?.outlineLevel !== undefined ? { outlineLevel: existing.outlineLevel } : {}),
      ...(existing?.collapsed !== undefined ? { collapsed: existing.collapsed } : {}),
      ...(existing?.thickTop !== undefined ? { thickTop: existing.thickTop } : {}),
      ...(existing?.thickBottom !== undefined ? { thickBottom: existing.thickBottom } : {}),
    })
  }
  return [...merged.values()].toSorted((left, right) => left.index - right.index)
}

function buildCellValueMap(cells: readonly SheetCellSnapshot[]): Map<string, SheetCellSnapshot['value']> {
  const values = new Map<string, SheetCellSnapshot['value']>()
  for (const cell of cells) {
    try {
      const coordinate = cell.row !== undefined && cell.col !== undefined ? { r: cell.row, c: cell.col } : decodeA1CellRef(cell.address)
      values.set(cellKey(coordinate.r, coordinate.c), cell.value)
    } catch {
      continue
    }
  }
  return values
}

function filteredRowsForCriteria(
  sheetName: string,
  filter: WorkbookAutoFilterSnapshot,
  values: ReadonlyMap<string, SheetCellSnapshot['value']>,
): Set<number> | null {
  if (filter.sheetName !== sheetName || !filter.criteria || filter.criteria.length === 0) {
    return null
  }
  let range
  try {
    range = decodeA1RangeRef(`${filter.startAddress}:${filter.endAddress}`)
  } catch {
    return null
  }
  if (range.e.r <= range.s.r) {
    return null
  }
  const criteria = filter.criteria.filter((criterion) => criterion.filters || criterion.customFilters)
  if (criteria.length === 0 || criteria.some((criterion) => !isSupportedCriterion(criterion))) {
    return null
  }

  const filteredRows = new Set<number>()
  for (let row = range.s.r + 1; row <= range.e.r; row += 1) {
    const visible = criteria.every((criterion) => {
      const col = range.s.c + criterion.colId
      if (col < range.s.c || col > range.e.c) {
        return true
      }
      return criterionMatches(criterion, values.get(cellKey(row, col)))
    })
    if (!visible) {
      filteredRows.add(row)
    }
  }
  return filteredRows
}

function isSupportedCriterion(criterion: WorkbookAutoFilterColumnSnapshot): boolean {
  return criterion.filters !== undefined || criterion.customFilters !== undefined
}

function criterionMatches(criterion: WorkbookAutoFilterColumnSnapshot, value: SheetCellSnapshot['value']): boolean {
  if (criterion.filters && !valueFilterMatches(criterion.filters.blank === true, criterion.filters.values, value)) {
    return false
  }
  if (criterion.customFilters && !customFiltersMatch(criterion.customFilters.and === true, criterion.customFilters.filters, value)) {
    return false
  }
  return true
}

function valueFilterMatches(blankMatches: boolean, acceptedValues: readonly string[], value: SheetCellSnapshot['value']): boolean {
  if (value === undefined || value === null || value === '') {
    return blankMatches
  }
  const normalized = normalizeFilterValue(value)
  return acceptedValues.some((accepted) => normalizeFilterValue(accepted) === normalized)
}

function customFiltersMatch(
  requireAll: boolean,
  filters: ReadonlyArray<NonNullable<WorkbookAutoFilterColumnSnapshot['customFilters']>['filters'][number]>,
  value: SheetCellSnapshot['value'],
): boolean {
  if (filters.length === 0) {
    return true
  }
  const matches = filters.map((filter) => customFilterMatches(filter.operator ?? 'equal', filter.value, value))
  return requireAll ? matches.every(Boolean) : matches.some(Boolean)
}

function customFilterMatches(operator: string, expected: string, value: SheetCellSnapshot['value']): boolean {
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

function normalizeFilterValue(value: NonNullable<SheetCellSnapshot['value']> | string): string {
  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE'
  }
  return String(value)
}

function cellKey(row: number, col: number): string {
  return `${String(row)}:${String(col)}`
}

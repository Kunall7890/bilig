import type {
  LiteralInput,
  WorkbookAutoFilterColumnSnapshot,
  WorkbookAutoFilterCustomCriterionSnapshot,
  WorkbookAutoFilterSnapshot,
  WorkbookAxisMetadataSnapshot,
  WorkbookSnapshot,
  WorkbookTableSnapshot,
} from '@bilig/protocol'
import { decodeA1CellRef } from './xlsx-a1-utils.js'

type SheetSnapshot = WorkbookSnapshot['sheets'][number]
type RowMetadataPatch = Omit<WorkbookAxisMetadataSnapshot, 'start' | 'count'>

interface FilterVisibilityRange {
  readonly filter: WorkbookAutoFilterSnapshot
  readonly bodyStartRow: number
  readonly bodyEndRow: number
  readonly startCol: number
}

const rowMetadataPatchKeys = [
  'size',
  'hidden',
  'filterHidden',
  'styleIndex',
  'xlsxWidth',
  'xlsxHeight',
  'customFormat',
  'customWidth',
  'bestFit',
  'outlineLevel',
  'collapsed',
  'customHeight',
  'thickTop',
  'thickBottom',
] as const satisfies readonly (keyof RowMetadataPatch)[]

export function applyImportedAutoFilterRowVisibility(
  sheet: SheetSnapshot,
  tables: readonly WorkbookTableSnapshot[] | undefined,
): SheetSnapshot {
  const ranges = filterVisibilityRanges(sheet, tables)
  if (ranges.length === 0) {
    return sheet
  }

  const valuesByCoordinate = sheetCellValuesByCoordinate(sheet)
  const filterHiddenRows = new Set<number>()
  for (const range of ranges) {
    for (let row = range.bodyStartRow; row <= range.bodyEndRow; row += 1) {
      if (!autoFilterRowMatches(valuesByCoordinate, range.filter.criteria ?? [], row, range.startCol)) {
        filterHiddenRows.add(row)
      }
    }
  }

  if (filterHiddenRows.size === 0) {
    return sheet
  }

  const rowMetadata = applyFilterHiddenRowsToMetadata(sheet.metadata?.rowMetadata ?? [], filterHiddenRows)
  return {
    ...sheet,
    metadata: {
      ...sheet.metadata,
      rowMetadata,
    },
  }
}

function filterVisibilityRanges(sheet: SheetSnapshot, tables: readonly WorkbookTableSnapshot[] | undefined): FilterVisibilityRange[] {
  return [
    ...(sheet.metadata?.filters ?? []).flatMap((filter) => worksheetFilterVisibilityRange(filter)),
    ...(tables ?? []).flatMap((table) => tableFilterVisibilityRange(sheet.name, table)),
  ]
}

function worksheetFilterVisibilityRange(filter: WorkbookAutoFilterSnapshot): FilterVisibilityRange[] {
  if (!filter.criteria || filter.criteria.length === 0) {
    return []
  }
  const start = decodeA1CellRef(filter.startAddress)
  const end = decodeA1CellRef(filter.endAddress)
  const bodyStartRow = Math.min(start.r + 1, end.r + 1)
  if (bodyStartRow > end.r) {
    return []
  }
  return [
    {
      filter,
      bodyStartRow,
      bodyEndRow: end.r,
      startCol: start.c,
    },
  ]
}

function tableFilterVisibilityRange(sheetName: string, table: WorkbookTableSnapshot): FilterVisibilityRange[] {
  const filter = table.autoFilter
  if (table.sheetName !== sheetName || !filter?.criteria || filter.criteria.length === 0) {
    return []
  }
  const start = decodeA1CellRef(table.startAddress)
  const end = decodeA1CellRef(table.endAddress)
  const bodyStartRow = table.headerRow ? start.r + 1 : start.r
  const bodyEndRow = table.totalsRow ? end.r - 1 : end.r
  if (bodyStartRow > bodyEndRow) {
    return []
  }
  return [
    {
      filter,
      bodyStartRow,
      bodyEndRow,
      startCol: start.c,
    },
  ]
}

function sheetCellValuesByCoordinate(sheet: SheetSnapshot): ReadonlyMap<string, LiteralInput | undefined> {
  const values = new Map<string, LiteralInput | undefined>()
  for (const cell of sheet.cells) {
    const row = cell.row
    const col = cell.col
    if (row !== undefined && col !== undefined) {
      values.set(coordinateKey(row, col), cell.value)
      continue
    }
    const decoded = decodeA1CellRef(cell.address)
    values.set(coordinateKey(decoded.r, decoded.c), cell.value)
  }
  return values
}

function coordinateKey(row: number, col: number): string {
  return `${String(row)}:${String(col)}`
}

function autoFilterRowMatches(
  valuesByCoordinate: ReadonlyMap<string, LiteralInput | undefined>,
  criteria: readonly WorkbookAutoFilterColumnSnapshot[],
  row: number,
  startCol: number,
): boolean {
  return criteria.every((criterion) => {
    const value = valuesByCoordinate.get(coordinateKey(row, startCol + criterion.colId))
    if (criterion.filters && !valueFilterMatches(value, criterion.filters)) {
      return false
    }
    if (criterion.customFilters && !customFiltersMatch(value, criterion.customFilters)) {
      return false
    }
    return true
  })
}

function valueFilterMatches(value: LiteralInput | undefined, criteria: NonNullable<WorkbookAutoFilterColumnSnapshot['filters']>): boolean {
  if (valueIsBlank(value)) {
    return criteria.blank === true
  }
  const text = filterText(value)
  return criteria.values.some((candidate) => textEqualsFilterValue(text, candidate))
}

function customFiltersMatch(
  value: LiteralInput | undefined,
  criteria: NonNullable<WorkbookAutoFilterColumnSnapshot['customFilters']>,
): boolean {
  const matches = criteria.filters.map((filter) => customFilterMatches(value, filter))
  return criteria.and === true ? matches.every(Boolean) : matches.some(Boolean)
}

function customFilterMatches(value: LiteralInput | undefined, criterion: WorkbookAutoFilterCustomCriterionSnapshot): boolean {
  const actual = filterText(value)
  const expected = criterion.value
  switch (criterion.operator ?? 'equal') {
    case 'equal':
      return textEqualsFilterValue(actual, expected)
    case 'notEqual':
      return !textEqualsFilterValue(actual, expected)
    case 'lessThan':
      return compareFilterValues(actual, expected) < 0
    case 'lessThanOrEqual':
      return compareFilterValues(actual, expected) <= 0
    case 'greaterThan':
      return compareFilterValues(actual, expected) > 0
    case 'greaterThanOrEqual':
      return compareFilterValues(actual, expected) >= 0
  }
}

function filterText(value: LiteralInput | undefined): string {
  if (value === undefined || value === null) {
    return ''
  }
  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE'
  }
  return String(value)
}

function valueIsBlank(value: LiteralInput | undefined): boolean {
  return value === undefined || value === null || value === ''
}

function textEqualsFilterValue(actual: string, expected: string): boolean {
  return actual.localeCompare(expected, undefined, { sensitivity: 'accent' }) === 0
}

function compareFilterValues(actual: string, expected: string): number {
  const left = Number(actual)
  const right = Number(expected)
  if (Number.isFinite(left) && Number.isFinite(right)) {
    return left - right
  }
  return actual.localeCompare(expected, undefined, { sensitivity: 'accent' })
}

function applyFilterHiddenRowsToMetadata(
  metadata: readonly WorkbookAxisMetadataSnapshot[],
  filterHiddenRows: ReadonlySet<number>,
): WorkbookAxisMetadataSnapshot[] {
  const output: WorkbookAxisMetadataSnapshot[] = []
  const consumedRows = new Set<number>()
  const sortedRows = [...filterHiddenRows].toSorted((left, right) => left - right)

  for (const record of metadata) {
    let cursor = record.start
    const end = record.start + record.count
    for (const row of sortedRows) {
      if (row < cursor) {
        continue
      }
      if (row >= end) {
        break
      }
      appendMetadataRecord(output, record, cursor, row - cursor)
      appendMetadataRecord(output, filterHiddenRecord(record, row), row, 1)
      consumedRows.add(row)
      cursor = row + 1
    }
    appendMetadataRecord(output, record, cursor, end - cursor)
  }

  for (const row of sortedRows) {
    if (!consumedRows.has(row)) {
      appendMetadataRecord(output, { start: row, count: 1, filterHidden: true }, row, 1)
    }
  }

  return coalesceRowMetadata(output)
}

function filterHiddenRecord(record: WorkbookAxisMetadataSnapshot, row: number): WorkbookAxisMetadataSnapshot {
  const next: WorkbookAxisMetadataSnapshot = {
    ...record,
    start: row,
    count: 1,
    filterHidden: true,
  }
  delete next.hidden
  return next
}

function appendMetadataRecord(
  output: WorkbookAxisMetadataSnapshot[],
  record: WorkbookAxisMetadataSnapshot,
  start: number,
  count: number,
): void {
  if (count <= 0) {
    return
  }
  output.push({
    ...record,
    start,
    count,
  })
}

function coalesceRowMetadata(records: readonly WorkbookAxisMetadataSnapshot[]): WorkbookAxisMetadataSnapshot[] {
  const output: WorkbookAxisMetadataSnapshot[] = []
  for (const record of records.toSorted((left, right) => left.start - right.start || left.count - right.count)) {
    const previous = output.at(-1)
    if (previous && previous.start + previous.count === record.start && metadataPatchesMatch(previous, record)) {
      previous.count += record.count
      continue
    }
    output.push({ ...record })
  }
  return output
}

function metadataPatchesMatch(left: WorkbookAxisMetadataSnapshot, right: WorkbookAxisMetadataSnapshot): boolean {
  return rowMetadataPatchKeys.every((key) => (left[key] ?? null) === (right[key] ?? null))
}

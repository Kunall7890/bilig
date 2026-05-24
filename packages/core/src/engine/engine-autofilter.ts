import type {
  CellRangeRef,
  LiteralInput,
  SheetMetadataSnapshot,
  WorkbookAutoFilterColumnSnapshot,
  WorkbookAutoFilterCustomCriterionSnapshot,
  WorkbookAutoFilterSnapshot,
  WorkbookAxisEntrySnapshot,
  WorkbookAxisMetadataSnapshot,
  WorkbookSnapshot,
  WorkbookTableSnapshot,
} from '@bilig/protocol'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { EngineOp } from '@bilig/workbook'

type RowGeometryPatch = Omit<WorkbookAxisMetadataSnapshot, 'start' | 'count' | 'size' | 'hidden' | 'filterHidden'>

interface FilterBounds {
  readonly headerRow: number
  readonly bodyStartRow: number
  readonly bodyEndRow: number
  readonly startCol: number
}

interface RowState {
  readonly size: number | null
  readonly hidden: boolean | null
  readonly filterHidden: boolean | null
  readonly geometry?: RowGeometryPatch
}

interface PendingRowUpdate extends RowState {
  readonly row: number
  readonly nextFilterHidden: boolean | null
}

const geometryKeys = [
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
] as const satisfies readonly (keyof RowGeometryPatch)[]

export function buildSetWorksheetAutoFilterOps(snapshot: WorkbookSnapshot, filter: WorkbookAutoFilterSnapshot): EngineOp[] {
  const sheet = snapshot.sheets.find((candidate) => candidate.name === filter.sheetName)
  if (!sheet) {
    throw new Error(`Unknown sheet: ${filter.sheetName}`)
  }
  const bounds = worksheetFilterBounds(filter)
  const visibilityOps =
    filter.criteria && filter.criteria.length > 0
      ? buildAutoFilterRowVisibilityOps(sheet, filter, bounds)
      : buildClearFilterRowVisibilityOps(sheet, bounds.bodyStartRow, bounds.bodyEndRow)
  const existingFilter = sheet.metadata?.filters?.find(
    (candidate) =>
      candidate.sheetName === filter.sheetName &&
      candidate.startAddress === filter.startAddress &&
      candidate.endAddress === filter.endAddress,
  )
  if (visibilityOps.length === 0 && existingFilter && JSON.stringify(existingFilter) === JSON.stringify(filter)) {
    return []
  }
  return [...visibilityOps, { kind: 'setFilter', sheetName: filter.sheetName, range: structuredClone(filter) }]
}

export function buildClearWorksheetAutoFilterOps(snapshot: WorkbookSnapshot, range: CellRangeRef): EngineOp[] {
  const sheet = snapshot.sheets.find((candidate) => candidate.name === range.sheetName)
  if (!sheet) {
    return [{ kind: 'clearFilter', sheetName: range.sheetName, range: { ...range } }]
  }
  const bounds = worksheetFilterBounds(range)
  return [
    ...buildClearFilterRowVisibilityOps(sheet, bounds.bodyStartRow, bounds.bodyEndRow),
    { kind: 'clearFilter', sheetName: range.sheetName, range: { ...range } },
  ]
}

export function buildApplyTableAutoFilterOps(
  snapshot: WorkbookSnapshot,
  sheetName: string,
  tableName: string,
  criteria: readonly WorkbookAutoFilterColumnSnapshot[],
): EngineOp[] {
  const sheet = snapshot.sheets.find((candidate) => candidate.name === sheetName)
  if (!sheet) {
    throw new Error(`Unknown sheet: ${sheetName}`)
  }
  const table = snapshot.workbook.metadata?.tables?.find((candidate) => candidate.name === tableName && candidate.sheetName === sheetName)
  if (!table) {
    throw new Error(`Unknown table: ${tableName}`)
  }
  const filter: WorkbookAutoFilterSnapshot = {
    sheetName,
    startAddress: table.startAddress,
    endAddress: table.endAddress,
    ...(criteria.length > 0 ? { criteria: criteria.map((entry) => structuredClone(entry)) } : {}),
  }
  const bounds = tableFilterBounds(table)
  return [
    ...(criteria.length > 0
      ? buildAutoFilterRowVisibilityOps(sheet, filter, bounds)
      : buildClearFilterRowVisibilityOps(sheet, bounds.bodyStartRow, bounds.bodyEndRow)),
    {
      kind: 'upsertTable',
      table: {
        ...structuredClone(table),
        autoFilter: filter,
      },
    },
  ]
}

function worksheetFilterBounds(range: CellRangeRef): FilterBounds {
  const start = parseCellAddress(range.startAddress, range.sheetName)
  const end = parseCellAddress(range.endAddress, range.sheetName)
  return {
    headerRow: start.row,
    bodyStartRow: Math.min(start.row + 1, end.row + 1),
    bodyEndRow: end.row,
    startCol: start.col,
  }
}

function tableFilterBounds(table: WorkbookTableSnapshot): FilterBounds {
  const start = parseCellAddress(table.startAddress, table.sheetName)
  const end = parseCellAddress(table.endAddress, table.sheetName)
  const bodyStartRow = table.headerRow ? start.row + 1 : start.row
  const bodyEndRow = table.totalsRow ? end.row - 1 : end.row
  return {
    headerRow: start.row,
    bodyStartRow,
    bodyEndRow,
    startCol: start.col,
  }
}

function buildAutoFilterRowVisibilityOps(
  sheet: WorkbookSnapshot['sheets'][number],
  filter: WorkbookAutoFilterSnapshot,
  bounds: FilterBounds,
): EngineOp[] {
  if (bounds.bodyStartRow > bounds.bodyEndRow) {
    return []
  }
  const cellsByAddress = new Map(sheet.cells.map((cell) => [cell.address, cell.value]))
  const updates: PendingRowUpdate[] = []
  for (let row = bounds.bodyStartRow; row <= bounds.bodyEndRow; row += 1) {
    const rowMatches = autoFilterRowMatches(cellsByAddress, filter.criteria ?? [], row, bounds.startCol)
    const state = rowStateAt(sheet.metadata, row)
    const nextFilterHidden = rowMatches ? null : true
    if (state.filterHidden === nextFilterHidden) {
      continue
    }
    updates.push({ ...state, row, nextFilterHidden })
  }
  return rowVisibilityUpdatesToOps(sheet.name, updates)
}

function buildClearFilterRowVisibilityOps(sheet: WorkbookSnapshot['sheets'][number], bodyStartRow: number, bodyEndRow: number): EngineOp[] {
  const updates: PendingRowUpdate[] = []
  for (let row = bodyStartRow; row <= bodyEndRow; row += 1) {
    const state = rowStateAt(sheet.metadata, row)
    if (state.filterHidden !== true) {
      continue
    }
    updates.push({ ...state, row, nextFilterHidden: null })
  }
  return rowVisibilityUpdatesToOps(sheet.name, updates)
}

function autoFilterRowMatches(
  cellsByAddress: ReadonlyMap<string, LiteralInput | undefined>,
  criteria: readonly WorkbookAutoFilterColumnSnapshot[],
  row: number,
  startCol: number,
): boolean {
  return criteria.every((criterion) => {
    const value = cellsByAddress.get(formatAddress(row, startCol + criterion.colId))
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

function rowVisibilityUpdatesToOps(sheetName: string, updates: readonly PendingRowUpdate[]): EngineOp[] {
  const ops: EngineOp[] = []
  let run: PendingRowUpdate[] = []
  const flush = (): void => {
    if (run.length === 0) {
      return
    }
    const first = run[0]!
    ops.push({
      kind: 'updateRowMetadata',
      sheetName,
      start: first.row,
      count: run.length,
      size: first.size,
      hidden: first.hidden,
      filterHidden: first.nextFilterHidden,
      ...(first.geometry ? { geometry: first.geometry } : {}),
    })
    run = []
  }
  for (const update of updates) {
    const previous = run.at(-1)
    if (
      previous &&
      update.row === previous.row + 1 &&
      update.size === previous.size &&
      update.hidden === previous.hidden &&
      update.nextFilterHidden === previous.nextFilterHidden &&
      geometryMatches(update.geometry, previous.geometry)
    ) {
      run.push(update)
      continue
    }
    flush()
    run.push(update)
  }
  flush()
  return ops
}

function rowStateAt(metadata: SheetMetadataSnapshot | undefined, row: number): RowState {
  const fromMetadata = metadata?.rowMetadata?.find((entry) => row >= entry.start && row < entry.start + entry.count)
  const fromEntry = metadata?.rows?.find((entry) => entry.index === row)
  const merged: WorkbookAxisEntrySnapshot & WorkbookAxisMetadataSnapshot = {
    id: fromEntry?.id ?? `row:${row}`,
    index: row,
    start: row,
    count: 1,
    ...fromMetadata,
    ...fromEntry,
  }
  const geometry = geometryPatch(merged)
  return {
    size: merged.size ?? null,
    hidden: merged.hidden ?? null,
    filterHidden: merged.filterHidden ?? null,
    ...(geometry ? { geometry } : {}),
  }
}

function geometryPatch(record: WorkbookAxisEntrySnapshot | WorkbookAxisMetadataSnapshot): RowGeometryPatch | undefined {
  const patch: Partial<RowGeometryPatch> = {}
  for (const key of geometryKeys) {
    if (record[key] !== undefined) {
      Object.assign(patch, { [key]: record[key] })
    }
  }
  return Object.keys(patch).length > 0 ? patch : undefined
}

function geometryMatches(left: RowGeometryPatch | undefined, right: RowGeometryPatch | undefined): boolean {
  return geometryKeys.every((key) => (left?.[key] ?? null) === (right?.[key] ?? null))
}

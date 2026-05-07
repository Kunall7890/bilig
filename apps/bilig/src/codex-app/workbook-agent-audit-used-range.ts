import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { CellRangeRef, SheetMetadataSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import type { WorkbookRuntime } from '../workbook-runtime/runtime-manager.js'
import { clampAuditLimit } from './workbook-agent-audit-limits.js'

interface CellBounds {
  minRow: number
  maxRow: number
  minCol: number
  maxCol: number
}

interface MetadataBoundsDriver {
  source:
    | 'styleRange'
    | 'formatRange'
    | 'filter'
    | 'sort'
    | 'validation'
    | 'conditionalFormat'
    | 'protectedRange'
    | 'commentThread'
    | 'note'
    | 'table'
    | 'pivot'
    | 'chart'
    | 'image'
    | 'shape'
    | 'spill'
    | 'definedName'
    | 'rowMetadata'
    | 'columnMetadata'
  startAddress: string
  endAddress: string
}

export interface WorkbookUsedRangeBloatSheetReport {
  sheetName: string
  populatedCellCount: number
  populatedRange: CellRangeRef | null
  compositeRange: CellRangeRef
  extraRows: number
  extraColumns: number
  populatedArea: number
  compositeArea: number
  bloatArea: number
  drivers: MetadataBoundsDriver[]
}

export interface WorkbookUsedRangeBloatReport {
  summary: {
    scannedSheetCount: number
    bloatedSheetCount: number
    totalBloatArea: number
    truncated: boolean
  }
  sheets: WorkbookUsedRangeBloatSheetReport[]
}

function boundsFromAddress(sheetName: string, address: string): CellBounds {
  const parsed = parseCellAddress(address, sheetName)
  return {
    minRow: parsed.row,
    maxRow: parsed.row,
    minCol: parsed.col,
    maxCol: parsed.col,
  }
}

function boundsFromRange(range: CellRangeRef): CellBounds {
  const start = parseCellAddress(range.startAddress, range.sheetName)
  const end = parseCellAddress(range.endAddress, range.sheetName)
  return {
    minRow: Math.min(start.row, end.row),
    maxRow: Math.max(start.row, end.row),
    minCol: Math.min(start.col, end.col),
    maxCol: Math.max(start.col, end.col),
  }
}

function mergeBounds(base: CellBounds | null, next: CellBounds | null): CellBounds | null {
  if (!next) {
    return base
  }
  if (!base) {
    return next
  }
  return {
    minRow: Math.min(base.minRow, next.minRow),
    maxRow: Math.max(base.maxRow, next.maxRow),
    minCol: Math.min(base.minCol, next.minCol),
    maxCol: Math.max(base.maxCol, next.maxCol),
  }
}

function boundsToRange(sheetName: string, bounds: CellBounds): CellRangeRef {
  return {
    sheetName,
    startAddress: formatAddress(bounds.minRow, bounds.minCol),
    endAddress: formatAddress(bounds.maxRow, bounds.maxCol),
  }
}

function boundsArea(bounds: CellBounds | null): number {
  if (!bounds) {
    return 0
  }
  return (bounds.maxRow - bounds.minRow + 1) * (bounds.maxCol - bounds.minCol + 1)
}

function collectSheetBloatDrivers(
  snapshot: WorkbookSnapshot,
  sheetName: string,
  metadata: SheetMetadataSnapshot | undefined,
): {
  bounds: CellBounds | null
  rowExtent: CellBounds | null
  columnExtent: CellBounds | null
  drivers: MetadataBoundsDriver[]
} {
  let bounds: CellBounds | null = null
  let rowExtent: CellBounds | null = null
  let columnExtent: CellBounds | null = null
  const drivers: MetadataBoundsDriver[] = []

  const addRangeDriver = (source: MetadataBoundsDriver['source'], range: CellRangeRef) => {
    bounds = mergeBounds(bounds, boundsFromRange(range))
    drivers.push({
      source,
      startAddress: range.startAddress,
      endAddress: range.endAddress,
    })
  }

  const addAddressDriver = (source: MetadataBoundsDriver['source'], address: string, endAddress = address) => {
    const range = {
      sheetName,
      startAddress: address,
      endAddress,
    } satisfies CellRangeRef
    addRangeDriver(source, range)
  }

  for (const styleRange of metadata?.styleRanges ?? []) {
    addRangeDriver('styleRange', styleRange.range)
  }
  for (const formatRange of metadata?.formatRanges ?? []) {
    addRangeDriver('formatRange', formatRange.range)
  }
  for (const filter of metadata?.filters ?? []) {
    addRangeDriver('filter', filter)
  }
  for (const sort of metadata?.sorts ?? []) {
    addRangeDriver('sort', sort.range)
  }
  for (const validation of metadata?.validations ?? []) {
    addRangeDriver('validation', validation.range)
  }
  for (const conditionalFormat of metadata?.conditionalFormats ?? []) {
    addRangeDriver('conditionalFormat', conditionalFormat.range)
  }
  for (const protectedRange of metadata?.protectedRanges ?? []) {
    addRangeDriver('protectedRange', protectedRange.range)
  }
  for (const commentThread of metadata?.commentThreads ?? []) {
    addAddressDriver('commentThread', commentThread.address)
  }
  for (const note of metadata?.notes ?? []) {
    addAddressDriver('note', note.address)
  }

  for (const table of snapshot.workbook.metadata?.tables ?? []) {
    if (table.sheetName === sheetName) {
      addRangeDriver('table', {
        sheetName,
        startAddress: table.startAddress,
        endAddress: table.endAddress,
      })
    }
  }
  for (const pivot of snapshot.workbook.metadata?.pivots ?? []) {
    if (pivot.sheetName !== sheetName) {
      continue
    }
    const start = parseCellAddress(pivot.address, sheetName)
    addAddressDriver(
      'pivot',
      pivot.address,
      formatAddress(start.row + Math.max(0, pivot.rows - 1), start.col + Math.max(0, pivot.cols - 1)),
    )
  }
  for (const chart of snapshot.workbook.metadata?.charts ?? []) {
    if (chart.sheetName === sheetName) {
      const start = parseCellAddress(chart.address, sheetName)
      addAddressDriver(
        'chart',
        chart.address,
        formatAddress(start.row + Math.max(0, chart.rows - 1), start.col + Math.max(0, chart.cols - 1)),
      )
    }
    if (chart.source.sheetName === sheetName) {
      addRangeDriver('chart', chart.source)
    }
  }
  for (const image of snapshot.workbook.metadata?.images ?? []) {
    if (image.sheetName !== sheetName) {
      continue
    }
    const start = parseCellAddress(image.address, sheetName)
    addAddressDriver(
      'image',
      image.address,
      formatAddress(start.row + Math.max(0, image.rows - 1), start.col + Math.max(0, image.cols - 1)),
    )
  }
  for (const shape of snapshot.workbook.metadata?.shapes ?? []) {
    if (shape.sheetName !== sheetName) {
      continue
    }
    const start = parseCellAddress(shape.address, sheetName)
    addAddressDriver(
      'shape',
      shape.address,
      formatAddress(start.row + Math.max(0, shape.rows - 1), start.col + Math.max(0, shape.cols - 1)),
    )
  }
  for (const spill of snapshot.workbook.metadata?.spills ?? []) {
    if (spill.sheetName !== sheetName) {
      continue
    }
    const start = parseCellAddress(spill.address, sheetName)
    addAddressDriver(
      'spill',
      spill.address,
      formatAddress(start.row + Math.max(0, spill.rows - 1), start.col + Math.max(0, spill.cols - 1)),
    )
  }
  for (const definedName of snapshot.workbook.metadata?.definedNames ?? []) {
    const value = definedName.value
    if (typeof value !== 'object' || value === null) {
      continue
    }
    if (value.kind === 'cell-ref' && value.sheetName === sheetName) {
      addAddressDriver('definedName', value.address)
    }
    if (value.kind === 'range-ref' && value.sheetName === sheetName) {
      addRangeDriver('definedName', {
        sheetName,
        startAddress: value.startAddress,
        endAddress: value.endAddress,
      })
    }
  }

  const rowRecords = [...(metadata?.rowMetadata ?? []), ...(metadata?.rows ?? [])]
  for (const record of rowRecords) {
    const start = 'index' in record ? record.index : record.start
    const count = 'count' in record ? record.count : 1
    const next = {
      minRow: start,
      maxRow: start + count - 1,
      minCol: 0,
      maxCol: 0,
    } satisfies CellBounds
    rowExtent = mergeBounds(rowExtent, next)
  }
  if (rowExtent) {
    drivers.push({
      source: 'rowMetadata',
      startAddress: formatAddress(rowExtent.minRow, 0),
      endAddress: formatAddress(rowExtent.maxRow, 0),
    })
  }

  const columnRecords = [...(metadata?.columnMetadata ?? []), ...(metadata?.columns ?? [])]
  for (const record of columnRecords) {
    const start = 'index' in record ? record.index : record.start
    const count = 'count' in record ? record.count : 1
    const next = {
      minRow: 0,
      maxRow: 0,
      minCol: start,
      maxCol: start + count - 1,
    } satisfies CellBounds
    columnExtent = mergeBounds(columnExtent, next)
  }
  if (columnExtent) {
    drivers.push({
      source: 'columnMetadata',
      startAddress: formatAddress(0, columnExtent.minCol),
      endAddress: formatAddress(0, columnExtent.maxCol),
    })
  }

  return {
    bounds,
    rowExtent,
    columnExtent,
    drivers,
  }
}

function buildContentBounds(sheet: WorkbookSnapshot['sheets'][number]): CellBounds | null {
  let bounds: CellBounds | null = null
  for (const cell of sheet.cells) {
    bounds = mergeBounds(bounds, boundsFromAddress(sheet.name, cell.address))
  }
  return bounds
}

export function scanWorkbookUsedRangeBloat(
  runtime: WorkbookRuntime,
  input: {
    sheetName?: string | undefined
    limit?: number | undefined
  } = {},
): WorkbookUsedRangeBloatReport {
  const snapshot = runtime.engine.exportSnapshot()
  const limit = clampAuditLimit(input.limit)
  const reports: WorkbookUsedRangeBloatSheetReport[] = []

  for (const sheet of snapshot.sheets) {
    if (input.sheetName !== undefined && sheet.name !== input.sheetName) {
      continue
    }
    const contentBounds = buildContentBounds(sheet)
    const drivers = collectSheetBloatDrivers(snapshot, sheet.name, sheet.metadata)
    let compositeBounds = mergeBounds(contentBounds, drivers.bounds)
    if (compositeBounds && drivers.rowExtent) {
      compositeBounds = {
        ...compositeBounds,
        minRow: Math.min(compositeBounds.minRow, drivers.rowExtent.minRow),
        maxRow: Math.max(compositeBounds.maxRow, drivers.rowExtent.maxRow),
      }
    } else if (!compositeBounds && drivers.rowExtent && drivers.columnExtent) {
      compositeBounds = {
        minRow: drivers.rowExtent.minRow,
        maxRow: drivers.rowExtent.maxRow,
        minCol: drivers.columnExtent.minCol,
        maxCol: drivers.columnExtent.maxCol,
      }
    }
    if (compositeBounds && drivers.columnExtent) {
      compositeBounds = {
        ...compositeBounds,
        minCol: Math.min(compositeBounds.minCol, drivers.columnExtent.minCol),
        maxCol: Math.max(compositeBounds.maxCol, drivers.columnExtent.maxCol),
      }
    }
    if (!compositeBounds) {
      continue
    }
    const populatedArea = boundsArea(contentBounds)
    const compositeArea = boundsArea(compositeBounds)
    if (
      contentBounds &&
      contentBounds.minRow === compositeBounds.minRow &&
      contentBounds.maxRow === compositeBounds.maxRow &&
      contentBounds.minCol === compositeBounds.minCol &&
      contentBounds.maxCol === compositeBounds.maxCol
    ) {
      continue
    }
    const extraRows =
      contentBounds === null
        ? compositeBounds.maxRow - compositeBounds.minRow + 1
        : Math.max(0, contentBounds.minRow - compositeBounds.minRow) + Math.max(0, compositeBounds.maxRow - contentBounds.maxRow)
    const extraColumns =
      contentBounds === null
        ? compositeBounds.maxCol - compositeBounds.minCol + 1
        : Math.max(0, contentBounds.minCol - compositeBounds.minCol) + Math.max(0, compositeBounds.maxCol - contentBounds.maxCol)
    reports.push({
      sheetName: sheet.name,
      populatedCellCount: sheet.cells.length,
      populatedRange: contentBounds ? boundsToRange(sheet.name, contentBounds) : null,
      compositeRange: boundsToRange(sheet.name, compositeBounds),
      extraRows,
      extraColumns,
      populatedArea,
      compositeArea,
      bloatArea: Math.max(0, compositeArea - populatedArea),
      drivers: drivers.drivers.slice(0, 12),
    })
  }

  const sortedReports = reports.toSorted((left, right) => {
    if (left.bloatArea !== right.bloatArea) {
      return right.bloatArea - left.bloatArea
    }
    if (left.extraRows !== right.extraRows) {
      return right.extraRows - left.extraRows
    }
    if (left.extraColumns !== right.extraColumns) {
      return right.extraColumns - left.extraColumns
    }
    return left.sheetName.localeCompare(right.sheetName)
  })

  return {
    summary: {
      scannedSheetCount:
        input.sheetName === undefined ? snapshot.sheets.length : snapshot.sheets.filter((sheet) => sheet.name === input.sheetName).length,
      bloatedSheetCount: sortedReports.length,
      totalBloatArea: sortedReports.reduce((sum, report) => sum + report.bloatArea, 0),
      truncated: sortedReports.length > limit,
    },
    sheets: sortedReports.slice(0, limit),
  }
}

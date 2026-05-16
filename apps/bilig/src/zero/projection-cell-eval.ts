import type { SpreadsheetEngine } from '@bilig/core'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import { ValueTag, type CellRangeRef } from '@bilig/protocol'
import type { CellEvalRow } from './projection.js'

interface RangeBounds {
  rowStart: number
  rowEnd: number
  colStart: number
  colEnd: number
}

function normalizeRangeBounds(sheetName: string, startAddress: string, endAddress: string): RangeBounds {
  const start = parseCellAddress(startAddress, sheetName)
  const end = parseCellAddress(endAddress, sheetName)
  return {
    rowStart: Math.min(start.row, end.row),
    rowEnd: Math.max(start.row, end.row),
    colStart: Math.min(start.col, end.col),
    colEnd: Math.max(start.col, end.col),
  }
}

function intersectBounds(left: RangeBounds, right: RangeBounds): RangeBounds | null {
  const rowStart = Math.max(left.rowStart, right.rowStart)
  const rowEnd = Math.min(left.rowEnd, right.rowEnd)
  const colStart = Math.max(left.colStart, right.colStart)
  const colEnd = Math.min(left.colEnd, right.colEnd)
  return rowStart <= rowEnd && colStart <= colEnd
    ? {
        rowStart,
        rowEnd,
        colStart,
        colEnd,
      }
    : null
}

function addressWithinBounds(row: number, col: number, bounds?: RangeBounds): boolean {
  if (!bounds) {
    return true
  }
  return row >= bounds.rowStart && row <= bounds.rowEnd && col >= bounds.colStart && col <= bounds.colEnd
}

function addAddressesForBounds(addresses: Set<string>, bounds: RangeBounds): void {
  for (let row = bounds.rowStart; row <= bounds.rowEnd; row += 1) {
    for (let col = bounds.colStart; col <= bounds.colEnd; col += 1) {
      addresses.add(formatAddress(row, col))
    }
  }
}

function collectCellEvalAddresses(engine: SpreadsheetEngine, sheetName: string, bounds?: RangeBounds): Set<string> {
  const sheet = engine.workbook.sheetsByName.get(sheetName)
  const addresses = new Set<string>()
  if (!sheet) {
    return addresses
  }

  sheet.grid.forEachCellEntry((_cellIndex, row, col) => {
    if (addressWithinBounds(row, col, bounds)) {
      addresses.add(formatAddress(row, col))
    }
  })

  for (const range of engine.workbook.listStyleRanges(sheetName)) {
    const rangeBounds = normalizeRangeBounds(sheetName, range.range.startAddress, range.range.endAddress)
    const clippedBounds = bounds ? intersectBounds(rangeBounds, bounds) : rangeBounds
    if (clippedBounds) {
      addAddressesForBounds(addresses, clippedBounds)
    }
  }

  for (const range of engine.workbook.listFormatRanges(sheetName)) {
    const rangeBounds = normalizeRangeBounds(sheetName, range.range.startAddress, range.range.endAddress)
    const clippedBounds = bounds ? intersectBounds(rangeBounds, bounds) : rangeBounds
    if (clippedBounds) {
      addAddressesForBounds(addresses, clippedBounds)
    }
  }

  return addresses
}

function materializeCellEvalRow(
  engine: SpreadsheetEngine,
  documentId: string,
  revision: number,
  updatedAt: string,
  sheetName: string,
  address: string,
  includeEmpty = false,
): CellEvalRow | null {
  const { row, col } = parseCellAddress(address, sheetName)
  const cell = engine.getCell(sheetName, address)
  if (
    !includeEmpty &&
    cell.value.tag === ValueTag.Empty &&
    cell.flags === 0 &&
    cell.styleId === undefined &&
    cell.numberFormatId === undefined &&
    cell.format === undefined
  ) {
    return null
  }
  return {
    workbookId: documentId,
    sheetName,
    address,
    rowNum: row,
    colNum: col,
    value: cell.value,
    flags: cell.flags,
    version: cell.version,
    styleId: cell.styleId ?? null,
    styleJson: engine.getCellStyle(cell.styleId) ?? null,
    formatId: cell.numberFormatId ?? null,
    formatCode: cell.format ?? null,
    calcRevision: revision,
    updatedAt,
  }
}

export function materializeCellEvalProjection(
  engine: SpreadsheetEngine,
  documentId: string,
  revision: number,
  updatedAt: string,
  changedCellIndices?: readonly number[],
): CellEvalRow[] {
  const entries: CellEvalRow[] = []

  if (changedCellIndices) {
    for (let i = 0; i < changedCellIndices.length; i += 1) {
      const cellIndex = changedCellIndices[i]!
      const qualifiedAddress = engine.workbook.getQualifiedAddress(cellIndex)
      const separatorIndex = qualifiedAddress.lastIndexOf('!')
      if (separatorIndex <= 0 || separatorIndex >= qualifiedAddress.length - 1) {
        continue
      }
      const sheetName = qualifiedAddress.slice(0, separatorIndex)
      const address = qualifiedAddress.slice(separatorIndex + 1)
      const row = materializeCellEvalRow(engine, documentId, revision, updatedAt, sheetName, address, true)
      if (row) {
        entries.push(row)
      }
    }
    return entries
  }

  for (const sheetName of engine.workbook.sheetsByName.keys()) {
    for (const address of collectCellEvalAddresses(engine, sheetName)) {
      const row = materializeCellEvalRow(engine, documentId, revision, updatedAt, sheetName, address)
      if (row) {
        entries.push(row)
      }
    }
  }

  return entries
}

export function materializeCellEvalRangeProjection(
  engine: SpreadsheetEngine,
  documentId: string,
  revision: number,
  updatedAt: string,
  range: CellRangeRef,
): CellEvalRow[] {
  const bounds = normalizeRangeBounds(range.sheetName, range.startAddress, range.endAddress)
  const entries: CellEvalRow[] = []
  for (const address of collectCellEvalAddresses(engine, range.sheetName, bounds)) {
    const row = materializeCellEvalRow(engine, documentId, revision, updatedAt, range.sheetName, address)
    if (row) {
      entries.push(row)
    }
  }
  return entries
}
